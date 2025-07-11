const axios = require("axios");
const fs = require("fs");
const path = require("path");
const { parse } = require("csv-parse/sync");
const xlsx = require("xlsx");
const fileDialog = require("node-file-dialog");

const API_KEY = "c501755c-827b-4a64-a0c0-7df0dd8a2ec6";
const LENDER_NAMES = [
  "Nationwide Finance Limited",
  "Sellersfunding International Portfolio LTD",
  "Capitalrise Finance Limited",
  "Peak Cashflow Limited",
  "Sevcap I Limited",
  "Seneca Trade Partners LTD",
  "Finbiz Funding Limited",
  "Swishfund LTD",
  "Optimum Sme Finance Limited",
  "Liquid Link Limited",
  "Reward Capital Limited",
];

const BASE_URL = "https://api.company-information.service.gov.uk";

// Prompt user for file path
async function promptFilePath() {
  const files = await fileDialog({
    type: "open-file",
    accept: [".csv", ".xlsx", ".xls"],
    multiple: false,
  });
  return files[0];
}

// Extract company numbers from CSV (column B, from row 3)
function extractCompanyNumbersFromCSV(filePath) {
  const content = fs.readFileSync(filePath, "utf8");
  const records = parse(content, { skip_empty_lines: true });
  const companyNumbers = [];
  // Start at row 3 (index 2), column B (index 1)
  for (let i = 2; i < records.length; i++) {
    const value = records[i][1];
    if (!value || String(value).trim() === "") break;
    companyNumbers.push(String(value).trim());
  }
  return companyNumbers;
}

// Extract company numbers from Excel (column B, from row 3)
function extractCompanyNumbersFromExcel(filePath) {
  const workbook = xlsx.readFile(filePath);
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const companyNumbers = [];
  let row = 3; // Start at row 3 (B3)
  while (true) {
    const cellAddress = `B${row}`;
    const cell = sheet[cellAddress];
    if (!cell || !cell.v || String(cell.v).trim() === "") break;
    companyNumbers.push(String(cell.v).trim());
    row++;
  }
  return companyNumbers;
}

// Get company numbers from file
async function getCompanyNumbersFromFile(filePath) {
  const ext = path.extname(filePath).toLowerCase();
  if (ext === ".csv") {
    return extractCompanyNumbersFromCSV(filePath);
  } else if (ext === ".xlsx" || ext === ".xls") {
    return extractCompanyNumbersFromExcel(filePath);
  } else {
    throw new Error("Unsupported file type");
  }
}

// Get charges with any lender in the list
async function getChargesWithLender(companyNumber, lenderNames) {
  try {
    const response = await axios.get(
      `${BASE_URL}/company/${companyNumber}/charges`,
      { auth: { username: API_KEY, password: "" } }
    );
    const charges = response.data.items || [];
    // Return all charges where any lender matches
    return charges.filter((charge) =>
      (charge.persons_entitled || []).some(
        (party) =>
          party.name &&
          lenderNames.some((lender) =>
            party.name.toLowerCase().includes(lender.toLowerCase())
          )
      )
    );
  } catch (error) {
    console.error(
      `Error fetching data for company ${companyNumber}:`,
      error.message
    );
    return [];
  }
}

// Get all charges for a company (no lender filter)
async function getAllCharges(companyNumber) {
  try {
    const response = await axios.get(
      `${BASE_URL}/company/${companyNumber}/charges`,
      { auth: { username: API_KEY, password: "" } }
    );
    return response.data.items || [];
  } catch (error) {
    console.error(
      `Error fetching data for company ${companyNumber}:`,
      error.message
    );
    return [];
  }
}

// Get company profile
async function getCompanyProfile(companyNumber) {
  try {
    const response = await axios.get(`${BASE_URL}/company/${companyNumber}`, {
      auth: { username: API_KEY, password: "" },
    });
    return response.data;
  } catch {
    return {};
  }
}

// Get company officers (directors)
async function getCompanyOfficers(companyNumber) {
  try {
    const response = await axios.get(
      `${BASE_URL}/company/${companyNumber}/officers`,
      { auth: { username: API_KEY, password: "" } }
    );
    return response.data.items || [];
  } catch {
    return [];
  }
}

// Main function
async function main() {
  // Prompt for file path ONCE
  const filePath = await promptFilePath();

  let companyNumbers = [];
  try {
    companyNumbers = await getCompanyNumbersFromFile(filePath);
    if (!companyNumbers.length) {
      console.error("No company numbers found in file.");
      process.exit(1);
    }
  } catch (err) {
    console.error("Error:", err.message);
    process.exit(1);
  }

  // Limit to 500 company numbers
  companyNumbers = companyNumbers.slice(0, 500);

  const outputRows = [];

  for (const number of companyNumbers) {
    console.log(`Checking company: ${number}`);
    //const charges = await getChargesWithLender(number, LENDER_NAMES);                 // Fetch charges with any lender in the list
    const charges = await getAllCharges(number);                                        // Fetch all charges without lender filter
    if (charges.length > 0) {
      const profile = await getCompanyProfile(number);
      const officers = await getCompanyOfficers(number);
      const director =
        officers.find((o) => o.officer_role === "director") || {};

      charges.forEach((charge) => {
        outputRows.push({
          "Company Name": profile.company_name || "",
          "Company Number": number,
          "Company Type": profile.type || "",
          "Incorporation Date": profile.date_of_creation || "",
          "Registered Office Address": profile.registered_office_address
            ? Object.values(profile.registered_office_address).join(", ")
            : "",
          "Director Name": director.name || "",
          "Dormant Latest Accounts?":
            profile.accounts?.last_accounts?.type === "dormant" ? "Yes" : "No",
          "Accounts Overdue?": profile.accounts?.overdue ? "Yes" : "No",
          "Confirmation Statement Overdue?": profile.confirmation_statement
            ?.overdue
            ? "Yes"
            : "No",
          "Charge Holders": (charge.persons_entitled || [])
            .map((p) => p.name)
            .join("; "),
        });
      });
    }
  }

  // Write output to Excel file
  if (outputRows.length > 0) {
    const wb = xlsx.utils.book_new();
    const ws = xlsx.utils.json_to_sheet(outputRows);
    xlsx.utils.book_append_sheet(wb, ws, "Matched Charges");
    const downloadsDir = path.join(require("os").homedir(), "Downloads");
    const outPath = path.join(downloadsDir, "matched_charges.xlsx");
    xlsx.writeFile(wb, outPath);
    console.log(`\n✅ Excel file created: ${outPath}`);
  } else {
    console.log("\nNo matching charges found, so no Excel file was created.");
  }
}

main();
