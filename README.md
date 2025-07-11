# Companies House Checker

This script checks UK Companies House data for a list of company numbers and exports selected company and charge details to an Excel file.

## Prerequisites

You need the following installed on your machine:

- **Node.js** (v16 or newer recommended): [Download Node.js](https://nodejs.org/)
- **npm** (comes with Node.js)
- **Git** (for cloning the repository): [Download Git](https://git-scm.com/)

## Installation Steps

1. **Clone the repository**

   Open a terminal (Command Prompt, PowerShell, or Terminal) and run:

   ```sh
   git clone [https://github.com/NikoDP20/company-charge-checker.git]
   cd companies-house-checker
   ```

2. **Install dependencies**

   In the project folder, run:

   ```sh
   npm install
   ```

   This will install all required packages, including:
   - `axios`
   - `csv-parse`
   - `xlsx`
   - `node-file-dialog`

3. **Set your Companies House API Key**

   The script expects your API key to be set in the `API_KEY` variable at the top of `checkCharges.js`.  
   You can get a free API key from [Companies House Developer Hub](https://developer.company-information.service.gov.uk/developer/applications).

## Usage

1. **Prepare your input file**

   - The input file should be a `.csv`, `.xlsx`, or `.xls` file.
   - Company numbers should be in column B, starting from row 3.

2. **Run the script**

   In the project folder, run:

   ```sh
   node checkCharges.js
   ```

   - A file dialog will open for you to select your input file.
   - The script will process up to 500 company numbers and create an Excel file called `matched_charges.xlsx` in your Downloads folder.

## Troubleshooting

- If you get errors about missing modules, make sure you ran `npm install` in the project folder.
- If you get an error about your API key, check that it is correct and active.

---
