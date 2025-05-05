// Import required modules using CommonJS syntax
const fs = require('fs');
const path = require('path');
const readline = require('readline');
const XLSX = require('xlsx'); // You'll need to install this: npm install xlsx

// Create a simple console-based interface
const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout
});

// Main function to process Excel files
async function processExcelFiles() {
  console.log("\n=== EXCEL MATCHER TOOL ===\n");
  
  // Get Excel 1 file path
  const excel1Path = await askQuestion("Enter path to Excel1 file (e.g., C:\\path\\to\\passwordchange.xlsm): ");
  if (!fs.existsSync(excel1Path)) {
    console.error(`Error: File not found at ${excel1Path}`);
    rl.close();
    return;
  }
  
  // Get Excel 2 file path
  const excel2Path = await askQuestion("Enter path to Excel2 file (e.g., C:\\path\\to\\checklist.xlsm): ");
  if (!fs.existsSync(excel2Path)) {
    console.error(`Error: File not found at ${excel2Path}`);
    rl.close();
    return;
  }
  
  try {
    // Read Excel 1
    console.log("Reading Excel 1...");
    const workbook1 = XLSX.readFile(excel1Path);
    const sheet1Name = workbook1.SheetNames[0]; // Using first sheet by default
    console.log(`Using sheet: ${sheet1Name}`);
    const sheet1 = workbook1.Sheets[sheet1Name];
    const data1 = XLSX.utils.sheet_to_json(sheet1);
    
    if (data1.length === 0) {
      console.error("Error: Excel 1 has no data");
      rl.close();
      return;
    }
    
    // Display column headers from Excel 1
    console.log("\nAvailable columns in Excel 1:");
    const headers1 = Object.keys(data1[0]);
    headers1.forEach((header, index) => {
      console.log(`${index + 1}. ${header}`);
    });
    
    // Get input column name
    const inputColumnName = await askQuestion("\nEnter name of input column from Excel 1: ");
    if (!headers1.includes(inputColumnName)) {
      console.error(`Error: Column "${inputColumnName}" not found in Excel 1`);
      rl.close();
      return;
    }
    
    // Get search column name
    const searchColumnName = await askQuestion("Enter name of search column from Excel 1: ");
    if (!headers1.includes(searchColumnName)) {
      console.error(`Error: Column "${searchColumnName}" not found in Excel 1`);
      rl.close();
      return;
    }
    
    // Read Excel 2
    console.log("\nReading Excel 2...");
    const workbook2 = XLSX.readFile(excel2Path);
    
    // Display available sheets in Excel 2
    console.log("\nAvailable sheets in Excel 2:");
    workbook2.SheetNames.forEach((sheet, index) => {
      console.log(`${index + 1}. ${sheet}`);
    });
    
    // Get sheet name for Excel 2
    const excel2SheetName = await askQuestion("\nEnter name of sheet to search in Excel 2: ");
    if (!workbook2.SheetNames.includes(excel2SheetName)) {
      console.error(`Error: Sheet "${excel2SheetName}" not found in Excel 2`);
      rl.close();
      return;
    }
    
    const sheet2 = workbook2.Sheets[excel2SheetName];
    const data2 = XLSX.utils.sheet_to_json(sheet2);
    
    if (data2.length === 0) {
      console.error(`Error: Sheet "${excel2SheetName}" in Excel 2 has no data`);
      rl.close();
      return;
    }
    
    console.log("\nProcessing data...");
    
    // Find matches and extract data
    const results = [];
    
    for (const row of data1) {
      if (!(inputColumnName in row) || !(searchColumnName in row)) {
        console.log(`Warning: Missing columns in row: ${JSON.stringify(row)}`);
        continue;
      }
      
      const inputValue = row[inputColumnName];
      const searchValue = row[searchColumnName];
      
      // Find matching row in Excel 2
      const matchingRow = data2.find(item => item[searchColumnName] === searchValue);
      
      let outputValue = 'No match found';
      if (matchingRow) {
        outputValue = matchingRow[inputColumnName] || 'Column not found in matching row';
      }
      
      results.push({
        input: inputValue,
        searchValue: searchValue,
        output: outputValue
      });
    }
    
    // Display results
    console.log("\n=== RESULTS ===");
    console.log(`Found ${results.length} matches`);
    
    if (results.length > 0) {
      console.log("\nFirst 5 results:");
      results.slice(0, 5).forEach((result, index) => {
        console.log(`${index + 1}. Input: ${result.input}, Search: ${result.searchValue}, Output: ${result.output}`);
      });
      
      // Save results to file
      const outputPath = path.join(process.cwd(), 'excel_matcher_results.txt');
      const outputContent = results.map(r => r.output).join('\n');
      fs.writeFileSync(outputPath, outputContent);
      
      console.log(`\nAll results saved to: ${outputPath}`);
      console.log(`Total matches: ${results.length}`);
    }
    
  } catch (error) {
    console.error(`Error processing files: ${error.message}`);
  }
  
  rl.close();
}

// Helper function to ask questions
function askQuestion(question) {
  return new Promise(resolve => {
    rl.question(question, answer => {
      resolve(answer);
    });
  });
}

// Run the program
processExcelFiles();