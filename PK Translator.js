/* Franklin Durbin
*  PG- 10288 4/10/2023
*  A script to assist with the language data copying for the PK.
*/

function stripID(id) {
    id = id.replace('{', '');
    id = id.replace('}', '');
    return id;
}

// Import Module
const xlsx = require('xlsx');
const fs = require('fs');
// File Strings
var PKimport = ''; // Formatted File to import into the tool
var marketing = ''; // Formatted human readable document
// Language
var language = ''; // Title of the language column
// Gather sheet data
// Files to import, must be encoded to utf16le
var PKImportSheet = xlsx.readFile(PKimport, {raw: true, encoding: 'utf16le'});
var marketingDoc = xlsx.readFile(marketing, {raw: true, encoding: 'utf16le'});

// Get the sheets into readable JSON data
var PKImportSheetNames = PKImportSheet.SheetNames;
var marketingDocSheetNames = marketingDoc.SheetNames;
var PKImportSheetData = PKImportSheet.Sheets[PKImportSheetNames[0]];
var marketingDocSheetData = marketingDoc.Sheets[marketingDocSheetNames[0]];
var PKImportData = xlsx.utils.sheet_to_json(PKImportSheetData, { encoding: 'utf16le' });
var marketingDocData = xlsx.utils.sheet_to_json(marketingDocSheetData, { encoding: 'utf16le' });

console.log("Original file row count: " + PKImportData.length)

// Get the correct translations
// CS-499 Algorithm and Data Structure enhancment
// adding an index and counter to remove the proper row from the array.
// This will create smaller and smaller arrays to search from as we iterate through.
PKImportData.forEach(PKrow => {
    if (PKrow.FIELD == "DisplayName") {
        var i = 0;
        marketingDocData.forEach(marketingRow => {
            if (marketingRow.ID == stripID(PKrow.ID)) {
                PKrow[language] = marketingRow[language];
                delete marketingDocData[i];
            }
            i++
        })
    }

    if (PKrow["PK VALUE"]) {
        if (PKrow["PK VALUE"].includes("\"")) {
            PKrow["PK VALUE"] = PKrow["PK VALUE"].replaceAll("\"", "\"\"");
        }
    }
})

// Format the data
var data = [];
// Send the header
data.push("\"TABLE\",\"ID\",\"FIELD\",\"PK VALUE\",\"" + language + "\"");
// Generate the empty cells
PKImportData.forEach(obj => {
    switch (Object.values(obj).length) {
        case 3:
            obj["PK VALUE"] = "";
            obj[language] = "";
            break;
        case 4:
            obj[language] = "";
            break;
        default:
            break;
    }
    // Format the connectors
     var string = "\"" + Object.values(obj).join("\",\"") + "\""
     data.push(string);
})
// Create the file write stream
const file = fs.createWriteStream(PKimport.replace("Resources", "Results"), { encoding: 'utf16le' })
// Catch errors
file.on('error', (err) => {
    console.log(err);
})
// Start writing line by line
data.forEach(line => {
    file.write(line + "\r\n");
})
console.log("Original file row count: " + PKImportData.length)
console.log("Updated file row count: " + data.length)
// Close file writting stream
file.end();