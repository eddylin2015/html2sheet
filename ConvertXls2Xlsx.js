const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');
function htmlToExcel(inputFile, outputFile) {
    try {
        // Read HTML file
        const htmlContent = fs.readFileSync(inputFile, 'utf8');
        
        // Parse HTML content (SheetJS automatically detects tables)
        const workbook = XLSX.read(htmlContent, { type: 'string' });
        
        // Write to Excel file
        XLSX.writeFile(workbook, outputFile);
        
        console.log(`Successfully converted ${inputFile} to ${outputFile}`);
        console.log(`Sheets created: ${workbook.SheetNames.join(', ')}`);
        
        return workbook;
    } catch (error) {
        console.error('Error:', error.message);
        throw error;
    }
}
async function searchFiles(directory, searchTerm) {
    try {
        const files = await fs.promises.readdir(directory);
        const results = files.filter(file => 
            file.toLowerCase().includes(searchTerm.toLowerCase())
        );
        return results;
    } catch (error) {
        console.error('Error reading directory:', error);
        return [];
    }
}

let dir_='./2526REG'
searchFiles(dir_, '2025')
    .then(files => {console.log('Found files:', files)
      for(let f of files){
         htmlToExcel(dir_+"/"+f, f.slice(-7)+'x');
      }
    });






