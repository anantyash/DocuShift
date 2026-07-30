const fs = require('fs-extra');
const path = require('path');
const { execSync } = require('child_process');

const testFolder = path.join(__dirname, '../temp/uploads');
const testHindiFile = path.join(testFolder, 'हिंदी_परीक्षण_दस्तावेज़.docx');

// Copy Sample_Test_Doc.docx to Hindi name
const sampleFile = path.join(testFolder, 'Sample_Test_Doc.docx');
if (fs.existsSync(sampleFile)) {
  fs.copySync(sampleFile, testHindiFile);
  console.log('Created Hindi test document:', testHindiFile);
}

// Test API via curl with UTF-8 support
try {
  const curlCmd = `curl.exe -X POST -F "files=@${testHindiFile}" http://localhost:3000/api/convert`;
  const output = execSync(curlCmd, { encoding: 'utf8' });
  console.log('--- API Conversion Output ---');
  console.log(output);
} catch (err) {
  console.error('Curl test error:', err.message);
}
