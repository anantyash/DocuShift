const http = require('http');
const fs = require('fs-extra');
const path = require('path');

const sampleFile = path.join(__dirname, '../temp/uploads/Sample_Test_Doc.docx');
let fileBuffer;

if (fs.existsSync(sampleFile)) {
  fileBuffer = fs.readFileSync(sampleFile);
} else {
  // Find any docx in uploads
  const uploadsDir = path.join(__dirname, '../temp/uploads');
  const files = fs.readdirSync(uploadsDir).filter(f => f.endsWith('.docx'));
  if (files.length > 0) {
    fileBuffer = fs.readFileSync(path.join(uploadsDir, files[0]));
  } else {
    fileBuffer = Buffer.from('PK\x03\x04DummyDocxData');
  }
}

const filename = 'हिंदी_परीक्षण_दस्तावेज़.docx';

const boundary = '----WebKitFormBoundary7MA4YWxkTrZu0gW';
let body = '';

body += `--${boundary}\r\n`;
body += `Content-Disposition: form-data; name="files"; filename="${filename}"\r\n`;
body += `Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document\r\n\r\n`;

const headerBuffer = Buffer.from(body, 'utf8');
const footerBuffer = Buffer.from(`\r\n--${boundary}--\r\n`, 'utf8');

const payload = Buffer.concat([headerBuffer, fileBuffer, footerBuffer]);

const req = http.request({
  hostname: 'localhost',
  port: 3000,
  path: '/api/convert',
  method: 'POST',
  headers: {
    'Content-Type': `multipart/form-data; boundary=${boundary}`,
    'Content-Length': payload.length
  }
}, (res) => {
  let responseData = '';
  res.on('data', (chunk) => responseData += chunk);
  res.on('end', () => {
    console.log('=== HTTP MULTIPART UNICODE RESPONSE ===');
    console.log(responseData);
  });
});

req.on('error', (err) => console.error('HTTP Request Error:', err));
req.write(payload);
req.end();
