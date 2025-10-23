const http = require('http');
const fs = require('fs');

const data = JSON.stringify({
  type: 'bar',
  data: {
    labels: ['Q1', 'Q2', 'Q3', 'Q4'],
    datasets: [{
      label: 'Revenue',
      data: [45000, 52000, 61000, 58000]
    }]
  },
  options: {
    title: 'Clean Test - No Errors'
  }
});

const options = {
  hostname: 'localhost',
  port: 3000,
  path: '/api/chart-to-pptx',
  method: 'POST',
  headers: {
    'Content-Type': 'application/json',
    'Content-Length': data.length
  }
};

const req = http.request(options, (res) => {
  console.log(`Status: ${res.statusCode}`);
  
  const chunks = [];
  res.on('data', (chunk) => chunks.push(chunk));
  res.on('end', () => {
    const buffer = Buffer.concat(chunks);
    fs.writeFileSync('test-clean.pptx', buffer);
    console.log('✅ Successfully created test-clean.pptx');
  });
});

req.on('error', (e) => {
  console.error(`Error: ${e.message}`);
});

req.write(data);
req.end();

