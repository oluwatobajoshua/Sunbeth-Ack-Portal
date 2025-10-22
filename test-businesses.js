// Simple test script to verify business functionality
const http = require('http');

const makeRequest = (method, path, data = null) => {
  return new Promise((resolve, reject) => {
    const options = {
      hostname: 'localhost',
      port: 4000,
      path,
      method,
      headers: {
        'Content-Type': 'application/json'
      }
    };

    const req = http.request(options, (res) => {
      let body = '';
      res.on('data', chunk => body += chunk);
      res.on('end', () => {
        try {
          const result = body ? JSON.parse(body) : {};
          resolve({ status: res.statusCode, data: result });
        } catch (e) {
          resolve({ status: res.statusCode, data: body });
        }
      });
    });

    req.on('error', reject);

    if (data) {
      req.write(JSON.stringify(data));
    }
    req.end();
  });
};

async function testBusinesses() {
  try {
    console.log('Testing businesses API...');
    
    // Test health endpoint
    console.log('\n1. Testing health endpoint...');
    const health = await makeRequest('GET', '/api/health');
    console.log('Health:', health);

    // Test get businesses
    console.log('\n2. Getting businesses...');
    const businesses = await makeRequest('GET', '/api/businesses');
    console.log('Businesses:', businesses);

    // Test create business
    console.log('\n3. Creating test business...');
    const newBusiness = await makeRequest('POST', '/api/businesses', {
      name: 'Test Business',
      code: 'TEST',
      isActive: true,
      description: 'Test business for validation'
    });
    console.log('Created business:', newBusiness);

    // Test get businesses again
    console.log('\n4. Getting businesses after creation...');
    const businessesAfter = await makeRequest('GET', '/api/businesses');
    console.log('Businesses after creation:', businessesAfter);

    console.log('\n✅ Business API tests completed successfully!');
  } catch (error) {
    console.error('❌ Error testing businesses:', error);
  }
}

testBusinesses();