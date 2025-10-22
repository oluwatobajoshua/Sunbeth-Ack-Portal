// Debug script to test batch creation relationships
const BASE_URL = 'http://localhost:4000';

async function testBatchCreation() {
    console.log('=== Testing Batch Creation Relationships ===\n');
    
    try {
        // Step 1: Create a batch
        console.log('Step 1: Creating batch...');
        const batchResponse = await fetch(`${BASE_URL}/api/batches`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                name: 'Debug Test Batch',
                description: 'Testing batch relationships',
                status: 1
            })
        });
        
        if (!batchResponse.ok) {
            throw new Error(`Batch creation failed: ${batchResponse.status}`);
        }
        
        const batchResult = await batchResponse.json();
        const batchId = batchResult.id || batchResult.batchId;
        console.log('✅ Batch created:', { id: batchId, ...batchResult });
        
        // Step 2: Verify batch exists
        console.log('\nStep 2: Verifying batch exists...');
        const verifyResponse = await fetch(`${BASE_URL}/api/batches`);
        const allBatches = await verifyResponse.json();
        const createdBatch = allBatches.find(b => b.id == batchId);
        console.log('✅ Batch verification:', createdBatch ? 'FOUND' : 'NOT FOUND', createdBatch);
        
        // Step 3: Add documents
        console.log('\nStep 3: Adding documents to batch...');
        const documentsPayload = {
            documents: [
                {
                    title: 'Test Document 1',
                    url: 'https://example.com/doc1.pdf',
                    version: 1,
                    requiresSignature: false
                },
                {
                    title: 'Test Document 2', 
                    url: 'https://example.com/doc2.pdf',
                    version: 1,
                    requiresSignature: true
                }
            ]
        };
        
        const docsResponse = await fetch(`${BASE_URL}/api/batches/${batchId}/documents`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(documentsPayload)
        });
        
        if (!docsResponse.ok) {
            const errorText = await docsResponse.text();
            console.error('❌ Documents creation failed:', docsResponse.status, errorText);
        } else {
            const docsResult = await docsResponse.json();
            console.log('✅ Documents added:', docsResult);
        }
        
        // Step 4: Add recipients 
        console.log('\nStep 4: Adding recipients to batch...');
        const recipientsPayload = {
            recipients: [
                {
                    businessId: 1,
                    user: 'user1@example.com',
                    email: 'user1@example.com',
                    displayName: 'Test User 1',
                    department: 'IT'
                },
                {
                    businessId: 1,
                    user: 'user2@example.com', 
                    email: 'user2@example.com',
                    displayName: 'Test User 2',
                    department: 'HR'
                }
            ]
        };
        
        const recipientsResponse = await fetch(`${BASE_URL}/api/batches/${batchId}/recipients`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(recipientsPayload)
        });
        
        if (!recipientsResponse.ok) {
            const errorText = await recipientsResponse.text();
            console.error('❌ Recipients creation failed:', recipientsResponse.status, errorText);
        } else {
            const recipientsResult = await recipientsResponse.json();
            console.log('✅ Recipients added:', recipientsResult);
        }
        
        // Step 5: Verify relationships
        console.log('\nStep 5: Verifying relationships...');
        
        // Check documents
        const batchDocsResponse = await fetch(`${BASE_URL}/api/batches/${batchId}/documents`);
        if (batchDocsResponse.ok) {
            const batchDocs = await batchDocsResponse.json();
            console.log(`✅ Documents linked to batch ${batchId}:`, batchDocs.length, 'documents');
            batchDocs.forEach(doc => console.log(`  - ${doc.title} (batchId: ${doc.batchId})`));
        }
        
        // Check recipients
        const batchRecipientsResponse = await fetch(`${BASE_URL}/api/batches/${batchId}/recipients`);
        if (batchRecipientsResponse.ok) {
            const batchRecipients = await batchRecipientsResponse.json();
            console.log(`✅ Recipients linked to batch ${batchId}:`, batchRecipients.length, 'recipients');
            batchRecipients.forEach(rec => console.log(`  - ${rec.email} (batchId: ${rec.batchId})`));
        }
        
        // Step 6: Check all documents and recipients tables
        console.log('\nStep 6: Checking all documents and recipients...');
        const allDocsResponse = await fetch(`${BASE_URL}/api/documents`);
        if (allDocsResponse.ok) {
            const allDocs = await allDocsResponse.json();
            console.log('All documents in database:', allDocs.length);
            allDocs.forEach(doc => console.log(`  - ${doc.title} (batchId: ${doc.batchId || 'NULL'})`));
        }
        
        const allRecipientsResponse = await fetch(`${BASE_URL}/api/recipients`);
        if (allRecipientsResponse.ok) {
            const allRecipients = await allRecipientsResponse.json();
            console.log('All recipients in database:', allRecipients.length);
            allRecipients.forEach(rec => console.log(`  - ${rec.email} (batchId: ${rec.batchId || 'NULL'})`));
        }
        
    } catch (error) {
        console.error('❌ Test failed:', error.message);
    }
}

// Run the test
testBatchCreation();