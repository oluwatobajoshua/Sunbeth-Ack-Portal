export {};

describe('dataverseService (mock)', () => {
  beforeAll(() => {
    process.env.REACT_APP_USE_MOCK = 'true';
  });

  test('getBatches returns array of batches', async () => {
    jest.isolateModules(() => {
      jest.mock('axios');
  const { getBatches } = require('../services/dbService');
      return getBatches().then((batches: any) => {
        expect(Array.isArray(batches)).toBe(true);
        expect(batches.length).toBeGreaterThan(0);
        expect(batches[0]).toHaveProperty('toba_batchid');
      });
    });
  });

  test('getDocumentsByBatch returns docs array', async () => {
    jest.isolateModules(() => {
      jest.mock('axios');
  const { getDocumentsByBatch } = require('../services/dbService');
      return getDocumentsByBatch('1').then((docs: any) => {
        expect(Array.isArray(docs)).toBe(true);
        expect(docs[0]).toHaveProperty('toba_documentid');
      });
    });
  });
});
 
