// Manual mock for axios to avoid importing ESM axios during Jest runs
const mockGet = jest.fn().mockResolvedValue({ data: { value: [] } });
const mockPost = jest.fn().mockResolvedValue({ data: {} });

module.exports = {
  get: mockGet,
  post: mockPost,
  // support default import interop
  default: {
    get: mockGet,
    post: mockPost,
  }
};
