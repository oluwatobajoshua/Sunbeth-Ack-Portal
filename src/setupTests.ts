// CRA automatically loads this file before running tests.
// Extend Jest's expect with @testing-library/jest-dom matchers.
import '@testing-library/jest-dom';

// Mock SweetAlert2 for Jest/jsdom to avoid CSS parsing issues and DOM injections
jest.mock('sweetalert2', () => {
	const fire = jest.fn().mockResolvedValue({ isConfirmed: true, isDenied: false, isDismissed: false });
	const mixin = jest.fn(() => ({ fire }));
	return {
		__esModule: true,
		default: { mixin, fire },
		mixin,
	};
});

// Mock axios ESM for Jest environment (use manual mock in __mocks__/axios.js)
jest.mock('axios');
