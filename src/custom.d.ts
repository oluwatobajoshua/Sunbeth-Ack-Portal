declare module '*.css';
declare module '*.scss';

// Allow importing images (png, jpg, svg) used by the app
declare module '*.png';
declare module '*.jpg';
declare module '*.jpeg';
declare module '*.svg';

// Ambient module for docx-preview (JS library without TS types)
declare module 'docx-preview';

// Extend RequestInit with optional busy tracking hints used by our fetch wrapper
declare global {
	interface RequestInit {
		/** Override the overlay label for this request */
		busyLabel?: string;
		/** Suppress automatic overlay for this request */
		busySilence?: boolean;
	}
}
