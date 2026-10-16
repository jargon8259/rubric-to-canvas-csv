import { initApp } from './app.ts';

const rootEl = document.getElementById('app');
if (rootEl) initApp(document, rootEl, (globalThis as any).__RBA_CONFIG__ || {});
