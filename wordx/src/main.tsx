import { createRoot } from 'react-dom/client';
import { initializeSharedRuntime } from '@vibeoffice/core';
import { WordxEditor } from './index.js';

initializeSharedRuntime();
const rootElement = document.getElementById('root');
if (!rootElement) throw new Error('Missing #root mount element');
createRoot(rootElement).render(<WordxEditor engine={{ url: '/wordx/' }} />);
