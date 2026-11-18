import { createRoot } from 'react-dom/client';
import { initializeSharedRuntime } from '@vibeoffice/core';
import { ExcelxEditor } from './index.js';

initializeSharedRuntime();
const rootElement = document.getElementById('root');
if (!rootElement) throw new Error('Missing #root mount element');
createRoot(rootElement).render(<ExcelxEditor engine={{ url: '/excelx/' }} />);
