import { createRoot } from 'react-dom/client';
import { PptxEditor } from './index.js';

const rootElement = document.getElementById('root');
if (!rootElement) throw new Error('Missing #root mount element');
createRoot(rootElement).render(<PptxEditor engine={{ url: '/lectern/' }} />);
