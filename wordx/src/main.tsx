import { createRoot } from 'react-dom/client';
import { WordxEditor } from './index.js';

const rootElement = document.getElementById('root');
if (!rootElement) throw new Error('Missing #root mount element');
createRoot(rootElement).render(<WordxEditor engine={{ url: '/quire/' }} />);
