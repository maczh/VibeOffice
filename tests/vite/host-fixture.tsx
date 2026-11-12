import { createRoot } from 'react-dom/client';
import { useState } from 'react';
import { WordxEditor } from '@vibeoffice/wordx';
import { ExcelxEditor } from '@vibeoffice/excelx';
import { PptxEditor } from '@vibeoffice/pptx';
import { useVibeOffice, VibeOfficeProvider } from '@vibeoffice/core';

type EditorId = 'wordx' | 'excelx' | 'pptx';

function HostFixture() {
  const office = useVibeOffice();
  const [editor, setEditor] = useState<EditorId>('wordx');
  const setLook = (look: typeof office.look) => {
    document.documentElement.dataset.look = look;
    document.documentElement.dataset.layout = look === 'paper2016' ? 'ribbon' : 'bars';
    window.dispatchEvent(new Event('vo-look'));
  };

  return (
    <main className="fixture">
      <header className="fixture-head">
        <h1>{office.translate('VibeOffice editors')}</h1>
        <label>
          {office.translate('Look')}{' '}
          <select aria-label={office.translate('Look')} value={office.look} onChange={(event) => setLook(event.target.value as typeof office.look)}>
            {office.looks.map(({ id, name }) => <option key={id} value={id}>{office.translate(name)}</option>)}
          </select>
        </label>
        <label>
          <select aria-label={office.translate('VibeOffice editors')} value={editor} onChange={(event) => setEditor(event.target.value as EditorId)}>
            <option value="wordx">Wordx</option>
            <option value="excelx">Excelx</option>
            <option value="pptx">PPTX</option>
          </select>
        </label>
      </header>
      <section className="fixture-editors">
        {editor === 'wordx' && <WordxEditor key="wordx" engine={{ url: '/wordx/' }} />}
        {editor === 'excelx' && <ExcelxEditor key="excelx" engine={{ url: '/excelx/' }} />}
        {editor === 'pptx' && <PptxEditor key="pptx" engine={{ url: '/pptx/' }} />}
      </section>
    </main>
  );
}

const root = document.getElementById('root');
if (!root) throw new Error('Missing #root mount element');
createRoot(root).render(<VibeOfficeProvider><HostFixture /></VibeOfficeProvider>);
