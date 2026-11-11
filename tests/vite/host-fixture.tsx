import { createRoot } from 'react-dom/client';
import { WordxEditor } from '@vibeoffice/wordx';
import { ExcelxEditor } from '@vibeoffice/excelx';
import { PptxEditor } from '@vibeoffice/pptx';
import { useVibeOffice } from '@vibeoffice/core';

function HostFixture() {
  const office = useVibeOffice();
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
      </header>
      <section className="fixture-editors">
        <WordxEditor engine={{ url: '/quire/' }} />
        <ExcelxEditor engine={{ url: '/ledger/' }} />
        <PptxEditor engine={{ url: '/lectern/' }} />
      </section>
    </main>
  );
}

const root = document.getElementById('root');
if (!root) throw new Error('Missing #root mount element');
createRoot(root).render(<HostFixture />);
