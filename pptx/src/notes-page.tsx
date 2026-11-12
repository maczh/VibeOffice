import { useEffect, useState } from 'react';
import { createPortal } from 'react-dom';

interface NotesPageSlide { id?: string; notes?: string }
interface NotesPagePresentation { slides: NotesPageSlide[]; W: number; H: number; firstNum?: number }
interface NotesPageRuntime {
  pres?: NotesPagePresentation;
  ed?: { idx: number };
  panes?: { notesView?: { el?: HTMLElement; reactMounted?: boolean; updateNotes(slide: NotesPageSlide, text: string): void } };
  render?: { thumb(presentation: NotesPagePresentation, slide: NotesPageSlide, width: number, options: { index: number }): HTMLElement };
  bus?: { on(event: string, listener: () => void): () => void };
  t?(message: string): string;
}

function runtime(): NotesPageRuntime | undefined { return (window as Window & { L?: NotesPageRuntime }).L; }

/** React-owned Notes Page canvas with engine-backed preview art and notes history. */
export function PptxNotesPageLayer({ host }: { host: HTMLElement }) {
  const [mount, setMount] = useState<HTMLElement | null>(null);
  const [revision, setRevision] = useState(0);
  useEffect(() => {
    let slot: HTMLElement | null = null;
    const attach = () => {
      const target = host.querySelector<HTMLElement>('.notespage');
      if (!target || target === slot) return;
      slot?.classList.remove('vibeoffice-react-owned');
      slot = target;
      slot.classList.add('vibeoffice-react-owned');
      const notesView = runtime()?.panes?.notesView;
      if (notesView) notesView.reactMounted = true;
      setMount(slot);
    };
    const off = runtime()?.bus?.on('pptx-notespage-render', () => setRevision((value) => value + 1));
    const observer = new MutationObserver(attach);
    observer.observe(host, { childList: true, subtree: true });
    attach();
    return () => {
      observer.disconnect();
      off?.();
      slot?.classList.remove('vibeoffice-react-owned');
      const notesView = runtime()?.panes?.notesView;
      if (notesView) notesView.reactMounted = false;
    };
  }, [host]);

  const L = runtime();
  const presentation = L?.pres;
  const index = L?.ed?.idx ?? 0;
  const slide = presentation?.slides[index];
  const thumb = L?.render?.thumb;
  void revision;
  if (!mount || !presentation || !slide || !thumb) return null;

  const availableWidth = Math.max(300, mount.clientWidth - 60);
  const pageWidth = Math.min(availableWidth, 640);
  const pageHeight = pageWidth * (10 / 7.5);
  const slideWidth = pageWidth * 0.67;
  const slideHeight = (slideWidth * presentation.H) / presentation.W;
  const notesTop = pageHeight * 0.075 + slideHeight + pageHeight * 0.05;
  const markup = thumb(presentation, slide, slideWidth, { index }).outerHTML;
  const page = <div className="np-page" style={{ width: pageWidth, height: pageHeight }}>
    <div className="np-slide" style={{ left: (pageWidth - slideWidth) / 2, top: pageHeight * 0.075, width: slideWidth, height: slideHeight }} dangerouslySetInnerHTML={{ __html: markup }} />
    <div className="np-notes" contentEditable suppressContentEditableWarning spellCheck role="textbox" aria-label={L?.t?.('Notes') || 'Notes'} style={{ left: pageWidth * 0.1, top: notesTop, width: pageWidth * 0.8, height: pageHeight - notesTop - pageHeight * 0.08, fontSize: Math.max(11, pageWidth / 48) }} onInput={(event) => L?.panes?.notesView?.updateNotes(slide, event.currentTarget.innerText.replace(/\n$/, ''))} onPaste={(event) => { event.preventDefault(); document.execCommand('insertText', false, event.clipboardData.getData('text/plain')); }}>{slide.notes || ''}</div>
    <div className="pv-hdr" style={{ right: pageWidth * 0.05, bottom: pageHeight * 0.03 }}>{String(index + (presentation.firstNum || 1))}</div>
  </div>;
  return createPortal(page, mount, `pptx-notes-page-${slide.id || index}`);
}
