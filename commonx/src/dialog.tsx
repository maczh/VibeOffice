import { useEffect, useId, useRef, type CSSProperties, type ReactNode } from 'react';

export interface OfficeDialogProps {
  title: string;
  children: ReactNode;
  className?: string;
  style?: CSSProperties;
  onRequestClose?: () => void;
}

/** Native, focus-trapping dialog primitive shared by the suite and editor packages. */
export function OfficeDialog({ title, children, className, style, onRequestClose }: OfficeDialogProps) {
  const dialogRef = useRef<HTMLDialogElement>(null);
  const closeRef = useRef(onRequestClose);
  const titleId = useId();
  closeRef.current = onRequestClose;

  useEffect(() => {
    const dialog = dialogRef.current;
    if (!dialog) return;
    dialog.showModal();
    const onCancel = () => closeRef.current?.();
    dialog.addEventListener('cancel', onCancel);
    return () => {
      dialog.removeEventListener('cancel', onCancel);
      if (dialog.open) dialog.close();
    };
  }, []);

  return <dialog ref={dialogRef} style={style} className={['vibeoffice-dialog', className].filter(Boolean).join(' ')} aria-modal="true" aria-labelledby={titleId}>
    <h2 id={titleId}>{title}</h2>
    {children}
  </dialog>;
}
