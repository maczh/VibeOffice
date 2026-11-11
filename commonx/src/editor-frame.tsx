import { useEffect, useRef, type HTMLAttributes, type ReactNode } from 'react';
import './editor.css';

export type EditorKind = 'wordx' | 'excelx' | 'pptx';

export interface EditorFrameProps extends Omit<HTMLAttributes<HTMLElement>, 'children'> {
  appName: string;
  editor: EditorKind;
  children?: ReactNode;
}

/** A mounted document engine owns its DOM subtree and returns a disposer for listeners/resources. */
export interface EditorMountContext {
  host: HTMLElement;
  editor: EditorKind;
}

export type EditorMount = (context: EditorMountContext) => void | (() => void) | Promise<void | (() => void)>;

export interface LegacyEngineFrameOptions {
  url: string | URL;
  title: string;
  launch?: 'new' | 'sample';
}

/** Mount an existing same-origin app page in React-host mode; the page will not consult legacy Recent Files storage. */
export function createLegacyEngineFrame({ url, title, launch = 'new' }: LegacyEngineFrameOptions): EditorMount {
  const source = new URL(url, typeof window === 'undefined' ? 'http://localhost/' : window.location.href);
  if (!['http:', 'https:'].includes(source.protocol)) throw new TypeError('The editor runtime URL must use HTTP or HTTPS');
  source.searchParams.delete('open');
  source.searchParams.delete('recent');
  source.searchParams.delete('draft');
  source.searchParams.set('reactHost', '');
  source.searchParams.set(launch, '');

  return ({ host }) => {
    const frame = host.ownerDocument.createElement('iframe');
    frame.title = title;
    frame.src = source.href;
    frame.setAttribute('loading', 'eager');
    frame.style.cssText = 'display:block;width:100%;height:100%;min-height:600px;border:0';
    host.replaceChildren(frame);
    return () => frame.remove();
  };
}

export interface EngineEditorFrameProps extends EditorFrameProps {
  mount?: EditorMount;
}

/** Bridge for legacy engines that manage their own DOM, without letting them replace React's outer host. */
export function EngineEditorFrame({ mount, children, ...props }: EngineEditorFrameProps) {
  const hostRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const host = hostRef.current;
    if (!host || !mount) return;
    let disposed = false;
    let dispose: void | (() => void);
    Promise.resolve(mount({ host, editor: props.editor })).then((cleanup) => {
      if (disposed) {
        if (typeof cleanup === 'function') cleanup();
      } else dispose = cleanup;
    });
    return () => {
      disposed = true;
      if (dispose) dispose();
      host.replaceChildren();
    };
  }, [mount, props.editor]);

  return <EditorFrame {...props}><div className="vibe-editor__engine" ref={hostRef}>{children}</div></EditorFrame>;
}

export function EditorFrame({ appName, editor, className, children, ...attributes }: EditorFrameProps) {
  const classes = ['vibe-editor', className].filter(Boolean).join(' ');

  return (
    <section
      {...attributes}
      aria-label={`${appName} editor`}
      className={classes}
      data-vibe-editor={editor}
    >
      {children}
    </section>
  );
}
