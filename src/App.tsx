import { useCallback, useEffect, useRef, useState } from 'react';
import { SuperDoc } from 'superdoc';
import { createSuperDocUI } from 'superdoc/ui';
import type { SelectionTarget, SuperDocUI, ViewportRect } from 'superdoc/ui';
import 'superdoc/style.css';

export default function App() {
  const mountRef = useRef<HTMLDivElement>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const superdocRef = useRef<SuperDoc | null>(null);
  const uiRef = useRef<SuperDocUI | null>(null);
  const preservedSelectionRef = useRef<SelectionTarget | null>(null);
  const exportingRef = useRef(false);
  const [document, setDocument] = useState<string | File>('/sample.docx');
  const [ready, setReady] = useState(false);
  const [exporting, setExporting] = useState(false);
  const [attachingMetadata, setAttachingMetadata] = useState(false);
  const [showMetadataOutlines, setShowMetadataOutlines] = useState(true);
  const [metadataOutlines, setMetadataOutlines] = useState<Array<ViewportRect & { key: string }>>([]);
  const [metadataStatus, setMetadataStatus] = useState('Select text within one paragraph to attach metadata.');

  useEffect(() => {
    if (!mountRef.current) return;

    let active = true;
    let opened = false;
    let loadFailed = false;
    let ui: SuperDocUI | null = null;
    setReady(false);
    const superdoc = new SuperDoc({
      selector: mountRef.current,
      document,
      onReady: () => {
        if (!active || loadFailed) return;
        opened = true;
        ui = createSuperDocUI({ superdoc });
        uiRef.current = ui;
        setReady(true);
      },
      onException: ({ error }) => {
        console.error('SuperDoc could not open the document.', error);
        if (!opened) {
          loadFailed = true;
          if (active) setReady(false);
        }
      },
    });
    superdocRef.current = superdoc;

    return () => {
      active = false;
      ui?.destroy();
      if (uiRef.current === ui) uiRef.current = null;
      if (superdocRef.current === superdoc) superdocRef.current = null;
      superdoc.destroy();
    };
  }, [document]);

  const refreshMetadataOutlines = useCallback(async () => {
    const superdoc = superdocRef.current;
    const ui = uiRef.current;
    const doc = superdoc?.activeEditor?.doc;
    if (!showMetadataOutlines || !ui || !doc) {
      setMetadataOutlines([]);
      return;
    }

    try {
      const entries = await doc.metadata.list({ resolvedOnly: true });
      const outlines = entries.items.flatMap((entry) => {
        const geometry = ui.metadata.getRect({ id: entry.id });
        if (!geometry.success) return [];
        return geometry.rects.map((rect, index) => ({ ...rect, key: `${entry.id}-${index}` }));
      });
      setMetadataOutlines(outlines);
      setMetadataStatus(`Found ${entries.items.length} metadata anchor${entries.items.length === 1 ? '' : 's'}.`);
    } catch (error) {
      setMetadataOutlines([]);
      setMetadataStatus(error instanceof Error ? error.message : String(error));
    }
  }, [showMetadataOutlines]);

  useEffect(() => {
    if (!ready || !showMetadataOutlines) return;
    const refresh = () => void refreshMetadataOutlines();
    refresh();
    window.addEventListener('resize', refresh);
    window.addEventListener('scroll', refresh, true);
    const interval = window.setInterval(refresh, 300);
    return () => {
      window.removeEventListener('resize', refresh);
      window.removeEventListener('scroll', refresh, true);
      window.clearInterval(interval);
    };
  }, [ready, refreshMetadataOutlines, showMetadataOutlines]);

  useEffect(() => {
    if (!ready) return;

    const handleEnter = (event: KeyboardEvent) => {
      if (event.key !== 'Enter' || event.shiftKey || event.altKey || event.ctrlKey || event.metaKey || event.isComposing) {
        return;
      }

      const superdoc = superdocRef.current;
      const editor = superdoc?.activeEditor;
      const doc = editor?.doc;
      const capture = uiRef.current?.selection.capture();
      const caret = capture?.selectionTarget ?? preservedSelectionRef.current;
      if (!editor || !doc) {
        console.warn('[metadata-enter] editor or document unavailable');
        return;
      }
      if (!caret || caret.start.kind !== 'text' || caret.end.kind !== 'text') {
        console.warn('[metadata-enter] no text caret', { capture, preserved: preservedSelectionRef.current });
        return;
      }
      if (caret.start.blockId !== caret.end.blockId || caret.start.offset !== caret.end.offset) {
        console.warn('[metadata-enter] selection is not a collapsed caret', caret);
        return;
      }
      const caretPoint = caret.start;

      const entries = doc.metadata.list({ resolvedOnly: true }).items;
      const entriesAtCaret = doc.metadata.list({ resolvedOnly: true, within: caret }).items;
      let match = entriesAtCaret.length === 1 ? entriesAtCaret[0] : entries.find((entry) => {
        const resolved = doc.metadata.resolve({ id: entry.id });
        const target = resolved?.target;
        if (!target || target.start.kind !== 'text' || target.end.kind !== 'text') return false;
        return target.start.blockId === caretPoint.blockId
          && target.end.blockId === caretPoint.blockId
          && caretPoint.offset > target.start.offset
          && caretPoint.offset < target.end.offset;
      });
      if (!match && entriesAtCaret.length > 0) {
        match = entriesAtCaret[0];
        console.info('[metadata-enter] using metadata resolved at the caret', {
          caret,
          resolved: doc.metadata.resolve({ id: match.id }),
        });
      }
      if (!match) {
        console.warn('[metadata-enter] caret is not strictly inside a metadata range', { caret, entries, entriesAtCaret });
        return;
      }

      const metadata = doc.metadata.get({ id: match.id });
      const resolved = doc.metadata.resolve({ id: match.id });
      const target = resolved?.target;
      if (!metadata || !target || target.start.kind !== 'text' || target.end.kind !== 'text') return;

      const logicalId =
        typeof metadata.payload === 'object'
        && metadata.payload !== null
        && 'logicalId' in metadata.payload
        && typeof metadata.payload.logicalId === 'string'
          ? metadata.payload.logicalId
          : crypto.randomUUID();
      const payload = typeof metadata.payload === 'object' && metadata.payload !== null
        ? { ...metadata.payload, logicalId }
        : { value: metadata.payload, logicalId };
      const leftLength = caret.start.offset - target.start.offset;
      const rightLength = target.end.offset - caret.start.offset;
      const originalStart = target.start;

      event.preventDefault();
      event.stopPropagation();
      const removed = doc.metadata.remove({ id: match.id });
      console.info('[metadata-enter] remove anchor', removed);
      if (!removed.success) return;

      // The metadata removal dispatch must settle before V1 can split the run.
      window.setTimeout(() => {
        const restoredSelection = capture ? uiRef.current?.selection.restore(capture) : null;
        console.info('[metadata-enter] restore caret after anchor removal', restoredSelection);
        if (capture && !restoredSelection?.success) {
          const restoredAnchor = doc.metadata.attach({
            id: match.id,
            namespace: metadata.namespace,
            payload: metadata.payload,
            target,
          });
          console.error('[metadata-enter] caret restore failed; restored original anchor', restoredAnchor);
          return;
        }

        const splitRun = editor.commands.splitRunToParagraph?.() ?? false;
        const split = splitRun || editor.commands.splitBlock();
        console.info('[metadata-enter] split paragraph', { split, splitRun });
        if (!split) {
          const restored = doc.metadata.attach({
            id: match.id,
            namespace: metadata.namespace,
            payload: metadata.payload,
            target,
          });
          console.error('[metadata-enter] split failed; restored original anchor', restored);
          return;
        }

        const afterSplit = uiRef.current?.selection.capture()?.selectionTarget;
        if (!afterSplit || afterSplit.start.kind !== 'text' || afterSplit.start.blockId === originalStart.blockId) {
          console.error('[metadata-enter] native paragraph split did not complete');
          return;
        }
        const newStart = afterSplit.start;

        const left = doc.metadata.attach({
          id: `${match.id}-left-${crypto.randomUUID()}`,
          namespace: metadata.namespace,
          payload,
          target: {
            kind: 'selection',
            start: originalStart,
            end: { kind: 'text', blockId: originalStart.blockId, offset: originalStart.offset + leftLength },
          },
        });
        console.info('[metadata-enter] attach left anchor', left);

        const right = doc.metadata.attach({
          id: `${match.id}-right-${crypto.randomUUID()}`,
          namespace: metadata.namespace,
          payload,
          target: {
            kind: 'selection',
            start: { kind: 'text', blockId: newStart.blockId, offset: 0 },
            end: { kind: 'text', blockId: newStart.blockId, offset: rightLength },
          },
        });
        console.info('[metadata-enter] attach right anchor', right, { logicalId });
        void refreshMetadataOutlines();
      }, 0);
    };

    window.addEventListener('keydown', handleEnter, true);
    return () => window.removeEventListener('keydown', handleEnter, true);
  }, [ready, refreshMetadataOutlines]);

  function importDocument(event: React.ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    if (!file) return;
    setDocument(file);
    event.target.value = '';
  }

  async function exportDocument() {
    if (exportingRef.current) return;
    exportingRef.current = true;
    setExporting(true);
    try {
      await superdocRef.current?.export({ exportType: ['docx'], exportedName: 'sample-edited' });
    } catch (error) {
      console.error('SuperDoc could not export the document.', error);
    } finally {
      exportingRef.current = false;
      if (superdocRef.current) setExporting(false);
    }
  }

  function preserveSelection(event: React.MouseEvent<HTMLButtonElement>) {
    event.preventDefault();
    const capture = uiRef.current?.selection.capture();
    if (capture?.selectionTarget) {
      preservedSelectionRef.current = capture.selectionTarget;
    }
  }

  async function wrapSelectionWithMetadata() {
    const superdoc = superdocRef.current;
    const doc = superdoc?.activeEditor?.doc;
    const live = uiRef.current?.selection.capture();
    const target = live?.selectionTarget ?? preservedSelectionRef.current;
    preservedSelectionRef.current = null;
    if (!doc || !target || target.start.kind !== 'text' || target.end.kind !== 'text') {
      setMetadataStatus('Select text within one paragraph.');
      return;
    }
    if (target.start.blockId !== target.end.blockId || target.start.offset === target.end.offset) {
      setMetadataStatus('Metadata requires a non-empty selection within one paragraph.');
      return;
    }

    setAttachingMetadata(true);
    try {
      const result = await doc.metadata.attach({
        id: `demo-finding-${crypto.randomUUID()}`,
        namespace: 'urn:superdoc:react-demo:findings',
        target,
        payload: { kind: 'manual-test', label: 'React demo selection', createdAt: new Date().toISOString() },
      });
      setMetadataStatus(result.success ? `Attached metadata: ${result.id}` : result.failure.message);
      await refreshMetadataOutlines();
      superdoc.focus();
    } catch (error) {
      setMetadataStatus(error instanceof Error ? error.message : String(error));
    } finally {
      setAttachingMetadata(false);
    }
  }

  function splitMetadataAtCaret() {
    console.info('[metadata-enter] manual split button invoked');
    mountRef.current?.dispatchEvent(new KeyboardEvent('keydown', {
      bubbles: true,
      cancelable: true,
      key: 'Enter',
    }));
  }

  return (
    <main>
      <div className='document-actions'>
        <input
          ref={fileInputRef}
          accept='.docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document'
          hidden
          onChange={importDocument}
          type='file'
        />
        <button className='action-button' onClick={() => fileInputRef.current?.click()} type='button'>
          Import document
        </button>
        <button
          className='action-button action-button--primary'
          disabled={!ready || exporting}
          onClick={() => void exportDocument()}
          type='button'
        >
          Export document
        </button>
        <button
          className='action-button'
          disabled={!ready || attachingMetadata}
          onClick={() => void wrapSelectionWithMetadata()}
          onMouseDown={preserveSelection}
          type='button'
        >
          Wrap selection with metadata
        </button>
        <button
          aria-pressed={showMetadataOutlines}
          className='action-button action-button--toggle'
          disabled={!ready}
          onClick={() => setShowMetadataOutlines((visible) => !visible)}
          type='button'
        >
          {showMetadataOutlines ? 'Hide metadata outlines' : 'Show metadata outlines'}
        </button>
        <button
          className='action-button'
          disabled={!ready}
          onClick={splitMetadataAtCaret}
          onMouseDown={preserveSelection}
          type='button'
        >
          Split metadata at caret
        </button>
      </div>
      <p aria-live='polite' className='metadata-status'>{metadataStatus}</p>
      {showMetadataOutlines && (
        <div aria-hidden='true' className='metadata-outline-layer'>
          {metadataOutlines.map(({ key, left, top, width, height }) => (
            <span className='metadata-outline' key={key} style={{ height, left, top, width }} />
          ))}
        </div>
      )}
      <div ref={mountRef} />
    </main>
  );
}
