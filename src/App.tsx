import { useCallback, useEffect, useRef, useState } from 'react';
import { SuperDoc } from 'superdoc';
import { createSuperDocUI } from 'superdoc/ui';
import type { SelectionTarget, SuperDocUI, ViewportRect } from 'superdoc/ui';
import 'superdoc/style.css';

type DemoMetadataPayload = {
  kind: 'manual-test';
  label: string;
  createdAt: string;
  logicalId: string;
};

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

  // ============================================================================
  // METADATA RANGE SPLITTING WORKAROUND
  // V1 has no Document API block split, so move the paragraph tail into a
  // newly created paragraph, then anchor both ranges with one logical ID.
  // ============================================================================
  const splitMetadataRangeAtCaret = useCallback(() => {
    // Read the active document and collapsed caret selection.
    const doc = superdocRef.current?.activeEditor?.doc;
    const caret = uiRef.current?.selection.capture()?.selectionTarget ?? preservedSelectionRef.current;
    if (!doc || !caret || caret.start.kind !== 'text' || caret.end.kind !== 'text') return false;
    if (caret.start.offset !== caret.end.offset) return false;
    const caretPoint = caret.start;

    // Find the metadata anchor containing the caret.
    const entry = doc.metadata.list({ resolvedOnly: true, within: caret }).items[0];
    if (!entry) return false;

    // Resolve the original payload and anchored text range.
    const metadata = doc.metadata.get({ id: entry.id })!;
    const target = doc.metadata.resolve({ id: entry.id })!.target;
    if (target.start.kind !== 'text' || target.end.kind !== 'text') return false;

    // Read the containing paragraph and its structural address.
    const payload = metadata.payload as DemoMetadataPayload;
    const paragraph = doc.extract({}).blocks.find((block) => block.nodeId === caretPoint.blockId)!;
    const block = doc.getNodeById({ nodeId: caretPoint.blockId }).address;
    if (block.kind !== 'block') return false;

    // Remove the original SDT anchor before changing paragraph structure.
    doc.metadata.remove({ id: entry.id });
    // Remove all paragraph text after the caret from the original block.
    doc.delete({
      behavior: 'exact',
      target: {
        kind: 'selection',
        start: caretPoint,
        end: { kind: 'text', blockId: caretPoint.blockId, offset: paragraph.text.length },
      },
    });

    // Recreate the removed text in a new paragraph after the original block.
    const created = doc.create.paragraph({
      at: { kind: 'after', target: block },
      text: paragraph.text.slice(caretPoint.offset),
    });
    if (!created.success) return true;

    // Anchor the metadata portion remaining in the original paragraph.
    doc.metadata.attach({
      id: crypto.randomUUID(),
      namespace: metadata.namespace,
      payload,
      target: {
        kind: 'selection',
        start: target.start,
        end: caretPoint,
      },
    });
    // Anchor the metadata remainder in the new paragraph with the same payload.
    doc.metadata.attach({
      id: crypto.randomUUID(),
      namespace: metadata.namespace,
      payload,
      target: {
        kind: 'selection',
        start: { kind: 'text', blockId: created.paragraph.nodeId, offset: 0 },
        end: {
          kind: 'text',
          blockId: created.paragraph.nodeId,
          offset: target.end.offset - caretPoint.offset,
        },
      },
    });

    // Repaint the visual metadata outlines after both anchors are attached.
    void refreshMetadataOutlines();
    return true;
  }, [refreshMetadataOutlines]);

  // ============================================================================
  // ENTER KEY INTERCEPTION WORKAROUND
  // V1 blocks a native paragraph split while the caret is inside an inline SDT.
  // Capture Enter first and delegate the actual split to the Document API flow.
  // ============================================================================
  useEffect(() => {
    if (!ready) return;

    const handleEnter = (event: KeyboardEvent) => {
      if (event.key !== 'Enter' || event.shiftKey || event.altKey || event.ctrlKey || event.metaKey || event.isComposing) {
        return;
      }
      if (!splitMetadataRangeAtCaret()) return;
      event.preventDefault();
      event.stopPropagation();
    };

    window.addEventListener('keydown', handleEnter, true);
    return () => window.removeEventListener('keydown', handleEnter, true);
  }, [ready, splitMetadataRangeAtCaret]);

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
        payload: {
          kind: 'manual-test',
          label: 'React demo selection',
          createdAt: new Date().toISOString(),
          logicalId: crypto.randomUUID(),
        } satisfies DemoMetadataPayload,
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
    splitMetadataRangeAtCaret();
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
