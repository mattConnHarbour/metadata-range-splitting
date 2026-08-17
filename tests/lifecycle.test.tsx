import { act, StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import { expect, test, vi } from 'vitest';

const instances = vi.hoisted(() =>
  [] as Array<{
    destroyed: boolean;
    exportCalls: number;
    exportPromise: Promise<void>;
    document: string | File | undefined;
    fail: () => void;
    finishExport: () => void;
    ready: () => void;
  }>,
);

vi.mock('superdoc', () => ({
  SuperDoc: class {
    instance: (typeof instances)[number];

    constructor(config: {
      document?: string | File;
      onException?: (payload: { error: Error }) => void;
      onReady?: () => void;
    }) {
      let finishExport = () => {};
      const exportPromise = new Promise<void>((resolve) => {
        finishExport = resolve;
      });
      this.instance = {
        destroyed: false,
        document: config.document,
        exportCalls: 0,
        exportPromise,
        fail: () => config.onException?.({ error: new Error('load failed') }),
        finishExport,
        ready: () => config.onReady?.(),
      };
      instances.push(this.instance);
    }

    destroy() {
      this.instance.destroyed = true;
    }

    export() {
      this.instance.exportCalls += 1;
      return this.instance.exportPromise;
    }
  },
}));
vi.mock('superdoc/style.css', () => ({}));

async function renderApp() {
  (globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
  instances.length = 0;
  const container = document.createElement('div');
  document.body.append(container);
  const root = createRoot(container);
  const { default: App } = await import('../src/App');
  await act(async () => root.render(<StrictMode><App /></StrictMode>));
  return { container, root };
}

test('destroys every SuperDoc instance when React unmounts', async () => {
  const { container, root } = await renderApp();
  expect(instances).toHaveLength(2);
  expect(instances[0].destroyed).toBe(true);
  expect(instances[1].destroyed).toBe(false);

  await act(async () => root.unmount());
  expect(instances.every((instance) => instance.destroyed)).toBe(true);
  container.remove();
});

test('ignores readiness from the instance StrictMode replaced', async () => {
  const { container, root } = await renderApp();
  const button = container.querySelectorAll('button')[1];
  expect(button?.disabled).toBe(true);

  await act(async () => instances[0].ready());
  expect(button?.disabled).toBe(true);
  await act(async () => instances[1].ready());
  expect(button?.disabled).toBe(false);

  await act(async () => root.unmount());
  container.remove();
});

test('keeps export disabled when loading fails before readiness', async () => {
  const error = vi.spyOn(console, 'error').mockImplementation(() => {});
  const { container, root } = await renderApp();
  const button = container.querySelectorAll('button')[1];

  await act(async () => instances[1].fail());
  await act(async () => instances[1].ready());
  expect(button?.disabled).toBe(true);
  expect(error).toHaveBeenCalledOnce();

  await act(async () => root.unmount());
  error.mockRestore();
  container.remove();
});

test('allows only one export at a time', async () => {
  const { container, root } = await renderApp();
  const button = container.querySelectorAll('button')[1];
  await act(async () => instances[1].ready());

  await act(async () => {
    button?.click();
    button?.click();
  });
  expect(instances[1].exportCalls).toBe(1);
  expect(button?.disabled).toBe(true);

  await act(async () => instances[1].finishExport());
  expect(button?.disabled).toBe(false);

  await act(async () => root.unmount());
  container.remove();
});

test('replaces the editor document with an imported DOCX', async () => {
  const { container, root } = await renderApp();
  const input = container.querySelector<HTMLInputElement>('input[type="file"]');
  const file = new File(['docx'], 'imported.docx', {
    type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });
  Object.defineProperty(input, 'files', { configurable: true, value: [file] });

  await act(async () => input?.dispatchEvent(new Event('change', { bubbles: true })));

  expect(instances.at(-1)?.document).toBe(file);
  expect(instances.at(-2)?.destroyed).toBe(true);

  await act(async () => root.unmount());
  container.remove();
});
