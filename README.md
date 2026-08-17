# Metadata range splitting

Standalone React demo for splitting an anchored SuperDoc metadata range when a user presses Enter inside it.

The workaround removes the original SDT anchor, restores the captured caret, splits the paragraph, and attaches separate metadata records to both resulting paragraphs. Both records carry the same application-level `logicalId` in their payloads.

## Run it

Requires Node 22.12 or newer and pnpm 10.

```bash
pnpm install
pnpm dev
```

Open the printed URL, select text within one paragraph, and choose **Wrap selection with metadata**. Place the caret inside the outlined range and press Enter. The **Split metadata at caret** button invokes the same path for manual debugging.

## Verify it

```bash
pnpm typecheck
pnpm build
pnpm browsers
pnpm test
```

The demo intentionally uses `superdoc@1.46.1`, the latest V1 release.

See [Mount SuperDoc in React](https://docs.superdoc.dev/editor/frameworks/react) for the guided explanation.
