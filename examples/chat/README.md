# PPTX editor chat panel example

Open a deck from a chat message in a panel beside the conversation. The assistant is scripted, and the deck stays in your browser.

## Run the example

From the repository root, run:

```bash
bun install
bun run dev:chat
```

Open `http://localhost:5276`.

## Try it

1. Select **Open** on the deck card. The editor slides in beside the chat.
2. Edit a slide. The card in the chat shows the change.
3. Type `show slide 6`, `present`, or `download`.
4. Drag the divider, or select **Expand**. A wide panel moves the slide rail to the left.

## How it works

One `PptxEditor.Root` wraps the chat and the panel, so the chat reads and drives the deck. A tall panel stacks the slide, `<PptxEditor.SlideRail orientation="horizontal" />`, and the notes.

For more information, see [Show the editor beside a chat](../../docs/site/content/guides/chat-panel.mdx).
