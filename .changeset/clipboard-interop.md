---
'@pptx-editor.dev/core': minor
---

Copy and paste work with other applications and inside the editor:

- Copied text keeps its formatting in the editor, and leaves as HTML that presentation and document applications read.
- HTML text from other applications pastes with bold, italic, underline, size, color, font, alignment, list levels, and bullets or numbers. With no text selected, it goes in a new text box. The new `pasteText` command does this for hosts, with `textClipboardParagraphs` to build its paragraphs from HTML or plain text.
- A picture on the clipboard (a screenshot, or shapes copied in another application) pastes as a picture, at the size its resolution gives.
- Copied shapes leave as a picture of the shapes for other applications, or as text when they are text boxes.
- Paste works after a click on another slide in the slide list, and with the focus on the toolbar.
- In the slide list, Copy, Cut, and Paste act on slides. The new `pasteSlides` command puts slides of a presentation file after the current slide.
