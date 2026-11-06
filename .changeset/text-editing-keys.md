---
'@pptx-editor.dev/core': minor
'@pptx-editor.dev/react': minor
'@pptx-editor.dev/vue': minor
---

Lists and text are faster to edit from the keyboard.

- Enter starts a new list item at the same level, also on an empty item.
- Tab and Shift+Tab change the level only at the start of a list item or across paragraphs. Elsewhere, both keys type a tab.
- Backspace at the start of a list item removes its bullet or number and keeps its level. The text moves to where the bullet was, and the next Backspace joins the item to the paragraph above. The new `removeBullet` command does this step.
- Ctrl/Cmd+L, E, R, and J align paragraphs. Ctrl/Cmd+Shift+> and < step the font size. Ctrl/Cmd+] and [ step the list level. Shift+F3 steps the case.
- Alt/Ctrl+Up and Down move to a paragraph start. A word step forward stops at the start of the next word. Cmd+Backspace deletes the word before the caret.
- Enter on a selected shape selects all its text.
- AutoCorrect works while you type: curly quotes, symbols, fractions, en dashes, sentence capitals, and lists started by typing `1.`, `1)`, `*`, or `-` and a space. Backspace right after a new list keeps the typed text. The new `autoCorrect` option and prop turn AutoCorrect off.
