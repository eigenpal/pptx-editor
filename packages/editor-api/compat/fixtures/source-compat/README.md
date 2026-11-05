# Office.js compatibility fixtures

These files adapt Microsoft Office.js examples to check API compatibility at compile time. They do not run or ship in the published package.

## License scope

Microsoft sample portions retain the MIT License in [MICROSOFT-LICENSE.txt](MICROSOFT-LICENSE.txt). The EigenPal Pro License applies to EigenPal additions. The file headers identify both sources of code.

Microsoft's [sample repository metadata](https://github.com/OfficeDev/office-js-snippets/blob/9507659e7ed016bc37195f7fb1f153287a542378/package.json) identifies Microsoft as the author and MIT as the license. The [Office.js reference documentation license](https://github.com/OfficeDev/office-js-docs-reference/blob/e05728bff4b91ccef667eff35fbf7ec697629c8a/LICENSE-CODE) supplies the Microsoft copyright and MIT terms. The local license file preserves that text exactly.

## Reference examples

These links pin the Microsoft examples used for the API call patterns. They do not identify every local statement as copied code.

| Fixture | Microsoft examples |
| --- | --- |
| `hyperlinks.ts` | [Hyperlinks](https://github.com/OfficeDev/office-js-snippets/blob/9507659e7ed016bc37195f7fb1f153287a542378/samples/powerpoint/hyperlinks/manage-hyperlinks.yaml) |
| `selection.ts` | [Slide selection](https://github.com/OfficeDev/office-js-snippets/blob/9507659e7ed016bc37195f7fb1f153287a542378/samples/powerpoint/slide-management/get-set-slides.yaml), [Shape selection](https://github.com/OfficeDev/office-js-snippets/blob/9507659e7ed016bc37195f7fb1f153287a542378/samples/powerpoint/shapes/get-set-shapes.yaml), [Text selection](https://github.com/OfficeDev/office-js-snippets/blob/9507659e7ed016bc37195f7fb1f153287a542378/samples/powerpoint/text/get-set-textrange.yaml) |
| `shapes-fill-and-line.ts` | [Shapes](https://github.com/OfficeDev/office-js-snippets/blob/9507659e7ed016bc37195f7fb1f153287a542378/samples/powerpoint/shapes/shapes.yaml), [Shape groups](https://github.com/OfficeDev/office-js-snippets/blob/9507659e7ed016bc37195f7fb1f153287a542378/samples/powerpoint/shapes/group-ungroup-shapes.yaml) |
| `slides.ts` | [Add slides and list masters](https://github.com/OfficeDev/office-js-snippets/blob/9507659e7ed016bc37195f7fb1f153287a542378/samples/powerpoint/slide-management/add-slides.yaml), [Slide selection](https://github.com/OfficeDev/office-js-snippets/blob/9507659e7ed016bc37195f7fb1f153287a542378/samples/powerpoint/slide-management/get-set-slides.yaml) |
| `tables.ts` | [Tables](https://github.com/OfficeDev/office-js-snippets/blob/9507659e7ed016bc37195f7fb1f153287a542378/samples/powerpoint/shapes/add-modify-tables.yaml) |
| `tags.ts` | [Tags](https://github.com/OfficeDev/office-js-snippets/blob/9507659e7ed016bc37195f7fb1f153287a542378/samples/powerpoint/tags/tags.yaml) |
| `text-box-and-font.ts` | [Text boxes](https://github.com/OfficeDev/office-js-snippets/blob/9507659e7ed016bc37195f7fb1f153287a542378/samples/powerpoint/shapes/shapes.yaml), [Text ranges](https://github.com/OfficeDev/office-js-snippets/blob/9507659e7ed016bc37195f7fb1f153287a542378/samples/powerpoint/text/get-set-textrange.yaml) |

## Adaptations

The fixtures replace the `PowerPoint` namespace with `PptxEditor` and use `runtime.run` for batches. They remove host interface code and add exports, type assertions, and local test values. Some examples combine operations or add property loads to check more API signatures. For example, `logSlideMastersAndLayouts` adapts the nested master and layout logging example from `add-slides.yaml`. It changes collection access and separates the layout log messages.
