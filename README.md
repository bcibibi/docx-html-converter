# docx-html-converter

Convert a subset of HTML into `docx` components (`FileChild[]`) so you can build Word documents with the [`docx`](https://www.npmjs.com/package/docx) library.

## Features

- Converts HTML to `docx` document children.
- Supports common inline formatting (`bold`, `italic`, `underline`, `strike`, `sub`, `sup`).
- Supports paragraphs and nested lists (`ul` / `ol` / `li`).
- Supports inline image conversion from local files or HTTP(S) sources.
- Supports inline CSS parsing for selected text styles.
- Works in ESM and CommonJS builds.

## Installation

```bash
npm install docx-html-converter docx
```

`docx` is a peer dependency and must be installed in your project.

## Quick Start

```ts
import { Document, Packer } from "docx";
import { DocxHtmlConverter } from "docx-html-converter";
import { writeFileSync } from "node:fs";

const html = `
	<p>Hello <strong>World</strong></p>
	<ul>
		<li>First item</li>
		<li>Second item</li>
	</ul>
`;

const children = await DocxHtmlConverter.convert(html);

const doc = new Document({
	sections: [{ children }],
});

const buffer = await Packer.toBuffer(doc);
writeFileSync("output.docx", buffer);
```

## Numbered Lists (`ol`)

To render ordered lists correctly, pass a numbering reference and define the corresponding numbering config in your `Document`.

```ts
import { Document, Packer, convertInchesToTwip } from "docx";
import { DocxHtmlConverter } from "docx-html-converter";

const numberingRef = "my-numbering";
const html = "<ol><li>One</li><li>Two</li></ol>";

const children = await DocxHtmlConverter.convert(html, { numbering: numberingRef });

const doc = new Document({
	sections: [{ children }],
	numbering: {
		config: [
			{
				reference: numberingRef,
				levels: [
					{
						level: 0,
						format: "decimal",
						text: "%1.",
						alignment: "left",
						style: {
							paragraph: {
								indent: {
									left: convertInchesToTwip(0.5),
									hanging: convertInchesToTwip(0.25),
								},
							},
						},
					},
				],
			},
		],
	},
});

const buffer = await Packer.toBuffer(doc);
```

## Images

`<img>` tags are supported.

- `src` can be a local path or an `http://` / `https://` URL.
- `width` and `height` attributes are used for image size.
- `style="float:left"` and `style="float:right"` are supported.
- If an image cannot be loaded, a fallback text run is inserted: `[Image not found: ...]`.

Supported image types:

- `jpg`
- `png`
- `gif`
- `bmp`

## Supported HTML Tags

- `#text`
- `p`
- `span`
- `ul`
- `ol`
- `li`
- `b`, `strong`
- `i`, `em`
- `u`
- `strike`
- `sub`
- `sup`
- `img`

## Supported Inline CSS

The converter parses inline `style` attributes and maps these properties:

- `font-weight: bold`
- `font-style: italic`
- `text-decoration: underline | line-through`
- `vertical-align: sub | super`
- `font-size` (unit conversion to points)
- `color` (parsed and converted to hex)
- `float: left | right` (for `img`)

## API

### `DocxHtmlConverter.convert(html, options?)`

Converts HTML to an array of `FileChild` values.

```ts
type DocxHtmlConverterOptions = {
	numbering?: string;
	fileProvider?: DocxHtmlFileProviderConstructor;
};

function convert(
	html: string,
	options?: DocxHtmlConverterOptions
): Promise<FileChild[]>;
```

### `fileProvider` option

Use a custom file provider if you need custom image loading behavior (auth, S3, CDN signing, in-memory assets, etc.).

```ts
import { DocxHtmlFileProvider, DocxHtmlConverter } from "docx-html-converter";

class MyFileProvider extends DocxHtmlFileProvider {
	async getFileContent(): Promise<Buffer> {
		// Implement your own source loading here
		return Buffer.from([]);
	}
}

const children = await DocxHtmlConverter.convert("<img src='my://asset/logo' />", {
	fileProvider: MyFileProvider,
});
```

## Development

Build the package:

```bash
npm run build
```

Verify both module outputs:

```bash
npm run verify:modules
```

## License

MIT