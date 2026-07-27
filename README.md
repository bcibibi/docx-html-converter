# @bcibibi/docx-html-converter

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
npm install @bcibibi/docx-html-converter docx
```

`docx` is a peer dependency and must be installed in your project.

## Quick Start

```ts
import { Document, Packer } from "docx";
import { writeFileSync } from "node:fs";
import { DocxHtmlConverter } from "@bcibibi/docx-html-converter";

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

## Ordered Lists (`ol`)

To render ordered lists correctly, pass a numbering reference and define the corresponding numbering config in your `Document`.

```ts
import { Document, Packer, convertInchesToTwip } from "docx";
import { DocxHtmlConverter } from "@bcibibi/docx-html-converter";

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
- `width` and `height` attributes are used for image size. Default is `100x100` when missing.
- `style="float:left"` and `style="float:right"` are supported.
- If an image cannot be loaded, a fallback text run is inserted: `[Image not found: ...]`.

Supported image types (detected from file content):

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
- `font-size` (unit conversion to half-points used by `docx`)
- `color` (parsed and converted to hex)
- `float: left | right` (for `img`)

## API

### `DocxHtmlConverter.convert(html, options?)`

Converts HTML to an array of `FileChild` values.

```ts
import type { FileChild, IRunOptions } from "docx";

type DocxHtmlConverterOptions = {
	numbering?: string;
	options?: IRunOptions;
	fileReader?: (src: string) => Promise<Buffer>;
	/** @deprecated Use fileReader */
	fileProvider?: new (src: string) => {
		getFileContent(): Promise<Buffer>;
	};
};

function convert(
	html: string,
	options?: DocxHtmlConverterOptions
): Promise<FileChild[]>;
```

### `options` option

`options` lets you provide base `IRunOptions` that are applied before inline HTML/CSS styles are merged.

```ts
const children = await DocxHtmlConverter.convert("<p>Hello</p>", {
	options: { italics: true },
});
```

### `fileReader` option (recommended)

Use `fileReader` to customize image loading behavior (auth, signed URLs, S3, CDN, in-memory assets, etc.).

```ts
import { DocxHtmlConverter } from "@bcibibi/docx-html-converter";

const fileReader = async (src: string): Promise<Buffer> => {
	if (src.startsWith("my://")) {
		return Buffer.from([]);
	}
	throw new Error(`Unsupported source: ${src}`);
};

const children = await DocxHtmlConverter.convert("<img src='my://asset/logo' />", {
	fileReader,
});
```

### `fileProvider` option (deprecated)

`fileProvider` is still available for backward compatibility but is deprecated and will be removed in a future version.

```ts
import {
	DocxHtmlConverter,
	DocxHtmlFileProvider,
} from "@bcibibi/docx-html-converter";

class MyFileProvider extends DocxHtmlFileProvider {
	async getFileContent(): Promise<Buffer> {
		return Buffer.from([]);
	}
}

const children = await DocxHtmlConverter.convert("<img src='my://asset/logo' />", {
	fileProvider: MyFileProvider,
});
```

## Development

Install dependencies:

```bash
npm ci
```

Build the package:

```bash
npm run build
```

Run tests:

```bash
npm test
```

Verify both module outputs:

```bash
npm run verify:modules
```

## License

MIT