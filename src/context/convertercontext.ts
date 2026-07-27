import type { Node } from "dom-parser";
import { DefaultDocxHtmlFileProvider } from "../file/DefaultDocxHtmlFileProvider.js";
import type { DocxHtmlFileProviderConstructor } from "../file/DocxHtmlFileProvider.js";
import type { DocxHtmlFileReader } from "../types/file.js";



export class ConverterContext {

  constructor(public numbering : string, private FileProvider: DocxHtmlFileProviderConstructor | undefined, private _fileReader: DocxHtmlFileReader) { }

  get fileReader(): DocxHtmlFileReader {
    return this._fileReader;
  }

  getFileProvider(src: string): InstanceType<DocxHtmlFileProviderConstructor>| undefined {
    return this.FileProvider ? new this.FileProvider(src) as InstanceType<DocxHtmlFileProviderConstructor> : undefined;
  }

  getCss(node: Node): Record<string, string> {
    const style = node.getAttribute('style');
    if (!style) {
      return {};
    }
    return style.split(';').reduce((acc, rule) => {
      const [property, value] = rule.split(':').map(s => s.trim());
      if (property && value) {
        acc[property] = value;
      }
      return acc;
    }, {} as Record<string, string>);
  }

}