import type { Node } from "dom-parser";
import { DefaultDocxHtmlFileProvider } from "../file/DefaultDocxHtmlFileProvider";
import type { DocxHtmlFileProviderConstructor } from "../file/DocxHtmlFileProvider";



export class ConverterContext {

  constructor(public numbering : string, private FileProvider: DocxHtmlFileProviderConstructor) { }

  getFileProvider(src: string): InstanceType<DocxHtmlFileProviderConstructor> {
    return new this.FileProvider(src) as InstanceType<DocxHtmlFileProviderConstructor>;
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