

export type DocxHtmlFileType = "jpg" | "png" | "gif" | "bmp";

export type DocxHtmlFileReader = (src: string) => Promise<Buffer>;