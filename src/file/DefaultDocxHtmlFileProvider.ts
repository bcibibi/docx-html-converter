import fs from "fs";
import http from "http";
import https from "https";
import { DocxHtmlFileProvider } from "./DocxHtmlFileProvider";

export class DefaultDocxHtmlFileProvider extends DocxHtmlFileProvider {

    getFileContent(): Promise<Buffer> {
        if (this.src.startsWith("http://") || this.src.startsWith("https://")) {
            return this.getHTTPFileContent();
        } else {
            return this.getLocalFileContent();
        }
    }

    private getHTTPFileContent(): Promise<Buffer> {
        return new Promise((resolve, reject) => {
            const client = this.src.startsWith("https://") ? https : http;
            client.get(this.src, (res) => {
                const data: Uint8Array[] = [];
                res.on("data", (chunk) => {
                    data.push(chunk);
                });
                res.on("end", () => {
                    resolve(Buffer.concat(data));
                });
                res.on("error", (err) => {
                    reject(err);
                });
            });
        });
    }

    private getLocalFileContent(): Promise<Buffer> {
        return new Promise((resolve, reject) => {
            fs.readFile(this.src, (err, data) => {
                if (err) {
                    reject(err);
                } else {
                    resolve(data);
                }
            });
        });
    }

}