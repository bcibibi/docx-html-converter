import fs from "fs";
import http from "http";
import https from "https";


export function defaultFileReader(src: string): Promise<Buffer> {

    const getHTTPFileContent = () => new Promise<Buffer>((resolve, reject) => {
        const client = src.startsWith("https://") ? https : http;
        client.get(src, (res) => {
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

    const getLocalFileContent = () => new Promise<Buffer>((resolve, reject) => {
        fs.readFile(src, (err, data) => {
            if (err) {
                reject(err);
            } else {
                resolve(data);
            }
        });
    });


    if (src.startsWith("http://") || src.startsWith("https://")) {
        return getHTTPFileContent();
    } else {
        return getLocalFileContent();
    }

}