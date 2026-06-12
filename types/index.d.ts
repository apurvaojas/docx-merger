export = DocxMerger;

declare class DocxMerger {
    constructor(
        options: DocxMerger.Options,
        files: Array<string | Uint8Array | ArrayBuffer | Buffer>
    );

    save(type: 'nodebuffer', callback: (data: Buffer) => void): void;
    save(type: 'blob', callback: (data: Blob) => void): void;
    save(type: 'uint8array', callback: (data: Uint8Array) => void): void;
    save(type: 'arraybuffer', callback: (data: ArrayBuffer) => void): void;
    save(type: 'base64' | 'string' | 'binarystring', callback: (data: string) => void): void;

    save(type: 'nodebuffer'): Promise<Buffer>;
    save(type: 'blob'): Promise<Blob>;
    save(type: 'uint8array'): Promise<Uint8Array>;
    save(type: 'arraybuffer'): Promise<ArrayBuffer>;
    save(type: 'base64' | 'string' | 'binarystring'): Promise<string>;
}

declare namespace DocxMerger {
    interface Options {
        /** Insert a page break between merged files. Default: true */
        pageBreak?: boolean;
        /** Style handling mode. Only 'source' is implemented. */
        style?: 'source';
    }
}
