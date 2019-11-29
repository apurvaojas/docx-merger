import { getFileBinary } from '../utils';
import JSZip from 'jszip';

export class Loader {
    constructor() {

    }
   
    async unzipFiles(files: Array<Blob>): Promise<Array<JSZip>>{
        const loadPromises = files.map(f=> JSZip.loadAsync(f));
        return Promise.all(loadPromises);
    }

    async load(files: Array<Blob>|Array<string>, isBinary: Boolean = false): Promise<Array<JSZip> | Error> {
        if (!files.length) return Promise.reject(new Error("Please Provide Files."));

        // const isFilePath = typeof files[0] === 'string';
        // // console.log(files, isFilePath)

        if (!isBinary) {
            return this.getFiles(<Array<string>>files);
        } else {
            return this.unzipFiles(<Array<Blob>>files);
        }
    }

    async getFiles(files: Array<string>): Promise<Array<JSZip>>{
        const promises = files.map(f=> getFileBinary(f));
        const _files = await Promise.all(promises);
        return this.unzipFiles(_files);
    }


}