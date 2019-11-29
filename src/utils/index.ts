


const fs = require('fs');

const getFileFS = async (path: string) => {
    return new Promise(function (resolve, reject) {
        fs.readFile(path, 'binary',(err: Error, data: any) => {
            if (err) reject(err);
            resolve(data);
        });
    });
}

const getFileBinary = async (path: string): Promise<any> => {
    return getFileFS(path);
}


export {
    getFileBinary
}