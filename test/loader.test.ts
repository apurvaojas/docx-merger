import { expect, assert } from 'chai';
import { Loader } from '../src/core/loader';

import JSZip from 'jszip';
const loader = new Loader();
const path = require('path');
import * as sinon from 'sinon';
import { JSDOM } from 'jsdom';
import { SinonStub } from 'sinon';
import fs from 'fs';
const window = new JSDOM('').window;
global.window = window;

let stubIsBrowser: SinonStub;

// const file: Buffer = fs.readFileSync(path.resolve(__dirname, 'mock/template.docx'), { encoding: 'binary' });
// console.log(file);
describe('Loader Node Js', () => {
    it('Load files from File system', (done) => {
        loader.load([path.resolve(__dirname, 'mock/template.docx'), path.resolve(__dirname, 'mock/template1.docx')])
            .then((files) => {
                Array.isArray(files) && files.forEach(f => expect(f).to.satisfy((v: JSZip) => v instanceof JSZip));
                done();
            });
    });

    it('Should throw error when files path not provided', (done) => {
        loader.load([])
            .then((files) => {
            }).catch((err) => {
                expect(err).to.satisfy((err: Error) => err instanceof Error);
                done();
            });
    });

    it('Should throw error when files path incorrect', (done) => {
        loader.load(['icorrect file path'])
            .then((files) => {
            }).catch((err) => {
                expect(err).to.satisfy((err: Error) => err instanceof Error);
                done();
            });
    });

    it('Should Accept Binary file', (done) => {
        fs.readFile(path.resolve(__dirname, 'mock/template.docx'), 'binary', (err: Error, data: any) => {

            loader.load([data], true)
            .then((files) => {
                Array.isArray(files) && files.forEach(f => expect(f).to.satisfy((v: JSZip) => v instanceof JSZip));
                done();
            });
        });
        
    });
});

