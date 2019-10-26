import { mergeStyles, prepareStyles, generateStyles } from './merge-styles';
import { mergeContentTypes, mergeRelations, generateContentTypes, generateRelations } from './merge-relations-and-content-type';
import { prepareMediaFiles, copyMediaFiles } from './merge-media';
import { prepareNumbering, mergeNumbering, generateNumbering } from './merge-bullets-numberings';

const JSZip = require('jszip');

export default DocxMerger = (options, files) => {

    const _body = [];
    //const _header = [];
    //const _footer = [];
    //const _Basestyle = options.style || 'source';
    const _style = [];
    const _numbering = [];
    const _pageBreak = typeof options.pageBreak !== 'undefined' ? !!options.pageBreak : true;
    const _files = [];

    (files || []).forEach(function (file) {
        _files.push(new JSZip(file));
    });
    const _contentTypes = {};
    const _media = {};
    const _rel = {};

    let _builder = _body;

    const insertPageBreak = () => {
        const pb = '<w:p> \
					<w:r> \
						<w:br w:type="page"/> \
					</w:r> \
				  </w:p>';

        _builder.push(pb);
    };

    const insertRaw = xml => {
        _builder.push(xml);
    };

    const mergeBody = (files) => {
        _builder = _body;

        mergeContentTypes(files, _contentTypes);
        prepareMediaFiles(files, _media);
        mergeRelations(files, _rel);

        prepareNumbering(files);
        mergeNumbering(files, _numbering);

        prepareStyles(files, _style);
        mergeStyles(files, _style);

        files.forEach(function (zip, index) {
            //let zip = new JSZip(file);
            let xml = zip.file("word/document.xml").asText();
            xml = xml.substring(xml.indexOf("<w:body>") + 8);
            xml = xml.substring(0, xml.indexOf("</w:body>"));
            xml = xml.substring(0, xml.lastIndexOf("<w:sectPr"));

            insertRaw(xml);
            //if (_pageBreak && index < files.length-1)
            //insertPageBreak();
        });
    };

    const save = (type, callback) => {
        let zip = _files[0];
        let xml = zip.file("word/document.xml").asText();
        let startIndex = xml.indexOf("<w:body>") + 8;
        let endIndex = xml.lastIndexOf("<w:sectPr");

        xml = xml.replace(xml.slice(startIndex, endIndex), _body.join(''));

        generateContentTypes(zip, _contentTypes);
        copyMediaFiles(zip, _media, _files);
        generateRelations(zip, _rel);
        generateNumbering(zip, _numbering);
        generateStyles(zip, _style);

        zip.file("word/document.xml", xml);

        callback(zip.generate({
            type: type,
            compression: "DEFLATE",
            compressionOptions: {
                level: 4
            }
        }));
    };


    if (_files.length > 0) {
        mergeBody(_files);
    };
};

