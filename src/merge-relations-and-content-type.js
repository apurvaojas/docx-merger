const XMLSerializer = require('xmldom').XMLSerializer;
const DOMParser = require('xmldom').DOMParser;

export const mergeContentTypes = (files, _contentTypes) => {

    files.forEach(function (zip) {
        // let zip = new JSZip(file);
        const xmlString = zip.file("[Content_Types].xml").asText();
        const xml = new DOMParser().parseFromString(xmlString, 'text/xml');

        const childNodes = xml.getElementsByTagName('Types')[0].childNodes;

        for (let node in childNodes) {
            if (/^\d+$/.test(node) && childNodes[node].getAttribute) {
                const contentType = childNodes[node].getAttribute('ContentType');
                if (!_contentTypes[contentType])
                    _contentTypes[contentType] = childNodes[node].cloneNode();
            }
        }

    });
};

export const mergeRelations = (files, _rel) => {

    files.forEach(function (zip) {
        // let zip = new JSZip(file);
        const xmlString = zip.file("word/_rels/document.xml.rels").asText();
        const xml = new DOMParser().parseFromString(xmlString, 'text/xml');

        const childNodes = xml.getElementsByTagName('Relationships')[0].childNodes;

        for (let node in childNodes) {
            if (/^\d+$/.test(node) && childNodes[node].getAttribute) {
                const Id = childNodes[node].getAttribute('Id');
                if (!_rel[Id])
                    _rel[Id] = childNodes[node].cloneNode();
            }
        }

    });
};

export const generateContentTypes = (zip, _contentTypes) => {
    // body...
    let xmlString = zip.file("[Content_Types].xml").asText();
    const xml = new DOMParser().parseFromString(xmlString, 'text/xml');
    const serializer = new XMLSerializer();

    const types = xml.documentElement.cloneNode();

    for (let node in _contentTypes) {
        types.appendChild(_contentTypes[node]);
    }

    const startIndex = xmlString.indexOf("<Types");
    xmlString = xmlString.replace(xmlString.slice(startIndex), serializer.serializeToString(types));

    zip.file("[Content_Types].xml", xmlString);
};

export const generateRelations = (zip, _rel) => {
    // body...
    let xmlString = zip.file("word/_rels/document.xml.rels").asText();
    const xml = new DOMParser().parseFromString(xmlString, 'text/xml');
    const serializer = new XMLSerializer();

    const types = xml.documentElement.cloneNode();

    for (let node in _rel) {
        types.appendChild(_rel[node]);
    }

    const startIndex = xmlString.indexOf("<Relationships");
    xmlString = xmlString.replace(xmlString.slice(startIndex), serializer.serializeToString(types));

    zip.file("word/_rels/document.xml.rels", xmlString);
};
