const {XMLSerializer} = require('@xmldom/xmldom');
const {DOMParser} = require('@xmldom/xmldom');


const mergeContentTypes = function(files, _contentTypes) {
    files.forEach(async (zip) => {
        const xmlString = await zip.file("[Content_Types].xml").async('string');
        const xml = new DOMParser().parseFromString(xmlString, 'text/xml');

        const childNodes = xml.getElementsByTagName('Types')[0].childNodes;

        for (const node in childNodes) {
            if (/^\d+$/.test(node) && childNodes[node].getAttribute) {
                const contentType = childNodes[node].getAttribute('ContentType');
                if (!_contentTypes[contentType])
                    _contentTypes[contentType] = childNodes[node].cloneNode();
            }
        }
    });
};

const mergeRelations = function(files, _rel) {
    files.forEach(async (zip) => {
        const xmlString = await zip.file("word/_rels/document.xml.rels").async('string');
        const xml = new DOMParser().parseFromString(xmlString, 'text/xml');

        const childNodes = xml.getElementsByTagName('Relationships')[0].childNodes;

        for (const node in childNodes) {
            if (/^\d+$/.test(node) && childNodes[node].getAttribute) {
                const Id = childNodes[node].getAttribute('Id');
                if (!_rel[Id])
                    _rel[Id] = childNodes[node].cloneNode();
            }
        }
    });
};

const generateContentTypes = async function(zip, _contentTypes) {
    let xmlString = await zip.file("[Content_Types].xml").async('string');
    const xml = new DOMParser().parseFromString(xmlString, 'text/xml');
    const serializer = new XMLSerializer();

    const types = xml.documentElement.cloneNode();

    for (const node in _contentTypes) {
        types.appendChild(_contentTypes[node]);
    }

    const startIndex = xmlString.indexOf("<Types");
    xmlString = xmlString.replace(xmlString.slice(startIndex), serializer.serializeToString(types));

    zip.file("[Content_Types].xml", xmlString);
};

const generateRelations = async function(zip, _rel) {
    let xmlString = await zip.file("word/_rels/document.xml.rels").async('string');
    const xml = new DOMParser().parseFromString(xmlString, 'text/xml');
    const serializer = new XMLSerializer();

    const types = xml.documentElement.cloneNode();

    for (const node in _rel) {
        types.appendChild(_rel[node]);
    }

    const startIndex = xmlString.indexOf("<Relationships");
    xmlString = xmlString.replace(xmlString.slice(startIndex), serializer.serializeToString(types));

    zip.file("word/_rels/document.xml.rels", xmlString);
};


module.exports = {
    mergeContentTypes: mergeContentTypes,
    mergeRelations: mergeRelations,
    generateContentTypes: generateContentTypes,
    generateRelations: generateRelations
};