const {XMLSerializer} = require('@xmldom/xmldom');
const {DOMParser} = require('@xmldom/xmldom');


const prepareMediaFiles = function(files, media) {
    let count = 1;
    files.forEach(async function(zip, index) {
        const medFiles = zip.folder("word/media").files;
        for (const mfile in medFiles) {
            if (/^word\/media/.test(mfile) && mfile.length > 11) {
                media[count] = {};
                media[count].oldTarget = mfile;
                media[count].newTarget = mfile.replace(/[0-9]/, '_' + count).replace('word/', "");
                media[count].fileIndex = index;
                await updateMediaRelations(zip, count, media);
                await updateMediaContent(zip, count, media);
                count++;
            }
        }
    });
};

const updateMediaRelations = async function(zip, count, _media) {
    let xmlString = await zip.file("word/_rels/document.xml.rels").async('string');
    let xml = new DOMParser().parseFromString(xmlString, 'text/xml');

    const childNodes = xml.getElementsByTagName('Relationships')[0].childNodes;
    const serializer = new XMLSerializer();

    for (const node in childNodes) {
        if (/^\d+$/.test(node) && childNodes[node].getAttribute) {
            const target = childNodes[node].getAttribute('Target');
            if ('word/' + target === _media[count].oldTarget) {

                _media[count].oldRelID = childNodes[node].getAttribute('Id');

                childNodes[node].setAttribute('Target', _media[count].newTarget);
                childNodes[node].setAttribute('Id', _media[count].oldRelID + '_' + count);
            }
        }
    }

    const startIndex = xmlString.indexOf("<Relationships");
    xmlString = xmlString.replace(xmlString.slice(startIndex), serializer.serializeToString(xml.documentElement));

    zip.file("word/_rels/document.xml.rels", xmlString);
};

const updateMediaContent = async function(zip, count, _media) {
    let xmlString = await zip.file("word/document.xml").async('string');
    xmlString = xmlString.replace(new RegExp(_media[count].oldRelID + '"', 'g'), _media[count].oldRelID + '_' + count + '"');
    zip.file("word/document.xml", xmlString);
};

const copyMediaFiles = async function(base, _media, _files) {
    for (const media in _media) {
        const content = await _files[_media[media].fileIndex].file(_media[media].oldTarget).async('uint8array');
        base.file('word/' + _media[media].newTarget, content);
    }
};

module.exports = {
    prepareMediaFiles: prepareMediaFiles,
    updateMediaRelations: updateMediaRelations,
    updateMediaContent: updateMediaContent,
    copyMediaFiles: copyMediaFiles
};