const XMLSerializer = require('xmldom').XMLSerializer;
const DOMParser = require('xmldom').DOMParser;

export const prepareMediaFiles = (files, media) => {

    let count = 1;

    files.forEach(function (zip, index) {
        // let zip = new JSZip(file);
        const medFiles = zip.folder("word/media").files;

        for (let mfile in medFiles) {
            if (/^word\/media/.test(mfile) && mfile.length > 11) {
                // console.log(mfile);
                media[count] = {};
                media[count].oldTarget = mfile;
                media[count].newTarget = mfile.replace(/[0-9]/, '_' + count).replace('word/', "");
                media[count].fileIndex = index;
                updateMediaRelations(zip, count, media);
                updateMediaContent(zip, count, media);
                count++;
            }
        }
    });

    // console.log(JSON.stringify(media));
    // updateRelation(files);
};

export const updateMediaRelations = (zip, count, _media) => {

    let xmlString = zip.file("word/_rels/document.xml.rels").asText();
    const xml = new DOMParser().parseFromString(xmlString, 'text/xml');

    const childNodes = xml.getElementsByTagName('Relationships')[0].childNodes;
    const serializer = new XMLSerializer();

    for (let node in childNodes) {
        if (/^\d+$/.test(node) && childNodes[node].getAttribute) {
            let target = childNodes[node].getAttribute('Target');
            if ('word/' + target == _media[count].oldTarget) {

                _media[count].oldRelID = childNodes[node].getAttribute('Id');

                childNodes[node].setAttribute('Target', _media[count].newTarget);
                childNodes[node].setAttribute('Id', _media[count].oldRelID + '_' + count);
            }
        }
    }

    // console.log(serializer.serializeToString(xml.documentElement));

    const startIndex = xmlString.indexOf("<Relationships");
    xmlString = xmlString.replace(xmlString.slice(startIndex), serializer.serializeToString(xml.documentElement));

    zip.file("word/_rels/document.xml.rels", xmlString);

    // console.log( xmlString );
};

export const updateMediaContent = (zip, count, _media) => {

    let xmlString = zip.file("word/document.xml").asText();
    //const xml = new DOMParser().parseFromString(xmlString, 'text/xml');

    xmlString = xmlString.replace(new RegExp(_media[count].oldRelID + '"', 'g'), _media[count].oldRelID + '_' + count + '"');

    zip.file("word/document.xml", xmlString);
};

export const copyMediaFiles = (base, _media, _files) => {

    for (let media in _media) {
        const content = _files[_media[media].fileIndex].file(_media[media].oldTarget).asUint8Array();

        base.file('word/' + _media[media].newTarget, content);
    }
};
