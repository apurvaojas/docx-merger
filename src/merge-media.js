
var XMLSerializer = require('xmldom').XMLSerializer;
var DOMParser = require('xmldom').DOMParser;


var prepareMediaFiles = function(files, media) {

       var count = 1;

    files.forEach(function(zip, index) {
        // var zip = new JSZip(file);
        var medFiles = zip.folder("word/media").files;

        for (var mfile in medFiles) {
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

    // this.updateRelation(files);
};

var updateMediaRelations = function(zip, count, _media) {

    var xmlString = zip.file("word/_rels/document.xml.rels").asText();
    var xml = new DOMParser().parseFromString(xmlString, 'text/xml');

    var childNodes = xml.getElementsByTagName('Relationships')[0].childNodes;
    var serializer = new XMLSerializer();

    for (var node in childNodes) {
        if (/^\d+$/.test(node) && childNodes[node].getAttribute) {
            var target = childNodes[node].getAttribute('Target');
            if ('word/' + target == _media[count].oldTarget) {

                _media[count].oldRelID = childNodes[node].getAttribute('Id');

                childNodes[node].setAttribute('Target', _media[count].newTarget);
                childNodes[node].setAttribute('Id', _media[count].oldRelID + '_' + count);
            }
        }
    }

    // console.log(serializer.serializeToString(xml.documentElement));

    var startIndex = xmlString.indexOf("<Relationships");
    xmlString = xmlString.replace(xmlString.slice(startIndex), serializer.serializeToString(xml.documentElement));

    zip.file("word/_rels/document.xml.rels", xmlString);

    // console.log( xmlString );
};

var updateMediaContent = function(zip, count, _media) {

    var xmlString = zip.file("word/document.xml").asText();

    xmlString = xmlString.replace(new RegExp(_media[count].oldRelID + '"', 'g'), _media[count].oldRelID + '_' + count + '"');

    zip.file("word/document.xml", xmlString);
};

var copyMediaFiles = function(base, _media, _files) {

    for (var media in _media) {
        var content = _files[_media[media].fileIndex].file(_media[media].oldTarget).asUint8Array();

        base.file('word/' + _media[media].newTarget, content);
    }
};

module.exports = {
    prepareMediaFiles: prepareMediaFiles,
    updateMediaRelations: updateMediaRelations,
    updateMediaContent: updateMediaContent,
    copyMediaFiles: copyMediaFiles
};
