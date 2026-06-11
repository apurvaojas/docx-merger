
var XMLSerializer = require('@xmldom/xmldom').XMLSerializer;
var DOMParser = require('@xmldom/xmldom').DOMParser;


var prepareMediaFiles = function(files, media) {

       var count = 1;

    files.forEach(function(zip, index) {
        zip.names().forEach(function(mfile) {
            if (/^word\/media\//.test(mfile) && mfile.length > 11) {
                media[count] = {};
                media[count].oldTarget = mfile;
                media[count].newTarget = mfile.replace(/[0-9]/, '_' + count).replace('word/', "");
                media[count].fileIndex = index;
                updateMediaRelations(zip, count, media);
                updateMediaContent(zip, count, media);
                count++;
            }
        });
    });
};

var updateMediaRelations = function(zip, count, _media) {

    var xmlString = zip.getText("word/_rels/document.xml.rels");
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

    var startIndex = xmlString.indexOf("<Relationships");
    xmlString = xmlString.replace(xmlString.slice(startIndex), serializer.serializeToString(xml.documentElement));

    zip.setText("word/_rels/document.xml.rels", xmlString);
};

var updateMediaContent = function(zip, count, _media) {

    var xmlString = zip.getText("word/document.xml");

    xmlString = xmlString.replace(new RegExp(_media[count].oldRelID + '"', 'g'), _media[count].oldRelID + '_' + count + '"');

    zip.setText("word/document.xml", xmlString);
};

var copyMediaFiles = function(base, _media, _files) {

    for (var media in _media) {
        var content = _files[_media[media].fileIndex].getBytes(_media[media].oldTarget);

        base.setBytes('word/' + _media[media].newTarget, content);
    }
};

module.exports = {
    prepareMediaFiles: prepareMediaFiles,
    updateMediaRelations: updateMediaRelations,
    updateMediaContent: updateMediaContent,
    copyMediaFiles: copyMediaFiles
};
