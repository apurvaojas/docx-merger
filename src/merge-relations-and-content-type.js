
var XMLSerializer = require('@xmldom/xmldom').XMLSerializer;
var DOMParser = require('@xmldom/xmldom').DOMParser;


var mergeContentTypes = function(files, _contentTypes) {


    files.forEach(function(zip) {
        // var zip = new JSZip(file);
        var xmlString = zip.file("[Content_Types].xml").asText();
        var xml = new DOMParser().parseFromString(xmlString, 'text/xml');

        var childNodes = xml.getElementsByTagName('Types')[0].childNodes;

        for (var node in childNodes) {
            if (/^\d+$/.test(node) && childNodes[node].getAttribute) {
                var contentType = childNodes[node].getAttribute('ContentType');
                if (!_contentTypes[contentType])
                    _contentTypes[contentType] = childNodes[node].cloneNode();
            }
        }

    });
};

var mergeRelations = function(files, _rel) {

    files.forEach(function(zip) {
        // var zip = new JSZip(file);
        var xmlString = zip.file("word/_rels/document.xml.rels").asText();
        var xml = new DOMParser().parseFromString(xmlString, 'text/xml');

        var childNodes = xml.getElementsByTagName('Relationships')[0].childNodes;

        for (var node in childNodes) {
            if (/^\d+$/.test(node) && childNodes[node].getAttribute) {
                var Id = childNodes[node].getAttribute('Id');
                if (!_rel[Id])
                    _rel[Id] = childNodes[node].cloneNode();
            }
        }

    });
};

var generateContentTypes = function(zip, _contentTypes) {
    // body...
    var xmlString = zip.file("[Content_Types].xml").asText();
    var xml = new DOMParser().parseFromString(xmlString, 'text/xml');
    var serializer = new XMLSerializer();

    var types = xml.documentElement.cloneNode();

    for (var node in _contentTypes) {
        types.appendChild(_contentTypes[node]);
    }

    var startIndex = xmlString.indexOf("<Types");
    xmlString = xmlString.replace(xmlString.slice(startIndex), serializer.serializeToString(types));

    zip.file("[Content_Types].xml", xmlString);
};

var generateRelations = function(zip, _rel) {
    // body...
    var xmlString = zip.file("word/_rels/document.xml.rels").asText();
    var xml = new DOMParser().parseFromString(xmlString, 'text/xml');
    var serializer = new XMLSerializer();

    var types = xml.documentElement.cloneNode();

    for (var node in _rel) {
        types.appendChild(_rel[node]);
    }

    var startIndex = xmlString.indexOf("<Relationships");
    xmlString = xmlString.replace(xmlString.slice(startIndex), serializer.serializeToString(types));

    zip.file("word/_rels/document.xml.rels", xmlString);
};


module.exports = {
    mergeContentTypes: mergeContentTypes,
    mergeRelations: mergeRelations,
    generateContentTypes: generateContentTypes,
    generateRelations: generateRelations
};