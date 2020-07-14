var XMLSerializer = require('xmldom').XMLSerializer;
var DOMParser = require('xmldom').DOMParser;

var prepareStyles = function(files, style) {
    // var self = this;
    // var style = this._styles;
    var serializer = new XMLSerializer();

    files.forEach(function(zip, index) {
        var xmlString = zip.file("word/styles.xml").asText();
        var xml = new DOMParser().parseFromString(xmlString, 'text/xml');
        var nodes = xml.getElementsByTagName('w:style');

        for (var node in nodes) {
            if (/^\d+$/.test(node) && nodes[node].getAttribute) {
                var styleId = nodes[node].getAttribute('w:styleId');
                nodes[node].setAttribute('w:styleId', styleId + '_' + index);
                var basedonStyle = nodes[node].getElementsByTagName('w:basedOn')[0];
                if (basedonStyle) {
                    var basedonStyleId = basedonStyle.getAttribute('w:val');
                    basedonStyle.setAttribute('w:val', basedonStyleId + '_' + index);
                }

                var w_next = nodes[node].getElementsByTagName('w:next')[0];
                if (w_next) {
                    var w_next_ID = w_next.getAttribute('w:val');
                    w_next.setAttribute('w:val', w_next_ID + '_' + index);
                }

                var w_link = nodes[node].getElementsByTagName('w:link')[0];
                if (w_link) {
                    var w_link_ID = w_link.getAttribute('w:val');
                    w_link.setAttribute('w:val', w_link_ID + '_' + index);
                }

                var numId = nodes[node].getElementsByTagName('w:numId')[0];
                if (numId) {
                    var numId_ID = numId.getAttribute('w:val');
                    numId.setAttribute('w:val', numId_ID + index);
                }

                updateStyleRel_Content(zip, index, styleId);
            }
        }

        var startIndex = xmlString.indexOf("<w:styles ");
        xmlString = xmlString.replace(xmlString.slice(startIndex), serializer.serializeToString(xml.documentElement));

        zip.file("word/styles.xml", xmlString);
        // console.log(nodes);
    });
};

var mergeStyles = function(files, _styles, mainStylesInFirstFile    ) {

    files.forEach(function(zip, index) {
        var file = mainStylesInFirstFile ? files[files.length - 1 - index] : zip;
        var xml = file.file("word/styles.xml").asText();

        xml = xml.substring(xml.indexOf("<w:style "), xml.indexOf("</w:styles"));

        _styles.push(xml);

    });
};

var updateStyleRel_Content = function(zip, fileIndex, styleId) {


    var xmlString = zip.file("word/document.xml").asText();
    var xml = new DOMParser().parseFromString(xmlString, 'text/xml');

    xmlString = xmlString.replace(new RegExp('w:val="' + styleId + '"', 'g'), 'w:val="' + styleId + '_' + fileIndex + '"');

    // zip.file("word/document.xml", "");

    zip.file("word/document.xml", xmlString);
};

var generateStyles = function(zip, _style) {
    var xml = zip.file("word/styles.xml").asText();
    var startIndex = xml.indexOf("<w:style ");
    var endIndex = xml.indexOf("</w:styles>");

    // console.log(xml.substring(startIndex, endIndex))

    xml = xml.replace(xml.slice(startIndex, endIndex), _style.join(''));

    // console.log(xml.substring(xml.indexOf("</w:docDefaults>")+16, xml.indexOf("</w:styles>")))
    // console.log(this._style.join(''))
    // console.log(xml)

    zip.file("word/styles.xml", xml);
};

module.exports = {
    mergeStyles: mergeStyles,
    prepareStyles: prepareStyles,
    updateStyleRel_Content: updateStyleRel_Content,
    generateStyles: generateStyles
};