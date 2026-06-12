var XMLSerializer = require('@xmldom/xmldom').XMLSerializer;
var DOMParser = require('@xmldom/xmldom').DOMParser;
var xmlUtils = require('./xml-utils');

var prepareStyles = function(files, style, numberingMaps) {
    var serializer = new XMLSerializer();

    files.forEach(function(zip, index) {
        var xmlString = zip.getText("word/styles.xml");
        var xml = new DOMParser().parseFromString(xmlString, 'text/xml');
        var nodes = xml.getElementsByTagName('w:style');
        var renamedIds = [];

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
                    var numMap = (numberingMaps && numberingMaps[index] && numberingMaps[index].num) || {};
                    numId.setAttribute('w:val', numMap[numId_ID] || numId_ID);
                }

                renamedIds.push(styleId);
            }
        }

        xmlString = xmlUtils.replaceFrom(xmlString, "<w:styles ", serializer.serializeToString(xml.documentElement));

        zip.setText("word/styles.xml", xmlString);

        // Rewrite every style reference in the document body in a single pass,
        // instead of re-reading and re-scanning document.xml once per style.
        if (renamedIds.length) {
            var docString = zip.getText("word/document.xml");
            var pattern = new RegExp('w:val="(' + renamedIds.map(xmlUtils.escapeRegExp).join('|') + ')"', 'g');
            docString = docString.replace(pattern, function(m, id) {
                return 'w:val="' + id + '_' + index + '"';
            });
            zip.setText("word/document.xml", docString);
        }
    });
};

var mergeStyles = function(files, _styles) {

    files.forEach(function(zip) {

        var xml = zip.getText("word/styles.xml");

        xml = xml.substring(xml.indexOf("<w:style "), xml.indexOf("</w:styles"));

        _styles.push(xml);

    });
};

var generateStyles = function(zip, _style) {
    var xml = zip.getText("word/styles.xml");
    var startIndex = xml.indexOf("<w:style ");
    var endIndex = xml.indexOf("</w:styles>");

    xml = xmlUtils.replaceBetween(xml, startIndex, endIndex, _style.join(''));

    zip.setText("word/styles.xml", xml);
};

module.exports = {
    mergeStyles: mergeStyles,
    prepareStyles: prepareStyles,
    generateStyles: generateStyles
};
