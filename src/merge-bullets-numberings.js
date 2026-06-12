var XMLSerializer = require('@xmldom/xmldom').XMLSerializer;
var DOMParser = require('@xmldom/xmldom').DOMParser;
var xmlUtils = require('./xml-utils');


var prepareNumbering = function(files) {

    var serializer = new XMLSerializer();
    var nextAbstractId = 1;
    var nextNumId = 1;
    var maps = []; // per file index: { abs: {old: new}, num: {old: new} }

    files.forEach(function(zip, index) {
        var xmlString = zip.getText('word/numbering.xml');
        maps[index] = { abs: {}, num: {} };
        if (xmlString === null) {
            return;
        }
        var xml = new DOMParser().parseFromString(xmlString, 'text/xml');
        var absMap = maps[index].abs;
        var numMap = maps[index].num;

        var nodes = xml.getElementsByTagName('w:abstractNum');
        for (var i = 0; i < nodes.length; i++) {
            var absNode = nodes.item(i);
            var absID = absNode.getAttribute('w:abstractNumId');
            absMap[absID] = String(nextAbstractId++);
            absNode.setAttribute('w:abstractNumId', absMap[absID]);

            var pStyles = absNode.getElementsByTagName('w:pStyle');
            for (var p = 0; p < pStyles.length; p++) {
                var pStyleId = pStyles.item(p).getAttribute('w:val');
                pStyles.item(p).setAttribute('w:val', pStyleId + '_' + index);
            }
            var numStyleLinks = absNode.getElementsByTagName('w:numStyleLink');
            for (var n = 0; n < numStyleLinks.length; n++) {
                numStyleLinks.item(n).setAttribute('w:val', numStyleLinks.item(n).getAttribute('w:val') + '_' + index);
            }
            var styleLinks = absNode.getElementsByTagName('w:styleLink');
            for (var s = 0; s < styleLinks.length; s++) {
                styleLinks.item(s).setAttribute('w:val', styleLinks.item(s).getAttribute('w:val') + '_' + index);
            }
        }

        var numNodes = xml.getElementsByTagName('w:num');
        for (var j = 0; j < numNodes.length; j++) {
            var numNode = numNodes.item(j);
            var oldNumId = numNode.getAttribute('w:numId');
            numMap[oldNumId] = String(nextNumId++);
            numNode.setAttribute('w:numId', numMap[oldNumId]);

            var absRefs = numNode.getElementsByTagName('w:abstractNumId');
            for (var k = 0; k < absRefs.length; k++) {
                var oldRef = absRefs.item(k).getAttribute('w:val');
                if (absMap[oldRef]) absRefs.item(k).setAttribute('w:val', absMap[oldRef]);
            }
        }

        zip.setText("word/numbering.xml",
            xmlUtils.replaceFrom(xmlString, "<w:numbering ", serializer.serializeToString(xml.documentElement)));

        // rewrite the body references that v1 left dangling
        var docString = zip.getText('word/document.xml');
        docString = docString.replace(/(<w:numId w:val=")(\d+)(")/g, function (m, pre, id, post) {
            return pre + (numMap[id] || id) + post;
        });
        zip.setText('word/document.xml', docString);
    });

    return maps;
};

var mergeNumbering = function(files, _numbering) {

    files.forEach(function(zip) {
        var xml = zip.getText('word/numbering.xml');
        if (xml === null) {
          return;
        }

        xml = xml.substring(xml.indexOf("<w:abstractNum "), xml.indexOf("</w:numbering"));

        _numbering.push(xml);

    });
};

var generateNumbering = function(zip, _numbering) {
    var xml = zip.getText('word/numbering.xml');
    if (xml === null) {
      return;
    }
    var startIndex = xml.indexOf("<w:abstractNum ");
    var endIndex = xml.indexOf("</w:numbering>");

    xml = xmlUtils.replaceBetween(xml, startIndex, endIndex, _numbering.join(''));

    zip.setText("word/numbering.xml", xml);
};


module.exports = {
    prepareNumbering: prepareNumbering,
    mergeNumbering: mergeNumbering,
    generateNumbering: generateNumbering
};
