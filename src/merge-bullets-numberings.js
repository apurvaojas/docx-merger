var XMLSerializer = require('@xmldom/xmldom').XMLSerializer;
var DOMParser = require('@xmldom/xmldom').DOMParser;


var prepareNumbering = function(files) {

    var serializer = new XMLSerializer();

    files.forEach(function(zip, index) {
        var xmlBin = zip.file('word/numbering.xml');
        if (!xmlBin) {
            return;
        }
        var xmlString = xmlBin.asText();
        var xml = new DOMParser().parseFromString(xmlString, 'text/xml');
        var nodes = xml.getElementsByTagName('w:abstractNum');

        for (var node in nodes) {
            if (/^\d+$/.test(node) && nodes[node].getAttribute) {
                var absID = nodes[node].getAttribute('w:abstractNumId');
                nodes[node].setAttribute('w:abstractNumId', absID + index);
                var pStyles = nodes[node].getElementsByTagName('w:pStyle');
                for (var pStyle in pStyles) {
                    if (pStyles[pStyle].getAttribute) {
                        var pStyleId = pStyles[pStyle].getAttribute('w:val');
                        pStyles[pStyle].setAttribute('w:val', pStyleId + '_' + index);
                    }
                }
                var numStyleLinks = nodes[node].getElementsByTagName('w:numStyleLink');
                for (var numstyleLink in numStyleLinks) {
                    if (numStyleLinks[numstyleLink].getAttribute) {
                        var styleLinkId = numStyleLinks[numstyleLink].getAttribute('w:val');
                        numStyleLinks[numstyleLink].setAttribute('w:val', styleLinkId + '_' + index);
                    }
                }

                var styleLinks = nodes[node].getElementsByTagName('w:styleLink');
                for (var styleLink in styleLinks) {
                    if (styleLinks[styleLink].getAttribute) {
                        var styleLinkId = styleLinks[styleLink].getAttribute('w:val');
                        styleLinks[styleLink].setAttribute('w:val', styleLinkId + '_' + index);
                    }
                }

            }
        }

        var numNodes = xml.getElementsByTagName('w:num');

        for (var node in numNodes) {
            if (/^\d+$/.test(node) && numNodes[node].getAttribute) {
                var ID = numNodes[node].getAttribute('w:numId');
                numNodes[node].setAttribute('w:numId', ID + index);
                var absrefID = numNodes[node].getElementsByTagName('w:abstractNumId');
                for (var i in absrefID) {
                    if (absrefID[i].getAttribute) {
                        var iId = absrefID[i].getAttribute('w:val');
                        absrefID[i].setAttribute('w:val', iId + index);
                    }
                }


            }
        }



        var startIndex = xmlString.indexOf("<w:numbering ");
        xmlString = xmlString.replace(xmlString.slice(startIndex), serializer.serializeToString(xml.documentElement));

        zip.file("word/numbering.xml", xmlString);
        // console.log(nodes);
    });
};

var mergeNumbering = function(files, _numbering) {

    // this._builder = this._style;

    // console.log("MERGE__STYLES");


    files.forEach(function(zip) {
        var xmlBin = zip.file('word/numbering.xml');
        if (!xmlBin) {
          return;
        }
        var xml = xmlBin.asText();

        xml = xml.substring(xml.indexOf("<w:abstractNum "), xml.indexOf("</w:numbering"));

        _numbering.push(xml);

    });
};

var generateNumbering = function(zip, _numbering) {
    var xmlBin = zip.file('word/numbering.xml');
    if (!xmlBin) {
      return;
    }
    var xml = xmlBin.asText();
    var startIndex = xml.indexOf("<w:abstractNum ");
    var endIndex = xml.indexOf("</w:numbering>");

    // console.log(xml.substring(startIndex, endIndex))

    xml = xml.replace(xml.slice(startIndex, endIndex), _numbering.join(''));

    // console.log(xml.substring(xml.indexOf("</w:docDefaults>")+16, xml.indexOf("</w:styles>")))
    // console.log(this._style.join(''))
    // console.log(xml)

    zip.file("word/numbering.xml", xml);
};


module.exports = {
    prepareNumbering: prepareNumbering,
    mergeNumbering: mergeNumbering,
    generateNumbering: generateNumbering
};