var XMLSerializer = require('xmldom').XMLSerializer;
var DOMParser = require('xmldom').DOMParser;


var prepareNumbering = function (files) {

  var serializer = new XMLSerializer();
  let abstractNumId = 0;
  const numberingAttributes = []
  files.forEach(function (zip, index) {
    if (index > 1) return;
    var xmlBin = zip.file('word/numbering.xml');
    if (!xmlBin) {
      return;
    }
    var xmlString = xmlBin.asText();
    var xml = new DOMParser().parseFromString(xmlString, 'text/xml');
    var nodes = xml.getElementsByTagName('w:abstractNum');

    for (var node in nodes) {
      if (/^\d+$/.test(node) && nodes[node].getAttribute) {
        nodes[node].setAttribute('w:abstractNumId', ++abstractNumId);
        var pStyles = nodes[node].getElementsByTagName('w:pStyle');
        for (var pStyle in pStyles) {
          if (pStyle !== "_node" && pStyles[pStyle].getAttribute) {
            var pStyleId = pStyles[pStyle].getAttribute('w:val');
            pStyles[pStyle].setAttribute('w:val', pStyleId + '_' + index);
          }
        }
        var numStyleLinks = nodes[node].getElementsByTagName('w:numStyleLink');
        for (var numstyleLink in numStyleLinks) {
          if (numstyleLink !== "_node" && numStyleLinks[numstyleLink].getAttribute) {
            var styleLinkId = numStyleLinks[numstyleLink].getAttribute('w:val');
            numStyleLinks[numstyleLink].setAttribute('w:val', styleLinkId + '_' + index);
          }
        }
        var styleLinks = nodes[node].getElementsByTagName('w:styleLink');
        for (var styleLink in styleLinks) {
          if (styleLink !== "_node" && styleLinks[styleLink].getAttribute) {
            var styleLinkId = styleLinks[styleLink].getAttribute('w:val');
            styleLinks[styleLink].setAttribute('w:val', styleLinkId + '_' + index);
          }
        }

      }
    }

    var numNodes = xml.getElementsByTagName('w:num');
    for (var node in numNodes) {
      if (numNodes[node].attributes && numNodes[node].getAttribute) {
        var ID = numNodes[node].getAttribute('w:numId');
        numNodes[node].setAttribute('w:numId', ID + '_' + index);
        var absrefID = numNodes[node].getElementsByTagName('w:abstractNumId');
        absrefID['0'].setAttribute('w:val', abstractNumId);

      }
    }



    var startIndex = xmlString.indexOf("<w:numbering ");
    xmlString = xmlString.replace(xmlString.slice(startIndex), serializer.serializeToString(xml.documentElement));
    zip.file("word/numbering.xml", xmlString);
  });
};

var mergeNumbering = function (files, _numbering, _numberingAttributes) {
  let Ignorable = [];
  files.forEach(function (zip) {
    var xmlBin = zip.file('word/numbering.xml');
    if (!xmlBin) return;
    var xml = xmlBin.asText();
    var xmlNum = xml.substring(xml.indexOf("<w:abstractNum "), xml.indexOf("</w:numbering"));
    _numbering.push(xmlNum);

    const xmlAttr = xml.substring(xml.indexOf("xmlns:"), xml.indexOf(">", xml.indexOf("xmlns")));
    let start;
    let end;
    while (xmlAttr.indexOf('xmlns', start) != -1) {
      start = xmlAttr.indexOf('xmlns', start)
      end = xmlAttr.indexOf('"', xmlAttr.indexOf('"', start) + 1);
      const attr = xmlAttr.slice(start, end + 1);
      !_numberingAttributes.includes(attr) && _numberingAttributes.push(attr);
      start = end + 1;
    }
    start = xmlAttr.indexOf('mc:Ignorable')
    if (~start) {
      end = xmlAttr.indexOf('"', xmlAttr.indexOf('"', start) + 1);
      const b = xmlAttr.slice(xmlAttr.indexOf('mc:Ignorable', start) + 14, end);
      b.split(' ').forEach(i => { if (!Ignorable.includes(i)) Ignorable.push(i) })
    }
  });
  _numberingAttributes.push('mc:Ignorable="' + Ignorable.join(' ') + '"');
};

var generateNumbering = function (zip, _numbering, _numberingAttributes) {
  var xmlBin = zip.file('word/numbering.xml');
  if (!xmlBin) return;
  var xml = xmlBin.asText();
  var startIndex = xml.indexOf("<w:abstractNum ");
  var endIndex = xml.indexOf("</w:numbering>");
  xml = xml.replace(xml.slice(startIndex, endIndex), _numbering.join(''));
  startIndex = xml.indexOf("<w:numbering ") + 13;
  endIndex = xml.indexOf(">", startIndex);
  xml = xml.replace(xml.slice(startIndex, endIndex), _numberingAttributes.join(' '));
  zip.file("word/numbering.xml", xml);
};


module.exports = {
  prepareNumbering: prepareNumbering,
  mergeNumbering: mergeNumbering,
  generateNumbering: generateNumbering
};