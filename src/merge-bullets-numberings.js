var XMLSerializer = require('xmldom').XMLSerializer;
var DOMParser = require('xmldom').DOMParser;

const prepareNumberings = (files) => {
  const serializer = new XMLSerializer();
  const numberings = {
    abstractNum: [],
    num: [],
    numIds: [],
    attributes: [],
    ignorable: []
  };
  let abstractNumId = 0;
  let numId = 0;
  files.forEach((zip, index) => {
    const xmlFile = zip.file('word/numbering.xml');
    if (!xmlFile) {
      numberings.abstractNum.push('');
      numberings.num.push('');
      numberings.numIds.push(null);
      return;
    }
    const abstractNumIds = {};
    const numIds = {};

    let xmlString = zip.file('word/numbering.xml').asText();
    getAttributes(xmlString, numberings.attributes, numberings.ignorable);
    const xml = new DOMParser().parseFromString(xmlString, 'text/xml');
    let nodes = xml.getElementsByTagName('w:abstractNum');
    for (var node in nodes) {
      if (/^\d+$/.test(node) && nodes[node].getAttribute) {
        const oldValue = nodes[node].getAttribute('w:abstractNumId');
        nodes[node].setAttribute('w:abstractNumId', ++abstractNumId);
        abstractNumIds[oldValue] = abstractNumId;;
      }
    }
    nodes = xml.getElementsByTagName('w:num');
    for (var node in nodes) {
      if (/^\d+$/.test(node) && nodes[node].getAttribute) {
        const abstractNumIdnode = nodes[node].getElementsByTagName('w:abstractNumId')[0];
        abstractNumIdnode.setAttribute('w:val', abstractNumIds[abstractNumIdnode.getAttribute('w:val')]);

        const oldValue = nodes[node].getAttribute('w:numId');
        nodes[node].setAttribute('w:numId', ++numId);
        numIds[oldValue] = numId;
      }
    }
    const result = serializer.serializeToString(xml.documentElement);
    numberings.abstractNum.push(result.slice(result.indexOf("<w:abstractNum "), result.indexOf("<w:num ")));
    numberings.num.push(result.slice(result.indexOf("<w:num "), result.indexOf("</w:numbering>")));
    numberings.numIds.push(numIds);
  })
  return numberings
}

const getAttributes = (xmlString, attributes, ignorable) => {
  const xmlAttr = xmlString.substring(xmlString.indexOf("xmlns:"), xmlString.indexOf(">", xmlString.indexOf("xmlns")));
  let start;
  let end;
  while (xmlAttr.indexOf('xmlns', start) != -1) {
    start = xmlAttr.indexOf('xmlns', start)
    end = xmlAttr.indexOf('"', xmlAttr.indexOf('"', start) + 1);
    const attr = xmlAttr.slice(start, end + 1);
    !attributes.includes(attr) && attributes.push(attr);
    start = end + 1;
  }
  start = xmlAttr.indexOf('mc:Ignorable')
  if (!~start) return;
  end = xmlAttr.indexOf('"', xmlAttr.indexOf('"', start) + 1);
  const b = xmlAttr.slice(xmlAttr.indexOf('mc:Ignorable', start) + 14, end);
  b.split(' ').forEach(i => { if (!ignorable.includes(i)) ignorable.push(i) })
}


// const updateFile = (zip, filePath, tagName, attributeName, valueObj) => {
//   const serializer = new XMLSerializer();
//   let xmlString = zip.file(filePath).asText();
//   const xml = new DOMParser().parseFromString(xmlString, 'text/xml');
//   let nodes = xml.getElementsByTagName(tagName);
//   for (var node in nodes) {
//     if (/^\d+$/.test(node) && nodes[node].getAttribute) {
//       const oldValue = nodes[node].getAttribute(attributeName);
//       nodes[node].setAttribute(attributeName, valueObj[oldValue]);
//     }
//   }
//   zip.file(filePath, serializer.serializeToString(xml));
// };


// var mergeNumbering = function (files, _numbering, _numberingAttributes) {
//   let Ignorable = [];
//   files.forEach(function (zip) {
//     var xmlBin = zip.file('word/numbering.xml');
//     if (!xmlBin) return;
//     var xml = xmlBin.asText();
//     const xmlNum = xml.substring(xml.indexOf("<w:abstractNum "), xml.indexOf("</w:numbering"));
//     _numbering.push(xmlNum);

//     const xmlAttr = xml.substring(xml.indexOf("xmlns:"), xml.indexOf(">", xml.indexOf("xmlns")));
//     let start;
//     let end;
//     while (xmlAttr.indexOf('xmlns', start) != -1) {
//       start = xmlAttr.indexOf('xmlns', start)
//       end = xmlAttr.indexOf('"', xmlAttr.indexOf('"', start) + 1);
//       const attr = xmlAttr.slice(start, end + 1);
//       !_numberingAttributes.includes(attr) && _numberingAttributes.push(attr);
//       start = end + 1;
//     }
//     start = xmlAttr.indexOf('mc:Ignorable')
//     if (~start) {
//       end = xmlAttr.indexOf('"', xmlAttr.indexOf('"', start) + 1);
//       const b = xmlAttr.slice(xmlAttr.indexOf('mc:Ignorable', start) + 14, end);
//       b.split(' ').forEach(i => { if (!Ignorable.includes(i)) Ignorable.push(i) })
//     }
//   });
//   _numberingAttributes.push('mc:Ignorable="' + Ignorable.join(' ') + '"');
// };

var generateNumbering = function (zip, numberings) {
  const xmlBin = zip.file('word/numbering.xml');

  if (!xmlBin) return;
  // console.log(numberings);
  // numberings.num.map(i => console.log(i))
  let xmlString = xmlBin.asText();
  xmlString = xmlString.slice(0, xmlString.indexOf('<w:numbering '));
  xmlString = xmlString
    + '<w:numbering ' + numberings.attributes.join(' ')
    + ' mc:Ignorable="' + numberings.ignorable.join(' ') + '">'
    + numberings.abstractNum.join('')
    + numberings.num.join('')
    + '</w:numbering>';
  // console.log(xmlString);



  // var startIndex = xml.indexOf("<w:abstractNum ");
  // var endIndex = xml.indexOf("</w:numbering>");
  // xml = xml.replace(xml.slice(startIndex, endIndex), _numbering.join(''));
  // startIndex = xml.indexOf("<w:numbering ") + 13;
  // endIndex = xml.indexOf(">", startIndex);
  // xml = xml.replace(xml.slice(startIndex, endIndex), _numberingAttributes.join(' '));
  // zip.file("word/numbering.xml", xml);
};


module.exports = {
  prepareNumberings,
  generateNumbering,
  // mergeNumbering: mergeNumbering,
  // generateNumbering: generateNumbering
};