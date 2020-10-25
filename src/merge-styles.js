var XMLSerializer = require('xmldom').XMLSerializer;
var DOMParser = require('xmldom').DOMParser;

const prepareStyles = (files) => {
  const styles = [];
  const serializer = new XMLSerializer();
  const styleIdList = [];
  files.forEach((zip, index) => {
    var xmlString = zip.file("word/styles.xml").asText();
    var xml = new DOMParser().parseFromString(xmlString, 'text/xml');
    var nodes = xml.getElementsByTagName('w:style');
    const deletedNode = [];
    for (var node in nodes) {
      if (/^\d+$/.test(node) && nodes[node].getAttribute) {
        const styleId = nodes[node].attributes.getNamedItem('w:styleId').value;
        if (styleIdList.includes(styleId)) deletedNode.push(node);
        else styleIdList.push(nodes[node].attributes.getNamedItem('w:styleId').value);
      }
    }
    deletedNode.forEach(node => nodes[node].parentNode.removeChild(nodes[node]));
    let str = serializer.serializeToString(xml.documentElement);
    str = str.slice(str.indexOf("<w:style "), str.indexOf("</w:styles>"))
    styles.push(str);
  })
  // updateStyles(files[0], styles);
  return styles;
}

const generateStyles = (zip, styles) => {
  var xmlString = zip.file("word/styles.xml").asText();
  xmlString = xmlString.replace("</w:styles>", styles.join('') + "</w:styles>");
  zip.file("word/styles.xml", xmlString);
};

module.exports = {
  prepareStyles,
  generateStyles
};