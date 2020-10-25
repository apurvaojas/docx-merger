var XMLSerializer = require('xmldom').XMLSerializer;
var DOMParser = require('xmldom').DOMParser;

const prepareBodies = (files) => {
  const bodies = [];
  files.forEach((zip, index) => {
    var xmlStr = zip.file("word/document.xml").asText();
    xmlStr = xmlStr.substring(xmlStr.indexOf("<w:body>") + 8, xmlStr.lastIndexOf("</w:body>"));
    if (!index) { bodies.push(xmlStr); return; }
    xmlStr = xmlStr.replace(/<w:sectPr.*? <\/w:sectPr>/g,'');
    xmlStr = xmlStr.replace(/<w:bookmark(Start|End).*?"\/>/g, '');
    xmlStr = xmlStr.replace(/<w:commentRange(Start|End).*?"\/>/g, '');
    xmlStr = xmlStr.replace(/<w:commentReference.*?"\/>/g, '');
    xmlStr = xmlStr.replace(/<w:headerReference.*?"\/>/g, '');
    xmlStr = xmlStr.replace(/<w:footerReference.*?"\/>/g, '');
    // if (pageBreak && index < files.length - 1) xml += '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
    bodies.push(xmlStr);
  });
  return bodies;
}

const generateBody = (zip, bodies) => {
  let xml = zip.file("word/document.xml").asText();
  const lastPart = xml.substring(xml.lastIndexOf("<w:sectPr"));
  xml = xml.replace(lastPart, bodies.join('')) + lastPart;
  zip.file("word/document.xml", xml);
}

module.exports = {
  prepareBodies,
  generateBody
};