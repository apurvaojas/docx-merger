const {XMLSerializer} = require('@xmldom/xmldom');
const {DOMParser} = require('@xmldom/xmldom');

const prepareStyles = function(files, style) {
    const serializer = new XMLSerializer();

    files.forEach(async function(zip, index) {
        let xmlString = await zip.file("word/styles.xml").async('string');
        let xml = new DOMParser().parseFromString(xmlString, 'text/xml');
        const nodes = xml.getElementsByTagName('w:style');

        for (const node in nodes) {
            if (/^\d+$/.test(node) && nodes[node].getAttribute) {
                const styleId = nodes[node].getAttribute('w:styleId');
                nodes[node].setAttribute('w:styleId', styleId + '_' + index);
                const basedonStyle = nodes[node].getElementsByTagName('w:basedOn')[0];
                if (basedonStyle) {
                    const basedonStyleId = basedonStyle.getAttribute('w:val');
                    basedonStyle.setAttribute('w:val', basedonStyleId + '_' + index);
                }

                const w_next = nodes[node].getElementsByTagName('w:next')[0];
                if (w_next) {
                    const w_next_ID = w_next.getAttribute('w:val');
                    w_next.setAttribute('w:val', w_next_ID + '_' + index);
                }

                const w_link = nodes[node].getElementsByTagName('w:link')[0];
                if (w_link) {
                    const w_link_ID = w_link.getAttribute('w:val');
                    w_link.setAttribute('w:val', w_link_ID + '_' + index);
                }

                const numId = nodes[node].getElementsByTagName('w:numId')[0];
                if (numId) {
                    const numId_ID = numId.getAttribute('w:val');
                    numId.setAttribute('w:val', numId_ID + index);
                }

                await updateStyleRel_Content(zip, index, styleId);
            }
        }

        const startIndex = xmlString.indexOf("<w:styles ");
        xmlString = xmlString.replace(xmlString.slice(startIndex), serializer.serializeToString(xml.documentElement));

        zip.file("word/styles.xml", xmlString);
    });
};

const mergeStyles = function(files, _styles) {
    files.forEach(async function(zip) {
        let xmlString = await zip.file("word/styles.xml").async('string');
        xmlString = xmlString.substring(xmlString.indexOf("<w:style "), xmlString.indexOf("</w:styles"));
        _styles.push(xmlString);
    });
};

const updateStyleRel_Content = async function(zip, fileIndex, styleId) {
    let xmlString = await zip.file("word/document.xml").async('string');
    xmlString = xmlString.replace(new RegExp('w:val="' + styleId + '"', 'g'), 'w:val="' + styleId + '_' + fileIndex + '"');
    zip.file("word/document.xml", xmlString);
};

const generateStyles = async function(zip, _style) {
    let xmlString = await zip.file("word/styles.xml").async('string');
    const startIndex = xmlString.indexOf("<w:style ");
    const endIndex = xmlString.indexOf("</w:styles>");

    xmlString = xmlString.replace(xmlString.slice(startIndex, endIndex), _style.join(''));

    zip.file("word/styles.xml", xmlString);
};

module.exports = {
    mergeStyles: mergeStyles,
    prepareStyles: prepareStyles,
    updateStyleRel_Content: updateStyleRel_Content,
    generateStyles: generateStyles
};