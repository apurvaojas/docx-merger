const XMLSerializer = require('xmldom').XMLSerializer;
const DOMParser = require('xmldom').DOMParser;

export const prepareStyles = (files, style) => {
    // let style = _styles;
    const serializer = new XMLSerializer();

    files.forEach(function (zip, index) {
        let xmlString = zip.file("word/styles.xml").asText();
        const xml = new DOMParser().parseFromString(xmlString, 'text/xml');
        const nodes = xml.getElementsByTagName('w:style');

        for (let node in nodes) {
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

                updateStyleRel_Content(zip, index, styleId);
            }
        }

        const startIndex = xmlString.indexOf("<w:styles ");
        xmlString = xmlString.replace(xmlString.slice(startIndex), serializer.serializeToString(xml.documentElement));

        zip.file("word/styles.xml", xmlString);
        // console.log(nodes);
    });
};

export const mergeStyles = (files, _styles) => {

    files.forEach(function (zip) {
        let xml = zip.file("word/styles.xml").asText();
        xml = xml.substring(xml.indexOf("<w:style "), xml.indexOf("</w:styles"));
        _styles.push(xml);
    });
};

export const updateStyleRel_Content = (zip, fileIndex, styleId) => {
    let xmlString = zip.file("word/document.xml").asText();
    //const xml = new DOMParser().parseFromString(xmlString, 'text/xml');

    xmlString = xmlString.replace(new RegExp('w:val="' + styleId + '"', 'g'), 'w:val="' + styleId + '_' + fileIndex + '"');

    // zip.file("word/document.xml", "");
    zip.file("word/document.xml", xmlString);
};

export const generateStyles = (zip, _style) => {
    let xml = zip.file("word/styles.xml").asText();
    const startIndex = xml.indexOf("<w:style ");
    const endIndex = xml.indexOf("</w:styles>");

    // console.log(xml.substring(startIndex, endIndex))

    xml = xml.replace(xml.slice(startIndex, endIndex), _style.join(''));

    // console.log(xml.substring(xml.indexOf("</w:docDefaults>")+16, xml.indexOf("</w:styles>")))
    // console.log(_style.join(''))
    // console.log(xml)

    zip.file("word/styles.xml", xml);
};
