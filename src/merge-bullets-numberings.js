const XMLSerializer = require('xmldom').XMLSerializer;
const DOMParser = require('xmldom').DOMParser;

export const prepareNumbering = (files) => {

    const serializer = new XMLSerializer();

    files.forEach(function (zip, index) {
        const xmlBin = zip.file('word/numbering.xml');
        if (!xmlBin) {
            return;
        }
        let xmlString = xmlBin.asText();
        const xml = new DOMParser().parseFromString(xmlString, 'text/xml');
        const nodes = xml.getElementsByTagName('w:abstractNum');

        for (let node in nodes) {
            if (/^\d+$/.test(node) && nodes[node].getAttribute) {
                const absID = nodes[node].getAttribute('w:abstractNumId');
                nodes[node].setAttribute('w:abstractNumId', absID + index);
                const pStyles = nodes[node].getElementsByTagName('w:pStyle');
                for (let pStyle in pStyles) {
                    if (pStyles[pStyle].getAttribute) {
                        let pStyleId = pStyles[pStyle].getAttribute('w:val');
                        pStyles[pStyle].setAttribute('w:val', pStyleId + '_' + index);
                    }
                }
                const numStyleLinks = nodes[node].getElementsByTagName('w:numStyleLink');
                for (let numstyleLink in numStyleLinks) {
                    if (numStyleLinks[numstyleLink].getAttribute) {
                        const styleLinkId = numStyleLinks[numstyleLink].getAttribute('w:val');
                        numStyleLinks[numstyleLink].setAttribute('w:val', styleLinkId + '_' + index);
                    }
                }

                const styleLinks = nodes[node].getElementsByTagName('w:styleLink');
                for (let styleLink in styleLinks) {
                    if (styleLinks[styleLink].getAttribute) {
                        const styleLinkId = styleLinks[styleLink].getAttribute('w:val');
                        styleLinks[styleLink].setAttribute('w:val', styleLinkId + '_' + index);
                    }
                }

            }
        }

        const numNodes = xml.getElementsByTagName('w:num');

        for (let node in numNodes) {
            if (/^\d+$/.test(node) && numNodes[node].getAttribute) {
                const ID = numNodes[node].getAttribute('w:numId');
                numNodes[node].setAttribute('w:numId', ID + index);
                const absrefID = numNodes[node].getElementsByTagName('w:abstractNumId');
                for (let i in absrefID) {
                    if (absrefID[i].getAttribute) {
                        const iId = absrefID[i].getAttribute('w:val');
                        absrefID[i].setAttribute('w:val', iId + index);
                    }
                }


            }
        }

        const startIndex = xmlString.indexOf("<w:numbering ");
        xmlString = xmlString.replace(xmlString.slice(startIndex), serializer.serializeToString(xml.documentElement));

        zip.file("word/numbering.xml", xmlString);
        // console.log(nodes);
    });
};

export const mergeNumbering = (files, _numbering) => {

    // _builder = _style;
    // console.log("MERGE__STYLES");

    files.forEach(function (zip) {
        const xmlBin = zip.file('word/numbering.xml');
        if (!xmlBin) {
            return;
        }
        let xml = xmlBin.asText();

        xml = xml.substring(xml.indexOf("<w:abstractNum "), xml.indexOf("</w:numbering"));

        _numbering.push(xml);

    });
};

export const generateNumbering = (zip, _numbering) => {
    const xmlBin = zip.file('word/numbering.xml');
    if (!xmlBin) {
        return;
    }
    let xml = xmlBin.asText();
    const startIndex = xml.indexOf("<w:abstractNum ");
    const endIndex = xml.indexOf("</w:numbering>");

    // console.log(xml.substring(startIndex, endIndex))

    xml = xml.replace(xml.slice(startIndex, endIndex), _numbering.join(''));

    // console.log(xml.substring(xml.indexOf("</w:docDefaults>")+16, xml.indexOf("</w:styles>")))
    // console.log(_style.join(''))
    // console.log(xml)

    zip.file("word/numbering.xml", xml);
};
