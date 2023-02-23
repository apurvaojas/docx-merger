const {XMLSerializer} = require('@xmldom/xmldom');
const {DOMParser} = require('@xmldom/xmldom');


const prepareNumbering = function(files) {
    const serializer = new XMLSerializer();

    files.forEach(async function(zip, index) {
        const xmlBin = zip.file('word/numbering.xml');
        if (!xmlBin) {
            return;
        }
        let xmlString = await xmlBin.async('string');
        const xml = new DOMParser().parseFromString(xmlString, 'text/xml');
        const nodes = xml.getElementsByTagName('w:abstractNum');

        for (const node in nodes) {
            if (/^\d+$/.test(node) && nodes[node].getAttribute) {
                const absID = nodes[node].getAttribute('w:abstractNumId');
                nodes[node].setAttribute('w:abstractNumId', absID + index);
                const pStyles = nodes[node].getElementsByTagName('w:pStyle');
                for (const pStyle in pStyles) {
                    if (pStyles[pStyle].getAttribute) {
                        const pStyleId = pStyles[pStyle].getAttribute('w:val');
                        pStyles[pStyle].setAttribute('w:val', pStyleId + '_' + index);
                    }
                }
                const numStyleLinks = nodes[node].getElementsByTagName('w:numStyleLink');
                for (const numstyleLink in numStyleLinks) {
                    if (numStyleLinks[numstyleLink].getAttribute) {
                        const styleLinkId = numStyleLinks[numstyleLink].getAttribute('w:val');
                        numStyleLinks[numstyleLink].setAttribute('w:val', styleLinkId + '_' + index);
                    }
                }

                const styleLinks = nodes[node].getElementsByTagName('w:styleLink');
                for (const styleLink in styleLinks) {
                    if (styleLinks[styleLink].getAttribute) {
                        const styleLinkId = styleLinks[styleLink].getAttribute('w:val');
                        styleLinks[styleLink].setAttribute('w:val', styleLinkId + '_' + index);
                    }
                }

            }
        }

        const numNodes = xml.getElementsByTagName('w:num');

        for (const node in numNodes) {
            if (/^\d+$/.test(node) && numNodes[node].getAttribute) {
                const ID = numNodes[node].getAttribute('w:numId');
                numNodes[node].setAttribute('w:numId', ID + index);
                const absrefID = numNodes[node].getElementsByTagName('w:abstractNumId');
                for (const i in absrefID) {
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
    });
};

const mergeNumbering = function(files, _numbering) {
    files.forEach(async function(zip) {
        const xmlBin = zip.file('word/numbering.xml');
        if (!xmlBin) {
          return;
        }
        let xmlString = await xmlBin.async('string');
        xmlString = xmlString.substring(xmlString.indexOf("<w:abstractNum "), xmlString.indexOf("</w:numbering"));
        _numbering.push(xmlString);
    });
};

const generateNumbering = async function(zip, _numbering) {
    const xmlBin = zip.file('word/numbering.xml');
    if (!xmlBin) {
      return;
    }
    let xmlString = await xmlBin.async('string');
    const startIndex = xmlString.indexOf("<w:abstractNum ");
    const endIndex = xmlString.indexOf("</w:numbering>");

    xmlString = xmlString.replace(xmlString.slice(startIndex, endIndex), _numbering.join(''));

    zip.file("word/numbering.xml", xmlString);
};


module.exports = {
    prepareNumbering: prepareNumbering,
    mergeNumbering: mergeNumbering,
    generateNumbering: generateNumbering
};