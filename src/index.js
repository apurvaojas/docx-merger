var JSZip = require('jszip');
// var DOMParser = require('xmldom').DOMParser;
// var XMLSerializer = require('xmldom').XMLSerializer;

var { prepareStyles, generateStyles } = require('./merge-styles');
var { prepareMediaFiles } = require('./merge-media');
var { mergeContentTypes, generateContentTypes } = require('./merge-relations-and-content-type');
var { prepareNumberings, generateNumbering } = require('./merge-bullets-numberings');
const { prepareBodies, generateBody } = require('./merge-body');

function DocxMerger(options, filesIn) {
  let addNumbering = false;
  let numberingAdded = false;
  const files = filesIn.map(file => new JSZip(file))
  const bodies = prepareBodies(files);
  const styles = prepareStyles(files);
  const media = prepareMediaFiles(files)
  const numberings = prepareNumberings(files);

  const pageBreak = !!options.pageBreak;

  if (!numberings.numIds[0] && numberings.numIds.filter(i => !!i).length) addNumbering = true;

  numberings.numIds.forEach((i, index) => {
    if (!i) return;
    if (addNumbering && !numberingAdded) {
      numberingAdded = true;
      const numbering = files[index].file('word/numbering.xml').asUint8Array();
      files[0].file('word/numbering.xml', numbering);
    }

    Object.keys(i).map(oldVal => {
      bodies[index] = bodies[index].replaceAll(`w:numId w:val="${oldVal}"`, `w:numId w:val="${i[oldVal]}"`);
      styles[index] = styles[index].replaceAll(`w:numId w:val="${oldVal}"`, `w:numId w:val="${i[oldVal]}"`);
    })
  });


  Object.keys(media).forEach(key => {
    const { oldRelID, newRelID, fileIndex } = media[key];
    bodies[fileIndex] = bodies[fileIndex].replaceAll(`r:embed="${oldRelID}"`, `r:embed="${newRelID}"`);
  });


  this.save = function (type, callback) {
    const zip = files[0];
    generateContentTypes(zip, mergeContentTypes(files), addNumbering);
    generateBody(zip, bodies);
    generateStyles(zip, styles);
    generateNumbering(zip, numberings);
    callback(zip.generate({
      type: type,
      compression: "DEFLATE",
      compressionOptions: {
        level: 4
      }
    }));
  };
}


module.exports = DocxMerger;
