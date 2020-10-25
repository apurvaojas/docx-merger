
var XMLSerializer = require('xmldom').XMLSerializer;
var DOMParser = require('xmldom').DOMParser;


var prepareMediaFiles = function (files) {
	const media = {};
	let count = 1;
	files.forEach(function (zip, index) {
		var medFiles = zip.folder("word/media/").files;
		for (var mfile in medFiles) {
			if (/^word\/media\/\S/.test(mfile)) {
				if (!index) { ++count; continue };
				const oldRelID = getOldRelID(zip, "word/_rels/document.xml.rels", mfile);
				if (!oldRelID) continue;
				media[count] = {};
				media[count].oldTarget = mfile;
				media[count].fileIndex = index;
				media[count].oldRelID = oldRelID;
				media[count].newTarget = `media/image${count}.${mfile.split('.')[mfile.split('.').length - 1]}`
				++count;
			}
		}
	});
	updateMediaRelations(files[0], media);
	copyMediaFiles(files, media);
	updateContent(files, media);
	// updateContentType(files, media);
	return media;
};

const getOldRelID = (zip, filePath, oldTarget) => {
	const xmlString = zip.file(filePath).asText();
	var xml = new DOMParser().parseFromString(xmlString, 'text/xml');
	var childNodes = xml.getElementsByTagName('Relationships')[0].childNodes;
	for (var node in childNodes) {
		if (/^\d+$/.test(node) && childNodes[node].getAttribute) {
			const target = childNodes[node].getAttribute('Target');
			if ('word/' + target == oldTarget) return childNodes[node].getAttribute('Id');
		}
	}
	return null;
}

// const updateContentType = (files, media) => {
// 	console.log(files, media);
// }

const updateMediaRelations = (zip, media) => {
	let xmlString = zip.file("word/_rels/document.xml.rels").asText();
	const xml = new DOMParser().parseFromString(xmlString, 'text/xml');
	const childNodes = xml.getElementsByTagName('Relationship');
	const arr = Object.keys(childNodes)
		.map(node => childNodes[node])
		.filter(item => item.getAttribute)
		.map(item => item.getAttribute('Id'))
		.filter(id => id.slice(0, 3) == 'rId' && +id.slice(3))
		.map(id => +id.slice(3))
	let id = Math.max(...arr);
	let addRelStr = ""
	Object.keys(media).map(key => {
		++id
		const { newTarget } = media[key];
		media[key].newRelID = `rId${id}`
		addRelStr += `<Relationship Id="rId${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="${newTarget}"/>`;
	})
	xmlString = xmlString.replace('</Relationships>', addRelStr + '</Relationships>');
	zip.file("word/_rels/document.xml.rels", xmlString);
}

const copyMediaFiles = (files, media) => {
	Object.keys(media).map(key => {
		var content = files[media[key].fileIndex].file(media[key].oldTarget).asUint8Array();
		files[0].file('word/' + media[key].newTarget, content);
	})
};
const updateContent = (files, media) => {
	// console.log(media);
	const fileIndexes = Object.keys(media).map(key => media[key].fileIndex);
	files.forEach((zip, index) => {
		if (!index || !fileIndexes.includes(index)) return;
		let xmlString = zip.file("word/document.xml").asText();
		Object.keys(media).forEach(key => {
			const { oldRelID, newRelID, fileIndex } = media[key];
			// console.log(fileIndex + ' === ' + index);
			if (fileIndex != index) return;
			// console.log('fileIndex =>', index, 'oldRelID => ', oldRelID, '==> newRelID => ', newRelID);
			// console.log(xmlString.indexOf(`r:embed="${oldRelID}"`));
			xmlString = xmlString.replace(`r:embed="${oldRelID}"`, `r:embed="${newRelID}"`);
		});
		zip.file("word/document.xml", xmlString);
	})
}



// var updateMediaRelations1 = function (zip, count, _media) {

// 	var xmlString = zip.file("word/_rels/document.xml.rels").asText();
// 	var xml = new DOMParser().parseFromString(xmlString, 'text/xml');

// 	var childNodes = xml.getElementsByTagName('Relationships')[0].childNodes;
// 	var serializer = new XMLSerializer();
// 	console.log(childNodes);
// 	for (var node in childNodes) {
// 		if (/^\d+$/.test(node) && childNodes[node].getAttribute) {
// 			var target = childNodes[node].getAttribute('Target');
// 			if ('word/' + target == _media[count].oldTarget) {

// 				_media[count].oldRelID = childNodes[node].getAttribute('Id');

// 				childNodes[node].setAttribute('Target', _media[count].newTarget);
// 				childNodes[node].setAttribute('Id', _media[count].oldRelID + '_' + count);
// 			}
// 		}
// 	}

// 	// console.log(serializer.serializeToString(xml.documentElement));

// 	var startIndex = xmlString.indexOf("<Relationships");
// 	xmlString = xmlString.replace(xmlString.slice(startIndex), serializer.serializeToString(xml.documentElement));

// 	zip.file("word/_rels/document.xml.rels", xmlString);

// 	// console.log( xmlString );
// };

var updateMediaContent = function (zip, count, _media) {

	var xmlString = zip.file("word/document.xml").asText();
	var xml = new DOMParser().parseFromString(xmlString, 'text/xml');

	xmlString = xmlString.replace(new RegExp(_media[count].oldRelID + '"', 'g'), _media[count].oldRelID + '_' + count + '"');

	zip.file("word/document.xml", xmlString);
};


module.exports = {
	prepareMediaFiles: prepareMediaFiles,
	// updateMediaRelations: updateMediaRelations,
	// updateMediaContent: updateMediaContent,
	copyMediaFiles: copyMediaFiles
};