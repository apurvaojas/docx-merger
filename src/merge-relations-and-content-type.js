
var XMLSerializer = require('xmldom').XMLSerializer;
var DOMParser = require('xmldom').DOMParser;


var mergeContentTypes = function (files) {
	const serializer = new XMLSerializer();
	const contentTypes = {};
	const contentTypesId = [];
	files.forEach((zip, index) => {
		var xmlString = zip.file("[Content_Types].xml").asText();
		var xml = new DOMParser().parseFromString(xmlString, 'text/xml');
		var childNodes = xml.getElementsByTagName('Types')[0].childNodes;
		for (var node in childNodes) {
			if (!/^\d+$/.test(node) || !childNodes[node].getAttribute) continue;
			const nodeEl = childNodes[node].getAttribute('Extension');
			if (!index) contentTypesId.push(nodeEl);
			if (nodeEl && index && !contentTypesId.includes(nodeEl) && !contentTypes[nodeEl]) {
				contentTypes[nodeEl] = childNodes[node].cloneNode()
				contentTypesId.push(nodeEl);
			};
		}
	});
	return contentTypes;
};

var mergeRelations = function (files, _rel) {

	files.forEach(function (zip) {
		// var zip = new JSZip(file);
		console.log(zip.file("word/_rels/document.xml.rels"));
		var xmlString = zip.file("word/_rels/document.xml.rels").asText();
		var xml = new DOMParser().parseFromString(xmlString, 'text/xml');

		var childNodes = xml.getElementsByTagName('Relationships')[0].childNodes;

		for (var node in childNodes) {
			if (/^\d+$/.test(node) && childNodes[node].getAttribute) {
				var Id = childNodes[node].getAttribute('Id');
				if (!_rel[Id])
					_rel[Id] = childNodes[node].cloneNode();
			}
		}

	});
};

var generateContentTypes = function (zip, contentTypes, addNumbering) {
	const serializer = new XMLSerializer();
	let xmlString = zip.file("[Content_Types].xml").asText();
	// console.log(addNumbering);
	if (addNumbering) { xmlString = xmlString.replace('</Types>', '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>') }
	const xml = new DOMParser().parseFromString(xmlString, 'text/xml');
	const types = xml.getElementsByTagName('Types')[0]
	for (var node in contentTypes) {
		types.appendChild(contentTypes[node]);
	}
	var startIndex = xmlString.indexOf("<Types");
	xmlString = xmlString.replace(xmlString.slice(startIndex), serializer.serializeToString(types));
	zip.file("[Content_Types].xml", xmlString);
};

var generateRelations = function (zip, _rel) {
	// body...
	var xmlString = zip.file("word/_rels/document.xml.rels").asText();
	var xml = new DOMParser().parseFromString(xmlString, 'text/xml');
	var serializer = new XMLSerializer();

	var types = xml.documentElement.cloneNode();

	for (var node in _rel) {
		types.appendChild(_rel[node]);
	}

	var startIndex = xmlString.indexOf("<Relationships");
	xmlString = xmlString.replace(xmlString.slice(startIndex), serializer.serializeToString(types));

	zip.file("word/_rels/document.xml.rels", xmlString);
};


module.exports = {
	mergeContentTypes: mergeContentTypes,
	mergeRelations: mergeRelations,
	generateContentTypes: generateContentTypes,
	generateRelations: generateRelations
};