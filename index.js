
var JSZip = require('jszip');

exports.DocxMerger = function(options, files) {

	this._body = [];
	this._header = [];
	this._footer = [];
	this._style = options.style || 'source';
	this._pageBreak = options.pageBreak || true;
	this._files = files || [];



	this._builder = this._body;

	this.insertPageBreak = function(){
		var pb = '<w:p> \
					<w:r> \
						<w:br w:type="page"/> \
					</w:r> \
				  </w:p>';

		this._builder.push(pb);
	};

	this.insertRaw = function(xml){

		this._builder.push(xml);
	};

	this.mergeBody = function(files){

		var self = this;

		files.forEach(function (file) {
					var zip = new JSZip(file);
					var xml = Utf8ArrayToString(zip.file("word/document.xml")._data.getContent());
					xml = xml.substring(xml.indexOf("<w:body>") + 8);
					xml = xml.substring(0, xml.indexOf("</w:body>"));
					xml = xml.substring(0, xml.indexOf("<w:sectPr"));
					self.insertRaw(xml);
					if(self._pageBreak)
						self.insertPageBreak();
				});

		var _xml = this._body.join('');
	};

	this.save = function( type, callback){

			var zip = new JSZip(this._files[0]);

			var xml = Utf8ArrayToString(zip.file("word/document.xml")._data.getContent());
			var startIndex = xml.indexOf("<w:body>") + 8;
			var endIndex = xml.indexOf("<w:sectPr");

			xml = xml.replace(xml.slice(startIndex,endIndex), this._body.join(''));

		zip.file("word/document.xml",xml);

		callback(zip.generate({type: type}))



		

	}

	if(this._files.length > 0){
		this.mergeBody(this._files);
	}
}



function Utf8ArrayToString(array) {
	var out, i, len, c;
	var char2, char3;

	out = "";
	len = array.length;
	i = 0;
	while(i < len) {
		c = array[i++];
		switch(c >> 4)
		{
			case 0: case 1: case 2: case 3: case 4: case 5: case 6: case 7:
			// 0xxxxxxx
			out += String.fromCharCode(c);
			break;
			case 12: case 13:
			// 110x xxxx   10xx xxxx
			char2 = array[i++];
			out += String.fromCharCode(((c & 0x1F) << 6) | (char2 & 0x3F));
			break;
			case 14:
				// 1110 xxxx  10xx xxxx  10xx xxxx
				char2 = array[i++];
				char3 = array[i++];
				out += String.fromCharCode(((c & 0x0F) << 12) |
					((char2 & 0x3F) << 6) |
					((char3 & 0x3F) << 0));
				break;
		}
	}

	return out;
}