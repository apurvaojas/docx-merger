var JSZip = require('jszip');
var DOMParser = require('xmldom').DOMParser;
var XMLSerializer = require('xmldom').XMLSerializer;


function DocxMerger(options, files) {

    this._body = [];
    this._header = [];
    this._footer = [];
    this._Basestyle = options.style || 'source';
    this._style = [];
    this._pageBreak = options.pageBreak || true;
    this._files = [];
    var self = this;
    (files || []).forEach(function(file) {
        self._files.push(new JSZip(file));
    });
    this._contentTypes = {};

    this._media = {};
    this._rel = {};

    this._builder = this._body;

    this.insertPageBreak = function() {
        var pb = '<w:p> \
					<w:r> \
						<w:br w:type="page"/> \
					</w:r> \
				  </w:p>';

        this._builder.push(pb);
    };

    this.insertRaw = function(xml) {

        this._builder.push(xml);
    };

    this.mergeBody = function(files) {

        var self = this;
        this._builder = this._body;

        this.mergeContentTypes(files);
        this.prepareMediaFiles(files);
        this.mergeRelations(files);

        this.prepareStyles(files);
        this.mergeStyles(files);

        files.forEach(function(zip) {
            //var zip = new JSZip(file);
            var xml = zip.file("word/document.xml").asText();
            xml = xml.substring(xml.indexOf("<w:body>") + 8);
            xml = xml.substring(0, xml.indexOf("</w:body>"));
            xml = xml.substring(0, xml.lastIndexOf("<w:sectPr"));

            self.insertRaw(xml);
            if (self._pageBreak)
                self.insertPageBreak();
        });
    };

    this.mergeStyles = function(files) {

        // this._builder = this._style;

        // console.log("MERGE__STYLES");
        var self = this;

        files.forEach(function(zip) {

            var xml = zip.file("word/styles.xml").asText();

            xml = xml.substring(xml.indexOf("<w:style "), xml.indexOf("</w:styles"));

            self._style.push(xml);

        });
    };

    this.prepareStyles = function(files){
        var self = this;
        var style = this._styles;
        var serializer = new XMLSerializer();

        files.forEach(function(zip, index){
            var xmlString = zip.file("word/styles.xml").asText();
            var xml = new DOMParser().parseFromString(xmlString, 'text/xml');
            var nodes = xml.getElementsByTagName('w:style');

            for (var node in nodes) {
                if (/^\d+$/.test(node) && nodes[node].getAttribute) {
                    var styleId = nodes[node].getAttribute('w:styleId');
                    nodes[node].setAttribute('w:styleId', styleId+'_'+index);
                    var basedonStyle = nodes[node].getElementsByTagName('w:basedOn')[0];
                    if(basedonStyle){
                        var basedonStyleId = basedonStyle.getAttribute('w:val');
                        basedonStyle.setAttribute('w:val',basedonStyleId+'_'+index);
                    }

                    var w_next = nodes[node].getElementsByTagName('w:next')[0];
                    if(w_next){
                        var w_next_ID = w_next.getAttribute('w:val');
                        w_next.setAttribute('w:val',w_next_ID+'_'+index);
                    }

                    var w_link = nodes[node].getElementsByTagName('w:link')[0];
                    if(w_link){
                        var w_link_ID = w_link.getAttribute('w:val');
                        w_link.setAttribute('w:val', w_link_ID+'_'+index);
                    }
                    
                    self.updateStyleRel_Content(zip, index, styleId);
                }
            }

            var startIndex = xmlString.indexOf("<w:styles ");
            xmlString = xmlString.replace(xmlString.slice(startIndex), serializer.serializeToString(xml.documentElement));

            zip.file("word/styles.xml", xmlString);
            // console.log(nodes);
        });
    };


    this.updateStyleRel_Content = function(zip, fileIndex, styleId) {
        var self = this;

        var xmlString = zip.file("word/document.xml").asText();
        var xml = new DOMParser().parseFromString(xmlString, 'text/xml');

        xmlString = xmlString.replace(new RegExp('w:val="'+styleId+'"', 'g'), 'w:val="'+styleId+'_'+fileIndex+'"');

        // zip.file("word/document.xml", "");

        zip.file("word/document.xml", xmlString);
    };

    this.prepareMediaFiles = function(files) {
        var self = this;
        var media = this._media,
            count = 1;

        files.forEach(function(zip, index) {
            // var zip = new JSZip(file);
            var medFiles = zip.folder("word/media").files;

            for (var mfile in medFiles) {
                if (/^word\/media/.test(mfile) && mfile.length > 11) {
                     // console.log(mfile);
                    media[count] = {};
                    media[count].oldTarget = mfile;
                    media[count].newTarget = mfile.replace(/[0-9]/, '_'+count).replace('word/', "");
                    media[count].fileIndex = index;
                    self.updateMediaRelations(zip, count);
                    self.updateMediaContent(zip, count);
                    count++;
                }
            }
        });

        // console.log(JSON.stringify(media));

        // this.updateRelation(files);
    };

    this.updateMediaRelations = function(zip, count) {

        var self = this;

        var xmlString = zip.file("word/_rels/document.xml.rels").asText();
        var xml = new DOMParser().parseFromString(xmlString, 'text/xml');

        var childNodes = xml.getElementsByTagName('Relationships')[0].childNodes;
        var serializer = new XMLSerializer();

        for (var node in childNodes) {
            if (/^\d+$/.test(node) && childNodes[node].getAttribute) {
                var target = childNodes[node].getAttribute('Target');
                if ('word/' + target == self._media[count].oldTarget) {

                    self._media[count].oldRelID = childNodes[node].getAttribute('Id');

                    childNodes[node].setAttribute('Target', self._media[count].newTarget);
                    childNodes[node].setAttribute('Id', self._media[count].oldRelID+'_'+count);
                }
            }
        }

        // console.log(serializer.serializeToString(xml.documentElement));

        var startIndex = xmlString.indexOf("<Relationships");
        xmlString = xmlString.replace(xmlString.slice(startIndex), serializer.serializeToString(xml.documentElement));

        zip.file("word/_rels/document.xml.rels", xmlString);

        // console.log( xmlString );
    };

    this.updateMediaContent = function(zip, count) {
        var self = this;

        var xmlString = zip.file("word/document.xml").asText();
        var xml = new DOMParser().parseFromString(xmlString, 'text/xml');

        xmlString = xmlString.replace(new RegExp(this._media[count].oldRelID+'"', 'g'), this._media[count].oldRelID+'_'+count+'"');

        zip.file("word/document.xml", xmlString);
    };

    this.mergeContentTypes = function(files) {

        var self = this;

        files.forEach(function(zip) {
            // var zip = new JSZip(file);
            var xmlString = zip.file("[Content_Types].xml").asText();
            var xml = new DOMParser().parseFromString(xmlString, 'text/xml');

            var childNodes = xml.getElementsByTagName('Types')[0].childNodes;

            for (var node in childNodes) {
                if (/^\d+$/.test(node) && childNodes[node].getAttribute) {
                    var contentType = childNodes[node].getAttribute('ContentType');
                    if (!self._contentTypes[contentType])
                        self._contentTypes[contentType] = childNodes[node].cloneNode();
                }
            }

        });
    };

    this.mergeRelations = function(files) {

        var self = this;

        files.forEach(function(zip) {
            // var zip = new JSZip(file);
            var xmlString = zip.file("word/_rels/document.xml.rels").asText();
            var xml = new DOMParser().parseFromString(xmlString, 'text/xml');

            var childNodes = xml.getElementsByTagName('Relationships')[0].childNodes;

            for (var node in childNodes) {
                if (/^\d+$/.test(node) && childNodes[node].getAttribute) {
                    var Id = childNodes[node].getAttribute('Id');
                    if (!self._rel[Id])
                        self._rel[Id] = childNodes[node].cloneNode();
                }
            }

        });
    };

    this.save = function(type, callback) {

        var zip = this._files[0];

        var xml = zip.file("word/document.xml").asText();
        var startIndex = xml.indexOf("<w:body>") + 8;
        var endIndex = xml.lastIndexOf("<w:sectPr");

        xml = xml.replace(xml.slice(startIndex, endIndex), this._body.join(''));

        this.generateContentTypes(zip);
        this.copyMediaFiles(zip);
        this.generateRelations(zip);
        this.generateStyles(zip);

        zip.file("word/document.xml", xml);

        callback(zip.generate({ type: type }));
    };

    this.copyMediaFiles = function(base) {
        var self = this;
        // this._media.forEach(function (media) {
        for (var media in this._media) {
            var content = self._files[this._media[media].fileIndex].file(this._media[media].oldTarget).asUint8Array();

            base.file('word/' + this._media[media].newTarget, content);
        }
    };

    this.generateContentTypes = function(zip) {
        // body...
        var xmlString = zip.file("[Content_Types].xml").asText();
        var xml = new DOMParser().parseFromString(xmlString, 'text/xml');
        var serializer = new XMLSerializer();

        var types = xml.documentElement.cloneNode();

        for (var node in this._contentTypes) {
            types.appendChild(this._contentTypes[node]);
        }

        var startIndex = xmlString.indexOf("<Types");
        xmlString = xmlString.replace(xmlString.slice(startIndex), serializer.serializeToString(types));

        zip.file("[Content_Types].xml", xmlString);
    };

    this.generateRelations = function(zip) {
        // body...
        var xmlString = zip.file("word/_rels/document.xml.rels").asText();
        var xml = new DOMParser().parseFromString(xmlString, 'text/xml');
        var serializer = new XMLSerializer();

        var types = xml.documentElement.cloneNode();

        for (var node in this._rel) {
            types.appendChild(this._rel[node]);
        }

        var startIndex = xmlString.indexOf("<Relationships");
        xmlString = xmlString.replace(xmlString.slice(startIndex), serializer.serializeToString(types));

        zip.file("word/_rels/document.xml.rels", xmlString);
    };

    this.generateStyles = function (zip) {
        var xml = zip.file("word/styles.xml").asText();
        var startIndex = xml.indexOf("<w:style ");
        var endIndex = xml.indexOf("</w:styles>");

        // console.log(xml.substring(startIndex, endIndex))

        xml = xml.replace(xml.slice(startIndex, endIndex), this._style.join(''));

        // console.log(xml.substring(xml.indexOf("</w:docDefaults>")+16, xml.indexOf("</w:styles>")))
        // console.log(this._style.join(''))
        // console.log(xml)

        zip.file("word/styles.xml", xml);
    };

    if (this._files.length > 0) {

        this.mergeBody(this._files);
    }
}


module.exports = DocxMerger;


function Utf8ArrayToString(array) {
    var out, i, len, c;
    var char2, char3;

    out = "";
    len = array.length;
    i = 0;
    while (i < len) {
        c = array[i++];
        switch (c >> 4) {
            case 0:
            case 1:
            case 2:
            case 3:
            case 4:
            case 5:
            case 6:
            case 7:
                // 0xxxxxxx
                out += String.fromCharCode(c);
                break;
            case 12:
            case 13:
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
