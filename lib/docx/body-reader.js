exports.createBodyReader = createBodyReader;
exports._readNumberingProperties = readNumberingProperties;

var _ = require("underscore");

var documents = require("../documents");
var Result = require("../results").Result;
var warning = require("../results").warning;
var uris = require("./uris");

var preserverUtils = require('../styles/preserver-utils');

function createBodyReader(options) {
    return {
        readXmlElement: function (element) {
            return new BodyReader(options).readXmlElement(element);
        },
        readXmlElements: function (elements) {
            return new BodyReader(options).readXmlElements(elements);
        }
    };
}

function BodyReader(options) {
    var complexFieldStack = [];
    var currentInstrText = [];
    var relationships = options.relationships;
    var contentTypes = options.contentTypes;
    var docxFile = options.docxFile;
    var files = options.files;
    var numbering = options.numbering;
    var styles = options.styles;

    function readXmlElements(elements) {
        var results = elements.map(readXmlElement);
        return combineResults(results);
    }

    function readXmlElement(element) {
        if (element.type === "element") {
            var handler = xmlElementReaders[element.name];
            if (handler) {
                return handler(element);
            } else if (!Object.prototype.hasOwnProperty.call(ignoreElements, element.name)) {
                var message = warning("An unrecognised element was ignored: " + element.name);
                return emptyResultWithMessages([message]);
            }
        }
        return emptyResult();
    }

    function readParagraphIndent(element) {
        return {
            start: element.attributes["w:start"] || element.attributes["w:left"],
            end: element.attributes["w:end"] || element.attributes["w:right"],
            firstLine: element.attributes["w:firstLine"],
            hanging: element.attributes["w:hanging"]
        };
    }

    function readRunProperties(element) {
        return readRunStyle(element).map(function (style) {
            return {
                type: "runProperties",
                styleId: style.styleId,
                styleName: style.name,
                verticalAlignment: element.firstOrEmpty("w:vertAlign").attributes["w:val"],
                font: element.firstOrEmpty("w:rFonts").attributes["w:ascii"],
                isBold: readBooleanElement(element.first("w:b")),
                // isUnderline: readBooleanElement(element.first("w:u")),
                isUnderline: readBooleanElement(element.first("w:u")) && element.firstOrEmpty("w:u").attributes["w:val"] !== 'none',
                //
                isItalic: readBooleanElement(element.first("w:i")),
                isStrikethrough: readBooleanElement(element.first("w:strike")),
                isSmallCaps: readBooleanElement(element.first("w:smallCaps")),
                // 新增
                color: element.firstOrEmpty("w:color").attributes["w:val"],
                fontSize: element.firstOrEmpty('w:sz').attributes['w:val'],
            };
        });
    }

    function readBooleanElement(element) {
        if (element) {
            var value = element.attributes["w:val"];
            return value !== "false" && value !== "0";
        } else {
            return false;
        }
    }

    function readParagraphStyle(element) {
        return readStyle(element, "w:pStyle", "Paragraph", styles.findParagraphStyleById);
    }

    function readRunStyle(element) {
        return readStyle(element, "w:rStyle", "Run", styles.findCharacterStyleById);
    }

    function readTableStyle(element) {
        return readStyle(element, "w:tblStyle", "Table", styles.findTableStyleById);
    }

    function readStyle(element, styleTagName, styleType, findStyleById) {
        var messages = [];
        var styleElement = element.first(styleTagName);
        var styleId = null;
        var name = null;
        if (styleElement) {
            styleId = styleElement.attributes["w:val"];
            if (styleId) {
                var style = findStyleById(styleId);
                if (style) {
                    name = style.name;
                } else {
                    messages.push(undefinedStyleWarning(styleType, styleId));
                }
            }
        }
        return elementResultWithMessages({
            styleId: styleId,
            name: name
        }, messages);
    }

    var unknownComplexField = {
        type: "unknown"
    };

    function readFldChar(element) {
        var type = element.attributes["w:fldCharType"];
        if (type === "begin") {
            complexFieldStack.push(unknownComplexField);
            currentInstrText = [];
        } else if (type === "end") {
            complexFieldStack.pop();
        } else if (type === "separate") {
            var href = parseHyperlinkFieldCode(currentInstrText.join(''));
            var complexField = href === null ? unknownComplexField : {
                type: "hyperlink",
                href: href
            };
            complexFieldStack.pop();
            complexFieldStack.push(complexField);
        }
        return emptyResult();
    }

    function currentHyperlinkHref() {
        var topHyperlink = _.last(complexFieldStack.filter(function (complexField) {
            return complexField.type === "hyperlink";
        }));
        return topHyperlink ? topHyperlink.href : null;
    }

    function parseHyperlinkFieldCode(code) {
        var result = /\s*HYPERLINK "(.*)"/.exec(code);
        if (result) {
            return result[1];
        } else {
            return null;
        }
    }

    function readInstrText(element) {
        currentInstrText.push(element.text());
        return emptyResult();
    }

    function noteReferenceReader(noteType) {
        return function (element) {
            var noteId = element.attributes["w:id"];
            return elementResult(new documents.NoteReference({
                noteType: noteType,
                noteId: noteId
            }));
        };
    }

    function readCommentReference(element) {
        return elementResult(documents.commentReference({
            commentId: element.attributes["w:id"]
        }));
    }

    function readChildElements(element) {
        return readXmlElements(element.children);
    }

    var xmlElementReaders = {
        "w:p": function (element) { //段落
            return readXmlElements(element.children)
                .map(function (children) {

                    var properties = _.find(children, isParagraphProperties);
                    return new documents.Paragraph(
                        children.filter(negate(isParagraphProperties)),
                        properties
                    );
                })
                .insertExtra();
        },
        "w:pPr": function (element) {
            return readParagraphStyle(element).map(function (style) {
                let sectPr = element.first('w:sectPr');
                let pageInfo;
                if (sectPr) { //页面信息
                    let size = sectPr.firstOrEmpty("w:pgSz").attributes;
                    let width = (+size["w:w"]) / 1440 * 2.54;
                    let height = (+size["w:h"]) / 1440 * 2.54;
                    pageInfo = {
                        width,
                        height
                    }
                }
                return {
                    type: "paragraphProperties",
                    styleId: style.styleId,
                    styleName: style.name,
                    alignment: element.firstOrEmpty("w:jc").attributes["w:val"],
                    numbering: readNumberingProperties(element.firstOrEmpty("w:numPr"), numbering),
                    indent: readParagraphIndent(element.firstOrEmpty("w:ind")),
                    pageInfo: pageInfo
                };
            });
        },
        // 获取默认的页面信息，全局的页面信息
        "w:sectPr": function (element) {
            let size = element.firstOrEmpty("w:pgSz").attributes;
            let width = (+size["w:w"]) / 1440 * 2.54;
            let height = (+size["w:h"]) / 1440 * 2.54;
            let node = {
                type: 'page',
                options: {
                    width,
                    height
                }
            }
            return elementResult(node);
        },
        "w:r": function (element) {
            return readXmlElements(element.children)
                .map(function (children) {
                    var properties = _.find(children, isRunProperties);
                    children = children.filter(negate(isRunProperties));

                    var hyperlinkHref = currentHyperlinkHref();
                    if (hyperlinkHref !== null) {
                        children = [new documents.Hyperlink(children, {
                            href: hyperlinkHref
                        })];
                    }

                    return new documents.Run(children, properties);
                });
        },
        "w:rPr": readRunProperties,
        "w:fldChar": readFldChar,
        "w:instrText": readInstrText,
        "w:t": function (element) {
            return elementResult(new documents.Text(element.text()));
        },
        "w:tab": function (element) {
            return elementResult(new documents.Tab());
        },
        "w:noBreakHyphen": function () {
            return elementResult(new documents.Text("\u2011"));
        },
        "w:hyperlink": function (element) {
            var relationshipId = element.attributes["r:id"];
            var anchor = element.attributes["w:anchor"];
            return readXmlElements(element.children).map(function (children) {
                function create(options) {
                    var targetFrame = element.attributes["w:tgtFrame"] || null;

                    return new documents.Hyperlink(
                        children,
                        _.extend({
                            targetFrame: targetFrame
                        }, options)
                    );
                }

                if (relationshipId) {
                    var href = relationships.findTargetByRelationshipId(relationshipId);
                    if (anchor) {
                        href = uris.replaceFragment(href, anchor);
                    }
                    return create({
                        href: href
                    });
                } else if (anchor) {
                    return create({
                        anchor: anchor
                    });
                } else {
                    return children;
                }
            });
        },
        "w:tbl": readTable,
        "w:tr": readTableRow,
        "w:tc": readTableCell,
        "w:footnoteReference": noteReferenceReader("footnote"),
        "w:endnoteReference": noteReferenceReader("endnote"),
        "w:commentReference": readCommentReference,
        "w:br": function (element) {
            var breakType = element.attributes["w:type"];
            if (breakType == null || breakType === "textWrapping") {
                return elementResult(documents.lineBreak);
            } else if (breakType === "page") {
                return elementResult(documents.pageBreak);
            } else if (breakType === "column") {
                return elementResult(documents.columnBreak);
            } else {
                return emptyResultWithMessages([warning("Unsupported break type: " + breakType)]);
            }
        },
        "w:bookmarkStart": function (element) {
            var name = element.attributes["w:name"];
            if (name === "_GoBack") {
                return emptyResult();
            } else {
                return elementResult(new documents.BookmarkStart({
                    name: name
                }));
            }
        },

        "mc:AlternateContent": function (element) {
            return readChildElements(element.first("mc:Fallback"));
        },

        "w:sdt": function (element) {
            return readXmlElements(element.firstOrEmpty("w:sdtContent").children);
        },

        "w:ins": readChildElements,
        "w:object": readChildElements,
        "w:smartTag": readChildElements,
        "w:drawing": readChildElements, // readChildElements,
        "w:pict": function (element) {
            return readChildElements(element).toExtra();
        },
        //
        "v:roundrect": readChildElements,
        "v:shape": readChildElements,
        "v:textbox": readChildElements,
        "w:txbxContent": readChildElements,
        "wp:inline": readDrawingElement,
        "wp:anchor": readDrawingElement,
        "v:imagedata": readImageData,
        "v:group": readChildElements,
        "v:rect": readChildElements,
        // 分隔线
        "v:line": function (element) { //暂定，转成html-docx-js支持的table来实现线段的绘制，目前只能支持水平线
            return elementResult(new documents.Table([], {
                preservedStyles: {
                    cssText: 'border-collapse:unset;border-top:2pt solid #ff0000',
                }
            }))
        },
    };

    return {
        readXmlElement: readXmlElement,
        readXmlElements: readXmlElements
    };


    function readTable(element) {
        // var propertiesResult = readTableProperties(element.firstOrEmpty("w:tblPr"));
        var tableProperties = element.firstOrEmpty("w:tblPr");
        var propertiesResult = readTableProperties(tableProperties);
        // 新增table样式解析
        var preservedStyles = preserverUtils.extractPreservableTableStyles('table', tableProperties);
        //
        return readXmlElements(element.children)
            .flatMap(calculateRowSpans)
            .flatMap(function (children) {
                return propertiesResult.map(function (properties) {
                    //
                    properties.preservedStyles = preservedStyles;
                    //
                    return documents.Table(children, properties);
                });
            });
    }

    function readTableProperties(element) {
        return readTableStyle(element).map(function (style) {
            return {
                styleId: style.styleId,
                styleName: style.name
            };
        });
    }

    function readTableRow(element) {
        var properties = element.firstOrEmpty("w:trPr");
        var isHeader = !!properties.first("w:tblHeader");
        // 新增行样式解析
        var trProperties = element.firstOrEmpty('w:tblPrEx');
        var preservableStyles = preserverUtils.extractPreservableTableStyles('table', trProperties);

        // 行高样式解析 todo
        var trHeight = properties.first("w:trHeight");
        if (trHeight) {

        }
        //
        return readXmlElements(element.children).map(function (children) {
            return documents.TableRow(children, {
                isHeader: isHeader,
                preservedStyles: preservableStyles // 保留行样式
            });
        });
    }

    function readTableCell(element) {
        return readXmlElements(element.children).map(function (children) {
            var properties = element.firstOrEmpty("w:tcPr");

            // 读取单元格样式
            var preservedStyles = preserverUtils.extractPreservableTableStyles('cell', properties);
            // 宽度
            var width = properties.firstOrEmpty("w:tcW").attributes["w:w"];
            if (width) { // 有问题
                if (!preservedStyles) {
                    preservedStyles = {};
                }
                preservedStyles.cssText = `width:${parseInt(width, 10)}px`;
            }
            // 合并列
            var gridSpan = properties.firstOrEmpty("w:gridSpan").attributes["w:val"];
            var colSpan = gridSpan ? parseInt(gridSpan, 10) : 1;
            var cell = documents.TableCell(children, {
                colSpan: colSpan
            });
            //
            // 读取单元格样式
            var cell = documents.TableCell(children, {
                colSpan: colSpan,
                preservedStyles: preservedStyles
            });
            //
            cell._vMerge = readVMerge(properties);
            return cell;
        });
    }

    function readVMerge(properties) {
        var element = properties.first("w:vMerge");
        if (element) {
            var val = element.attributes["w:val"];
            return val === "continue" || !val;
        } else {
            return null;
        }
    }

    function calculateRowSpans(rows) {
        var unexpectedNonRows = _.any(rows, function (row) {
            return row.type !== documents.types.tableRow;
        });
        if (unexpectedNonRows) {
            return elementResultWithMessages(rows, [warning(
                "unexpected non-row element in table, cell merging may be incorrect"
            )]);
        }
        var unexpectedNonCells = _.any(rows, function (row) {
            return _.any(row.children, function (cell) {
                return cell.type !== documents.types.tableCell;
            });
        });
        if (unexpectedNonCells) {
            return elementResultWithMessages(rows, [warning(
                "unexpected non-cell element in table row, cell merging may be incorrect"
            )]);
        }

        var columns = {};

        rows.forEach(function (row) {
            var cellIndex = 0;
            row.children.forEach(function (cell) {
                if (cell._vMerge && columns[cellIndex]) {
                    columns[cellIndex].rowSpan++;
                } else {
                    columns[cellIndex] = cell;
                    cell._vMerge = false;
                }
                cellIndex += cell.colSpan;
            });
        });

        rows.forEach(function (row) {
            row.children = row.children.filter(function (cell) {
                return !cell._vMerge;
            });
            row.children.forEach(function (cell) {
                delete cell._vMerge;
            });
        });

        return elementResult(rows);
    }

    function readDrawingElement(element) {
        var pic = element
            .getElementsByTagName("a:graphic")
            .getElementsByTagName("a:graphicData")
            .getElementsByTagName("pic:pic")

        var blips = pic.getElementsByTagName("pic:blipFill")
            .getElementsByTagName("a:blip");

        // 读取图片大小
        // var ext = pic.getElementsByTagName("pic:spPr")
        //     .getElementsByTagName("a:xfrm")
        //     .getElementsByTagName("a:ext");


        return combineResults(blips.map(readBlip.bind(null, element)));
    }

    function readBlip(element, blip) {
        var properties = element.first("wp:docPr").attributes;
        var altText = isBlank(properties.descr) ? properties.title : properties.descr;
        // 读取图片大小
        var extent = element.first("wp:extent").attributes;
        return readImage(findBlipImageFile(blip), {
            altText,
            extent
        });
    }

    function isBlank(value) {
        return value == null || /^\s*$/.test(value);
    }

    function findBlipImageFile(blip) {
        var embedRelationshipId = blip.attributes["r:embed"];
        var linkRelationshipId = blip.attributes["r:link"];
        if (embedRelationshipId) {
            return findEmbeddedImageFile(embedRelationshipId);
        } else {
            var imagePath = relationships.findTargetByRelationshipId(linkRelationshipId);
            return {
                path: imagePath,
                read: files.read.bind(files, imagePath)
            };
        }
    }

    function readImageData(element) {
        var relationshipId = element.attributes['r:id'];

        if (relationshipId) {
            return readImage(
                findEmbeddedImageFile(relationshipId),
                element.attributes["o:title"]);
        } else {
            return emptyResultWithMessages([warning("A v:imagedata element without a relationship ID was ignored")]);
        }
    }

    function findEmbeddedImageFile(relationshipId) {
        var path = uris.uriToZipEntryName("word", relationships.findTargetByRelationshipId(relationshipId));
        return {
            path: path,
            read: docxFile.read.bind(docxFile, path)
        };
    }

    function readImage(imageFile, options) {
        var contentType = contentTypes.findContentType(imageFile.path);
        // 合并选项
        var image = documents.Image(Object.assign({
            readImage: imageFile.read,
            contentType: contentType
        }, options));
        var warnings = supportedImageTypes[contentType] ? [] : warning("Image of type " + contentType + " is unlikely to display in web browsers");
        return elementResultWithMessages(image, warnings);
    }

    function undefinedStyleWarning(type, styleId) {
        return warning(
            type + " style with ID " + styleId + " was referenced but not defined in the document");
    }
}


function readNumberingProperties(element, numbering) {
    var level = element.firstOrEmpty("w:ilvl").attributes["w:val"];
    var numId = element.firstOrEmpty("w:numId").attributes["w:val"];
    if (level === undefined || numId === undefined) {
        return null;
    } else {
        return numbering.findLevel(numId, level);
    }
}

var supportedImageTypes = {
    "image/png": true,
    "image/gif": true,
    "image/jpeg": true,
    "image/svg+xml": true,
    "image/tiff": true
};
// 未读取，且已知要忽略的元素
var ignoreElements = {
    "office-word:wrap": true,
    "v:shadow": true,
    "v:shapetype": true,
    "w:annotationRef": true,
    "w:bookmarkEnd": true,
    "w:sectPr": true, // 页面信息
    "w:proofErr": true,
    "w:lastRenderedPageBreak": true,
    "w:commentRangeStart": true,
    "w:commentRangeEnd": true,
    "w:del": true,
    "w:footnoteRef": true,
    "w:endnoteRef": true,
    "w:tblPr": true, // 表格或表格行的样式解析
    "w:tblGrid": true,
    "w:trPr": true, // 表格行样式解析
    "w:tcPr": true, // 单元格样式解析
    "w:tblPrEx": true // 新增忽略元素
};

function isParagraphProperties(element) {
    return element.type === "paragraphProperties";
}

function isRunProperties(element) {
    return element.type === "runProperties";
}

function negate(predicate) {
    return function (value) {
        return !predicate(value);
    };
}

function emptyResultWithMessages(messages) {
    return new ReadResult(null, null, messages);
}

function emptyResult() {
    return new ReadResult(null);
}

function elementResult(element) {
    return new ReadResult(element);
}

function elementResultWithMessages(element, messages) {
    return new ReadResult(element, null, messages);
}

function ReadResult(element, extra, messages) {
    this.value = element || [];
    this.extra = extra;
    this._result = new Result({
        element: this.value,
        extra: extra
    }, messages);
    this.messages = this._result.messages;
}

ReadResult.prototype.toExtra = function () {
    return new ReadResult(null, joinElements(this.extra, this.value), this.messages);
};

ReadResult.prototype.insertExtra = function () {
    var extra = this.extra;
    if (extra && extra.length) {
        return new ReadResult(joinElements(this.value, extra), null, this.messages);
    } else {
        return this;
    }
};

ReadResult.prototype.map = function (func) {
    var result = this._result.map(function (value) {
        return func(value.element);
    });
    return new ReadResult(result.value, this.extra, result.messages);
};

ReadResult.prototype.flatMap = function (func) {
    var result = this._result.flatMap(function (value) {
        return func(value.element)._result;
    });
    return new ReadResult(result.value.element, joinElements(this.extra, result.value.extra), result.messages);
};

function combineResults(results) {
    var result = Result.combine(_.pluck(results, "_result"));
    return new ReadResult(
        _.flatten(_.pluck(result.value, "element")),
        _.filter(_.flatten(_.pluck(result.value, "extra")), identity),
        result.messages
    );
}

function joinElements(first, second) {
    return _.flatten([first, second]);
}

function identity(value) {
    return value;
}