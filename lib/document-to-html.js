var _ = require("underscore");

var promises = require("./promises");
var documents = require("./documents");
var htmlPaths = require("./styles/html-paths");
var results = require("./results");
var images = require("./images");
var Html = require("./html");
var writers = require("./writers");
//  新增样式解析渲染工具
var preserverUtils = require('./styles/preserver-utils');

exports.DocumentConverter = DocumentConverter;


function DocumentConverter(options) {
    return {
        convertToHtml: function (element, docxFile) {
            var comments = _.indexBy(
                element.type === documents.types.document ? element.comments : [],
                "commentId"
            );
            var conversion = new DocumentConversion(options, comments);
            return conversion.convertToHtml(element, docxFile);
        }
    };
}

function DocumentConversion(options, comments) {
    var noteNumber = 1;

    var noteReferences = [];

    var referencedComments = [];

    options = _.extend({
        ignoreEmptyParagraphs: true
    }, options);
    var idPrefix = options.idPrefix === undefined ? "" : options.idPrefix;
    var ignoreEmptyParagraphs = options.ignoreEmptyParagraphs;

    var defaultParagraphStyle = htmlPaths.topLevelElement("p");
    //
    var _stylePreservations = preserverUtils.processOptions(options);
    //
    var styleMap = options.styleMap || [];
    //  文档信息
    var documentInfo = preserverUtils.currentDocumentStyle.documentInfo = {};

    // 
    function convertToHtml(document, docxFile) {
        var messages = [];
        var html = elementToHtml(document, messages, {});

        var deferredNodes = [];
        walkHtml(html, function (node) {
            if (node.type === "deferred") {
                deferredNodes.push(node);
            }

        });
        var deferredValues = {};
        return promises.mapSeries(deferredNodes, function (deferred) {
            return deferred.value().then(function (value) {
                deferredValues[deferred.id] = value;
            });
        }).then(function () {
            function replaceDeferred(nodes) {
                return flatMap(nodes, function (node) {
                    if (node.type === "deferred") {
                        return deferredValues[node.id];
                    } else if (node.children) {
                        return [
                            _.extend({}, node, {
                                children: replaceDeferred(node.children)
                            })
                        ];
                    } else {
                        return [node];
                    }
                });
            }
            var writer = writers.writer({
                prettyPrint: options.prettyPrint,
                outputFormat: options.outputFormat
            });
            Html.write(writer, Html.simplify(replaceDeferred(html)));
            let res = new results.Result(writer.asString(), messages, documentInfo);
            if (!res.value) {
                let htmlchunk = docxFile.read('word/afchunk.mht', 'utf-8')._rejectionHandler0

                res.value = /<body>([\s\S]+)<\/body>/.exec(htmlchunk)[1].replace(/3D/g, '')
                // 临时处理，当文件上传到服务器再下载时无法正常加载base64图片的问题。
                // @author wyf
                let matches = htmlchunk.match(/(Content-Type: image)([\s\S]+?)(------=mhtDocumentPart)/g)
                matches && matches.forEach(matched => {
                    let arr = matched.split(/\n+/g)
                    if (arr[1].split(': ')[1] !== 'base64') {
                        return
                    }
                    // 图片路径替换成base64字符串
                    res.value = res.value.replace(arr[2].split(': ')[1], 'data:' + arr[1].split(': ')[1] + ';base64,' + arr[3])
                })
            }
            return res;
        });
    }

    function convertElements(elements, messages, options) {
        return flatMap(elements, function (element) {
            return elementToHtml(element, messages, options);
        });
    }

    function elementToHtml(element, messages, options) {
        if (!options) {
            throw new Error("options not set");
        }
        var handler = elementConverters[element.type];
        if (handler) {
            return handler(element, messages, options);
        } else {
            return [];
        }
    }

    function convertParagraph(element, messages, options) {
        if (element.pageInfo) {
            documentInfo.pages = documentInfo.pages || [];
            documentInfo.pages.push(element.pageInfo);
        }
        return htmlPathForParagraph(element, messages).wrap(function () {
            var content = convertElements(element.children, messages, options);

            if (ignoreEmptyParagraphs) {
                return content;
            } else {
                if (!content.length) { // 空行支持，与 html-docx-js 搭配，todo完善
                    content.push({
                        type: "text",
                        value: "  "
                    })
                }
                return [Html.forceWrite].concat(content);
            }
        });
    }

    function htmlPathForParagraph(element, messages) {
        var style = findStyle(element);
        if (style) {
            return style.to;
        } else {
            if (element.styleId) {
                messages.push(unrecognisedStyleWarning("paragraph", element));
            }
            // mocheer
            // 支持style段落样式配置，主要是text-align
            if (element.style) {
                let to = htmlPaths.topLevelElement("p")
                to._elements[0].attributes.style = element.style;
                return to;
            }

            return defaultParagraphStyle;
        }
    }

    function convertRun(run, messages, options) {
        var nodes = function () {
            return convertElements(run.children, messages, options);
        };
        var paths = [];
        if (run.isSmallCaps) {
            paths.push(findHtmlPathForRunProperty("smallCaps"));
        }
        if (run.isStrikethrough) {
            paths.push(findHtmlPathForRunProperty("strikethrough", "s"));
        }

        if (run.isUnderline) {
            paths.push(findHtmlPathForRunProperty("underline", 'u'));
        }

        if (run.verticalAlignment === documents.verticalAlignment.subscript) {
            paths.push(htmlPaths.element("sub", {}, {
                fresh: false
            }));
        }
        if (run.verticalAlignment === documents.verticalAlignment.superscript) {
            paths.push(htmlPaths.element("sup", {}, {
                fresh: false
            }));
        }

        if (run.isItalic) {
            paths.push(findHtmlPathForRunProperty("italic", "em"));
        }
        if (run.isBold) {
            paths.push(findHtmlPathForRunProperty("bold", "strong"));
        }
        // 新增样式
        let styleString = '';
        if (run.font) {
            styleString += `font-family:${run.font};`;
        }
        // 新增颜色样式
        if (run.color && run.color !== 'auto') { //docx 转换的color默认值是auto
            styleString += `color:#${run.color};`;
        }
        // 新增字体大小
        if (run.fontSize) {
            styleString += `font-size:${(run.fontSize / 2)}pt;`;
        }
        // span
        if (styleString) {
            paths.push(htmlPaths.element('span', {
                style: styleString
            }, {
                fresh: false
            }));
        }

        var stylePath = htmlPaths.empty;
        var style = findStyle(run);
        if (style) {
            stylePath = style.to;
        } else if (run.styleId) {
            messages.push(unrecognisedStyleWarning("run", run));
        }

        paths.push(stylePath);

        paths.forEach(function (path) {
            nodes = path.wrap.bind(path, nodes);
        });
        return nodes();
    }

    function findHtmlPathForRunProperty(elementType, defaultTagName) {
        var path = findHtmlPath({
            type: elementType
        });
        if (path) {
            return path;
        } else if (defaultTagName) {
            return htmlPaths.element(defaultTagName, {}, {
                fresh: false
            });
        } else {
            return htmlPaths.empty;
        }
    }

    function findHtmlPath(element, defaultPath) {
        var style = findStyle(element);
        return style ? style.to : defaultPath;
    }

    function findStyle(element) {
        for (var i = 0; i < styleMap.length; i++) {
            if (styleMap[i].from.matches(element)) {
                return styleMap[i];
            }
        }
    }

    function recoveringConvertImage(convertImage) {
        return function (image, messages) {
            return promises.attempt(function () {
                return convertImage(image, messages);
            }).caught(function (error) {
                messages.push(results.error(error));
                return [];
            });
        };
    }

    function noteHtmlId(note) {
        return referentHtmlId(note.noteType, note.noteId);
    }

    function noteRefHtmlId(note) {
        return referenceHtmlId(note.noteType, note.noteId);
    }

    function referentHtmlId(referenceType, referenceId) {
        return htmlId(referenceType + "-" + referenceId);
    }

    function referenceHtmlId(referenceType, referenceId) {
        return htmlId(referenceType + "-ref-" + referenceId);
    }

    function htmlId(suffix) {
        return idPrefix + suffix;
    }

    var defaultTablePath = htmlPaths.elements([
        htmlPaths.element("table", {}, {
            fresh: true
        })
    ]);
    /**
     * 
     * @param {*} element 
     * @param {*} messages 
     * @param {*} options 
     */
    function convertTable(element, messages, options) {
        // return findHtmlPath(element, defaultTablePath).wrap(function () {
        //     return convertTableChildren(element, messages, options);
        // });
        var htmlPath = findHtmlPath(element, defaultTablePath); //默认，所有table公用一个defaultTablePath 对象
        //
        if (_stylePreservations.applyTableStyles) {
            options.cellHorizontalEdges = element.preservedStyles ? (element.preservedStyles.horizontalEdges || null) : null;
            options.cellVerticalEdges = element.preservedStyles ? (element.preservedStyles.verticalEdges || null) : null;
            if (!_stylePreservations.ignoreTableElementBorders) {
                if (htmlPath._elements && htmlPath._elements.length > 0) {
                    if (element.preservedStyles) {
                        htmlPath = htmlPaths.elements([
                            htmlPaths.element("table", {}, {
                                fresh: true
                            })
                        ]);
                        let attributes = htmlPath._elements[0].attributes;
                        //  表格table不应用样式，只用于tr和td继承
                        let style = preserverUtils.convertPreservableStylesToCssString(element.preservedStyles, _stylePreservations.reduceCellBorderStylesUsed, 'table');
                        // 新会区河长制-水质通报的红色分割行，特殊处理
                        if(element.children.length==0){
                            attributes.style = style; 
                        }
                        // 
                        // todo 重写 表格整体边框样式
                        // style = T.styleToObj(style);
                        // let flag = ['border-bottom', 'border-left', 'border-right', 'border-top'].filter(e => style[e] && style[e].includes('none')).length > 0;
                        // if (flag) {
                        //     attributes.class = "noBorder"
                        // }

                    }
                }
            }
        }

        //
        return htmlPath.wrap(function () {
            return convertTableChildren(element, messages, options);
        });
    }

    function convertTableChildren(element, messages, options) {
        var bodyIndex = _.findIndex(element.children, function (child) {
            return !child.type === documents.types.tableRow || !child.isHeader;
        });
        if (bodyIndex === -1) {
            bodyIndex = element.children.length;
        }
        var children;
        if (bodyIndex === 0) {
            children = convertElements(
                element.children,
                messages,
                _.extend({}, options, {
                    isTableHeader: false
                })
            );
        } else {
            var headRows = convertElements(
                element.children.slice(0, bodyIndex),
                messages,
                _.extend({}, options, {
                    isTableHeader: true
                })
            );
            var bodyRows = convertElements(
                element.children.slice(bodyIndex),
                messages,
                _.extend({}, options, {
                    isTableHeader: false
                })
            );
            children = [
                Html.freshElement("thead", {}, headRows),
                Html.freshElement("tbody", {}, bodyRows)
            ];
        }
        return [Html.forceWrite].concat(children);
    }

    function convertTableRow(element, messages, options) {
        var children = convertElements(element.children, messages, options);
        var attributes = {};
        //新增行样式解析
        if (_stylePreservations.applyTableStyles && element.preservedStyles) {
            if (options.cellHorizontalEdges) {
                element.preservedStyles.horizontalEdges = options.cellHorizontalEdges;
            }
            if (options.cellVerticalEdges) {
                element.preservedStyles.verticalEdges = options.cellVerticalEdges;
            }

            let style = preserverUtils.convertPreservableStylesToCssString(element.preservedStyles, _stylePreservations.reduceCellBorderStylesUsed, 'row');
            // 行样式解析，这里包括行高，不能去掉
            attributes.style = style;
            //todo 重写
            // style = T.styleToObj(style)
            // let flag = ['border-bottom', 'border-left', 'border-right', 'border-top'].filter(e => style[e] && !style[e].includes('none')).length === 0;
            // if (flag) {
            //     attributes.class = "noBorder"
            // }

        }

        return [
            Html.freshElement("tr", attributes, [Html.forceWrite].concat(children))
        ];
    }

    function convertTableCell(element, messages, options) {
        var tagName = options.isTableHeader ? "th" : "td";
        var children = convertElements(element.children, messages, options);
        var attributes = {};
        if (element.colSpan !== 1) {
            attributes.colspan = element.colSpan.toString();
        }
        if (element.rowSpan !== 1) {
            attributes.rowspan = element.rowSpan.toString();
        }
        // mocheer
        // 保留表格的单元格边框等样式
        if (_stylePreservations.applyTableStyles) {
            // 山东省水雨情快报特殊处理，暂未深究
            let isHeadingCell = !(element.children.length === 1 && !element.children[0].children.length && element.children[0].styleId === "Heading1");
            if (element.preservedStyles) {
                if (options.cellHorizontalEdges) {
                    element.preservedStyles.horizontalEdges = options.cellHorizontalEdges;
                }
                if (options.cellVerticalEdges) {
                    element.preservedStyles.verticalEdges = options.cellVerticalEdges;
                }
                if (isHeadingCell) {
                    attributes.style = preserverUtils.convertPreservableStylesToCssString(element.preservedStyles, _stylePreservations.reduceCellBorderStylesUsed, 'cell');
                }

            } else {//从行内继承
                // 当内容为空的时候不继承样式，这个机制有待完善
                if (isHeadingCell) {
                    attributes.style = preserverUtils.convertPreservableStylesToCssString(preserverUtils.getCurTableRowStyles(), _stylePreservations.reduceCellBorderStylesUsed, 'cell');
                } else {
                    console.log(element)
                }
                //
            }

        }
        //
        return [
            Html.freshElement(tagName, attributes, [Html.forceWrite].concat(children))
        ];

    }

    function convertCommentReference(reference, messages, options) {
        return findHtmlPath(reference, htmlPaths.ignore).wrap(function () {
            var comment = comments[reference.commentId];
            var count = referencedComments.length + 1;
            var label = "[" + commentAuthorLabel(comment) + count + "]";
            referencedComments.push({
                label: label,
                comment: comment
            });
            // TODO: remove duplication with note references
            return [
                Html.freshElement("a", {
                    href: "#" + referentHtmlId("comment", reference.commentId),
                    id: referenceHtmlId("comment", reference.commentId)
                }, [Html.text(label)])
            ];
        });
    }

    function convertComment(referencedComment, messages, options) {
        // TODO: remove duplication with note references

        var label = referencedComment.label;
        var comment = referencedComment.comment;
        var body = convertElements(comment.body, messages, options).concat([
            Html.nonFreshElement("p", {}, [
                Html.text(" "),
                Html.freshElement("a", {
                    "href": "#" + referenceHtmlId("comment", comment.commentId)
                }, [
                    Html.text("↑")
                ])
            ])
        ]);

        return [
            Html.freshElement(
                "dt", {
                "id": referentHtmlId("comment", comment.commentId)
            },
                [Html.text("Comment " + label)]
            ),
            Html.freshElement("dd", {}, body)
        ];
    }

    function convertBreak(element, messages, options) {
        return htmlPathForBreak(element).wrap(function () {
            return [];
        });
    }
    /**
     * 页面信息
     * @param {*} element 
     * @param {*} messages 
     * @param {*} options 
     */
    function convertPage(element, messages, options) {
        documentInfo.pageInfo = element.options;
        return [];
    }

    function htmlPathForBreak(element) {
        var style = findStyle(element);
        if (style) {
            return style.to;
        } else if (element.breakType === "line") {
            return htmlPaths.topLevelElement("br");
        } else {
            return htmlPaths.empty;
        }
    }

    var elementConverters = {
        "document": function (document, messages, options) {
            var children = convertElements(document.children, messages, options);
            var notes = noteReferences.map(function (noteReference) {
                return document.notes.resolve(noteReference);
            });
            var notesNodes = convertElements(notes, messages, options);
            return children.concat([
                Html.freshElement("ol", {}, notesNodes),
                Html.freshElement("dl", {}, flatMap(referencedComments, function (referencedComment) {
                    return convertComment(referencedComment, messages, options);
                }))
            ]);
        },
        "paragraph": convertParagraph,
        "run": convertRun,
        "text": function (element, messages, options) {
            return [Html.text(element.value)];
        },
        "tab": function (element, messages, options) {
            return [Html.text("\t")];
        },
        "hyperlink": function (element, messages, options) {
            var href = element.anchor ? "#" + htmlId(element.anchor) : element.href;
            var attributes = {
                href: href
            };
            if (element.targetFrame != null) {
                attributes.target = element.targetFrame;
            }

            var children = convertElements(element.children, messages, options);
            return [Html.nonFreshElement("a", attributes, children)];
        },
        "bookmarkStart": function (element, messages, options) {
            var anchor = Html.freshElement("a", {
                id: htmlId(element.name)
            }, [Html.forceWrite]);
            return [anchor];
        },
        "noteReference": function (element, messages, options) {
            noteReferences.push(element);
            var anchor = Html.freshElement("a", {
                href: "#" + noteHtmlId(element),
                id: noteRefHtmlId(element)
            }, [Html.text("[" + (noteNumber++) + "]")]);

            return [Html.freshElement("sup", {}, [anchor])];
        },
        "note": function (element, messages, options) {
            var children = convertElements(element.body, messages, options);
            var backLink = Html.elementWithTag(htmlPaths.element("p", {}, {
                fresh: false
            }), [
                Html.text(" "),
                Html.freshElement("a", {
                    href: "#" + noteRefHtmlId(element)
                }, [Html.text("↑")])
            ]);
            var body = children.concat([backLink]);

            return Html.freshElement("li", {
                id: noteHtmlId(element)
            }, body);
        },
        "commentReference": convertCommentReference,
        "comment": convertComment,
        "image": deferredConversion(recoveringConvertImage(options.convertImage || images.dataUri)),
        "table": convertTable,
        "tableRow": convertTableRow,
        "tableCell": convertTableCell,
        "break": convertBreak,
        //新增
        "page": convertPage
    };
    return {
        convertToHtml: convertToHtml
    };
}

var deferredId = 1;

function deferredConversion(func) {
    return function (element, messages, options) {
        return [{
            type: "deferred",
            id: deferredId++,
            value: function () {
                return func(element, messages, options);
            }
        }];
    };
}

function unrecognisedStyleWarning(type, element) {
    return results.warning(
        "Unrecognised " + type + " style: '" + element.styleName + "'" +
        " (Style ID: " + element.styleId + ")"
    );
}

function flatMap(values, func) {
    return _.flatten(values.map(func), true);
}

function walkHtml(nodes, callback) {
    nodes.forEach(function (node) {
        callback(node);
        if (node.children) {
            walkHtml(node.children, callback);
        }
    });
}

var commentAuthorLabel = exports.commentAuthorLabel = function commentAuthorLabel(comment) {
    return comment.authorInitials || "";
};