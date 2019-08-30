var _ = require("underscore");

var promises = require("./promises");
var Html = require("./html");

exports.imgElement = imgElement;

function imgElement(func) {
    return function (element, messages) {
        return promises.when(func(element)).then(function (result) {
            var attributes = _.clone(result);
            if (element.altText) {
                attributes.alt = element.altText;
            }
            // 读取图片大小
            if (element.extent) {
                var {
                    cx,
                    cy
                } = element.extent;
                var emusPerInch = 914400;
                // var emusPerCm = 360000;
                var width = cx / emusPerInch * 96
                var height = cy / emusPerInch * 96
                attributes.style = `width:${width}px;height:${height}px;`
            }

            return [Html.freshElement("img", attributes)];
        });
    };
}

// Undocumented, but retained for backwards-compatibility with 0.3.x
exports.inline = exports.imgElement;

exports.dataUri = imgElement(function (element) {
    return element.read("base64").then(function (imageBuffer) {
        return {
            src: "data:" + element.contentType + ";base64," + imageBuffer
        };
    });
});