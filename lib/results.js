var _ = require("underscore");


exports.Result = Result;
exports.success = success;
exports.warning = warning;
exports.error = error;


function Result(value, messages, options) {
    this.value = value;
    this.messages = messages || [];
    this.options = options; // 用于保存页面样式信息
}

Result.prototype.map = function (func) {
    return new Result(func(this.value), this.messages);
};

Result.prototype.flatMap = function (func) {
    var funcResult = func(this.value);
    return new Result(funcResult.value, combineMessages([this, funcResult]), funcResult.options);
};

Result.prototype.flatMapThen = function (func) {
    var that = this;
    return func(this.value).then(function (otherResult) {
        return new Result(otherResult.value, combineMessages([that, otherResult]), otherResult.options);
    });
};

Result.combine = function (results) {
    var values = _.flatten(_.pluck(results, "value"));
    var messages = combineMessages(results);
    return new Result(values, messages);
};

function success(value) {
    return new Result(value, []);
}

function warning(message) {
    return {
        type: "warning",
        message: message
    };
}

function error(exception) {
    return {
        type: "error",
        message: exception.message,
        error: exception
    };
}

function combineMessages(results) {
    var messages = [];
    _.flatten(_.pluck(results, "messages"), true).forEach(function (message) {
        if (!containsMessage(messages, message)) {
            messages.push(message);
        }
    });
    return messages;
}

function containsMessage(messages, message) {
    return _.find(messages, isSameMessage.bind(null, message)) !== undefined;
}

function isSameMessage(first, second) {
    return first.type === second.type && first.message === second.message;
}