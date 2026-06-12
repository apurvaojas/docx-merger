'use strict';

// String.prototype.replace(string, replacement) interprets $&, $', $`, $n in the
// replacement — document content can legally contain those. Always splice instead.
function replaceBetween(str, startIndex, endIndex, replacement) {
    return str.slice(0, startIndex) + replacement + str.slice(endIndex);
}

// replace everything from the first occurrence of marker to end of string
function replaceFrom(str, marker, replacement) {
    var startIndex = str.indexOf(marker);
    if (startIndex === -1) return str;
    return str.slice(0, startIndex) + replacement;
}

function escapeRegExp(s) {
    return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

module.exports = { replaceBetween: replaceBetween, replaceFrom: replaceFrom, escapeRegExp: escapeRegExp };
