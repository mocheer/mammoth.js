// @see https://github.com/pirate-matt/mammoth-with-style-preservation.js
// @see https://github.com/mwilliamson/mammoth.js/compare/master...pirate-matt:master?diff=split#diff-68b60b2443cf3be90e7e7223aaf3d383L100

exports.processOptions = _processOptions;
exports.extractPreservableTableStyles = _extractPreservableTableStyles;
exports.convertPreservableStylesToCssString = _convertPreservableStylesToCssString;
// 新增获取文档解析的样式、临时保存的样式
const currentDocumentStyle = {
  //当前表格的全局样式，用于继承
  curTableStyles: null,
  curTableRowStyles: null
};
exports.currentDocumentStyle = currentDocumentStyle
//
var _ = require('underscore');

function _processOptions(options) {
  options = options || {};

  var preservationOptions = options.stylePreservations || {};

  if (preservationOptions === 'all') {
    preservationOptions = {
      useColorSpans: true,
      useFontSizeSpans: true,
      useStrictFontSize: true,
      applyTableStyles: true,
      reduceCellBorderStylesUsed: false,
      ignoreTableElementBorders: false
    };
  }

  if (preservationOptions === 'default') {
    preservationOptions = {
      useColorSpans: true,
      useFontSizeSpans: true,
      useStrictFontSize: false,
      applyTableStyles: true,
      reduceCellBorderStylesUsed: false,
      ignoreTableElementBorders: false
    };
  }

  if (typeof preservationOptions !== 'object') {
    preservationOptions = {};
  }

  return preservationOptions;
}

// 用于提取table和td的样式
// used for both table defaults (<table>) and cells (<td>)
function _extractPreservableTableStyles(elementType, elementProperties, options) {
  if (elementType !== 'table' && elementType !== 'cell') {
    return '';
  }
  options = options || {};

  var fill = elementProperties.firstOrEmpty('w:shd').attributes['w:fill'];
  var cellMargins = elementProperties.firstOrEmpty(elementType === 'cell' ? 'w:tcMar' : 'w:tblCellMar');
  var borders = elementProperties.firstOrEmpty(elementType === 'cell' ? 'w:tcBorders' : 'w:tblBorders');


  var styles = {
    fill: fill && fill !== 'auto' ? fill : null,
    cellMarginTop: _extractCellMarginStyles('top', cellMargins),
    cellMarginLeft: _extractCellMarginStyles('left', cellMargins),
    cellMarginBottom: _extractCellMarginStyles('bottom', cellMargins),
    cellMarginRight: _extractCellMarginStyles('right', cellMargins),
    borderTop: _extractBorderStyles('top', borders),
    borderLeft: _extractBorderStyles('left', borders),
    borderBottom: _extractBorderStyles('bottom', borders),
    borderRight: _extractBorderStyles('right', borders),
    horizontalEdges: _extractBorderStyles('insideH', borders),
    verticalEdges: _extractBorderStyles('insideV', borders)
  };

  styles = _reduceBorderStyles(styles);

  if (
    styles.fill === null &&
    styles.cellMarginTop === null && styles.cellMarginLeft === null && styles.cellMarginBottom === null && styles.cellMarginRight === null &&
    styles.borderTop === null && styles.borderLeft === null && styles.borderBottom === null && styles.borderRight === null
  ) {
    styles = null;
  }

  return styles;
}


function _reduceBorderStyles(styles) {
  var directions = ['Top', 'Left', 'Bottom', 'Right'];
  var widthsAndCounts = {};
  var stylesAndCounts = {};
  var colorsAndCounts = {};

  _.each(directions, function (direction) {
    var borderKey = 'border' + direction;

    if (styles[borderKey]) {
      var directionsStyles = styles[borderKey];
      if (directionsStyles.width) {
        _incrementOrStartCount(widthsAndCounts, directionsStyles.width);
      }
      if (directionsStyles.style) {
        _incrementOrStartCount(stylesAndCounts, directionsStyles.style);
      }
      if (directionsStyles.color) {
        _incrementOrStartCount(colorsAndCounts, directionsStyles.color);
      }
    }
  });

  var sortedWidths = _.sortBy(_.values(widthsAndCounts), 'count').reverse();
  var sortedStyles = _.sortBy(_.values(stylesAndCounts), 'count').reverse();
  var sortedColors = _.sortBy(_.values(colorsAndCounts), 'count').reverse();

  styles.simplifiedBorder = {
    width: sortedWidths.length ? sortedWidths[0].val : null,
    style: sortedStyles.length ? sortedStyles[0].val : null,
    color: sortedColors.length ? sortedColors[0].val : null
  };

  styles.simplifiedBorder = (styles.simplifiedBorder.width || styles.simplifiedBorder.style || styles.simplifiedBorder.color) ? styles.simplifiedBorder : null;

  return styles;
}


function _incrementOrStartCount(resultObj, key) {
  if (resultObj[key]) {
    resultObj[key].count += 1;
  } else {
    resultObj[key] = {
      val: key,
      count: 1
    };
  }
}


function _extractBorderStyles(borderElementKey, borders) {
  var borderElement = borders.firstOrEmpty('w:' + borderElementKey);

  var borderWidth = borderElement.attributes['w:sz'];
  var borderStyle = borderElement.attributes['w:val'];
  var borderColor = borderElement.attributes['w:color'];

  var borderStyles = {
    width: borderWidth || null,
    style: borderStyle || null,
    color: borderColor || null
  };

  return (borderWidth || borderStyle || borderColor ? borderStyles : null);
}


function _extractCellMarginStyles(marginElementKey, margins) {
  var directionalMargin = margins.firstOrEmpty('w:' + marginElementKey);

  return directionalMargin.attributes['w:w'] || null;
}


/**
 * 用于table和tableRow,tableCell的样式解析，返回 dom 的 css 字符串
 * @param {*} elementStyles 
 * @param {*} reduceBorderStyles 
 */
function _convertPreservableStylesToCssString(elementStyles, reduceBorderStyles, type) {
  let curTableStyles;
  let curTableRowStyles;
  switch (type) {
    case 'table':
      curTableStyles = currentDocumentStyle.curTableStyles = elementStyles;
    case 'row':
      curTableStyles = currentDocumentStyle.curTableStyles || {};
      curTableRowStyles = currentDocumentStyle.curTableRowStyles =   elementStyles = Object.assign({}, elementStyles);
      curTableRowStyles.borderBottom = curTableRowStyles.borderBottom || curTableStyles.borderBottom;
      curTableRowStyles.borderTop = curTableRowStyles.borderTop || curTableStyles.borderTop;
      curTableRowStyles.borderLeft = curTableRowStyles.borderLeft || curTableStyles.borderLeft;
      curTableRowStyles.borderRight = curTableRowStyles.borderRight || curTableStyles.borderRight;
      break;
    default:
      curTableRowStyles = currentDocumentStyle.curTableRowStyles || {};
      elementStyles.borderBottom = elementStyles.borderBottom || curTableRowStyles.borderBottom ;
      elementStyles.borderTop = elementStyles.borderTop || curTableRowStyles.borderTop ;
      elementStyles.borderLeft = elementStyles.borderLeft ||  curTableRowStyles.borderLeft ;
      elementStyles.borderRight = elementStyles.borderRight || curTableRowStyles.borderRight;

      // 以下有问题
      if (!elementStyles.borderBottom && !elementStyles.borderTop && !elementStyles.borderLeft && !elementStyles.borderRight) {
        elementStyles.borderBottom = elementStyles.borderTop = elementStyles.borderLeft = elementStyles.borderRight = {
          width: 8,
          style: 'thick'
        }
      }
      break;
  }

  var cssString = '';

  // @FUTURE: feature toggle each of these?
  cssString += elementStyles.fill ? ('background-color: #' + elementStyles.fill + ';') : '';

  // 不准确，或者Table Grid 的不向下传递? 待测试验证
  // NOTE: these "table-wide" edges cascade down from `w:tblBorders`, and will be overriden by the cascading nature of css if any
  //       conflicting top/left/bottom/right borders are specfied w/in the `w:tcBorders` docx xml element
  // if (elementStyles.horizontalEdges) {
  //   cssString += _convertBorderStylesToCssString('top', elementStyles.horizontalEdges)
  //   cssString += _convertBorderStylesToCssString('bottom', elementStyles.horizontalEdges);
  // }
  // if (elementStyles.verticalEdges) {
  //   cssString += _convertBorderStylesToCssString('left', elementStyles.verticalEdges);
  //   cssString += _convertBorderStylesToCssString('right', elementStyles.verticalEdges);
  // }

  if (reduceBorderStyles) {
    cssString += elementStyles.simplifiedBorder ? _convertBorderStylesToCssString('', elementStyles.simplifiedBorder) : '';
  } else {
    cssString += elementStyles.borderTop ? _convertBorderStylesToCssString('top', elementStyles.borderTop) : '';
    cssString += elementStyles.borderLeft ? _convertBorderStylesToCssString('left', elementStyles.borderLeft) : '';
    cssString += elementStyles.borderBottom ? _convertBorderStylesToCssString('bottom', elementStyles.borderBottom) : '';
    cssString += elementStyles.borderRight ? _convertBorderStylesToCssString('right', elementStyles.borderRight) : '';
  }

  cssString += elementStyles.cellMarginTop ? 'padding-top: ' + (elementStyles.cellMarginTop / 20) + 'px;' : '';
  cssString += elementStyles.cellMarginLeft ? 'padding-left: ' + (elementStyles.cellMarginLeft / 20) + 'px;' : '';
  cssString += elementStyles.cellMarginBottom ? 'padding-bottom: ' + (elementStyles.cellMarginBottom / 20) + 'px;' : '';
  cssString += elementStyles.cellMarginRight ? 'padding-right: ' + (elementStyles.cellMarginRight / 20) + 'px;' : '';


  // 特殊处理的表格高度设置，目前用于分隔线等
  if (elementStyles.cssText) {
    cssString += elementStyles.cssText
  }

  return cssString;
}


var _docxBorderStylesToCssStyles = {
  single: 'solid',
  dashDotStroked: null,
  dashed: 'dashed',
  dashSmallGap: null,
  dotDash: 'dashed',
  dotDotDash: 'dotted',
  dotted: 'dotted',
  double: 'double',
  doubleWave: 'double',
  inset: 'inset',
  // nil: 'hidden', // 修改nil值对应的样式
  nil: 'unset',
  none: 'none',
  outset: 'outset',
  thick: 'solid',
  thickThinLargeGap: 'double',
  thickThinMediumGap: 'double',
  thickThinSmallGap: 'double',
  thinThickLargeGap: 'double',
  thinThickMediumGap: 'double',
  thinThickSmallGap: 'double',
  thinThickThinLargeGap: 'double',
  thinThickThinMediumGap: 'double',
  thinThickThinSmallGap: 'double',
  threeDEmboss: null,
  threeDEngrave: null,
  triple: null,
  wave: null
};

function _convertBorderStylesToCssString(whichBorder, borderStyles) {
  var css = 'border' + (whichBorder ? '-' + whichBorder : '') + ':';

  css += borderStyles.width ? ' ' + (borderStyles.width / 8) + 'pt' : ''; // border widths are stored in eights
  css += borderStyles.style && _docxBorderStylesToCssStyles[borderStyles.style] ? ' ' + _docxBorderStylesToCssStyles[borderStyles.style] : '';
  css += borderStyles.color ? ' #' + (borderStyles.color === 'auto' ? '000' : borderStyles.color) : '';

  css += ';';

  return css;
}