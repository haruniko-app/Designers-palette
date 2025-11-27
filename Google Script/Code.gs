// Code.gs (v41.0 軽量監視・高速化版)

// 承認用テスト関数（エディターから実行可能）
function testAuth() {
  var presentation = SlidesApp.getActivePresentation();
  Logger.log('プレゼンテーション名: ' + presentation.getName());
  return '承認完了';
}

function onOpen() {
SlidesApp.getUi().createMenu('Designer\'s Palette').addItem('開く', 'showSidebar').addToUi();
}

function showSidebar() {
var html = HtmlService.createHtmlOutputFromFile('Sidebar').setTitle('Designer\'s Palette').setWidth(340);
SlidesApp.getUi().showSidebar(html);
}

/**
* 【新機能】選択状態のIDだけを返す軽量関数
* 画像データを取得しないため、一瞬で応答します。ポーリング監視用。
*/
function checkSelection() {
try {
var selection = SlidesApp.getActivePresentation().getSelection()                                   ;
var pageElementRange = selection.getPageElementRange()                                             ;
if (!pageElementRange) return null                                                                 ;

var elements = pageElementRange.getPageElements() ;
if (elements.length === 0) return null            ;

var element = elements[0]                                                         ;
// 全ての要素のIDを返す（画像、図形、線など）
return element.getObjectId() ;
} catch (e) {
return null                  ;
}
}

function checkSelectionForImage() {
try {
var selection = SlidesApp.getActivePresentation().getSelection()                                   ;
var pageElementRange = selection.getPageElementRange()                                             ;
if (!pageElementRange) return null                                                                 ;

var elements = pageElementRange.getPageElements() ;
if (elements.length === 0) return null            ;

var element = elements[0]                                                         ;
// 画像のみ返す
if (element.getPageElementType() !== SlidesApp.PageElementType.IMAGE) return null ;

return element.getObjectId() ;
} catch (e) {
return null                  ;
}
}

// --- 以下は既存の関数 (一部トリミング取得ロジックを含む) ---

function findElementRecursive(elements, targetId) {
if (!elements) return null                                           ;
for (var i = 0                                                       ; i < elements.length; i++) {
var el = elements[i]                                                 ;
if (el.objectId === targetId) return el                              ;
if (el.elementGroup && el.elementGroup.children) {
var found = findElementRecursive(el.elementGroup.children, targetId) ;
if (found) return found                                              ;
}
if (el.group && el.group.children) {
var foundOld = findElementRecursive(el.group.children, targetId)     ;
if (foundOld) return foundOld                                        ;
}
}
return null                                                          ;
}

function getSelectedImage() {
try {
var selection = SlidesApp.getActivePresentation().getSelection()            ;
var pageElementRange = selection.getPageElementRange()                      ;
if (!pageElementRange || pageElementRange.getPageElements().length === 0) {
throw new Error('画像を選択してください。');
}
var element = pageElementRange.getPageElements()[0]                         ;
if (element.getPageElementType() !== SlidesApp.PageElementType.IMAGE) {
throw new Error('選択された要素は画像ではありません。');
}

var image = element.asImage() ;
var blob = image.getBlob()    ;

if (blob.getBytes().length > 10 * 1024 * 1024) {
throw new Error('画像サイズが大きすぎます（10MB以下推奨）。');
}

var base64 = Utilities.base64Encode(blob.getBytes()) ;
var cropInfo = null                                  ;

try {
var presentationId = SlidesApp.getActivePresentation().getId()                      ;
var slideId = selection.getCurrentPage().getObjectId()                              ;
var objectId = image.getObjectId()                                                  ;
var page = Slides.Presentations.Pages.get(presentationId, slideId, { fields: "*" }) ;
var apiElement = findElementRecursive(page.pageElements, objectId)                  ;

if (apiElement && apiElement.image && apiElement.image.imageProperties && apiElement.image.imageProperties.cropProperties) {
cropInfo = apiElement.image.imageProperties.cropProperties                                                                   ;
}
} catch (e) {
console.warn("トリミング情報取得失敗: " + e.message)                                                              ;
}

return {
base64: 'data:' + blob.getContentType() + ';base64,' + base64,
width: image.getWidth(),
height: image.getHeight(),
objectId: image.getObjectId(),
cropInfo: cropInfo
}                                                              ;
} catch (e) {
throw new Error(e.message)                                     ;
}
}

function replaceImage(base64Data) {
try {
var contentType = base64Data.substring(5, base64Data.indexOf(';'));
var extension = contentType.split('/')[1];
var data = base64Data.split(',')[1];
var blob = Utilities.newBlob(Utilities.base64Decode(data), contentType, 'processed.' + extension) ;

var selection = SlidesApp.getActivePresentation().getSelection() ;
var elements = selection.getPageElementRange().getPageElements() ;

if (elements.length > 0 && elements[0].getPageElementType() === SlidesApp.PageElementType.IMAGE) {
elements[0].asImage().replace(blob)                                                                ;
} else {
var slide = SlidesApp.getActivePresentation().getSelection().getCurrentPage()                      ;
slide.insertImage(blob)                                                                            ;
}
return "完了";
} catch (e) {
throw new Error('配置エラー: ' + e.message)                                                   ;
}
}

// =================================================================
// 整列機能
// =================================================================

/**
 * 選択された要素を整列する汎用関数
 * @param {string} alignmentType - 整列タイプ (left, center, right, top, middle, bottom)
 * @param {string} referenceType - 基準タイプ (first, last, largest, smallest, slide)
 */
function alignElements(alignmentType, referenceType) {
  try {
    var selection = SlidesApp.getActivePresentation().getSelection();
    var pageElementRange = selection.getPageElementRange();
    var slide = selection.getCurrentPage();

    // スライド基準の場合は要素選択不要
    if (referenceType === 'slide') {
      if (!pageElementRange) {
        throw new Error('要素を選択してください');
      }
      var elements = pageElementRange.getPageElements();
      if (elements.length === 0) {
        throw new Error('要素を選択してください');
      }

      // スライドサイズを取得
      var presentation = SlidesApp.getActivePresentation();
      var slideWidth = presentation.getPageWidth();
      var slideHeight = presentation.getPageHeight();

      // スライド基準で整列
      for (var i = 0; i < elements.length; i++) {
        var element = elements[i];
        var width = element.getWidth();
        var height = element.getHeight();

        switch(alignmentType) {
          case 'left':
            element.setLeft(0);
            break;
          case 'center':
            element.setLeft((slideWidth - width) / 2);
            break;
          case 'right':
            element.setLeft(slideWidth - width);
            break;
          case 'top':
            element.setTop(0);
            break;
          case 'middle':
            element.setTop((slideHeight - height) / 2);
            break;
          case 'bottom':
            element.setTop(slideHeight - height);
            break;
        }
      }
      return '整列完了（スライド基準）';
    }

    if (!pageElementRange) {
      throw new Error('要素を選択してください');
    }

    var elements = pageElementRange.getPageElements();
    if (elements.length === 0) {
      throw new Error('要素を選択してください');
    }

    if (elements.length === 1) {
      return '複数の要素を選択してください';
    }

    // 基準要素を決定
    var referenceElement = null;
    var referenceIndex = 0;

    switch(referenceType) {
      case 'first':
        referenceElement = elements[0];
        referenceIndex = 0;
        break;
      case 'last':
        referenceElement = elements[elements.length - 1];
        referenceIndex = elements.length - 1;
        break;
      case 'largest':
        var maxArea = 0;
        for (var i = 0; i < elements.length; i++) {
          var area = elements[i].getWidth() * elements[i].getHeight();
          if (area > maxArea) {
            maxArea = area;
            referenceElement = elements[i];
            referenceIndex = i;
          }
        }
        break;
      case 'smallest':
        var minArea = Infinity;
        for (var i = 0; i < elements.length; i++) {
          var area = elements[i].getWidth() * elements[i].getHeight();
          if (area < minArea) {
            minArea = area;
            referenceElement = elements[i];
            referenceIndex = i;
          }
        }
        break;
      default:
        referenceElement = elements[0];
        referenceIndex = 0;
    }

    // 基準位置を取得
    var refLeft = referenceElement.getLeft();
    var refTop = referenceElement.getTop();
    var refWidth = referenceElement.getWidth();
    var refHeight = referenceElement.getHeight();

    var referenceValue;
    switch(alignmentType) {
      case 'left':
        referenceValue = refLeft;
        break;
      case 'center':
        referenceValue = refLeft + refWidth / 2;
        break;
      case 'right':
        referenceValue = refLeft + refWidth;
        break;
      case 'top':
        referenceValue = refTop;
        break;
      case 'middle':
        referenceValue = refTop + refHeight / 2;
        break;
      case 'bottom':
        referenceValue = refTop + refHeight;
        break;
    }

    // 基準位置に揃える（基準要素以外）
    for (var i = 0; i < elements.length; i++) {
      if (i === referenceIndex) continue;

      var element = elements[i];
      var width = element.getWidth();
      var height = element.getHeight();

      switch(alignmentType) {
        case 'left':
          element.setLeft(referenceValue);
          break;
        case 'center':
          element.setLeft(referenceValue - width / 2);
          break;
        case 'right':
          element.setLeft(referenceValue - width);
          break;
        case 'top':
          element.setTop(referenceValue);
          break;
        case 'middle':
          element.setTop(referenceValue - height / 2);
          break;
        case 'bottom':
          element.setTop(referenceValue - height);
          break;
      }
    }

    var refLabels = {
      'first': '最初の要素',
      'last': '最後の要素',
      'largest': '最大の要素',
      'smallest': '最小の要素'
    };
    return '整列完了（' + (refLabels[referenceType] || '基準') + '）';
  } catch (e) {
    throw new Error(e.message);
  }
}

/**
 * 左揃え
 */
function alignLeft() {
  return alignElements('LEFT');
}

/**
 * 中央揃え（水平）
 */
function alignCenter() {
  return alignElements('CENTER');
}

/**
 * 右揃え
 */
function alignRight() {
  return alignElements('RIGHT');
}

/**
 * 上揃え
 */
function alignTop() {
  return alignElements('TOP');
}

/**
 * 中央揃え（垂直）
 */
function alignMiddle() {
  return alignElements('MIDDLE');
}

/**
 * 下揃え
 */
function alignBottom() {
  return alignElements('BOTTOM');
}

/**
 * 水平方向に均等配置
 */
function distributeHorizontally() {
  try {
    var selection = SlidesApp.getActivePresentation().getSelection();
    var pageElementRange = selection.getPageElementRange();

    if (!pageElementRange) {
      throw new Error('要素を選択してください');
    }

    var elements = pageElementRange.getPageElements();
    if (elements.length < 3) {
      throw new Error('均等配置には3つ以上の要素を選択してください');
    }

    // 要素を左端順にソート
    var sortedElements = [];
    for (var i = 0; i < elements.length; i++) {
      sortedElements.push({
        element: elements[i],
        left: elements[i].getLeft(),
        width: elements[i].getWidth()
      });
    }
    sortedElements.sort(function(a, b) { return a.left - b.left; });

    // 最初と最後の要素の位置を取得
    var firstLeft = sortedElements[0].left;
    var lastRight = sortedElements[sortedElements.length - 1].left + sortedElements[sortedElements.length - 1].width;

    // すべての要素の幅の合計を計算
    var totalWidth = 0;
    for (var i = 0; i < sortedElements.length; i++) {
      totalWidth += sortedElements[i].width;
    }

    // 要素間の間隔を計算
    var availableSpace = lastRight - firstLeft - totalWidth;
    var spacing = availableSpace / (sortedElements.length - 1);

    // 各要素を配置（最初と最後は動かさない）
    var currentLeft = firstLeft;
    for (var i = 0; i < sortedElements.length; i++) {
      if (i > 0 && i < sortedElements.length - 1) {
        sortedElements[i].element.setLeft(currentLeft);
      }
      currentLeft += sortedElements[i].width + spacing;
    }

    return '水平方向に均等配置完了';
  } catch (e) {
    throw new Error(e.message);
  }
}

/**
 * 垂直方向に均等配置
 */
function distributeVertically() {
  try {
    var selection = SlidesApp.getActivePresentation().getSelection();
    var pageElementRange = selection.getPageElementRange();

    if (!pageElementRange) {
      throw new Error('要素を選択してください');
    }

    var elements = pageElementRange.getPageElements();
    if (elements.length < 3) {
      throw new Error('均等配置には3つ以上の要素を選択してください');
    }

    // 要素を上端順にソート
    var sortedElements = [];
    for (var i = 0; i < elements.length; i++) {
      sortedElements.push({
        element: elements[i],
        top: elements[i].getTop(),
        height: elements[i].getHeight()
      });
    }
    sortedElements.sort(function(a, b) { return a.top - b.top; });

    // 最初と最後の要素の位置を取得
    var firstTop = sortedElements[0].top;
    var lastBottom = sortedElements[sortedElements.length - 1].top + sortedElements[sortedElements.length - 1].height;

    // すべての要素の高さの合計を計算
    var totalHeight = 0;
    for (var i = 0; i < sortedElements.length; i++) {
      totalHeight += sortedElements[i].height;
    }

    // 要素間の間隔を計算
    var availableSpace = lastBottom - firstTop - totalHeight;
    var spacing = availableSpace / (sortedElements.length - 1);

    // 各要素を配置（最初と最後は動かさない）
    var currentTop = firstTop;
    for (var i = 0; i < sortedElements.length; i++) {
      if (i > 0 && i < sortedElements.length - 1) {
        sortedElements[i].element.setTop(currentTop);
      }
      currentTop += sortedElements[i].height + spacing;
    }

    return '垂直方向に均等配置完了';
  } catch (e) {
    throw new Error(e.message);
  }
}

// =================================================================
// カラー機能
// =================================================================

/**
 * 選択された要素の色情報を取得
 * @returns {Object} 塗りと線の色情報
 */
function getSelectedElementColors() {
  try {
    var selection = SlidesApp.getActivePresentation().getSelection();
    var pageElementRange = selection.getPageElementRange();

    if (!pageElementRange) {
      throw new Error('要素を選択してください');
    }

    var elements = pageElementRange.getPageElements();
    if (elements.length === 0) {
      throw new Error('要素を選択してください');
    }

    var element = elements[0];
    var result = {
      elementType: element.getPageElementType().toString(),
      fillColor: null,
      fillAlpha: 1,
      strokeColor: null,
      strokeAlpha: 1,
      strokeWeight: null,
      strokeDashStyle: null,
      textColor: null,
      backgroundColor: null,
      fontSize: null,
      fontFamily: null,
      bold: null,
      italic: null,
      underline: null,
      strikethrough: null
    };

    // 図形の場合
    if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
      var shape = element.asShape();

      // 塗りつぶし色を取得
      var fill = shape.getFill();
      if (fill.getType() === SlidesApp.FillType.SOLID) {
        var solidFill = fill.getSolidFill();
        if (solidFill) {
          result.fillAlpha = solidFill.getAlpha();
          var color = solidFill.getColor();
          if (color.getColorType() === SlidesApp.ColorType.RGB) {
            result.fillColor = rgbColorToHex(color.asRgbColor());
          } else if (color.getColorType() === SlidesApp.ColorType.THEME) {
            // テーマカラーの場合はRGBに変換を試みる
            try {
              result.fillColor = rgbColorToHex(color.asRgbColor());
            } catch(e) {
              result.fillColor = '#CCCCCC'; // フォールバック
            }
          }
        }
      }

      // 枠線色を取得
      var border = shape.getBorder();
      if (border.isVisible()) {
        var lineFill = border.getLineFill();
        if (lineFill.getFillType() === SlidesApp.LineFillType.SOLID) {
          var lineSolidFill = lineFill.getSolidFill();
          result.strokeAlpha = lineSolidFill.getAlpha();
          var lineColor = lineSolidFill.getColor();
          if (lineColor.getColorType() === SlidesApp.ColorType.RGB) {
            result.strokeColor = rgbColorToHex(lineColor.asRgbColor());
          } else if (lineColor.getColorType() === SlidesApp.ColorType.THEME) {
            try {
              result.strokeColor = rgbColorToHex(lineColor.asRgbColor());
            } catch(e) {
              result.strokeColor = '#000000';
            }
          }
        }
        result.strokeWeight = border.getWeight();
        result.strokeDashStyle = border.getDashStyle().toString();
      }

      // テキスト色を取得（図形内のテキスト）
      try {
        var textRange = shape.getText();
        if (textRange && textRange.getLength() > 0) {
          // 最初の文字のスタイルを取得
          var firstCharRange = textRange.getRange(0, 1);
          var textStyle = firstCharRange.getTextStyle();
          var fgColor = textStyle.getForegroundColor();
          if (fgColor) {
            try {
              var colorType = fgColor.getColorType();
              if (colorType === SlidesApp.ColorType.RGB) {
                result.textColor = rgbColorToHex(fgColor.asRgbColor());
              } else if (colorType === SlidesApp.ColorType.THEME) {
                // テーマカラーの場合、RGBに解決を試みる
                var themeColor = fgColor.asThemeColor();
                var presentation = SlidesApp.getActivePresentation();
                var scheme = presentation.getMasters()[0].getColorScheme();
                var resolvedColor = scheme.getConcreteColor(themeColor.getThemeColorType());
                result.textColor = rgbColorToHex(resolvedColor.asRgbColor());
              }
            } catch(e) {
              // フォールバック: 直接RGBとして取得を試みる
              try {
                result.textColor = rgbColorToHex(fgColor.asRgbColor());
              } catch(e2) {}
            }
          }
          // 背景色（ハイライト）を取得
          var bgColor = textStyle.getBackgroundColor();
          if (bgColor) {
            try {
              var bgColorType = bgColor.getColorType();
              if (bgColorType === SlidesApp.ColorType.RGB) {
                result.backgroundColor = rgbColorToHex(bgColor.asRgbColor());
              } else if (bgColorType === SlidesApp.ColorType.THEME) {
                var bgThemeColor = bgColor.asThemeColor();
                var bgPresentation = SlidesApp.getActivePresentation();
                var bgScheme = bgPresentation.getMasters()[0].getColorScheme();
                var bgResolvedColor = bgScheme.getConcreteColor(bgThemeColor.getThemeColorType());
                result.backgroundColor = rgbColorToHex(bgResolvedColor.asRgbColor());
              }
            } catch(e) {
              try {
                result.backgroundColor = rgbColorToHex(bgColor.asRgbColor());
              } catch(e2) {}
            }
          }
          // フォント情報を取得
          result.fontSize = textStyle.getFontSize();
          result.fontFamily = textStyle.getFontFamily();
          result.bold = textStyle.isBold();
          result.italic = textStyle.isItalic();
          result.underline = textStyle.isUnderline();
          result.strikethrough = textStyle.isStrikethrough();

          // 行間・段落間隔を取得
          var paragraphs = textRange.getParagraphs();
          if (paragraphs.length > 0) {
            var paragraphStyle = paragraphs[0].getRange().getParagraphStyle();
            result.lineSpacing = paragraphStyle.getLineSpacing() / 100; // 100が1行なので割る
            result.spaceBefore = paragraphStyle.getSpaceAbove();
            result.spaceAfter = paragraphStyle.getSpaceBelow();
          }
        }
      } catch(e) {}
    }
    // 線の場合
    else if (element.getPageElementType() === SlidesApp.PageElementType.LINE) {
      var line = element.asLine();
      var lineFill = line.getLineFill();
      if (lineFill.getFillType() === SlidesApp.LineFillType.SOLID) {
        var lineSolidFill = lineFill.getSolidFill();
        result.strokeAlpha = lineSolidFill.getAlpha();
        var lineColor = lineSolidFill.getColor();
        if (lineColor.getColorType() === SlidesApp.ColorType.RGB) {
          result.strokeColor = rgbColorToHex(lineColor.asRgbColor());
        }
      }
      result.strokeWeight = line.getWeight();
      result.strokeDashStyle = line.getDashStyle().toString();
    }
    // 画像の場合
    else if (element.getPageElementType() === SlidesApp.PageElementType.IMAGE) {
      var image = element.asImage();

      // 枠線情報を取得
      var border = image.getBorder();
      if (border.isVisible()) {
        var lineFill = border.getLineFill();
        if (lineFill.getFillType() === SlidesApp.LineFillType.SOLID) {
          var lineSolidFill = lineFill.getSolidFill();
          result.strokeAlpha = lineSolidFill.getAlpha();
          var lineColor = lineSolidFill.getColor();
          if (lineColor.getColorType() === SlidesApp.ColorType.RGB) {
            result.strokeColor = rgbColorToHex(lineColor.asRgbColor());
          } else if (lineColor.getColorType() === SlidesApp.ColorType.THEME) {
            try {
              result.strokeColor = rgbColorToHex(lineColor.asRgbColor());
            } catch(e) {
              result.strokeColor = '#000000';
            }
          }
        }
        result.strokeWeight = border.getWeight();
        result.strokeDashStyle = border.getDashStyle().toString();
      }
    }

    return result;
  } catch (e) {
    throw new Error(e.message);
  }
}

/**
 * 選択された要素の塗りつぶし色を設定
 * @param {string} hexColor - HEX形式の色（#RRGGBB）またはnull（透明）
 * @param {number} alpha - 透明度（0-1、1が不透明）省略時は1
 */
function setFillColor(hexColor, alpha) {
  try {
    var selection = SlidesApp.getActivePresentation().getSelection();
    var pageElementRange = selection.getPageElementRange();

    if (!pageElementRange) {
      throw new Error('要素を選択してください');
    }

    var elements = pageElementRange.getPageElements();
    if (elements.length === 0) {
      throw new Error('要素を選択してください');
    }

    // alphaのデフォルト値は1（不透明）
    if (alpha === undefined || alpha === null) {
      alpha = 1;
    }

    var count = 0;
    for (var i = 0; i < elements.length; i++) {
      var element = elements[i];

      if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
        var shape = element.asShape();
        if (hexColor === null) {
          shape.getFill().setTransparent();
        } else {
          shape.getFill().setSolidFill(hexColor, alpha);
        }
        count++;
      }
    }

    if (count === 0) {
      throw new Error('塗りつぶし可能な要素がありません');
    }

    return '塗りつぶし色を設定しました';
  } catch (e) {
    throw new Error(e.message);
  }
}

/**
 * 選択された要素の枠線色を設定
 * @param {string} hexColor - HEX形式の色（#RRGGBB）またはnull（透明）
 * @param {number} alpha - 透明度（0-1、1が不透明）省略時は1
 */
function setStrokeColor(hexColor, alpha) {
  try {
    var selection = SlidesApp.getActivePresentation().getSelection();
    var pageElementRange = selection.getPageElementRange();

    if (!pageElementRange) {
      throw new Error('要素を選択してください');
    }

    var elements = pageElementRange.getPageElements();
    if (elements.length === 0) {
      throw new Error('要素を選択してください');
    }

    // alphaのデフォルト値は1（不透明）
    if (alpha === undefined || alpha === null) {
      alpha = 1;
    }

    var count = 0;
    for (var i = 0; i < elements.length; i++) {
      var element = elements[i];
      var elementType = element.getPageElementType();

      if (elementType === SlidesApp.PageElementType.SHAPE) {
        var shape = element.asShape();
        var border = shape.getBorder();
        if (hexColor === null) {
          border.setTransparent();
        } else {
          border.getLineFill().setSolidFill(hexColor, alpha);
        }
        count++;
      } else if (elementType === SlidesApp.PageElementType.LINE) {
        var line = element.asLine();
        if (hexColor !== null) {
          line.getLineFill().setSolidFill(hexColor, alpha);
        }
        count++;
      } else if (elementType === SlidesApp.PageElementType.IMAGE) {
        var image = element.asImage();
        var border = image.getBorder();
        if (hexColor === null) {
          border.setTransparent();
        } else {
          border.getLineFill().setSolidFill(hexColor, alpha);
        }
        count++;
      }
    }

    if (count === 0) {
      throw new Error('枠線を設定可能な要素がありません');
    }

    return '枠線色を設定しました';
  } catch (e) {
    throw new Error(e.message);
  }
}

/**
 * 選択された要素の枠線スタイルを設定
 * @param {number} weight - 線の太さ（ポイント）
 * @param {string} dashStyle - 線のスタイル（SOLID, DOT, DASH, DASH_DOT, LONG_DASH, LONG_DASH_DOT）
 */
function setStrokeStyle(weight, dashStyle) {
  try {
    var selection = SlidesApp.getActivePresentation().getSelection();
    var pageElementRange = selection.getPageElementRange();

    if (!pageElementRange) {
      throw new Error('要素を選択してください');
    }

    var elements = pageElementRange.getPageElements();
    if (elements.length === 0) {
      throw new Error('要素を選択してください');
    }

    var dashStyleEnum = SlidesApp.DashStyle[dashStyle] || SlidesApp.DashStyle.SOLID;

    var count = 0;
    for (var i = 0; i < elements.length; i++) {
      var element = elements[i];
      var elementType = element.getPageElementType();

      if (elementType === SlidesApp.PageElementType.SHAPE) {
        var shape = element.asShape();
        var border = shape.getBorder();
        if (weight !== null) border.setWeight(weight);
        if (dashStyle !== null) border.setDashStyle(dashStyleEnum);
        count++;
      } else if (elementType === SlidesApp.PageElementType.LINE) {
        var line = element.asLine();
        if (weight !== null) line.setWeight(weight);
        if (dashStyle !== null) line.setDashStyle(dashStyleEnum);
        count++;
      } else if (elementType === SlidesApp.PageElementType.IMAGE) {
        var image = element.asImage();
        var border = image.getBorder();
        if (weight !== null) border.setWeight(weight);
        if (dashStyle !== null) border.setDashStyle(dashStyleEnum);
        count++;
      }
    }

    if (count === 0) {
      throw new Error('枠線を設定可能な要素がありません');
    }

    return '枠線スタイルを設定しました';
  } catch (e) {
    throw new Error(e.message);
  }
}

/**
 * RgbColorをHEX文字列に変換
 * @param {RgbColor} rgbColor
 * @returns {string} HEX形式（#RRGGBB）
 */
function rgbColorToHex(rgbColor) {
  // asHexString()が使える場合はそれを使用（#RRGGBB形式で返る）
  if (rgbColor.asHexString) {
    var hex = rgbColor.asHexString();
    // #RRGGBBAA形式の場合は#RRGGBBに切り詰める
    if (hex && hex.length > 7) {
      hex = hex.substring(0, 7);
    }
    return hex.toUpperCase();
  }
  // フォールバック: 手動でRGB値から変換
  var r = Math.round(rgbColor.getRed() * 255);
  var g = Math.round(rgbColor.getGreen() * 255);
  var b = Math.round(rgbColor.getBlue() * 255);
  return '#' + ((1 << 24) + (r << 16) + (g << 8) + b).toString(16).slice(1).toUpperCase();
}

/**
 * 選択要素のタイプを確認（カラーパネル用の軽量チェック）
 * @returns {Object} 要素タイプと色設定可否
 */
function checkElementForColor() {
  try {
    var selection = SlidesApp.getActivePresentation().getSelection();
    var pageElementRange = selection.getPageElementRange();

    if (!pageElementRange) return null;

    var elements = pageElementRange.getPageElements();
    if (elements.length === 0) return null;

    var element = elements[0];
    var type = element.getPageElementType();

    return {
      type: type.toString(),
      canFill: type === SlidesApp.PageElementType.SHAPE,
      canStroke: type === SlidesApp.PageElementType.SHAPE || type === SlidesApp.PageElementType.LINE,
      count: elements.length
    };
  } catch (e) {
    return null;
  }
}

// =================================================================
// サイズ変更機能
// =================================================================

/**
 * 要素の幅を設定
 * @param {number} width - 幅（ポイント）
 * @param {boolean} keepRatio - 縦横比を維持するか
 */
function setElementWidth(width, keepRatio) {
  try {
    var selection = SlidesApp.getActivePresentation().getSelection();
    var pageElementRange = selection.getPageElementRange();
    if (!pageElementRange) throw new Error('要素を選択してください');

    var elements = pageElementRange.getPageElements();
    for (var i = 0; i < elements.length; i++) {
      var element = elements[i];
      if (keepRatio) {
        var ratio = element.getHeight() / element.getWidth();
        element.setWidth(width);
        element.setHeight(width * ratio);
      } else {
        element.setWidth(width);
      }
    }
    return '幅を設定しました';
  } catch (e) {
    throw new Error(e.message);
  }
}

/**
 * 要素の高さを設定
 * @param {number} height - 高さ（ポイント）
 * @param {boolean} keepRatio - 縦横比を維持するか
 */
function setElementHeight(height, keepRatio) {
  try {
    var selection = SlidesApp.getActivePresentation().getSelection();
    var pageElementRange = selection.getPageElementRange();
    if (!pageElementRange) throw new Error('要素を選択してください');

    var elements = pageElementRange.getPageElements();
    for (var i = 0; i < elements.length; i++) {
      var element = elements[i];
      if (keepRatio) {
        var ratio = element.getWidth() / element.getHeight();
        element.setHeight(height);
        element.setWidth(height * ratio);
      } else {
        element.setHeight(height);
      }
    }
    return '高さを設定しました';
  } catch (e) {
    throw new Error(e.message);
  }
}

/**
 * 基準要素を取得するヘルパー関数
 * @param {Array} elements - 要素の配列
 * @param {string} referenceType - 基準タイプ (first, last, largest, smallest)
 * @return {object} { element: 基準要素, index: インデックス }
 */
function getReferenceElement(elements, referenceType) {
  var referenceElement = null;
  var referenceIndex = 0;

  switch(referenceType) {
    case 'first':
      referenceElement = elements[0];
      referenceIndex = 0;
      break;
    case 'last':
      referenceElement = elements[elements.length - 1];
      referenceIndex = elements.length - 1;
      break;
    case 'largest':
      var maxArea = 0;
      for (var i = 0; i < elements.length; i++) {
        var area = elements[i].getWidth() * elements[i].getHeight();
        if (area > maxArea) {
          maxArea = area;
          referenceElement = elements[i];
          referenceIndex = i;
        }
      }
      break;
    case 'smallest':
      var minArea = Infinity;
      for (var i = 0; i < elements.length; i++) {
        var area = elements[i].getWidth() * elements[i].getHeight();
        if (area < minArea) {
          minArea = area;
          referenceElement = elements[i];
          referenceIndex = i;
        }
      }
      break;
    default:
      referenceElement = elements[0];
      referenceIndex = 0;
  }

  return { element: referenceElement, index: referenceIndex };
}

/**
 * 選択要素の幅を揃える
 * @param {string} referenceType - 基準タイプ
 */
function matchElementsWidth(referenceType) {
  try {
    var selection = SlidesApp.getActivePresentation().getSelection();
    var pageElementRange = selection.getPageElementRange();
    if (!pageElementRange) throw new Error('要素を選択してください');

    var elements = pageElementRange.getPageElements();
    if (elements.length < 2) throw new Error('2つ以上の要素を選択してください');

    var ref = getReferenceElement(elements, referenceType);
    var targetWidth = ref.element.getWidth();

    for (var i = 0; i < elements.length; i++) {
      if (i !== ref.index) {
        elements[i].setWidth(targetWidth);
      }
    }

    var refLabels = { 'first': '最初の要素', 'last': '最後の要素', 'largest': '最大の要素', 'smallest': '最小の要素' };
    return '幅を揃えました（' + (refLabels[referenceType] || '基準') + '）';
  } catch (e) {
    throw new Error(e.message);
  }
}

/**
 * 選択要素の高さを揃える
 * @param {string} referenceType - 基準タイプ
 */
function matchElementsHeight(referenceType) {
  try {
    var selection = SlidesApp.getActivePresentation().getSelection();
    var pageElementRange = selection.getPageElementRange();
    if (!pageElementRange) throw new Error('要素を選択してください');

    var elements = pageElementRange.getPageElements();
    if (elements.length < 2) throw new Error('2つ以上の要素を選択してください');

    var ref = getReferenceElement(elements, referenceType);
    var targetHeight = ref.element.getHeight();

    for (var i = 0; i < elements.length; i++) {
      if (i !== ref.index) {
        elements[i].setHeight(targetHeight);
      }
    }

    var refLabels = { 'first': '最初の要素', 'last': '最後の要素', 'largest': '最大の要素', 'smallest': '最小の要素' };
    return '高さを揃えました（' + (refLabels[referenceType] || '基準') + '）';
  } catch (e) {
    throw new Error(e.message);
  }
}

/**
 * 選択要素のサイズを揃える
 * @param {string} referenceType - 基準タイプ
 */
function matchElementsSize(referenceType) {
  try {
    var selection = SlidesApp.getActivePresentation().getSelection();
    var pageElementRange = selection.getPageElementRange();
    if (!pageElementRange) throw new Error('要素を選択してください');

    var elements = pageElementRange.getPageElements();
    if (elements.length < 2) throw new Error('2つ以上の要素を選択してください');

    var ref = getReferenceElement(elements, referenceType);
    var targetWidth = ref.element.getWidth();
    var targetHeight = ref.element.getHeight();

    for (var i = 0; i < elements.length; i++) {
      if (i !== ref.index) {
        elements[i].setWidth(targetWidth);
        elements[i].setHeight(targetHeight);
      }
    }

    var refLabels = { 'first': '最初の要素', 'last': '最後の要素', 'largest': '最大の要素', 'smallest': '最小の要素' };
    return 'サイズを揃えました（' + (refLabels[referenceType] || '基準') + '）';
  } catch (e) {
    throw new Error(e.message);
  }
}

/**
 * 要素を指定のアスペクト比に変形
 * @param {number} widthRatio - 幅の比率
 * @param {number} heightRatio - 高さの比率
 */
function setElementAspectRatio(widthRatio, heightRatio) {
  try {
    var selection = SlidesApp.getActivePresentation().getSelection();
    var pageElementRange = selection.getPageElementRange();
    if (!pageElementRange) throw new Error('要素を選択してください');

    var elements = pageElementRange.getPageElements();
    for (var i = 0; i < elements.length; i++) {
      var element = elements[i];
      var currentWidth = element.getWidth();
      var currentHeight = element.getHeight();
      var currentArea = currentWidth * currentHeight;

      // 面積を維持しながらアスペクト比を変更
      var newWidth = Math.sqrt(currentArea * widthRatio / heightRatio);
      var newHeight = Math.sqrt(currentArea * heightRatio / widthRatio);

      element.setWidth(newWidth);
      element.setHeight(newHeight);
    }
    return 'アスペクト比を変更しました';
  } catch (e) {
    throw new Error(e.message);
  }
}

// =================================================================
// 間隔調整機能
// =================================================================

/**
 * 水平方向の間隔を設定
 * @param {number} spacing - 間隔（ポイント）
 */
function setHorizontalSpacing(spacing) {
  try {
    var selection = SlidesApp.getActivePresentation().getSelection();
    var pageElementRange = selection.getPageElementRange();
    if (!pageElementRange) throw new Error('要素を選択してください');

    var elements = pageElementRange.getPageElements();
    if (elements.length < 2) throw new Error('2つ以上の要素を選択してください');

    // 左端順にソート
    var sortedElements = [];
    for (var i = 0; i < elements.length; i++) {
      sortedElements.push({
        element: elements[i],
        left: elements[i].getLeft(),
        width: elements[i].getWidth()
      });
    }
    sortedElements.sort(function(a, b) { return a.left - b.left; });

    // 最初の要素の位置から間隔を設定
    var currentLeft = sortedElements[0].left + sortedElements[0].width + spacing;
    for (var i = 1; i < sortedElements.length; i++) {
      sortedElements[i].element.setLeft(currentLeft);
      currentLeft += sortedElements[i].width + spacing;
    }

    return '水平間隔を設定しました';
  } catch (e) {
    throw new Error(e.message);
  }
}

/**
 * 垂直方向の間隔を設定
 * @param {number} spacing - 間隔（ポイント）
 */
function setVerticalSpacing(spacing) {
  try {
    var selection = SlidesApp.getActivePresentation().getSelection();
    var pageElementRange = selection.getPageElementRange();
    if (!pageElementRange) throw new Error('要素を選択してください');

    var elements = pageElementRange.getPageElements();
    if (elements.length < 2) throw new Error('2つ以上の要素を選択してください');

    // 上端順にソート
    var sortedElements = [];
    for (var i = 0; i < elements.length; i++) {
      sortedElements.push({
        element: elements[i],
        top: elements[i].getTop(),
        height: elements[i].getHeight()
      });
    }
    sortedElements.sort(function(a, b) { return a.top - b.top; });

    // 最初の要素の位置から間隔を設定
    var currentTop = sortedElements[0].top + sortedElements[0].height + spacing;
    for (var i = 1; i < sortedElements.length; i++) {
      sortedElements[i].element.setTop(currentTop);
      currentTop += sortedElements[i].height + spacing;
    }

    return '垂直間隔を設定しました';
  } catch (e) {
    throw new Error(e.message);
  }
}

// =================================================================
// 複製機能
// =================================================================

/**
 * 水平方向に複製
 * @param {number} count - 複製数
 * @param {number} spacing - 間隔（ポイント）
 */
function duplicateHorizontal(count, spacing) {
  try {
    var selection = SlidesApp.getActivePresentation().getSelection();
    var pageElementRange = selection.getPageElementRange();
    if (!pageElementRange) throw new Error('要素を選択してください');

    var elements = pageElementRange.getPageElements();
    var slide = selection.getCurrentPage().asSlide();

    for (var i = 0; i < elements.length; i++) {
      var element = elements[i];
      var originalLeft = element.getLeft();
      var originalTop = element.getTop();
      var width = element.getWidth();

      for (var j = 1; j <= count; j++) {
        var duplicate = element.duplicate();
        duplicate.setLeft(originalLeft + (width + spacing) * j);
        duplicate.setTop(originalTop);
      }
    }

    return '水平方向に複製しました';
  } catch (e) {
    throw new Error(e.message);
  }
}

/**
 * 垂直方向に複製
 * @param {number} count - 複製数
 * @param {number} spacing - 間隔（ポイント）
 */
function duplicateVertical(count, spacing) {
  try {
    var selection = SlidesApp.getActivePresentation().getSelection();
    var pageElementRange = selection.getPageElementRange();
    if (!pageElementRange) throw new Error('要素を選択してください');

    var elements = pageElementRange.getPageElements();

    for (var i = 0; i < elements.length; i++) {
      var element = elements[i];
      var originalLeft = element.getLeft();
      var originalTop = element.getTop();
      var height = element.getHeight();

      for (var j = 1; j <= count; j++) {
        var duplicate = element.duplicate();
        duplicate.setLeft(originalLeft);
        duplicate.setTop(originalTop + (height + spacing) * j);
      }
    }

    return '垂直方向に複製しました';
  } catch (e) {
    throw new Error(e.message);
  }
}

/**
 * グリッド状に複製
 * @param {number} count - 1行あたりの複製数
 * @param {number} spacing - 間隔（ポイント）
 */
function duplicateGrid(count, spacing) {
  try {
    var selection = SlidesApp.getActivePresentation().getSelection();
    var pageElementRange = selection.getPageElementRange();
    if (!pageElementRange) throw new Error('要素を選択してください');

    var elements = pageElementRange.getPageElements();

    for (var i = 0; i < elements.length; i++) {
      var element = elements[i];
      var originalLeft = element.getLeft();
      var originalTop = element.getTop();
      var width = element.getWidth();
      var height = element.getHeight();

      // 行数は列数と同じにする（正方形グリッド）
      for (var row = 0; row <= count; row++) {
        for (var col = 0; col <= count; col++) {
          if (row === 0 && col === 0) continue; // オリジナルはスキップ
          var duplicate = element.duplicate();
          duplicate.setLeft(originalLeft + (width + spacing) * col);
          duplicate.setTop(originalTop + (height + spacing) * row);
        }
      }
    }

    return 'グリッド状に複製しました';
  } catch (e) {
    throw new Error(e.message);
  }
}

// =================================================================
// 回転・反転機能
// =================================================================

/**
 * 要素を回転
 * @param {number} degrees - 回転角度（度）
 */
function rotateElement(degrees) {
  try {
    var selection = SlidesApp.getActivePresentation().getSelection();
    var pageElementRange = selection.getPageElementRange();
    if (!pageElementRange) throw new Error('要素を選択してください');

    var elements = pageElementRange.getPageElements();
    for (var i = 0; i < elements.length; i++) {
      var element = elements[i];
      var currentRotation = element.getRotation();
      element.setRotation(currentRotation + degrees);
    }

    return '回転しました';
  } catch (e) {
    throw new Error(e.message);
  }
}

/**
 * 水平方向に反転
 */
function flipHorizontal() {
  try {
    var selection = SlidesApp.getActivePresentation().getSelection();
    var pageElementRange = selection.getPageElementRange();
    if (!pageElementRange) throw new Error('要素を選択してください');

    var presentationId = SlidesApp.getActivePresentation().getId();
    var elements = pageElementRange.getPageElements();
    var requests = [];

    for (var i = 0; i < elements.length; i++) {
      var element = elements[i];
      var transform = element.getTransform();

      // 水平反転はscaleXを-1にする
      requests.push({
        updatePageElementTransform: {
          objectId: element.getObjectId(),
          applyMode: 'RELATIVE',
          transform: {
            scaleX: -1,
            scaleY: 1,
            shearX: 0,
            shearY: 0,
            translateX: 0,
            translateY: 0,
            unit: 'PT'
          }
        }
      });
    }

    if (requests.length > 0) {
      Slides.Presentations.batchUpdate({ requests: requests }, presentationId);
    }

    return '水平方向に反転しました';
  } catch (e) {
    throw new Error(e.message);
  }
}

/**
 * 垂直方向に反転
 */
function flipVertical() {
  try {
    var selection = SlidesApp.getActivePresentation().getSelection();
    var pageElementRange = selection.getPageElementRange();
    if (!pageElementRange) throw new Error('要素を選択してください');

    var presentationId = SlidesApp.getActivePresentation().getId();
    var elements = pageElementRange.getPageElements();
    var requests = [];

    for (var i = 0; i < elements.length; i++) {
      var element = elements[i];

      // 垂直反転はscaleYを-1にする
      requests.push({
        updatePageElementTransform: {
          objectId: element.getObjectId(),
          applyMode: 'RELATIVE',
          transform: {
            scaleX: 1,
            scaleY: -1,
            shearX: 0,
            shearY: 0,
            translateX: 0,
            translateY: 0,
            unit: 'PT'
          }
        }
      });
    }

    if (requests.length > 0) {
      Slides.Presentations.batchUpdate({ requests: requests }, presentationId);
    }

    return '垂直方向に反転しました';
  } catch (e) {
    throw new Error(e.message);
  }
}

// =================================================================
// 順序変更機能
// =================================================================

/**
 * 最前面へ移動
 */
function bringToFront() {
  try {
    var selection = SlidesApp.getActivePresentation().getSelection();
    var pageElementRange = selection.getPageElementRange();
    if (!pageElementRange) throw new Error('要素を選択してください');

    var elements = pageElementRange.getPageElements();
    for (var i = 0; i < elements.length; i++) {
      elements[i].bringToFront();
    }

    return '最前面に移動しました';
  } catch (e) {
    throw new Error(e.message);
  }
}

/**
 * 1つ前面へ移動
 */
function bringForward() {
  try {
    var selection = SlidesApp.getActivePresentation().getSelection();
    var pageElementRange = selection.getPageElementRange();
    if (!pageElementRange) throw new Error('要素を選択してください');

    var elements = pageElementRange.getPageElements();
    for (var i = 0; i < elements.length; i++) {
      elements[i].bringForward();
    }

    return '1つ前面に移動しました';
  } catch (e) {
    throw new Error(e.message);
  }
}

/**
 * 1つ背面へ移動
 */
function sendBackward() {
  try {
    var selection = SlidesApp.getActivePresentation().getSelection();
    var pageElementRange = selection.getPageElementRange();
    if (!pageElementRange) throw new Error('要素を選択してください');

    var elements = pageElementRange.getPageElements();
    for (var i = 0; i < elements.length; i++) {
      elements[i].sendBackward();
    }

    return '1つ背面に移動しました';
  } catch (e) {
    throw new Error(e.message);
  }
}

/**
 * 最背面へ移動
 */
function sendToBack() {
  try {
    var selection = SlidesApp.getActivePresentation().getSelection();
    var pageElementRange = selection.getPageElementRange();
    if (!pageElementRange) throw new Error('要素を選択してください');

    var elements = pageElementRange.getPageElements();
    for (var i = 0; i < elements.length; i++) {
      elements[i].sendToBack();
    }

    return '最背面に移動しました';
  } catch (e) {
    throw new Error(e.message);
  }
}

// =================================================================
// グループ化機能
// =================================================================

/**
 * 要素をグループ化
 */
function groupElements() {
  try {
    var selection = SlidesApp.getActivePresentation().getSelection();
    var pageElementRange = selection.getPageElementRange();
    if (!pageElementRange) throw new Error('要素を選択してください');

    var elements = pageElementRange.getPageElements();
    if (elements.length < 2) throw new Error('2つ以上の要素を選択してください');

    var slide = selection.getCurrentPage().asSlide();
    slide.group(elements);

    return 'グループ化しました';
  } catch (e) {
    throw new Error(e.message);
  }
}

/**
 * グループを解除
 */
function ungroupElements() {
  try {
    var selection = SlidesApp.getActivePresentation().getSelection();
    var pageElementRange = selection.getPageElementRange();
    if (!pageElementRange) throw new Error('要素を選択してください');

    var elements = pageElementRange.getPageElements();
    for (var i = 0; i < elements.length; i++) {
      var element = elements[i];
      if (element.getPageElementType() === SlidesApp.PageElementType.GROUP) {
        element.asGroup().ungroup();
      }
    }

    return 'グループを解除しました';
  } catch (e) {
    throw new Error(e.message);
  }
}

// =================================================================
// テキスト機能
// =================================================================

/**
 * 選択テキストの取得（ヘルパー関数）
 */
function getSelectedTextRange() {
  var selection = SlidesApp.getActivePresentation().getSelection();
  var selectionType = selection.getSelectionType();

  // テキスト選択の場合
  if (selectionType === SlidesApp.SelectionType.TEXT) {
    return selection.getTextRange();
  }

  // 要素選択の場合、テキストボックス/図形のテキスト全体を対象
  var pageElementRange = selection.getPageElementRange();
  if (pageElementRange) {
    var elements = pageElementRange.getPageElements();
    for (var i = 0; i < elements.length; i++) {
      var element = elements[i];
      if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
        try {
          var shape = element.asShape();
          var textRange = shape.getText();
          if (textRange && textRange.getLength() > 0) {
            return textRange;
          }
        } catch(e) {
          // テキストをサポートしない図形の場合はスキップ
        }
      }
    }
  }

  return null;
}

/**
 * テキストフォントを設定
 * @param {string} fontFamily - フォント名
 */
function setTextFont(fontFamily) {
  try {
    var textRange = getSelectedTextRange();
    if (!textRange) throw new Error('テキストを選択してください');

    textRange.getTextStyle().setFontFamily(fontFamily);
    return 'フォントを設定しました';
  } catch (e) {
    throw new Error(e.message);
  }
}

/**
 * テキストサイズを設定
 * @param {number} size - フォントサイズ（ポイント）
 */
function setTextFontSize(size) {
  try {
    var textRange = getSelectedTextRange();
    if (!textRange) throw new Error('テキストを選択してください');

    textRange.getTextStyle().setFontSize(size);
    return 'フォントサイズを設定しました';
  } catch (e) {
    throw new Error(e.message);
  }
}

/**
 * 太字を設定
 * @param {boolean} bold - 太字にするか
 */
function setTextBold(bold) {
  try {
    var textRange = getSelectedTextRange();
    if (!textRange) throw new Error('テキストを選択してください');

    textRange.getTextStyle().setBold(bold);
    return '太字を設定しました';
  } catch (e) {
    throw new Error(e.message);
  }
}

/**
 * 斜体を設定
 * @param {boolean} italic - 斜体にするか
 */
function setTextItalic(italic) {
  try {
    var textRange = getSelectedTextRange();
    if (!textRange) throw new Error('テキストを選択してください');

    textRange.getTextStyle().setItalic(italic);
    return '斜体を設定しました';
  } catch (e) {
    throw new Error(e.message);
  }
}

/**
 * 下線を設定
 * @param {boolean} underline - 下線を付けるか
 */
function setTextUnderline(underline) {
  try {
    var textRange = getSelectedTextRange();
    if (!textRange) throw new Error('テキストを選択してください');

    textRange.getTextStyle().setUnderline(underline);
    return '下線を設定しました';
  } catch (e) {
    throw new Error(e.message);
  }
}

/**
 * 取り消し線を設定
 * @param {boolean} strikethrough - 取り消し線を付けるか
 */
function setTextStrikethrough(strikethrough) {
  try {
    var textRange = getSelectedTextRange();
    if (!textRange) throw new Error('テキストを選択してください');

    textRange.getTextStyle().setStrikethrough(strikethrough);
    return '取り消し線を設定しました';
  } catch (e) {
    throw new Error(e.message);
  }
}

/**
 * 行間を設定
 * @param {number} spacing - 行間倍率（1, 1.15, 1.5, 2など）
 */
function setLineSpacing(spacing) {
  try {
    var selection = SlidesApp.getActivePresentation().getSelection();
    var selectionType = selection.getSelectionType();

    var shape;
    if (selectionType === SlidesApp.SelectionType.TEXT) {
      shape = selection.getPageElementRange().getPageElements()[0].asShape();
    } else if (selectionType === SlidesApp.SelectionType.PAGE_ELEMENT) {
      var elements = selection.getPageElementRange().getPageElements();
      if (elements.length === 0) throw new Error('要素を選択してください');
      var element = elements[0];
      if (element.getPageElementType() !== SlidesApp.PageElementType.SHAPE) {
        throw new Error('テキストボックスを選択してください');
      }
      shape = element.asShape();
    } else {
      throw new Error('テキストボックスを選択してください');
    }

    var textRange = shape.getText();
    var paragraphs = textRange.getParagraphs();

    for (var i = 0; i < paragraphs.length; i++) {
      var paragraph = paragraphs[i];
      var style = paragraph.getRange().getParagraphStyle();
      style.setLineSpacing(spacing * 100); // Google Slidesでは100が1行
    }

    return '行間を ' + spacing + ' に設定しました';
  } catch (e) {
    throw new Error(e.message);
  }
}

/**
 * 段落の前後の間隔を設定
 * @param {number} before - 段落前のスペース（ポイント）
 * @param {number} after - 段落後のスペース（ポイント）
 */
function setParagraphSpacing(before, after) {
  try {
    var selection = SlidesApp.getActivePresentation().getSelection();
    var selectionType = selection.getSelectionType();

    var shape;
    if (selectionType === SlidesApp.SelectionType.TEXT) {
      shape = selection.getPageElementRange().getPageElements()[0].asShape();
    } else if (selectionType === SlidesApp.SelectionType.PAGE_ELEMENT) {
      var elements = selection.getPageElementRange().getPageElements();
      if (elements.length === 0) throw new Error('要素を選択してください');
      var element = elements[0];
      if (element.getPageElementType() !== SlidesApp.PageElementType.SHAPE) {
        throw new Error('テキストボックスを選択してください');
      }
      shape = element.asShape();
    } else {
      throw new Error('テキストボックスを選択してください');
    }

    var textRange = shape.getText();
    var paragraphs = textRange.getParagraphs();

    for (var i = 0; i < paragraphs.length; i++) {
      var paragraph = paragraphs[i];
      var style = paragraph.getRange().getParagraphStyle();
      style.setSpaceAbove(before);
      style.setSpaceBelow(after);
    }

    return '段落間隔を設定しました';
  } catch (e) {
    throw new Error(e.message);
  }
}

/**
 * テキスト色を設定
 * @param {string} hexColor - HEX形式の色
 */
function setTextColor(hexColor) {
  try {
    var textRange = getSelectedTextRange();
    if (!textRange) throw new Error('テキストを選択してください');

    textRange.getTextStyle().setForegroundColor(hexColor);
    return 'テキスト色を設定しました';
  } catch (e) {
    throw new Error(e.message);
  }
}

/**
 * テキスト背景色（ハイライト）を設定
 * @param {string} hexColor - HEX形式の色またはnull（透明）
 */
function setTextBackgroundColor(hexColor) {
  try {
    var textRange = getSelectedTextRange();
    if (!textRange) throw new Error('テキストを選択してください');

    if (hexColor === null) {
      textRange.getTextStyle().setBackgroundColorTransparent();
    } else {
      textRange.getTextStyle().setBackgroundColor(hexColor);
    }
    return 'テキスト背景色を設定しました';
  } catch (e) {
    throw new Error(e.message);
  }
}

// =================================================================
// ガイド機能
// =================================================================

// =================================================================
// 類似オブジェクト検索・選択機能
// =================================================================

/**
 * 選択中のオブジェクトの属性を取得
 * @returns {Object} オブジェクトの属性情報
 */
function getSelectedElementAttributes() {
  try {
    var selection = SlidesApp.getActivePresentation().getSelection();
    var pageElementRange = selection.getPageElementRange();

    if (!pageElementRange) {
      throw new Error('オブジェクトを選択してください');
    }

    var elements = pageElementRange.getPageElements();
    if (elements.length === 0) {
      throw new Error('オブジェクトを選択してください');
    }

    var element = elements[0];
    var attrs = {
      elementType: element.getPageElementType().toString(),
      fillColor: null,
      strokeColor: null,
      strokeWeight: null,
      strokeDashStyle: null,
      textColor: null,
      fontFamily: null,
      fontSize: null,
      bold: null,
      italic: null
    };

    // 図形の場合
    if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
      var shape = element.asShape();

      // 塗りつぶし色
      var fill = shape.getFill();
      if (fill.getType() === SlidesApp.FillType.SOLID) {
        var solidFill = fill.getSolidFill();
        if (solidFill) {
          try {
            attrs.fillColor = rgbColorToHex(solidFill.getColor().asRgbColor());
          } catch(e) {}
        }
      }

      // 枠線
      var border = shape.getBorder();
      if (border.isVisible()) {
        var lineFill = border.getLineFill();
        if (lineFill.getFillType() === SlidesApp.LineFillType.SOLID) {
          try {
            attrs.strokeColor = rgbColorToHex(lineFill.getSolidFill().getColor().asRgbColor());
          } catch(e) {}
        }
        attrs.strokeWeight = border.getWeight();
        attrs.strokeDashStyle = border.getDashStyle().toString();
      }

      // テキスト属性
      try {
        var textRange = shape.getText();
        if (textRange && textRange.getLength() > 0) {
          var firstChar = textRange.getRange(0, 1);
          var textStyle = firstChar.getTextStyle();
          // 文字色を取得
          try {
            var fgColor = textStyle.getForegroundColor();
            if (fgColor) {
              var colorType = fgColor.getColorType();
              if (colorType === SlidesApp.ColorType.RGB) {
                attrs.textColor = rgbColorToHex(fgColor.asRgbColor());
              } else if (colorType === SlidesApp.ColorType.THEME) {
                // テーマカラーの場合もRGBに変換を試みる
                try {
                  attrs.textColor = rgbColorToHex(fgColor.asRgbColor());
                } catch(e2) {
                  attrs.textColor = '#000000'; // フォールバック
                }
              }
            }
          } catch(e) {}
          attrs.fontFamily = textStyle.getFontFamily();
          attrs.fontSize = textStyle.getFontSize();
          attrs.bold = textStyle.isBold();
          attrs.italic = textStyle.isItalic();
        }
      } catch(e) {}
    }
    // 線の場合
    else if (element.getPageElementType() === SlidesApp.PageElementType.LINE) {
      var line = element.asLine();
      var lineFill = line.getLineFill();
      if (lineFill.getFillType() === SlidesApp.LineFillType.SOLID) {
        try {
          attrs.strokeColor = rgbColorToHex(lineFill.getSolidFill().getColor().asRgbColor());
        } catch(e) {}
      }
      attrs.strokeWeight = line.getWeight();
      attrs.strokeDashStyle = line.getDashStyle().toString();
    }
    // 画像の場合
    else if (element.getPageElementType() === SlidesApp.PageElementType.IMAGE) {
      var image = element.asImage();
      var border = image.getBorder();
      if (border.isVisible()) {
        var lineFill = border.getLineFill();
        if (lineFill.getFillType() === SlidesApp.LineFillType.SOLID) {
          try {
            attrs.strokeColor = rgbColorToHex(lineFill.getSolidFill().getColor().asRgbColor());
          } catch(e) {}
        }
        attrs.strokeWeight = border.getWeight();
        attrs.strokeDashStyle = border.getDashStyle().toString();
      }
    }

    return attrs;
  } catch (e) {
    throw new Error(e.message);
  }
}

/**
 * 指定した属性に一致するオブジェクトを同一スライド内で検索・選択
 * @param {string} matchType - 検索タイプ (fillColor, strokeColor, textColor, fontFamily, fontSize, strokeWeight)
 * @returns {Object} 検索結果情報
 */
function selectSimilarElements(matchType) {
  try {
    var selection = SlidesApp.getActivePresentation().getSelection();
    var pageElementRange = selection.getPageElementRange();

    if (!pageElementRange) {
      throw new Error('基準となるオブジェクトを選択してください');
    }

    var selectedElements = pageElementRange.getPageElements();
    if (selectedElements.length === 0) {
      throw new Error('基準となるオブジェクトを選択してください');
    }

    // 基準オブジェクトの属性を取得
    var baseElement = selectedElements[0];
    var baseAttrs = getElementAttributes(baseElement);

    // 現在のスライドの全要素を取得
    var slide = selection.getCurrentPage().asSlide();
    var allElements = slide.getPageElements();

    // 条件に一致する要素を検索
    var matchedElements = [];
    var targetValue = null;

    switch(matchType) {
      case 'fillColor':
        targetValue = baseAttrs.fillColor;
        break;
      case 'strokeColor':
        targetValue = baseAttrs.strokeColor;
        break;
      case 'textColor':
        targetValue = baseAttrs.textColor;
        break;
      case 'fontFamily':
        targetValue = baseAttrs.fontFamily;
        break;
      case 'fontSize':
        targetValue = baseAttrs.fontSize;
        break;
      case 'strokeWeight':
        targetValue = baseAttrs.strokeWeight;
        break;
    }

    if (targetValue === null || targetValue === undefined) {
      throw new Error('選択したオブジェクトにこの属性がありません');
    }

    for (var i = 0; i < allElements.length; i++) {
      var elem = allElements[i];
      var elemAttrs = getElementAttributes(elem);
      var matches = false;

      switch(matchType) {
        case 'fillColor':
          matches = elemAttrs.fillColor === targetValue;
          break;
        case 'strokeColor':
          matches = elemAttrs.strokeColor === targetValue;
          break;
        case 'textColor':
          matches = elemAttrs.textColor === targetValue;
          break;
        case 'fontFamily':
          matches = elemAttrs.fontFamily === targetValue;
          break;
        case 'fontSize':
          matches = elemAttrs.fontSize === targetValue;
          break;
        case 'strokeWeight':
          matches = Math.abs((elemAttrs.strokeWeight || 0) - targetValue) < 0.1;
          break;
      }

      if (matches) {
        matchedElements.push(elem);
      }
    }

    if (matchedElements.length === 0) {
      return { count: 0, message: '一致するオブジェクトが見つかりませんでした' };
    }

    // 一致する要素を選択
    slide.selectAsCurrentPage();
    for (var j = 0; j < matchedElements.length; j++) {
      if (j === 0) {
        matchedElements[j].select();
      } else {
        matchedElements[j].select(false); // 既存の選択に追加
      }
    }

    var typeLabels = {
      'fillColor': '塗りつぶし色',
      'strokeColor': '枠線色',
      'textColor': '文字色',
      'fontFamily': 'フォント',
      'fontSize': 'フォントサイズ',
      'strokeWeight': '枠線の太さ'
    };

    return {
      count: matchedElements.length,
      message: typeLabels[matchType] + 'が同じオブジェクトを ' + matchedElements.length + ' 個選択しました'
    };
  } catch (e) {
    throw new Error(e.message);
  }
}

/**
 * 要素の属性を取得するヘルパー関数
 * @param {PageElement} element - ページ要素
 * @returns {Object} 属性情報
 */
function getElementAttributes(element) {
  var attrs = {
    fillColor: null,
    strokeColor: null,
    strokeWeight: null,
    textColor: null,
    fontFamily: null,
    fontSize: null
  };

  try {
    // 図形の場合
    if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
      var shape = element.asShape();

      // 塗りつぶし色
      var fill = shape.getFill();
      if (fill.getType() === SlidesApp.FillType.SOLID) {
        var solidFill = fill.getSolidFill();
        if (solidFill) {
          try {
            attrs.fillColor = rgbColorToHex(solidFill.getColor().asRgbColor());
          } catch(e) {}
        }
      }

      // 枠線
      var border = shape.getBorder();
      if (border.isVisible()) {
        var lineFill = border.getLineFill();
        if (lineFill.getFillType() === SlidesApp.LineFillType.SOLID) {
          try {
            attrs.strokeColor = rgbColorToHex(lineFill.getSolidFill().getColor().asRgbColor());
          } catch(e) {}
        }
        attrs.strokeWeight = border.getWeight();
      }

      // テキスト属性
      try {
        var textRange = shape.getText();
        if (textRange && textRange.getLength() > 0) {
          var firstChar = textRange.getRange(0, 1);
          var textStyle = firstChar.getTextStyle();
          // 文字色を取得
          try {
            var fgColor = textStyle.getForegroundColor();
            if (fgColor) {
              var colorType = fgColor.getColorType();
              if (colorType === SlidesApp.ColorType.RGB) {
                attrs.textColor = rgbColorToHex(fgColor.asRgbColor());
              } else if (colorType === SlidesApp.ColorType.THEME) {
                try {
                  attrs.textColor = rgbColorToHex(fgColor.asRgbColor());
                } catch(e2) {
                  attrs.textColor = '#000000';
                }
              }
            }
          } catch(e) {}
          attrs.fontFamily = textStyle.getFontFamily();
          attrs.fontSize = textStyle.getFontSize();
        }
      } catch(e) {}
    }
    // 線の場合
    else if (element.getPageElementType() === SlidesApp.PageElementType.LINE) {
      var line = element.asLine();
      var lineFill = line.getLineFill();
      if (lineFill.getFillType() === SlidesApp.LineFillType.SOLID) {
        try {
          attrs.strokeColor = rgbColorToHex(lineFill.getSolidFill().getColor().asRgbColor());
        } catch(e) {}
      }
      attrs.strokeWeight = line.getWeight();
    }
    // 画像の場合
    else if (element.getPageElementType() === SlidesApp.PageElementType.IMAGE) {
      var image = element.asImage();
      var border = image.getBorder();
      if (border.isVisible()) {
        var lineFill = border.getLineFill();
        if (lineFill.getFillType() === SlidesApp.LineFillType.SOLID) {
          try {
            attrs.strokeColor = rgbColorToHex(lineFill.getSolidFill().getColor().asRgbColor());
          } catch(e) {}
        }
        attrs.strokeWeight = border.getWeight();
      }
    }
  } catch(e) {}

  return attrs;
}

// =================================================================
// 書式・スタイルコピー履歴機能
// =================================================================

/**
 * 選択中のオブジェクトのスタイルを取得
 * @returns {Object} スタイル情報
 */
function copyElementStyle() {
  try {
    var selection = SlidesApp.getActivePresentation().getSelection();
    var pageElementRange = selection.getPageElementRange();

    if (!pageElementRange) {
      throw new Error('オブジェクトを選択してください');
    }

    var elements = pageElementRange.getPageElements();
    if (elements.length === 0) {
      throw new Error('オブジェクトを選択してください');
    }

    var element = elements[0];
    var style = {
      timestamp: new Date().toISOString(),
      elementType: element.getPageElementType().toString(),
      fillColor: null,
      fillAlpha: null,
      strokeColor: null,
      strokeWeight: null,
      strokeDashStyle: null,
      textColor: null,
      fontFamily: null,
      fontSize: null,
      bold: null,
      italic: null,
      underline: null,
      strikethrough: null
    };

    // 図形の場合
    if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
      var shape = element.asShape();

      // 塗りつぶし色
      var fill = shape.getFill();
      if (fill.getType() === SlidesApp.FillType.SOLID) {
        var solidFill = fill.getSolidFill();
        if (solidFill) {
          try {
            style.fillColor = rgbColorToHex(solidFill.getColor().asRgbColor());
            style.fillAlpha = solidFill.getAlpha();
          } catch(e) {}
        }
      }

      // 枠線
      var border = shape.getBorder();
      if (border.isVisible()) {
        var lineFill = border.getLineFill();
        if (lineFill.getFillType() === SlidesApp.LineFillType.SOLID) {
          try {
            style.strokeColor = rgbColorToHex(lineFill.getSolidFill().getColor().asRgbColor());
          } catch(e) {}
        }
        style.strokeWeight = border.getWeight();
        style.strokeDashStyle = border.getDashStyle().toString();
      }

      // テキスト属性
      try {
        var textRange = shape.getText();
        if (textRange && textRange.getLength() > 0) {
          var firstChar = textRange.getRange(0, 1);
          var textStyle = firstChar.getTextStyle();
          try {
            var fgColor = textStyle.getForegroundColor();
            if (fgColor) {
              var colorType = fgColor.getColorType();
              if (colorType === SlidesApp.ColorType.RGB) {
                style.textColor = rgbColorToHex(fgColor.asRgbColor());
              } else if (colorType === SlidesApp.ColorType.THEME) {
                var themeColor = fgColor.asThemeColor();
                var presentation = SlidesApp.getActivePresentation();
                var scheme = presentation.getMasters()[0].getColorScheme();
                var resolvedColor = scheme.getConcreteColor(themeColor.getThemeColorType());
                style.textColor = rgbColorToHex(resolvedColor.asRgbColor());
              }
            }
          } catch(e) {
            try {
              style.textColor = rgbColorToHex(fgColor.asRgbColor());
            } catch(e2) {}
          }
          style.fontFamily = textStyle.getFontFamily();
          style.fontSize = textStyle.getFontSize();
          style.bold = textStyle.isBold();
          style.italic = textStyle.isItalic();
          style.underline = textStyle.isUnderline();
          style.strikethrough = textStyle.isStrikethrough();
        }
      } catch(e) {}
    }
    // 線の場合
    else if (element.getPageElementType() === SlidesApp.PageElementType.LINE) {
      var line = element.asLine();
      var lineFill = line.getLineFill();
      if (lineFill.getFillType() === SlidesApp.LineFillType.SOLID) {
        try {
          style.strokeColor = rgbColorToHex(lineFill.getSolidFill().getColor().asRgbColor());
        } catch(e) {}
      }
      style.strokeWeight = line.getWeight();
      style.strokeDashStyle = line.getDashStyle().toString();
    }
    // 画像の場合
    else if (element.getPageElementType() === SlidesApp.PageElementType.IMAGE) {
      var image = element.asImage();
      var border = image.getBorder();
      if (border.isVisible()) {
        var lineFill = border.getLineFill();
        if (lineFill.getFillType() === SlidesApp.LineFillType.SOLID) {
          try {
            style.strokeColor = rgbColorToHex(lineFill.getSolidFill().getColor().asRgbColor());
          } catch(e) {}
        }
        style.strokeWeight = border.getWeight();
        style.strokeDashStyle = border.getDashStyle().toString();
      }
    }

    // プレビュー用のラベルを生成
    var labels = [];
    if (style.fillColor) labels.push('塗:' + style.fillColor);
    if (style.strokeColor) labels.push('線:' + style.strokeColor);
    if (style.textColor) labels.push('字:' + style.textColor);
    if (style.fontFamily) labels.push(style.fontFamily);
    if (style.fontSize) labels.push(style.fontSize + 'pt');
    style.label = labels.length > 0 ? labels.join(' / ') : style.elementType;

    return style;
  } catch (e) {
    throw new Error(e.message);
  }
}

/**
 * 保存されたスタイルを選択中のオブジェクトに適用
 * @param {Object} style - 適用するスタイル
 * @param {Object} options - 適用オプション（どの属性を適用するか）
 * @returns {string} 結果メッセージ
 */
function applyElementStyle(style, options) {
  try {
    var selection = SlidesApp.getActivePresentation().getSelection();
    var pageElementRange = selection.getPageElementRange();

    if (!pageElementRange) {
      throw new Error('オブジェクトを選択してください');
    }

    var elements = pageElementRange.getPageElements();
    if (elements.length === 0) {
      throw new Error('オブジェクトを選択してください');
    }

    var appliedCount = 0;
    var applyAll = !options || options.all;

    for (var i = 0; i < elements.length; i++) {
      var element = elements[i];
      var applied = false;

      // 図形の場合
      if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
        var shape = element.asShape();

        // 塗りつぶし色
        if ((applyAll || options.fill) && style.fillColor) {
          try {
            shape.getFill().setSolidFill(style.fillColor, style.fillAlpha || 1);
            applied = true;
          } catch(e) {}
        }

        // 枠線色
        if ((applyAll || options.stroke) && style.strokeColor) {
          try {
            var border = shape.getBorder();
            border.getLineFill().setSolidFill(style.strokeColor);
            applied = true;
          } catch(e) {}
        }

        // 枠線の太さ
        if ((applyAll || options.strokeWeight) && style.strokeWeight !== null) {
          try {
            shape.getBorder().setWeight(style.strokeWeight);
            applied = true;
          } catch(e) {}
        }

        // 枠線スタイル
        if ((applyAll || options.strokeDash) && style.strokeDashStyle) {
          try {
            shape.getBorder().setDashStyle(SlidesApp.DashStyle[style.strokeDashStyle]);
            applied = true;
          } catch(e) {}
        }

        // テキストスタイル
        var textRange = shape.getText();
        if (textRange && textRange.getLength() > 0) {
          var textStyle = textRange.getTextStyle();

          if ((applyAll || options.textColor) && style.textColor) {
            try {
              textStyle.setForegroundColor(style.textColor);
              applied = true;
            } catch(e) {}
          }

          if ((applyAll || options.font) && style.fontFamily) {
            try {
              textStyle.setFontFamily(style.fontFamily);
              applied = true;
            } catch(e) {}
          }

          if ((applyAll || options.fontSize) && style.fontSize) {
            try {
              textStyle.setFontSize(style.fontSize);
              applied = true;
            } catch(e) {}
          }

          if ((applyAll || options.textStyle) && style.bold !== null) {
            try {
              textStyle.setBold(style.bold);
              textStyle.setItalic(style.italic);
              textStyle.setUnderline(style.underline);
              textStyle.setStrikethrough(style.strikethrough);
              applied = true;
            } catch(e) {}
          }
        }
      }
      // 線の場合
      else if (element.getPageElementType() === SlidesApp.PageElementType.LINE) {
        var line = element.asLine();

        if ((applyAll || options.stroke) && style.strokeColor) {
          try {
            line.getLineFill().setSolidFill(style.strokeColor);
            applied = true;
          } catch(e) {}
        }

        if ((applyAll || options.strokeWeight) && style.strokeWeight !== null) {
          try {
            line.setWeight(style.strokeWeight);
            applied = true;
          } catch(e) {}
        }

        if ((applyAll || options.strokeDash) && style.strokeDashStyle) {
          try {
            line.setDashStyle(SlidesApp.DashStyle[style.strokeDashStyle]);
            applied = true;
          } catch(e) {}
        }
      }
      // 画像の場合
      else if (element.getPageElementType() === SlidesApp.PageElementType.IMAGE) {
        var image = element.asImage();

        if ((applyAll || options.stroke) && style.strokeColor) {
          try {
            image.getBorder().getLineFill().setSolidFill(style.strokeColor);
            applied = true;
          } catch(e) {}
        }

        if ((applyAll || options.strokeWeight) && style.strokeWeight !== null) {
          try {
            image.getBorder().setWeight(style.strokeWeight);
            applied = true;
          } catch(e) {}
        }
      }

      if (applied) appliedCount++;
    }

    return appliedCount + '個のオブジェクトにスタイルを適用しました';
  } catch (e) {
    throw new Error(e.message);
  }
}

// =================================================================
// ガイド機能
// =================================================================

/**
 * ルーラー表示を切り替え（Slides APIでは直接制御不可、UIメッセージのみ）
 */
function toggleRuler(show) {
  // Google Slides APIではルーラー表示の制御はサポートされていません
  // ユーザーに手動で設定するよう案内
  return 'ルーラーは「表示」メニューから手動で切り替えてください';
}

/**
 * ガイド表示を切り替え（Slides APIでは直接制御不可）
 */
function toggleGuides(show) {
  return 'ガイドは「表示」メニューから手動で切り替えてください';
}

/**
 * 垂直ガイドを追加
 */
function addVerticalGuide() {
  try {
    var presentationId = SlidesApp.getActivePresentation().getId();
    var selection = SlidesApp.getActivePresentation().getSelection();
    var pageId = selection.getCurrentPage().getObjectId();

    // スライドの中央に垂直ガイドを追加
    var slideWidth = SlidesApp.getActivePresentation().getPageWidth();

    Slides.Presentations.batchUpdate({
      requests: [{
        createSlideGuide: {
          slideId: pageId,
          guide: {
            orientation: 'VERTICAL',
            position: {
              magnitude: slideWidth / 2,
              unit: 'PT'
            }
          }
        }
      }]
    }, presentationId);

    return '垂直ガイドを追加しました';
  } catch (e) {
    throw new Error(e.message);
  }
}

/**
 * 水平ガイドを追加
 */
function addHorizontalGuide() {
  try {
    var presentationId = SlidesApp.getActivePresentation().getId();
    var selection = SlidesApp.getActivePresentation().getSelection();
    var pageId = selection.getCurrentPage().getObjectId();

    // スライドの中央に水平ガイドを追加
    var slideHeight = SlidesApp.getActivePresentation().getPageHeight();

    Slides.Presentations.batchUpdate({
      requests: [{
        createSlideGuide: {
          slideId: pageId,
          guide: {
            orientation: 'HORIZONTAL',
            position: {
              magnitude: slideHeight / 2,
              unit: 'PT'
            }
          }
        }
      }]
    }, presentationId);

    return '水平ガイドを追加しました';
  } catch (e) {
    throw new Error(e.message);
  }
}

/**
 * ガイド編集画面を開く（UIメッセージのみ）
 */
function editGuides() {
  return 'ガイドはスライド上で直接ドラッグして編集できます';
}

/**
 * すべてのガイドをクリア
 */
function clearGuides() {
  try {
    var presentationId = SlidesApp.getActivePresentation().getId();
    var selection = SlidesApp.getActivePresentation().getSelection();
    var pageId = selection.getCurrentPage().getObjectId();

    // 現在のページのガイド情報を取得
    var page = Slides.Presentations.Pages.get(presentationId, pageId);
    var requests = [];

    if (page.slideProperties && page.slideProperties.guides) {
      var guides = page.slideProperties.guides;
      for (var i = 0; i < guides.length; i++) {
        requests.push({
          deleteSlideGuide: {
            slideId: pageId,
            guideId: guides[i].guideId
          }
        });
      }
    }

    if (requests.length > 0) {
      Slides.Presentations.batchUpdate({ requests: requests }, presentationId);
      return 'ガイドをクリアしました';
    } else {
      return 'クリアするガイドがありません';
    }
  } catch (e) {
    throw new Error(e.message);
  }
}
