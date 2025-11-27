// Code.gs (v41.0 軽量監視・高速化版)

// 多言語メッセージ辞書
var MESSAGES = {
  ja: {
    auth_complete: '承認完了',
    complete: '完了',
    align_complete_slide: '整列完了（スライド基準）',
    align_complete: '整列完了',
    select_multiple: '複数の要素を選択してください',
    ref_slide: 'スライド基準',
    ref_first: '最初の要素基準',
    ref_last: '最後の要素基準',
    distribute_h_complete: '水平方向に均等配置完了',
    distribute_v_complete: '垂直方向に均等配置完了',
    fill_color_set: '塗りつぶし色を設定しました',
    stroke_color_set: '枠線色を設定しました',
    stroke_style_set: '枠線スタイルを設定しました',
    width_set: '幅を設定しました',
    height_set: '高さを設定しました',
    x_set: 'X座標を設定しました',
    y_set: 'Y座標を設定しました',
    width_matched: '幅を揃えました',
    height_matched: '高さを揃えました',
    size_matched: 'サイズを揃えました',
    aspect_changed: 'アスペクト比を変更しました',
    h_spacing_set: '水平間隔を設定しました',
    v_spacing_set: '垂直間隔を設定しました',
    h_duplicated: '水平方向に複製しました',
    v_duplicated: '垂直方向に複製しました',
    grid_duplicated: 'グリッド状に複製しました',
    rotated: '回転しました',
    flipped_h: '水平方向に反転しました',
    flipped_v: '垂直方向に反転しました',
    order_front: '最前面に移動しました',
    order_forward: '1つ前面に移動しました',
    order_backward: '1つ背面に移動しました',
    order_back: '最背面に移動しました',
    grouped: 'グループ化しました',
    ungrouped: 'グループを解除しました',
    font_set: 'フォントを設定しました',
    font_size_set: 'フォントサイズを設定しました',
    bold_set: '太字を設定しました',
    italic_set: '斜体を設定しました',
    underline_set: '下線を設定しました',
    strikethrough_set: '取り消し線を設定しました',
    line_spacing_set: '行間を設定しました',
    paragraph_spacing_set: '段落間隔を設定しました',
    text_color_set: 'テキスト色を設定しました',
    text_bg_set: 'テキスト背景色を設定しました',
    style_applied: '個のオブジェクトにスタイルを適用しました',
    guide_v_added: '垂直ガイドを追加しました',
    guide_h_added: '水平ガイドを追加しました',
    guide_cleared: 'ガイドをクリアしました'
  },
  en: {
    auth_complete: 'Authorization complete',
    complete: 'Complete',
    align_complete_slide: 'Aligned (slide reference)',
    align_complete: 'Aligned',
    select_multiple: 'Please select multiple elements',
    ref_slide: 'slide reference',
    ref_first: 'first element reference',
    ref_last: 'last element reference',
    distribute_h_complete: 'Distributed horizontally',
    distribute_v_complete: 'Distributed vertically',
    fill_color_set: 'Fill color set',
    stroke_color_set: 'Stroke color set',
    stroke_style_set: 'Stroke style set',
    width_set: 'Width set',
    height_set: 'Height set',
    x_set: 'X position set',
    y_set: 'Y position set',
    width_matched: 'Width matched',
    height_matched: 'Height matched',
    size_matched: 'Size matched',
    aspect_changed: 'Aspect ratio changed',
    h_spacing_set: 'Horizontal spacing set',
    v_spacing_set: 'Vertical spacing set',
    h_duplicated: 'Duplicated horizontally',
    v_duplicated: 'Duplicated vertically',
    grid_duplicated: 'Duplicated in grid',
    rotated: 'Rotated',
    flipped_h: 'Flipped horizontally',
    flipped_v: 'Flipped vertically',
    order_front: 'Moved to front',
    order_forward: 'Moved forward',
    order_backward: 'Moved backward',
    order_back: 'Moved to back',
    grouped: 'Grouped',
    ungrouped: 'Ungrouped',
    font_set: 'Font set',
    font_size_set: 'Font size set',
    bold_set: 'Bold set',
    italic_set: 'Italic set',
    underline_set: 'Underline set',
    strikethrough_set: 'Strikethrough set',
    line_spacing_set: 'Line spacing set',
    paragraph_spacing_set: 'Paragraph spacing set',
    text_color_set: 'Text color set',
    text_bg_set: 'Text background color set',
    style_applied: ' objects styled',
    guide_v_added: 'Vertical guide added',
    guide_h_added: 'Horizontal guide added',
    guide_cleared: 'Guides cleared'
  }
};

function msg(key, lang) {
  lang = lang || 'ja';
  return MESSAGES[lang][key] || MESSAGES['ja'][key] || key;
}

// 承認用テスト関数（エディターから実行可能）
function testAuth(lang) {
  var presentation = SlidesApp.getActivePresentation();
  Logger.log('プレゼンテーション名: ' + presentation.getName());
  return lang === 'en' ? 'Authorization complete' : '承認完了';
}

function onOpen() {
  SlidesApp.getUi().createMenu('Designer\'s Palette')
    .addItem('日本語版を開く', 'showSidebarJa')
    .addItem('Open English Version', 'showSidebarEn')
    .addToUi();
}

function showSidebar() {
  showSidebarJa();
}

function showSidebarJa() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Designer\'s Palette')
    .setWidth(340);
  // Inject language parameter
  var content = html.getContent();
  content = content.replace('let currentLang = \'ja\';', 'let currentLang = \'ja\'; localStorage.setItem(\'dp_lang\', \'ja\');');
  html.setContent(content);
  SlidesApp.getUi().showSidebar(html);
}

function showSidebarEn() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Designer\'s Palette')
    .setWidth(340);
  // Inject language parameter
  var content = html.getContent();
  content = content.replace('let currentLang = \'ja\';', 'let currentLang = \'en\'; localStorage.setItem(\'dp_lang\', \'en\');');
  html.setContent(content);
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

/**
 * 選択要素のサイズと位置を取得
 */
function getSelectedElementSizeAndPosition() {
  try {
    var selection = SlidesApp.getActivePresentation().getSelection();
    var pageElementRange = selection.getPageElementRange();
    if (!pageElementRange) return null;

    var elements = pageElementRange.getPageElements();
    if (elements.length === 0) return null;

    var element = elements[0];
    return {
      x: Math.round(element.getLeft()),
      y: Math.round(element.getTop()),
      width: Math.round(element.getWidth()),
      height: Math.round(element.getHeight())
    };
  } catch (e) {
    return null;
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

function getSelectedImage(lang) {
try {
var selection = SlidesApp.getActivePresentation().getSelection()            ;
var pageElementRange = selection.getPageElementRange()                      ;
if (!pageElementRange || pageElementRange.getPageElements().length === 0) {
throw new Error(lang === 'en' ? 'Please select an image.' : '画像を選択してください。');
}
var element = pageElementRange.getPageElements()[0]                         ;
if (element.getPageElementType() !== SlidesApp.PageElementType.IMAGE) {
throw new Error(lang === 'en' ? 'Selected element is not an image.' : '選択された要素は画像ではありません。');
}

var image = element.asImage() ;
var blob = image.getBlob()    ;

if (blob.getBytes().length > 10 * 1024 * 1024) {
throw new Error(lang === 'en' ? 'Image size is too large (10MB or less recommended).' : '画像サイズが大きすぎます（10MB以下推奨）。');
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
console.warn("Failed to get crop info: " + e.message)                                                              ;
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

function replaceImage(base64Data, lang) {
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
return lang === 'en' ? "Complete" : "完了";
} catch (e) {
throw new Error((lang === 'en' ? 'Placement error: ' : '配置エラー: ') + e.message)                ;
}
}

// =================================================================
// 整列機能
// =================================================================

/**
 * 選択された要素を整列する汎用関数
 * @param {string} alignmentType - 整列タイプ (left, center, right, top, middle, bottom)
 * @param {string} referenceType - 基準タイプ (first, last, largest, smallest, slide)
 * @param {string} lang - 言語コード (ja, en)
 */
function alignElements(alignmentType, referenceType, lang) {
  lang = lang || 'ja';
  try {
    var selection = SlidesApp.getActivePresentation().getSelection();
    var pageElementRange = selection.getPageElementRange();
    var slide = selection.getCurrentPage();

    // スライド基準の場合は要素選択不要
    if (referenceType === 'slide') {
      if (!pageElementRange) {
        throw new Error(lang === 'en' ? 'Please select an element' : '要素を選択してください');
      }
      var elements = pageElementRange.getPageElements();
      if (elements.length === 0) {
        throw new Error(lang === 'en' ? 'Please select an element' : '要素を選択してください');
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
      return msg('align_complete_slide', lang);
    }

    if (!pageElementRange) {
      throw new Error(lang === 'en' ? 'Please select an element' : '要素を選択してください');
    }

    var elements = pageElementRange.getPageElements();
    if (elements.length === 0) {
      throw new Error(lang === 'en' ? 'Please select an element' : '要素を選択してください');
    }

    if (elements.length === 1) {
      return msg('select_multiple', lang);
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

    var refLabels = lang === 'en' ? {
      'first': 'first element',
      'last': 'last element',
      'largest': 'largest element',
      'smallest': 'smallest element'
    } : {
      'first': '最初の要素',
      'last': '最後の要素',
      'largest': '最大の要素',
      'smallest': '最小の要素'
    };
    var baseMsg = lang === 'en' ? 'Aligned (' : '整列完了（';
    var suffix = lang === 'en' ? ')' : '）';
    return baseMsg + (refLabels[referenceType] || (lang === 'en' ? 'reference' : '基準')) + suffix;
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
 * @param {string} lang - 言語コード (ja, en)
 */
function distributeHorizontally(lang) {
  lang = lang || 'ja';
  try {
    var selection = SlidesApp.getActivePresentation().getSelection();
    var pageElementRange = selection.getPageElementRange();

    if (!pageElementRange) {
      throw new Error(lang === 'en' ? 'Please select an element' : '要素を選択してください');
    }

    var elements = pageElementRange.getPageElements();
    if (elements.length < 3) {
      throw new Error(lang === 'en' ? 'Select 3 or more elements to distribute' : '均等配置には3つ以上の要素を選択してください');
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

    return msg('distribute_h_complete', lang);
  } catch (e) {
    throw new Error(e.message);
  }
}

/**
 * 垂直方向に均等配置
 * @param {string} lang - 言語コード (ja, en)
 */
function distributeVertically(lang) {
  lang = lang || 'ja';
  try {
    var selection = SlidesApp.getActivePresentation().getSelection();
    var pageElementRange = selection.getPageElementRange();

    if (!pageElementRange) {
      throw new Error(lang === 'en' ? 'Please select an element' : '要素を選択してください');
    }

    var elements = pageElementRange.getPageElements();
    if (elements.length < 3) {
      throw new Error(lang === 'en' ? 'Select 3 or more elements to distribute' : '均等配置には3つ以上の要素を選択してください');
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

    return msg('distribute_v_complete', lang);
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
function getSelectedElementColors(lang) {
  try {
    var selection = SlidesApp.getActivePresentation().getSelection();
    var pageElementRange = selection.getPageElementRange();

    if (!pageElementRange) {
      throw new Error(lang === 'en' ? 'Please select an element' : '要素を選択してください');
    }

    var elements = pageElementRange.getPageElements();
    if (elements.length === 0) {
      throw new Error(lang === 'en' ? 'Please select an element' : '要素を選択してください');
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
      var fillType = fill.getType();
      if (fillType === SlidesApp.FillType.SOLID) {
        var solidFill = fill.getSolidFill();
        if (solidFill) {
          result.fillAlpha = solidFill.getAlpha();
          var color = solidFill.getColor();
          var colorType = color.getColorType();
          if (colorType === SlidesApp.ColorType.RGB) {
            result.fillColor = rgbColorToHex(color.asRgbColor());
          } else if (colorType === SlidesApp.ColorType.THEME) {
            // テーマカラーの場合はカラースキームから解決
            try {
              var themeColor = color.asThemeColor();
              var presentation = SlidesApp.getActivePresentation();
              var scheme = presentation.getMasters()[0].getColorScheme();
              var resolvedColor = scheme.getConcreteColor(themeColor.getThemeColorType());
              result.fillColor = rgbColorToHex(resolvedColor.asRgbColor());
            } catch(e) {
              // フォールバック: 直接RGBとして取得を試みる
              try {
                result.fillColor = rgbColorToHex(color.asRgbColor());
              } catch(e2) {
                result.fillColor = '#CCCCCC';
              }
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
          var lineColorType = lineColor.getColorType();
          if (lineColorType === SlidesApp.ColorType.RGB) {
            result.strokeColor = rgbColorToHex(lineColor.asRgbColor());
          } else if (lineColorType === SlidesApp.ColorType.THEME) {
            // テーマカラーの場合はカラースキームから解決
            try {
              var lineThemeColor = lineColor.asThemeColor();
              var presentation = SlidesApp.getActivePresentation();
              var scheme = presentation.getMasters()[0].getColorScheme();
              var resolvedLineColor = scheme.getConcreteColor(lineThemeColor.getThemeColorType());
              result.strokeColor = rgbColorToHex(resolvedLineColor.asRgbColor());
            } catch(e) {
              try {
                result.strokeColor = rgbColorToHex(lineColor.asRgbColor());
              } catch(e2) {
                result.strokeColor = '#000000';
              }
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
function setFillColor(hexColor, alpha, lang) {
  try {
    var selection = SlidesApp.getActivePresentation().getSelection();
    var pageElementRange = selection.getPageElementRange();

    if (!pageElementRange) {
      throw new Error(lang === 'en' ? 'Please select an element' : '要素を選択してください');
    }

    var elements = pageElementRange.getPageElements();
    if (elements.length === 0) {
      throw new Error(lang === 'en' ? 'Please select an element' : '要素を選択してください');
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
      throw new Error(lang === 'en' ? 'No fillable elements' : '塗りつぶし可能な要素がありません');
    }

    return lang === 'en' ? 'Fill color set' : '塗りつぶし色を設定しました';
  } catch (e) {
    throw new Error(e.message);
  }
}

/**
 * 選択された要素の枠線色を設定
 * @param {string} hexColor - HEX形式の色（#RRGGBB）またはnull（透明）
 * @param {number} alpha - 透明度（0-1、1が不透明）省略時は1
 */
function setStrokeColor(hexColor, alpha, lang) {
  try {
    var selection = SlidesApp.getActivePresentation().getSelection();
    var pageElementRange = selection.getPageElementRange();

    if (!pageElementRange) {
      throw new Error(lang === 'en' ? 'Please select an element' : '要素を選択してください');
    }

    var elements = pageElementRange.getPageElements();
    if (elements.length === 0) {
      throw new Error(lang === 'en' ? 'Please select an element' : '要素を選択してください');
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
      throw new Error(lang === 'en' ? 'No elements with stroke' : '枠線を設定可能な要素がありません');
    }

    return lang === 'en' ? 'Stroke color set' : '枠線色を設定しました';
  } catch (e) {
    throw new Error(e.message);
  }
}

/**
 * 選択された要素の枠線スタイルを設定
 * @param {number} weight - 線の太さ（ポイント）
 * @param {string} dashStyle - 線のスタイル（SOLID, DOT, DASH, DASH_DOT, LONG_DASH, LONG_DASH_DOT）
 */
function setStrokeStyle(weight, dashStyle, lang) {
  try {
    var selection = SlidesApp.getActivePresentation().getSelection();
    var pageElementRange = selection.getPageElementRange();

    if (!pageElementRange) {
      throw new Error(lang === 'en' ? 'Please select an element' : '要素を選択してください');
    }

    var elements = pageElementRange.getPageElements();
    if (elements.length === 0) {
      throw new Error(lang === 'en' ? 'Please select an element' : '要素を選択してください');
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
      throw new Error(lang === 'en' ? 'No elements with stroke' : '枠線を設定可能な要素がありません');
    }

    return lang === 'en' ? 'Stroke style set' : '枠線スタイルを設定しました';
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
function setElementWidth(width, keepRatio, lang) {
  try {
    var selection = SlidesApp.getActivePresentation().getSelection();
    var pageElementRange = selection.getPageElementRange();
    if (!pageElementRange) throw new Error(lang === 'en' ? 'Please select an element' : '要素を選択してください');

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
    return lang === 'en' ? 'Width set' : '幅を設定しました';
  } catch (e) {
    throw new Error(e.message);
  }
}

/**
 * 要素の高さを設定
 * @param {number} height - 高さ（ポイント）
 * @param {boolean} keepRatio - 縦横比を維持するか
 */
function setElementHeight(height, keepRatio, lang) {
  try {
    var selection = SlidesApp.getActivePresentation().getSelection();
    var pageElementRange = selection.getPageElementRange();
    if (!pageElementRange) throw new Error(lang === 'en' ? 'Please select an element' : '要素を選択してください');

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
    return lang === 'en' ? 'Height set' : '高さを設定しました';
  } catch (e) {
    throw new Error(e.message);
  }
}

/**
 * 要素のX座標を設定
 * @param {number} x - X座標（ポイント）
 */
function setElementPositionX(x, lang) {
  try {
    var selection = SlidesApp.getActivePresentation().getSelection();
    var pageElementRange = selection.getPageElementRange();
    if (!pageElementRange) throw new Error(lang === 'en' ? 'Please select an element' : '要素を選択してください');

    var elements = pageElementRange.getPageElements();
    for (var i = 0; i < elements.length; i++) {
      elements[i].setLeft(x);
    }
    return lang === 'en' ? 'X position set' : 'X座標を設定しました';
  } catch (e) {
    throw new Error(e.message);
  }
}

/**
 * 要素のY座標を設定
 * @param {number} y - Y座標（ポイント）
 */
function setElementPositionY(y, lang) {
  try {
    var selection = SlidesApp.getActivePresentation().getSelection();
    var pageElementRange = selection.getPageElementRange();
    if (!pageElementRange) throw new Error(lang === 'en' ? 'Please select an element' : '要素を選択してください');

    var elements = pageElementRange.getPageElements();
    for (var i = 0; i < elements.length; i++) {
      elements[i].setTop(y);
    }
    return lang === 'en' ? 'Y position set' : 'Y座標を設定しました';
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
 * @param {string} lang - 言語コード (ja, en)
 */
function matchElementsWidth(referenceType, lang) {
  lang = lang || 'ja';
  try {
    var selection = SlidesApp.getActivePresentation().getSelection();
    var pageElementRange = selection.getPageElementRange();
    if (!pageElementRange) throw new Error(lang === 'en' ? 'Please select an element' : '要素を選択してください');

    var elements = pageElementRange.getPageElements();
    if (elements.length < 2) throw new Error(lang === 'en' ? 'Select 2 or more elements' : '2つ以上の要素を選択してください');

    var ref = getReferenceElement(elements, referenceType);
    var targetWidth = ref.element.getWidth();

    for (var i = 0; i < elements.length; i++) {
      if (i !== ref.index) {
        elements[i].setWidth(targetWidth);
      }
    }

    var refLabels = lang === 'en' ?
      { 'first': 'first element', 'last': 'last element', 'largest': 'largest element', 'smallest': 'smallest element' } :
      { 'first': '最初の要素', 'last': '最後の要素', 'largest': '最大の要素', 'smallest': '最小の要素' };
    var baseMsg = lang === 'en' ? 'Width matched (' : '幅を揃えました（';
    var suffix = lang === 'en' ? ')' : '）';
    return baseMsg + (refLabels[referenceType] || (lang === 'en' ? 'reference' : '基準')) + suffix;
  } catch (e) {
    throw new Error(e.message);
  }
}

/**
 * 選択要素の高さを揃える
 * @param {string} referenceType - 基準タイプ
 * @param {string} lang - 言語コード (ja, en)
 */
function matchElementsHeight(referenceType, lang) {
  lang = lang || 'ja';
  try {
    var selection = SlidesApp.getActivePresentation().getSelection();
    var pageElementRange = selection.getPageElementRange();
    if (!pageElementRange) throw new Error(lang === 'en' ? 'Please select an element' : '要素を選択してください');

    var elements = pageElementRange.getPageElements();
    if (elements.length < 2) throw new Error(lang === 'en' ? 'Select 2 or more elements' : '2つ以上の要素を選択してください');

    var ref = getReferenceElement(elements, referenceType);
    var targetHeight = ref.element.getHeight();

    for (var i = 0; i < elements.length; i++) {
      if (i !== ref.index) {
        elements[i].setHeight(targetHeight);
      }
    }

    var refLabels = lang === 'en' ?
      { 'first': 'first element', 'last': 'last element', 'largest': 'largest element', 'smallest': 'smallest element' } :
      { 'first': '最初の要素', 'last': '最後の要素', 'largest': '最大の要素', 'smallest': '最小の要素' };
    var baseMsg = lang === 'en' ? 'Height matched (' : '高さを揃えました（';
    var suffix = lang === 'en' ? ')' : '）';
    return baseMsg + (refLabels[referenceType] || (lang === 'en' ? 'reference' : '基準')) + suffix;
  } catch (e) {
    throw new Error(e.message);
  }
}

/**
 * 選択要素のサイズを揃える
 * @param {string} referenceType - 基準タイプ
 * @param {string} lang - 言語コード (ja, en)
 */
function matchElementsSize(referenceType, lang) {
  lang = lang || 'ja';
  try {
    var selection = SlidesApp.getActivePresentation().getSelection();
    var pageElementRange = selection.getPageElementRange();
    if (!pageElementRange) throw new Error(lang === 'en' ? 'Please select an element' : '要素を選択してください');

    var elements = pageElementRange.getPageElements();
    if (elements.length < 2) throw new Error(lang === 'en' ? 'Select 2 or more elements' : '2つ以上の要素を選択してください');

    var ref = getReferenceElement(elements, referenceType);
    var targetWidth = ref.element.getWidth();
    var targetHeight = ref.element.getHeight();

    for (var i = 0; i < elements.length; i++) {
      if (i !== ref.index) {
        elements[i].setWidth(targetWidth);
        elements[i].setHeight(targetHeight);
      }
    }

    var refLabels = lang === 'en' ?
      { 'first': 'first element', 'last': 'last element', 'largest': 'largest element', 'smallest': 'smallest element' } :
      { 'first': '最初の要素', 'last': '最後の要素', 'largest': '最大の要素', 'smallest': '最小の要素' };
    var baseMsg = lang === 'en' ? 'Size matched (' : 'サイズを揃えました（';
    var suffix = lang === 'en' ? ')' : '）';
    return baseMsg + (refLabels[referenceType] || (lang === 'en' ? 'reference' : '基準')) + suffix;
  } catch (e) {
    throw new Error(e.message);
  }
}

/**
 * 要素を指定のアスペクト比に変形
 * @param {number} widthRatio - 幅の比率
 * @param {number} heightRatio - 高さの比率
 */
function setElementAspectRatio(widthRatio, heightRatio, lang) {
  try {
    var selection = SlidesApp.getActivePresentation().getSelection();
    var pageElementRange = selection.getPageElementRange();
    if (!pageElementRange) throw new Error(lang === 'en' ? 'Please select an element' : '要素を選択してください');

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
    return lang === 'en' ? 'Aspect ratio changed' : 'アスペクト比を変更しました';
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
function setHorizontalSpacing(spacing, lang) {
  try {
    var selection = SlidesApp.getActivePresentation().getSelection();
    var pageElementRange = selection.getPageElementRange();
    if (!pageElementRange) throw new Error(lang === 'en' ? 'Please select an element' : '要素を選択してください');

    var elements = pageElementRange.getPageElements();
    if (elements.length < 2) throw new Error(lang === 'en' ? 'Select 2 or more elements' : '2つ以上の要素を選択してください');

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

    return lang === 'en' ? 'Horizontal spacing set' : '水平間隔を設定しました';
  } catch (e) {
    throw new Error(e.message);
  }
}

/**
 * 垂直方向の間隔を設定
 * @param {number} spacing - 間隔（ポイント）
 */
function setVerticalSpacing(spacing, lang) {
  try {
    var selection = SlidesApp.getActivePresentation().getSelection();
    var pageElementRange = selection.getPageElementRange();
    if (!pageElementRange) throw new Error(lang === 'en' ? 'Please select an element' : '要素を選択してください');

    var elements = pageElementRange.getPageElements();
    if (elements.length < 2) throw new Error(lang === 'en' ? 'Select 2 or more elements' : '2つ以上の要素を選択してください');

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

    return lang === 'en' ? 'Vertical spacing set' : '垂直間隔を設定しました';
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
function duplicateHorizontal(count, spacing, lang) {
  try {
    var selection = SlidesApp.getActivePresentation().getSelection();
    var pageElementRange = selection.getPageElementRange();
    if (!pageElementRange) throw new Error(lang === 'en' ? 'Please select an element' : '要素を選択してください');

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

    return lang === 'en' ? 'Duplicated horizontally' : '水平方向に複製しました';
  } catch (e) {
    throw new Error(e.message);
  }
}

/**
 * 垂直方向に複製
 * @param {number} count - 複製数
 * @param {number} spacing - 間隔（ポイント）
 */
function duplicateVertical(count, spacing, lang) {
  try {
    var selection = SlidesApp.getActivePresentation().getSelection();
    var pageElementRange = selection.getPageElementRange();
    if (!pageElementRange) throw new Error(lang === 'en' ? 'Please select an element' : '要素を選択してください');

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

    return lang === 'en' ? 'Duplicated vertically' : '垂直方向に複製しました';
  } catch (e) {
    throw new Error(e.message);
  }
}

/**
 * グリッド状に複製
 * @param {number} count - 1行あたりの複製数
 * @param {number} spacing - 間隔（ポイント）
 */
function duplicateGrid(count, spacing, lang) {
  try {
    var selection = SlidesApp.getActivePresentation().getSelection();
    var pageElementRange = selection.getPageElementRange();
    if (!pageElementRange) throw new Error(lang === 'en' ? 'Please select an element' : '要素を選択してください');

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

    return lang === 'en' ? 'Duplicated in grid' : 'グリッド状に複製しました';
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
function rotateElement(degrees, lang) {
  try {
    var selection = SlidesApp.getActivePresentation().getSelection();
    var pageElementRange = selection.getPageElementRange();
    if (!pageElementRange) throw new Error(lang === 'en' ? 'Please select an element' : '要素を選択してください');

    var elements = pageElementRange.getPageElements();
    for (var i = 0; i < elements.length; i++) {
      var element = elements[i];
      var currentRotation = element.getRotation();
      element.setRotation(currentRotation + degrees);
    }

    return lang === 'en' ? 'Rotated' : '回転しました';
  } catch (e) {
    throw new Error(e.message);
  }
}

/**
 * 水平方向に反転
 */
function flipHorizontal(lang) {
  try {
    var selection = SlidesApp.getActivePresentation().getSelection();
    var pageElementRange = selection.getPageElementRange();
    if (!pageElementRange) throw new Error(lang === 'en' ? 'Please select an element' : '要素を選択してください');

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

    return lang === 'en' ? 'Flipped horizontally' : '水平方向に反転しました';
  } catch (e) {
    throw new Error(e.message);
  }
}

/**
 * 垂直方向に反転
 */
function flipVertical(lang) {
  try {
    var selection = SlidesApp.getActivePresentation().getSelection();
    var pageElementRange = selection.getPageElementRange();
    if (!pageElementRange) throw new Error(lang === 'en' ? 'Please select an element' : '要素を選択してください');

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

    return lang === 'en' ? 'Flipped vertically' : '垂直方向に反転しました';
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
function bringToFront(lang) {
  try {
    var selection = SlidesApp.getActivePresentation().getSelection();
    var pageElementRange = selection.getPageElementRange();
    if (!pageElementRange) throw new Error(lang === 'en' ? 'Please select an element' : '要素を選択してください');

    var elements = pageElementRange.getPageElements();
    for (var i = 0; i < elements.length; i++) {
      elements[i].bringToFront();
    }

    return lang === 'en' ? 'Moved to front' : '最前面に移動しました';
  } catch (e) {
    throw new Error(e.message);
  }
}

/**
 * 1つ前面へ移動
 */
function bringForward(lang) {
  try {
    var selection = SlidesApp.getActivePresentation().getSelection();
    var pageElementRange = selection.getPageElementRange();
    if (!pageElementRange) throw new Error(lang === 'en' ? 'Please select an element' : '要素を選択してください');

    var elements = pageElementRange.getPageElements();
    for (var i = 0; i < elements.length; i++) {
      elements[i].bringForward();
    }

    return lang === 'en' ? 'Moved forward' : '1つ前面に移動しました';
  } catch (e) {
    throw new Error(e.message);
  }
}

/**
 * 1つ背面へ移動
 */
function sendBackward(lang) {
  try {
    var selection = SlidesApp.getActivePresentation().getSelection();
    var pageElementRange = selection.getPageElementRange();
    if (!pageElementRange) throw new Error(lang === 'en' ? 'Please select an element' : '要素を選択してください');

    var elements = pageElementRange.getPageElements();
    for (var i = 0; i < elements.length; i++) {
      elements[i].sendBackward();
    }

    return lang === 'en' ? 'Moved backward' : '1つ背面に移動しました';
  } catch (e) {
    throw new Error(e.message);
  }
}

/**
 * 最背面へ移動
 */
function sendToBack(lang) {
  try {
    var selection = SlidesApp.getActivePresentation().getSelection();
    var pageElementRange = selection.getPageElementRange();
    if (!pageElementRange) throw new Error(lang === 'en' ? 'Please select an element' : '要素を選択してください');

    var elements = pageElementRange.getPageElements();
    for (var i = 0; i < elements.length; i++) {
      elements[i].sendToBack();
    }

    return lang === 'en' ? 'Moved to back' : '最背面に移動しました';
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
function groupElements(lang) {
  try {
    var selection = SlidesApp.getActivePresentation().getSelection();
    var pageElementRange = selection.getPageElementRange();
    if (!pageElementRange) throw new Error(lang === 'en' ? 'Please select an element' : '要素を選択してください');

    var elements = pageElementRange.getPageElements();
    if (elements.length < 2) throw new Error(lang === 'en' ? 'Select 2 or more elements' : '2つ以上の要素を選択してください');

    var slide = selection.getCurrentPage().asSlide();
    slide.group(elements);

    return lang === 'en' ? 'Grouped' : 'グループ化しました';
  } catch (e) {
    throw new Error(e.message);
  }
}

/**
 * グループを解除
 */
function ungroupElements(lang) {
  try {
    var selection = SlidesApp.getActivePresentation().getSelection();
    var pageElementRange = selection.getPageElementRange();
    if (!pageElementRange) throw new Error(lang === 'en' ? 'Please select an element' : '要素を選択してください');

    var elements = pageElementRange.getPageElements();
    for (var i = 0; i < elements.length; i++) {
      var element = elements[i];
      if (element.getPageElementType() === SlidesApp.PageElementType.GROUP) {
        element.asGroup().ungroup();
      }
    }

    return lang === 'en' ? 'Ungrouped' : 'グループを解除しました';
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
function setTextFont(fontFamily, lang) {
  try {
    var textRange = getSelectedTextRange();
    if (!textRange) throw new Error(lang === 'en' ? 'Please select text' : 'テキストを選択してください');

    textRange.getTextStyle().setFontFamily(fontFamily);
    return lang === 'en' ? 'Font set' : 'フォントを設定しました';
  } catch (e) {
    throw new Error(e.message);
  }
}

/**
 * テキストサイズを設定
 * @param {number} size - フォントサイズ（ポイント）
 */
function setTextFontSize(size, lang) {
  try {
    var textRange = getSelectedTextRange();
    if (!textRange) throw new Error(lang === 'en' ? 'Please select text' : 'テキストを選択してください');

    textRange.getTextStyle().setFontSize(size);
    return lang === 'en' ? 'Font size set' : 'フォントサイズを設定しました';
  } catch (e) {
    throw new Error(e.message);
  }
}

/**
 * 太字を設定
 * @param {boolean} bold - 太字にするか
 */
function setTextBold(bold, lang) {
  try {
    var textRange = getSelectedTextRange();
    if (!textRange) throw new Error(lang === 'en' ? 'Please select text' : 'テキストを選択してください');

    textRange.getTextStyle().setBold(bold);
    return lang === 'en' ? 'Bold set' : '太字を設定しました';
  } catch (e) {
    throw new Error(e.message);
  }
}

/**
 * 斜体を設定
 * @param {boolean} italic - 斜体にするか
 */
function setTextItalic(italic, lang) {
  try {
    var textRange = getSelectedTextRange();
    if (!textRange) throw new Error(lang === 'en' ? 'Please select text' : 'テキストを選択してください');

    textRange.getTextStyle().setItalic(italic);
    return lang === 'en' ? 'Italic set' : '斜体を設定しました';
  } catch (e) {
    throw new Error(e.message);
  }
}

/**
 * 下線を設定
 * @param {boolean} underline - 下線を付けるか
 */
function setTextUnderline(underline, lang) {
  try {
    var textRange = getSelectedTextRange();
    if (!textRange) throw new Error(lang === 'en' ? 'Please select text' : 'テキストを選択してください');

    textRange.getTextStyle().setUnderline(underline);
    return lang === 'en' ? 'Underline set' : '下線を設定しました';
  } catch (e) {
    throw new Error(e.message);
  }
}

/**
 * 取り消し線を設定
 * @param {boolean} strikethrough - 取り消し線を付けるか
 */
function setTextStrikethrough(strikethrough, lang) {
  try {
    var textRange = getSelectedTextRange();
    if (!textRange) throw new Error(lang === 'en' ? 'Please select text' : 'テキストを選択してください');

    textRange.getTextStyle().setStrikethrough(strikethrough);
    return lang === 'en' ? 'Strikethrough set' : '取り消し線を設定しました';
  } catch (e) {
    throw new Error(e.message);
  }
}

/**
 * 行間を設定
 * @param {number} spacing - 行間倍率（1, 1.15, 1.5, 2など）
 * @param {string} lang - 言語コード (ja, en)
 */
function setLineSpacing(spacing, lang) {
  lang = lang || 'ja';
  try {
    var selection = SlidesApp.getActivePresentation().getSelection();
    var selectionType = selection.getSelectionType();

    var shape;
    if (selectionType === SlidesApp.SelectionType.TEXT) {
      shape = selection.getPageElementRange().getPageElements()[0].asShape();
    } else if (selectionType === SlidesApp.SelectionType.PAGE_ELEMENT) {
      var elements = selection.getPageElementRange().getPageElements();
      if (elements.length === 0) throw new Error(lang === 'en' ? 'Please select an element' : '要素を選択してください');
      var element = elements[0];
      if (element.getPageElementType() !== SlidesApp.PageElementType.SHAPE) {
        throw new Error(lang === 'en' ? 'Please select a text box' : 'テキストボックスを選択してください');
      }
      shape = element.asShape();
    } else {
      throw new Error(lang === 'en' ? 'Please select a text box' : 'テキストボックスを選択してください');
    }

    var textRange = shape.getText();
    var paragraphs = textRange.getParagraphs();

    for (var i = 0; i < paragraphs.length; i++) {
      var paragraph = paragraphs[i];
      var style = paragraph.getRange().getParagraphStyle();
      style.setLineSpacing(spacing * 100); // Google Slidesでは100が1行
    }

    return lang === 'en' ? 'Line spacing set to ' + spacing : '行間を ' + spacing + ' に設定しました';
  } catch (e) {
    throw new Error(e.message);
  }
}

/**
 * 段落の前後の間隔を設定
 * @param {number} before - 段落前のスペース（ポイント）
 * @param {number} after - 段落後のスペース（ポイント）
 * @param {string} lang - 言語コード (ja, en)
 */
function setParagraphSpacing(before, after, lang) {
  lang = lang || 'ja';
  try {
    var selection = SlidesApp.getActivePresentation().getSelection();
    var selectionType = selection.getSelectionType();

    var shape;
    if (selectionType === SlidesApp.SelectionType.TEXT) {
      shape = selection.getPageElementRange().getPageElements()[0].asShape();
    } else if (selectionType === SlidesApp.SelectionType.PAGE_ELEMENT) {
      var elements = selection.getPageElementRange().getPageElements();
      if (elements.length === 0) throw new Error(lang === 'en' ? 'Please select an element' : '要素を選択してください');
      var element = elements[0];
      if (element.getPageElementType() !== SlidesApp.PageElementType.SHAPE) {
        throw new Error(lang === 'en' ? 'Please select a text box' : 'テキストボックスを選択してください');
      }
      shape = element.asShape();
    } else {
      throw new Error(lang === 'en' ? 'Please select a text box' : 'テキストボックスを選択してください');
    }

    var textRange = shape.getText();
    var paragraphs = textRange.getParagraphs();

    for (var i = 0; i < paragraphs.length; i++) {
      var paragraph = paragraphs[i];
      var style = paragraph.getRange().getParagraphStyle();
      style.setSpaceAbove(before);
      style.setSpaceBelow(after);
    }

    return msg('paragraph_spacing_set', lang);
  } catch (e) {
    throw new Error(e.message);
  }
}

/**
 * テキスト色を設定
 * @param {string} hexColor - HEX形式の色
 */
function setTextColor(hexColor, lang) {
  try {
    var textRange = getSelectedTextRange();
    if (!textRange) throw new Error(lang === 'en' ? 'Please select text' : 'テキストを選択してください');

    textRange.getTextStyle().setForegroundColor(hexColor);
    return lang === 'en' ? 'Text color set' : 'テキスト色を設定しました';
  } catch (e) {
    throw new Error(e.message);
  }
}

/**
 * テキスト背景色（ハイライト）を設定
 * @param {string} hexColor - HEX形式の色またはnull（透明）
 */
function setTextBackgroundColor(hexColor, lang) {
  try {
    var textRange = getSelectedTextRange();
    if (!textRange) throw new Error(lang === 'en' ? 'Please select text' : 'テキストを選択してください');

    if (hexColor === null) {
      textRange.getTextStyle().setBackgroundColorTransparent();
    } else {
      textRange.getTextStyle().setBackgroundColor(hexColor);
    }
    return lang === 'en' ? 'Text background color set' : 'テキスト背景色を設定しました';
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
function getSelectedElementAttributes(lang) {
  try {
    var selection = SlidesApp.getActivePresentation().getSelection();
    var pageElementRange = selection.getPageElementRange();

    if (!pageElementRange) {
      throw new Error(lang === 'en' ? 'Select an object' : 'オブジェクトを選択してください');
    }

    var elements = pageElementRange.getPageElements();
    if (elements.length === 0) {
      throw new Error(lang === 'en' ? 'Select an object' : 'オブジェクトを選択してください');
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
          var fillColor = solidFill.getColor();
          var fillColorType = fillColor.getColorType();
          if (fillColorType === SlidesApp.ColorType.RGB) {
            try {
              attrs.fillColor = rgbColorToHex(fillColor.asRgbColor());
            } catch(e) {}
          } else if (fillColorType === SlidesApp.ColorType.THEME) {
            // テーマカラーの場合はカラースキームから解決
            try {
              var themeColor = fillColor.asThemeColor();
              var presentation = SlidesApp.getActivePresentation();
              var scheme = presentation.getMasters()[0].getColorScheme();
              var resolvedColor = scheme.getConcreteColor(themeColor.getThemeColorType());
              attrs.fillColor = rgbColorToHex(resolvedColor.asRgbColor());
            } catch(e) {
              try {
                attrs.fillColor = rgbColorToHex(fillColor.asRgbColor());
              } catch(e2) {}
            }
          }
        }
      }

      // 枠線
      var border = shape.getBorder();
      if (border.isVisible()) {
        var lineFill = border.getLineFill();
        if (lineFill.getFillType() === SlidesApp.LineFillType.SOLID) {
          var strokeColor = lineFill.getSolidFill().getColor();
          var strokeColorType = strokeColor.getColorType();
          if (strokeColorType === SlidesApp.ColorType.RGB) {
            try {
              attrs.strokeColor = rgbColorToHex(strokeColor.asRgbColor());
            } catch(e) {}
          } else if (strokeColorType === SlidesApp.ColorType.THEME) {
            // テーマカラーの場合はカラースキームから解決
            try {
              var strokeThemeColor = strokeColor.asThemeColor();
              var presentation = SlidesApp.getActivePresentation();
              var scheme = presentation.getMasters()[0].getColorScheme();
              var resolvedStrokeColor = scheme.getConcreteColor(strokeThemeColor.getThemeColorType());
              attrs.strokeColor = rgbColorToHex(resolvedStrokeColor.asRgbColor());
            } catch(e) {
              try {
                attrs.strokeColor = rgbColorToHex(strokeColor.asRgbColor());
              } catch(e2) {}
            }
          }
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
function selectSimilarElements(matchType, lang) {
  try {
    var selection = SlidesApp.getActivePresentation().getSelection();
    var pageElementRange = selection.getPageElementRange();

    if (!pageElementRange) {
      throw new Error(lang === 'en' ? 'Select a base object' : '基準オブジェクトを選択してください');
    }

    var selectedElements = pageElementRange.getPageElements();
    if (selectedElements.length === 0) {
      throw new Error(lang === 'en' ? 'Select a base object' : '基準オブジェクトを選択してください');
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
      throw new Error(lang === 'en' ? 'Selected object does not have this attribute' : '選択されたオブジェクトにはこの属性がありません');
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
      return { count: 0, message: lang === 'en' ? 'No matching objects found' : '一致するオブジェクトが見つかりませんでした' };
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

    var typeLabelsEn = {
      'fillColor': 'fill color',
      'strokeColor': 'stroke color',
      'textColor': 'text color',
      'fontFamily': 'font',
      'fontSize': 'font size',
      'strokeWeight': 'stroke weight'
    };

    var typeLabelsJa = {
      'fillColor': '塗りつぶし色',
      'strokeColor': '枠線色',
      'textColor': '文字色',
      'fontFamily': 'フォント',
      'fontSize': 'フォントサイズ',
      'strokeWeight': '枠線の太さ'
    };

    var message = lang === 'en'
      ? 'Selected ' + matchedElements.length + ' objects with same ' + typeLabelsEn[matchType]
      : typeLabelsJa[matchType] + 'が同じオブジェクトを ' + matchedElements.length + ' 個選択しました';

    return {
      count: matchedElements.length,
      message: message
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
          var fillColor = solidFill.getColor();
          var fillColorType = fillColor.getColorType();
          if (fillColorType === SlidesApp.ColorType.RGB) {
            try {
              attrs.fillColor = rgbColorToHex(fillColor.asRgbColor());
            } catch(e) {}
          } else if (fillColorType === SlidesApp.ColorType.THEME) {
            // テーマカラーの場合はカラースキームから解決
            try {
              var themeColor = fillColor.asThemeColor();
              var presentation = SlidesApp.getActivePresentation();
              var scheme = presentation.getMasters()[0].getColorScheme();
              var resolvedColor = scheme.getConcreteColor(themeColor.getThemeColorType());
              attrs.fillColor = rgbColorToHex(resolvedColor.asRgbColor());
            } catch(e) {
              try {
                attrs.fillColor = rgbColorToHex(fillColor.asRgbColor());
              } catch(e2) {}
            }
          }
        }
      }

      // 枠線
      var border = shape.getBorder();
      if (border.isVisible()) {
        var lineFill = border.getLineFill();
        if (lineFill.getFillType() === SlidesApp.LineFillType.SOLID) {
          var strokeColor = lineFill.getSolidFill().getColor();
          var strokeColorType = strokeColor.getColorType();
          if (strokeColorType === SlidesApp.ColorType.RGB) {
            try {
              attrs.strokeColor = rgbColorToHex(strokeColor.asRgbColor());
            } catch(e) {}
          } else if (strokeColorType === SlidesApp.ColorType.THEME) {
            // テーマカラーの場合はカラースキームから解決
            try {
              var strokeThemeColor = strokeColor.asThemeColor();
              var presentation = SlidesApp.getActivePresentation();
              var scheme = presentation.getMasters()[0].getColorScheme();
              var resolvedStrokeColor = scheme.getConcreteColor(strokeThemeColor.getThemeColorType());
              attrs.strokeColor = rgbColorToHex(resolvedStrokeColor.asRgbColor());
            } catch(e) {
              try {
                attrs.strokeColor = rgbColorToHex(strokeColor.asRgbColor());
              } catch(e2) {}
            }
          }
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
function copyElementStyle(lang) {
  try {
    var selection = SlidesApp.getActivePresentation().getSelection();
    var pageElementRange = selection.getPageElementRange();

    if (!pageElementRange) {
      throw new Error(lang === 'en' ? 'Select an object' : 'オブジェクトを選択してください');
    }

    var elements = pageElementRange.getPageElements();
    if (elements.length === 0) {
      throw new Error(lang === 'en' ? 'Select an object' : 'オブジェクトを選択してください');
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
            var fillColor = solidFill.getColor();
            var fillColorType = fillColor.getColorType();
            if (fillColorType === SlidesApp.ColorType.RGB) {
              style.fillColor = rgbColorToHex(fillColor.asRgbColor());
            } else if (fillColorType === SlidesApp.ColorType.THEME) {
              // テーマカラーの場合はカラースキームから解決
              var themeColor = fillColor.asThemeColor();
              var presentation = SlidesApp.getActivePresentation();
              var scheme = presentation.getMasters()[0].getColorScheme();
              var resolvedColor = scheme.getConcreteColor(themeColor.getThemeColorType());
              style.fillColor = rgbColorToHex(resolvedColor.asRgbColor());
            }
            style.fillAlpha = solidFill.getAlpha();
          } catch(e) {
            // フォールバック: 直接RGBとして取得を試みる
            try {
              style.fillColor = rgbColorToHex(solidFill.getColor().asRgbColor());
              style.fillAlpha = solidFill.getAlpha();
            } catch(e2) {}
          }
        }
      }

      // 枠線
      var border = shape.getBorder();
      if (border.isVisible()) {
        var lineFill = border.getLineFill();
        if (lineFill.getFillType() === SlidesApp.LineFillType.SOLID) {
          try {
            var strokeColor = lineFill.getSolidFill().getColor();
            var strokeColorType = strokeColor.getColorType();
            if (strokeColorType === SlidesApp.ColorType.RGB) {
              style.strokeColor = rgbColorToHex(strokeColor.asRgbColor());
            } else if (strokeColorType === SlidesApp.ColorType.THEME) {
              // テーマカラーの場合はカラースキームから解決
              var themeColor = strokeColor.asThemeColor();
              var presentation = SlidesApp.getActivePresentation();
              var scheme = presentation.getMasters()[0].getColorScheme();
              var resolvedColor = scheme.getConcreteColor(themeColor.getThemeColorType());
              style.strokeColor = rgbColorToHex(resolvedColor.asRgbColor());
            }
          } catch(e) {
            // フォールバック
            try {
              style.strokeColor = rgbColorToHex(lineFill.getSolidFill().getColor().asRgbColor());
            } catch(e2) {}
          }
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
function applyElementStyle(style, options, lang) {
  try {
    var selection = SlidesApp.getActivePresentation().getSelection();
    var pageElementRange = selection.getPageElementRange();

    if (!pageElementRange) {
      throw new Error(lang === 'en' ? 'Select an object' : 'オブジェクトを選択してください');
    }

    var elements = pageElementRange.getPageElements();
    if (elements.length === 0) {
      throw new Error(lang === 'en' ? 'Select an object' : 'オブジェクトを選択してください');
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

    return lang === 'en'
      ? 'Style applied to ' + appliedCount + ' object(s)'
      : appliedCount + '個のオブジェクトにスタイルを適用しました';
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
function toggleRuler(show, lang) {
  // Google Slides APIではルーラー表示の制御はサポートされていません
  // ユーザーに手動で設定するよう案内
  return lang === 'en'
    ? 'Please toggle ruler manually from the View menu'
    : 'ルーラーは「表示」メニューから手動で切り替えてください';
}

/**
 * ガイド表示を切り替え（Slides APIでは直接制御不可）
 */
function toggleGuides(show, lang) {
  return lang === 'en'
    ? 'Please toggle guides manually from the View menu'
    : 'ガイドは「表示」メニューから手動で切り替えてください';
}

/**
 * 垂直ガイドを追加
 */
function addVerticalGuide(lang) {
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

    return lang === 'en' ? 'Vertical guide added' : '垂直ガイドを追加しました';
  } catch (e) {
    throw new Error(e.message);
  }
}

/**
 * 水平ガイドを追加
 */
function addHorizontalGuide(lang) {
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

    return lang === 'en' ? 'Horizontal guide added' : '水平ガイドを追加しました';
  } catch (e) {
    throw new Error(e.message);
  }
}

/**
 * ガイド編集画面を開く（UIメッセージのみ）
 */
function editGuides(lang) {
  return lang === 'en' ? 'Guides can be edited by dragging directly on the slide' : 'ガイドはスライド上で直接ドラッグして編集できます';
}

/**
 * すべてのガイドをクリア
 */
function clearGuides(lang) {
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
      return lang === 'en' ? 'Guides cleared' : 'ガイドをクリアしました';
    } else {
      return lang === 'en' ? 'No guides to clear' : 'クリアするガイドがありません';
    }
  } catch (e) {
    throw new Error(e.message);
  }
}
