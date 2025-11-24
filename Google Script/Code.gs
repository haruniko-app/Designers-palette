// Code.gs (v41.0 軽量監視・高速化版)

function onOpen() {
SlidesApp.getUi().createMenu('AI画像ツール').addItem('ツールを起動', 'showSidebar').addToUi();
}

function showSidebar() {
var html = HtmlService.createHtmlOutputFromFile('Sidebar').setTitle('AI画像拡張ツール').setWidth(340);
SlidesApp.getUi().showSidebar(html)                                                                           ;
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
// 画像以外は無視
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
 * @param {string} alignmentType - 整列タイプ (LEFT, CENTER, RIGHT, TOP, MIDDLE, BOTTOM)
 */
function alignElements(alignmentType) {
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

    // スライドのサイズを取得
    var slide = selection.getCurrentPage();
    var pageWidth = slide.getPageWidth();
    var pageHeight = slide.getPageHeight();

    // 各要素に対して整列を実行
    for (var i = 0; i < elements.length; i++) {
      var element = elements[i];
      var width = element.getWidth();
      var height = element.getHeight();

      // 整列タイプに応じた位置計算
      switch(alignmentType) {
        case 'LEFT':
          element.setLeft(0);
          break;
        case 'CENTER':
          element.setLeft((pageWidth - width) / 2);
          break;
        case 'RIGHT':
          element.setLeft(pageWidth - width);
          break;
        case 'TOP':
          element.setTop(0);
          break;
        case 'MIDDLE':
          element.setTop((pageHeight - height) / 2);
          break;
        case 'BOTTOM':
          element.setTop(pageHeight - height);
          break;
        default:
          throw new Error('不明な整列タイプ: ' + alignmentType);
      }
    }

    return '整列完了';
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
