const SS = SpreadsheetApp.getActiveSpreadsheet();
const SETTING_SHEET = SS.getSheetByName("Setting");
const DOC_ID = SETTING_SHEET.getRange("B1").getValue();
const GRADE_NUM = Number(SETTING_SHEET.getRange("B2").getValue().slice(0, 1));
const PROPERTIES = PropertiesService.getScriptProperties().getProperties();

function main() {
  // Googleドキュメントからテキストを配列で取得
  const texts = getTextsListFromGoogleDocument_(DOC_ID);

  // APIを叩いてルビが振られたテキストを配列で取得
  const rubyTexts = texts.map((text) => {
    if (text !== "") {
      const json = fetchAPI_(text, GRADE_NUM);
      const words = json.result.word;
      const rubyWords = words.map((word) => {
        if (word.subword) {
          const subwords = word.subword;
          const chars = subwords.map((subword) => checkIsKanji_(subword.surface) ? `|${subword.surface}|\{${subword.furigana}\}` : subword.surface);
          return chars.join("");
        } else {
          return word.furigana ? `|${word.surface}|\{${word.furigana}\}` : word.surface;
        }
      }
      );
      return rubyWords.join("");
    }
  });

  // ルビが振られたテキストをGoogleドキュメントに追記
  postRubyTextsToDoc_(rubyTexts);
  console.log("ルビふり完了。\nルビふりテキストが追記されたドキュメントのURL => ", DocumentApp.openById(DOC_ID).getUrl());
}


/**
 * APIを叩く
 * @param{string} queryText - ルビを振るテキスト
 * @param{number} gradeSetting - 学年の設定（初期値は小３）
 * @return{object} JSON.parse(response) - APIから返されたオブジェクト型
 */
function fetchAPI_(queryText, gradeSetting = 3) {
  // Yahoo!ルビふりAPI
  // https://developer.yahoo.co.jp/webapi/jlp/furigana/v2/furigana.html
  const clientId = PROPERTIES["YahooClientID"];
  const url = `https://jlp.yahooapis.jp/FuriganaService/V2/furigana?appid=${clientId}`;


  const headers = {
    "Content-Type": "application/json",
    "User-Agent": `Yahoo AppID: ${clientId}`,
  }

  const id = Utilities.getUuid();     //generate UUID: string
  const obj = {
    "id": id,
    "jsonrpc": "2.0",
    "method": "jlp.furiganaservice.furigana",
    "params": {
      "q": queryText,
      "grade": Number(gradeSetting)
    }
  };

  const options = {
    "headers": headers,
    "payload": JSON.stringify(obj)
  };
  try {
    const response = UrlFetchApp.fetch(url, options);
    return JSON.parse(response.getContentText());
  } catch (err) {
    console.error(err);
  }
}

/**
 * Googleドキュメントからテキスト配列を取得する
 * @param{string} docId - GoogleドキュメントのID
 * @return{string[]} result - テキストの配列
 */
function getTextsListFromGoogleDocument_(docId) {
  const doc = DocumentApp.openById(docId);
  const paragraphs = doc.getBody().getParagraphs();

  const result = paragraphs.map(text => text.getText());
  return result;
}

/**
 * ルビ付きのテキストをGoogleドキュメントに追加する
 * @param{string[]} rubyTexts - ルビふりされたテキストの配列 
 * @return{void}
 */
function postRubyTextsToDoc_(rubyTexts) {
  const body = DocumentApp.openById(DOC_ID).getBody();

  // 改ページ
  body.appendPageBreak();
  body.appendParagraph("=======ルビふり=======").setHeading(DocumentApp.ParagraphHeading.HEADING2);

  // テキストの転記
  try {
    rubyTexts.map((text) => body.appendParagraph(text));
  } catch (err) {
    console.error(err);
  }
}

/**
 * 漢字かどうか判断する関数（文字コードで判断する）
 * @param{string} chars - 判断する文字列
 * @return{boolean} isKanji - 漢字ならTrue、感じでなければFalse
 */
function checkIsKanji_(chars) {
  const charCode = chars.charCodeAt(0);
  const code = charCode < 0 ? charCode + 65536 : charCode;
  const isKanji = code >= 19968 && code <= 40959;
  return isKanji;
}




