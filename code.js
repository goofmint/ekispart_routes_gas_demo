// Ekispert API キー（実際に使用する際はご自身のキーに差し替えてください）
const API_KEY = 'YOUR_API_KEY';

// スプレッドシートを開いたときに自動実行される関数
function onOpen() {
  SpreadsheetApp.getUi() // UI操作用のオブジェクト取得
    .createMenu('Ekispert') // メニュー名を作成
    .addItem('ルート探索', 'runRoutes') // メニュー項目と実行関数を指定
    .addToUi(); // シートのUIに追加
}

/**
 * シート名が重複している場合、末尾に (2), (3)… をつけて一意な名前を作成する
 * @param {string} baseName - 元になるシート名
 * @returns {string} - 一意なシート名
 */
function getUniqueSheetName(baseName) {
  const ss = SpreadsheetApp.getActive(); // 現在のスプレッドシート取得
  let name = baseName;
  let counter = 1;

  // 同名シートが存在する限り、番号を振って検証
  while (ss.getSheetByName(name)) {
    counter++;
    name = `${baseName} (${counter})`;
  }
  return name; // 重複しないユニーク名を返す
}

/**
 * 新しいシートを追加する際に名前が被らないようにするラッパー
 * @param {string} baseName - 希望のシート名
 * @returns {Sheet} - 作成されたシートオブジェクト
 */
function addSheetWithIncrement(baseName) {
  const ss = SpreadsheetApp.getActive();
  const uniqueName = getUniqueSheetName(baseName); // 重複を避けた名前を取得
  return ss.insertSheet(uniqueName); // 新規シートを作成
}

/**
 * 入力された駅名リストをもとに経路検索を実行し、新しいシートに結果を出力
 */
function runRoutes() {
  const ss = SpreadsheetApp.getActive();
  const input = ss.getSheetByName('経路入力'); // 入力用シート
  const codes = []; // 駅コード一覧
  const names = []; // 駅名一覧

  // 出発、経由1、経由2、到着の4行分をループ
  for (let i = 1; i <= 4; i++) {
    const cell = input.getRange(i, 2).getValue(); // B列の値を取得
    // 出発(1行目)・到着(4行目)が未入力ならアラートして中断
    if (!cell && (i === 1 || i === 4)) {
      SpreadsheetApp.getUi().alert('出発駅／到着駅は必須です');
      return;
    }
    if (cell) {
      names.push(cell); // 駅名を格納
      const data = ss.getSheetByName('データ'); // コード管理用シート
      const code = data.createTextFinder(cell).findNext(); // 駅名を検索

      if (!code) {
        SpreadsheetApp.getUi().alert(`${i}行目の名前が見つかりません（${cell}）`);
        return;
      }
      // 検索結果の隣のセルにある駅コードを取得
      codes.push(data.getRange(code.getRow(), code.getColumn() + 1).getValue());
    }
  }

  // コードから経路APIを呼び出し
  const courses = fetchRoutes(codes.join(':'));
  const base = `${names[0]}から${names[names.length - 1]}`; // シート名のベース
  const outSheet = addSheetWithIncrement(base); // 新規シート

  let col = 1;
  courses.forEach((course, idx) => {
    // ルートごとにヘッダーを書き込み
    outSheet.getRange(1, col).setValue(`ルート${idx + 1}`);
    const price = course.Price.find(p => p.kind === 'FareSummary')?.Oneway || 0;
    outSheet.getRange(2, col).setValue(`合計：${price}円`);

    // Route が単体or配列の可能性に対応
    let routes = course.Route;
    if (!Array.isArray(routes)) routes = [routes];

    // データ開始行とヘッダー
    let row = 4;
    outSheet.getRange(row, col, 1, 3).setValues([['乗車駅', '路線', '降車駅']]);
    row++;

    // 各区間ごとに駅名＆路線名を書き込む
    routes.forEach(route => {
      let pts = route.Point;
      if (!Array.isArray(pts)) pts = [pts];

      let lines = route.Line;
      if (!Array.isArray(lines)) lines = [lines];

      for (let i = 0; i < lines.length; i++) {
        const start = pts[i].Station.Name;
        const end = pts[i + 1].Station.Name;
        const lineName = lines[i].Name;
        outSheet.getRange(row, col, 1, 3).setValues([[start, lineName, end]]);
        row++;
      }
    });

    col += 4; // 次のルートは4列右に出力
  });
}

/**
 * Price配列から「合計運賃(FareSummary)」を抽出して返す
 * @param {Array} prices - Priceオブジェクトの配列
 * @returns {number} - 運賃(円)
 */
function getPrice(prices) {
  return prices.filter(p => p.kind === 'FareSummary')[0].Oneway || 0;
}

/**
 * 駅検索用APIを呼び出して駅名とコードの配列を取得
 * @param {string} name - 駅名またはキーワード
 * @returns {Array} - {Station:{Name,code}} の配列
 */
function fetchStations(name) {
  const url = `https://api.ekispert.jp/v1/json/station/light?key=${API_KEY}` +
    `&name=${encodeURIComponent(name)}&nameMatchType=partial&type=train`;
  const res = UrlFetchApp.fetch(url); // API呼び出し
  if (res.getResponseCode() !== 200) return []; // エラー時は空配列
  console.log(res.getContentText()); // ログに生JSONを出力（開発時用）
  const points = JSON.parse(res.getContentText()).ResultSet.Point;
  return Array.isArray(points) ? points : [points]; // 配列化して返す
}

/**
 * 経路探索APIを呼び出して経路一覧を取得
 * @param {string} viaList - コロン区切りの駅コード（例: code1:code2...）
 * @returns {Array} - Courseオブジェクトの配列（ルート情報）
 */
function fetchRoutes(viaList) {
  const url = `https://api.ekispert.jp/v1/json/search/course/extreme` +
    `?key=${API_KEY}&viaList=${viaList}&searchType=plain`;
  const res = UrlFetchApp.fetch(url);
  if (res.getResponseCode() !== 200) return [];
  let courses = JSON.parse(res.getContentText()).ResultSet.Course;
  return Array.isArray(courses) ? courses : [courses];
}

/**
 * onEdit用関数：駅名入力時に補完候補リストとドロップダウンを設定
 * @param {Object} e - onEdit イベントオブジェクト
 */
function onEditHandler(e) {
  const sheet = e.range.getSheet();
  if (sheet.getName() !== '経路入力') return; // 別シートでは実行しない

  const row = e.range.getRow();
  const col = e.range.getColumn();
  if (col !== 2 || row < 1 || row > 4) return; // B1〜B4のみ対象

  const val = e.value;
  const dataSh = SpreadsheetApp.getActive().getSheetByName('データ');
  console.log(`編集: row=${row}, 値=${val}`);

  // 入力行に対応した列位置を計算：B1→A,B / B2→C,D / ...
  const startCol = (row - 1) * 2 + 1;

  if (!val) {
    e.range.clearDataValidations(); // ドロップダウン解除
    // 該当セル範囲をクリア
    dataSh.getRange(2, startCol, dataSh.getMaxRows() - 1, 2).clearContent();
    return;
  }

  const pts = fetchStations(val); // APIから駅情報取得
  const names = pts.map(p => p.Station.Name); // 駅名のみリスト化

  // 駅名・コードを縦に書き込み
  const output = pts.map(p => [p.Station.Name, p.Station.code]);
  dataSh.getRange(2, startCol, output.length, 2).clearContent().setValues(output);

  // ドロップダウン候補に駅名をセット
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(names, true) // リストから選べる
    .setAllowInvalid(false) // 無効入力を禁止
    .build();
  e.range.clearDataValidations();
  e.range.setDataValidation(rule);
}

/**
 * 列番号を A, B, …, AA … の列名文字列に変換する関数
 * @param {number} column - 列番号 (1スタート)
 * @returns {string} - 列名文字列
 */
function columnToLetter(column) {
  let temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(65 + temp) + letter;
    column = Math.floor((column - temp - 1) / 26);
  }
  return letter;
}
