const SOURCES_SHEET_NAME = "地区リストのシート名";
const TARGET_SHEETS_ID = "集計した花粉量をのせるシートID";
const TARGETS_SHEET_NAME = "集計した花粉量をのせるシート名";
const DISTRICTS_SHEET_DISTRICT_COLMN_NAME = "B";
const DISTRICTS_SHEET_CITYCODE_COLMN_NAME = "C";

function myFunction() {
  // 現在日時
  let date = new Date(); //現在日時のDateオブジェクトを作る
  let today = Utilities.formatDate(date, 'JST', 'yyyy年MM月dd日 HH時mm分ss秒');
  console.log(`today:${today}`);

  // シートから[地区名,市区町村コード]のリストを取得
  let districts = get_target_list();

  // [地区名,市区町村コード]のリストから、昨日の[市区町村コード,花粉飛散量]のリストを取得
  date.setDate(date.getDate()-1);  
  let district_pollen_list = [["更新日時",today],["県庁所在地","前日の花粉飛散量"]];
  for(let i=0;i < districts.length;i++) {
    let ret = get_pollen_level_for_district(districts[i][1].toString().padStart(5,"0"),date);
    district_pollen_list.push([districts[i][0],ret])
    console.log(`${districts[i][0]}:${ret}`);
  }

  // ターゲットとなるシートへ書き込み
  set_pollen_list(district_pollen_list);

}

function set_pollen_list(data) {
  const sheet = SpreadsheetApp.openById(TARGET_SHEETS_ID).getSheetByName(TARGETS_SHEET_NAME);
  console.log(`sheet:${sheet}`);


  // 書き込むデータは[地区名,花粉飛散量]
  console.log(`data:${data}`);
  console.log(`data.length:${data.length}`);
  console.log(`data[0]:${data[0]}`);
  console.log(`data[1]:${data[1]}`);
  console.log(`data[2]:${data[2]}`);

  console.log(`A1:B${data.length}`);
  var range = sheet.getRange(`A1:B${data.length}`);
  console.log(`range:${range}`);
  var values = range.setValues(data);
  console.log(`values:${values}`);

}

function get_pollen_level_for_district(district,date) {
  targetDay = Utilities.formatDate(date, 'JST', 'yyyyMMdd');
  let url = `https://wxtech.weathernews.com/opendata/v1/pollen?citycode=${district}&start=${targetDay}&end=${targetDay}`;
  let response = UrlFetchApp.fetch(url).getContentText();

  // csvで落ちてくる時系列のデータをリストにする
  let pollenHourly = response.split(/\r\n|\n/).slice(1,-1);

  // 花粉量が0以上の数値を合計する
  let pollenSum = 0;
  for(i = 0; i < pollenHourly.length;i++){
    pollenSum += Number(pollenHourly[i].split(",")[2]) >= 0 ? Number(pollenHourly[i].split(",")[2]) : 0;
  };

  return pollenSum;
}

function get_target_list() {
  const sheet = SpreadsheetApp.openById(TARGET_SHEETS_ID).getSheetByName(SOURCES_SHEET_NAME);
  const rows = sheet.getLastRow();
  console.log(rows);

  // 地区名と市区町村コードの範囲を別々に取得
  const district_range = `${DISTRICTS_SHEET_DISTRICT_COLMN_NAME}2:${DISTRICTS_SHEET_DISTRICT_COLMN_NAME}${rows}`;
  const city_range = `${DISTRICTS_SHEET_CITYCODE_COLMN_NAME}2:${DISTRICTS_SHEET_CITYCODE_COLMN_NAME}${rows}`;

  // 地区名リストと市区町村コードリストを[地区名,市区町村コード]のリストに
  let citie_codes = sheet.getRange(city_range).getValues();
  let districts = sheet.getRange(district_range).getValues();
  data = [];
  for(i = 0;i < (rows-1);i++) {
    data.push([districts[i],citie_codes[i]]);
  }

  return data;
}
