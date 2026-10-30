/**
 * Webアプリを起動して画面表示する入口 WebアプリのURLを開くとGASがdoGet()を自動で実行する。
 * 戻り値 - GoogleAppsScript.HTML.HtmlOutput
 */
function doGet() {
  return HtmlService.createTemplateFromFile("Browser/view_home")
    .evaluate().setTitle("店舗運営支援アプリ");
}