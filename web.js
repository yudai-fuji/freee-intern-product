// Webアプリを起動して画面表示する入口

function doGet() {
  return HtmlService.createTemplateFromFile("view_home")
    .evaluate().setTitle("店舗運営支援アプリ")
}