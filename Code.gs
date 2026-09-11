/**
 * Точка входа веб-приложения.
 */
function doGet(e) {
  const template = HtmlService.createTemplateFromFile('BRO');

  template.initialSliceJson = jsonForHtml_(
    e && e.parameter
      ? e.parameter.slice
      : ''
  );

  return template
    .evaluate()
    .setTitle('Commercial Contracts Dashboard v39');
}


/**
 * Подключает HTML-файлы Styles, Assets и Client.
 */
function include(fileName) {
  return HtmlService
    .createHtmlOutputFromFile(fileName)
    .getContent();
}


/**
 * Возвращает постоянную ссылку действующего веб-приложения.
 */
function getDashboardBaseUrl() {
  return ScriptApp.getService().getUrl();
}


/**
 * Безопасно передаёт строковый параметр в HTML-шаблон.
 */
function jsonForHtml_(value) {
  return JSON.stringify(String(value || ''))
    .replace(/</g, '\\u003c');
}
