function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Control de asistencia')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
