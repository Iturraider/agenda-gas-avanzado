function doGet()
{
    return HtmlService.createTemplateFromFile('web').evaluate().setTitle('Agenda Google Apps Script');
}

function obtenerDatosHtml(nombre)
{
    return HtmlService.createHtmlOutputFromFile(nombre).getContent();
}

function obtenerContactos()
{
    let hoja = SpreadsheetApp.openById('1S9NN_xYUtjsVz_gA-9FNYf9FHG75xLrenbje5CzdIdM').getActiveSheet();
    let datos = hoja.getDataRange().getValues();
    return datos;
}