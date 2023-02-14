const hoja = SpreadsheetApp.openById('1S9NN_xYUtjsVz_gA-9FNYf9FHG75xLrenbje5CzdIdM').getActiveSheet();

function doGet()
{
    return HtmlService.createTemplateFromFile('index').evaluate().setTitle('Agenda Google Apps Script');
}

function doPost()
{
    return HtmlService.createTemplateFromFile('index').evaluate().setTitle('Agenda Google Apps Script');
}

function obtenerDatosHtml(nombre)
{
    return HtmlService.createHtmlOutputFromFile(nombre).getContent();
}

function obtenerContactos()
{
    return hoja.getDataRange().getValues();
}

function insertarContactos(nombre, correo)
{
    hoja.appendRow([nombre, correo]);
}