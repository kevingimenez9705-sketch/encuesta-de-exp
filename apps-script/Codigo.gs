/**
 * Web App que recibe las respuestas de index.html y las guarda en la hoja.
 * Escribe cada dato bajo la columna con su mismo nombre y crea las columnas
 * que falten, así las preguntas nuevas se registran sin tocar este código.
 */
const NOMBRE_HOJA = 'Respuestas';

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const datos = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName(NOMBRE_HOJA) || ss.insertSheet(NOMBRE_HOJA);

    let encabezados = hoja.getLastColumn()
      ? hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0]
      : [];
    const faltantes = Object.keys(datos).filter(k => !encabezados.includes(k));
    if (faltantes.length) {
      hoja.getRange(1, encabezados.length + 1, 1, faltantes.length).setValues([faltantes]);
      encabezados = encabezados.concat(faltantes);
    }

    const fila = encabezados.map(h => (h in datos ? datos[h] : ''));
    hoja.appendRow(fila);
    return respuesta({ ok: true });
  } catch (err) {
    return respuesta({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function respuesta(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
