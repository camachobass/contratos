function organizarContratosFirmados() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Contratos");
  const datos = hoja.getDataRange().getValues();
  const encabezados = datos[0];

  const iNombre = encabezados.indexOf("Nombre trabajador");
  const iLinkFinal = encabezados.indexOf("Link Final");
  const iEstado = encabezados.indexOf("Estado contrato");

  const carpetaFirmados = DriveApp.getFolderById("1DwQRZ4_689QhXt72SuyxcprMalVt4L6L");
  const carpetaClasificados = DriveApp.getFolderById("13Qltz1lpapzbpFctgp7m6grqrjqMiCP6");
  const carpetaContratosParaFirmar = DriveApp.getFolderById("1wWsr-hEDMrUl5r2nm1urp4Fj7UxJOsy9");

  // Crear o acceder a subcarpeta "Procesados" en 'Firmados'
  let carpetaProcesadosFirmados;
  const carpetas = carpetaFirmados.getFoldersByName("Procesados");
  carpetaProcesadosFirmados = carpetas.hasNext() ? carpetas.next() : carpetaFirmados.createFolder("Procesados");

  // Crear o acceder a subcarpeta "Procesados" en 'Contratos para firmar'
  let carpetaProcesadosParaFirmar;
  const carpetasCF = carpetaContratosParaFirmar.getFoldersByName("Procesados");
  carpetaProcesadosParaFirmar = carpetasCF.hasNext() ? carpetasCF.next() : carpetaContratosParaFirmar.createFolder("Procesados");

  const archivos = carpetaFirmados.getFiles();
  const nombresEnHoja = datos.map(fila => fila[iNombre]);

  while (archivos.hasNext()) {
    const archivo = archivos.next();
    const nombreArchivo = archivo.getName();

    // Buscar coincidencia con trabajador
    const index = nombresEnHoja.findIndex((n, idx) => idx > 0 && n && nombreArchivo.toLowerCase().includes(String(n).toLowerCase()));
    if (index === -1) continue;

    const nombreTrabajador = datos[index][iNombre];
    const yaFirmado = datos[index][iEstado] === "Firmado" && datos[index][iLinkFinal];

    // Buscar o crear subcarpeta del trabajador
    let subcarpeta;
    const carpetasTrabajador = carpetaClasificados.getFoldersByName(nombreTrabajador);
    subcarpeta = carpetasTrabajador.hasNext() ? carpetasTrabajador.next() : carpetaClasificados.createFolder(nombreTrabajador);

    // Crear copia del firmado
    const copia = archivo.makeCopy(archivo.getName(), subcarpeta);

    // Actualizar hoja
    hoja.getRange(index + 1, iLinkFinal + 1).setValue(copia.getUrl());
    hoja.getRange(index + 1, iEstado + 1).setValue("Firmado");

    // Mover archivo original a "Procesados" en Firmados
    try {
      carpetaProcesadosFirmados.addFile(archivo);
      carpetaFirmados.removeFile(archivo);
    } catch (e) {
      // Si no se puede mover, renombrar como ARCHIVADO
      const nombreActual = archivo.getName();
      if (!nombreActual.startsWith("ARCHIVADO - ") && yaFirmado) {
        try {
          archivo.setName(`ARCHIVADO - ${nombreActual}`);
        } catch (err) {
          Logger.log(`No se pudo renombrar ${nombreActual}: ${err.message}`);
        }
      }
    }

    // Mover contrato original (sin firmar) de "Contratos para firmar" a su carpeta "Procesados"
    try {
      const archivosCF = carpetaContratosParaFirmar.getFiles();
      while (archivosCF.hasNext()) {
        const archivoCF = archivosCF.next();
        const nombreCF = archivoCF.getName();
        if (nombreCF.toLowerCase().includes(nombreTrabajador.toLowerCase())) {
          carpetaProcesadosParaFirmar.addFile(archivoCF);
          carpetaContratosParaFirmar.removeFile(archivoCF);
          break;
        }
      }
    } catch (e) {
      Logger.log(`No se pudo mover contrato sin firmar de ${nombreTrabajador}: ${e.message}`);
    }
  }
}
