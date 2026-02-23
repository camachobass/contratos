function organizarContratosFirmados() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Contratos");
  if (!hoja) {
    throw new Error('No existe la hoja "Contratos".');
  }

  const datos = hoja.getDataRange().getValues();
  if (datos.length < 2) {
    return;
  }

  const encabezados = datos[0];
  const indices = obtenerIndicesEncabezados_(encabezados, [
    "Nombre trabajador",
    "Cédula",
    "Link Final",
    "Estado contrato"
  ]);

  const iNombre = indices["Nombre trabajador"];
  const iCedula = indices["Cédula"];
  const iLinkFinal = indices["Link Final"];
  const iEstado = indices["Estado contrato"];

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
  const trabajadores = datos
    .map((fila, idx) => {
      if (idx === 0) return null;
      return {
        fila: idx,
        nombre: String(fila[iNombre] || "").trim(),
        cedula: extraerSoloDigitos_(fila[iCedula])
      };
    })
    .filter(Boolean);

  while (archivos.hasNext()) {
    const archivo = archivos.next();
    const nombreArchivo = archivo.getName();

    const trabajador = encontrarTrabajadorPorArchivo_(nombreArchivo, trabajadores);
    if (!trabajador) continue;

    const index = trabajador.fila;
    const nombreTrabajador = trabajador.nombre;
    const cedulaTrabajador = trabajador.cedula;
    const yaFirmado = datos[index][iEstado] === "Firmado" && datos[index][iLinkFinal];

    // Buscar o crear subcarpeta del trabajador
    let subcarpeta;
    const carpetasTrabajador = carpetaClasificados.getFoldersByName(nombreTrabajador);
    subcarpeta = carpetasTrabajador.hasNext() ? carpetasTrabajador.next() : carpetaClasificados.createFolder(nombreTrabajador);

    // Evita duplicados si el archivo ya fue clasificado antes.
    const copia = obtenerArchivoPorNombreEnCarpeta_(subcarpeta, archivo.getName()) || archivo.makeCopy(archivo.getName(), subcarpeta);

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
        const coincideCedula = cedulaTrabajador && nombreCF.indexOf(cedulaTrabajador) !== -1;
        const coincideNombre = coincideNombreArchivo_(nombreCF, nombreTrabajador);
        if (coincideCedula || coincideNombre) {
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

function encontrarTrabajadorPorArchivo_(nombreArchivo, trabajadores) {
  const porCedula = trabajadores.find(function(trabajador) {
    return trabajador.cedula && nombreArchivo.indexOf(trabajador.cedula) !== -1;
  });
  if (porCedula) {
    return porCedula;
  }

  return trabajadores.find(function(trabajador) {
    return coincideNombreArchivo_(nombreArchivo, trabajador.nombre);
  }) || null;
}

function obtenerArchivoPorNombreEnCarpeta_(carpeta, nombreArchivo) {
  const archivos = carpeta.getFilesByName(nombreArchivo);
  return archivos.hasNext() ? archivos.next() : null;
}
