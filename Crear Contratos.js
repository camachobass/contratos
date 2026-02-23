function generarContratos() {
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
    "Fecha inicio",
    "Fecha fin",
    "Dirección trabajador",
    "Teléfono",
    "Correo",
    "Salario",
    "Fecha firma",
    "Funciones Especiales",
    "Creado",
    "Link"
  ]);

  const iNombre = indices["Nombre trabajador"];
  const iCedula = indices["Cédula"];
  const iFechaInicio = indices["Fecha inicio"];
  const iFechaFin = indices["Fecha fin"];
  const iDireccion = indices["Dirección trabajador"];
  const iTelefono = indices["Teléfono"];
  const iCorreo = indices["Correo"];
  const iSalario = indices["Salario"];
  const iFechaFirma = indices["Fecha firma"];
  const iFunciones = indices["Funciones Especiales"];
  const iCreado = indices["Creado"];
  const iLink = indices["Link"];

  const plantillaId = '1pBCfXUOR-bvDe9a9-NKidd_UZajNjaKs13ff9ynn770';
  const carpetaDestino = DriveApp.getFolderById('1wWsr-hEDMrUl5r2nm1urp4Fj7UxJOsy9');
  validarAccesoDrive_([
    { tipo: "archivo", nombre: "Plantilla de contrato", id: plantillaId },
    { tipo: "carpeta", nombre: "Contratos para firmar", id: "1wWsr-hEDMrUl5r2nm1urp4Fj7UxJOsy9" }
  ]);

  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];

    const creado = fila[iCreado];
    if (!fila[iNombre] || esFilaVacia_(fila) || valorBooleano_(creado)) continue;

    if (!esFechaValida_(fila[iFechaInicio]) || !esFechaValida_(fila[iFechaFin]) || !esFechaValida_(fila[iFechaFirma])) {
      Logger.log(`Fila ${i + 1}: fecha invalida, se omite la generacion.`);
      continue;
    }

    const nombre = String(fila[iNombre]).trim();
    const cedula = extraerSoloDigitos_(fila[iCedula]);
    const salarioNum = parseInt(String(fila[iSalario]).replace(/[^\d]/g, '')) || 0;
    const salarioTexto = convertirNumeroALetras(salarioNum).toUpperCase();
    const salarioFinal = `${salarioTexto} ($${salarioNum.toLocaleString("es-CO")})`;

    const fechaInicio = formatearFecha(fila[iFechaInicio]);
    const fechaFin = formatearFecha(fila[iFechaFin]);
    const fechaFirma = formatearFecha(fila[iFechaFirma]);
    const fechaCompacta = fechaFormatoCompacto(fila[iFechaFirma]);

    const nombreArchivo = cedula
      ? `${nombre} - ${cedula} - ${fechaCompacta}`
      : `${nombre} - ${fechaCompacta}`;

    eliminarArchivosPreviosContrato_(carpetaDestino, nombre, cedula);

    // Crear nuevo contrato
    const copiaDoc = DriveApp.getFileById(plantillaId).makeCopy(nombreArchivo);
    const documento = DocumentApp.openById(copiaDoc.getId());
    const body = documento.getBody();

    body.replaceText('{{NOMBRE_TRABAJADOR}}', nombre);
    body.replaceText('{{CEDULA}}', String(fila[iCedula] || ""));
    body.replaceText('{{FECHA_INICIO}}', fechaInicio);
    body.replaceText('{{FECHA_FIN}}', fechaFin);
    body.replaceText('{{DIRECCION}}', String(fila[iDireccion] || ""));
    body.replaceText('{{TELEFONO}}', String(fila[iTelefono] || ""));
    body.replaceText('{{CORREO}}', String(fila[iCorreo] || ""));
    body.replaceText('{{SALARIO}}', salarioFinal);
    body.replaceText('{{FECHA_FIRMA}}', fechaFirma);
    body.replaceText('{{FUNCIONES_ESPECIFICAS}}', String(fila[iFunciones] || ""));

    documento.saveAndClose();

    // Convertir a PDF y mover a carpeta final
    const pdfBlob = DriveApp.getFileById(copiaDoc.getId()).getAs(MimeType.PDF).setName(nombreArchivo + '.pdf');
    const archivoPDF = carpetaDestino.createFile(pdfBlob);
    DriveApp.getFileById(copiaDoc.getId()).setTrashed(true); // Eliminar el archivo .docx temporal

    // Actualizar hoja
    hoja.getRange(i + 1, iCreado + 1).setValue(true);
    hoja.getRange(i + 1, iLink + 1).setValue(archivoPDF.getUrl());
  }
}




function convertirNumeroALetras(numero) {
  var UNIDADES = ["", "uno", "dos", "tres", "cuatro", "cinco", "seis", "siete", "ocho", "nueve"];
  var DECENAS = ["", "diez", "veinte", "treinta", "cuarenta", "cincuenta", "sesenta", "setenta", "ochenta", "noventa"];
  var CENTENAS = ["", "ciento", "doscientos", "trescientos", "cuatrocientos", "quinientos", "seiscientos", "setecientos", "ochocientos", "novecientos"];

  function convertirMenorMil(n) {
    var centenas = Math.floor(n / 100);
    var decenas = Math.floor((n % 100) / 10);
    var unidades = n % 10;
    var texto = "";

    if (n === 100) return "cien";
    if (centenas > 0) texto += CENTENAS[centenas] + " ";
    if (decenas === 1 && unidades > 0) {
      var especiales = ["once", "doce", "trece", "catorce", "quince", "dieciséis", "diecisiete", "dieciocho", "diecinueve"];
      return texto + especiales[unidades - 1];
    } else if (decenas > 0) {
      texto += DECENAS[decenas];
      if (unidades > 0) texto += " y " + UNIDADES[unidades];
    } else {
      texto += UNIDADES[unidades];
    }
    return texto.trim();
  }

  function seccion(n, divisor, singular, plural) {
    var cientos = Math.floor(n / divisor);
    var resto = n - (cientos * divisor);
    var texto = "";

    if (cientos === 0) return "";
    if (cientos === 1) texto = singular;
    else texto = convertirNumero(cientos) + " " + plural;

    return texto;
  }

  function convertirNumero(n) {
    if (n === 0) return "cero";

    var millones = seccion(n, 1000000, "un millón", "millones");
    var miles = seccion(n % 1000000, 1000, "mil", "mil");
    var resto = convertirMenorMil(n % 1000);

    var resultado = [millones, miles, resto].filter(Boolean).join(" ");
    return resultado.trim();
  }

  return convertirNumero(numero) + " pesos";
}

function formatearFecha(fecha) {
  if (!esFechaValida_(fecha)) {
    return "";
  }

  const meses = [
    "enero", "febrero", "marzo", "abril", "mayo", "junio",
    "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"
  ];
  const f = new Date(fecha);
  const dia = f.getDate();
  const mes = meses[f.getMonth()];
  const año = f.getFullYear();
  return `${dia} de ${mes} de ${año}`;
}

function fechaFormatoCompacto(fecha) {
  if (!esFechaValida_(fecha)) {
    return "";
  }

  const f = new Date(fecha);
  const año = String(f.getFullYear()).slice(-2);
  const mes = String(f.getMonth() + 1).padStart(2, '0');
  const dia = String(f.getDate()).padStart(2, '0');
  return `${año}${mes}${dia}`;
}

function obtenerIndicesEncabezados_(encabezados, requeridos) {
  const indices = {};
  const faltantes = [];

  requeridos.forEach(function(nombre) {
    const idx = encabezados.indexOf(nombre);
    if (idx === -1) {
      faltantes.push(nombre);
      return;
    }
    indices[nombre] = idx;
  });

  if (faltantes.length > 0) {
    throw new Error("Faltan encabezados obligatorios: " + faltantes.join(", "));
  }

  return indices;
}

function esFilaVacia_(fila) {
  return fila.every(function(celda) {
    return celda === "" || celda === null;
  });
}

function valorBooleano_(valor) {
  return valor === true || String(valor).toUpperCase() === "TRUE";
}

function esFechaValida_(valor) {
  const fecha = new Date(valor);
  return !isNaN(fecha.getTime());
}

function extraerSoloDigitos_(valor) {
  return String(valor || "").replace(/\D/g, "");
}

function normalizarTexto_(valor) {
  return String(valor || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

function coincideNombreArchivo_(nombreArchivo, nombre) {
  const archivoNorm = " " + normalizarTexto_(nombreArchivo) + " ";
  const nombreNorm = normalizarTexto_(nombre);
  if (!nombreNorm) {
    return false;
  }
  return archivoNorm.indexOf(" " + nombreNorm + " ") !== -1;
}

function eliminarArchivosPreviosContrato_(carpetaDestino, nombre, cedula) {
  const archivosAntiguos = carpetaDestino.getFiles();
  while (archivosAntiguos.hasNext()) {
    const archivo = archivosAntiguos.next();
    const nombreArchivo = archivo.getName();
    const coincideCedula = cedula && nombreArchivo.indexOf(cedula) !== -1;
    const coincideNombre = coincideNombreArchivo_(nombreArchivo, nombre);
    if (coincideCedula || coincideNombre) {
      archivo.setTrashed(true);
    }
  }
}

function validarAccesoDrive_(recursos) {
  recursos.forEach(function(recurso) {
    try {
      if (recurso.tipo === "archivo") {
        DriveApp.getFileById(recurso.id).getName();
        return;
      }
      if (recurso.tipo === "carpeta") {
        DriveApp.getFolderById(recurso.id).getName();
        return;
      }
      throw new Error("Tipo de recurso no soportado: " + recurso.tipo);
    } catch (e) {
      throw new Error(
        "Sin acceso a " + recurso.tipo + " [" + recurso.nombre + "] con ID " + recurso.id +
        ". Verifica permisos de la cuenta ejecutora. Detalle: " + e.message
      );
    }
  });
}
