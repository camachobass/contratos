function generarContratos() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Contratos");
  const datos = hoja.getDataRange().getValues();
  const encabezados = datos[0];

  const iNombre = encabezados.indexOf("Nombre trabajador");
  const iCedula = encabezados.indexOf("Cédula");
  const iFechaInicio = encabezados.indexOf("Fecha inicio");
  const iFechaFin = encabezados.indexOf("Fecha fin");
  const iDireccion = encabezados.indexOf("Dirección trabajador");
  const iTelefono = encabezados.indexOf("Teléfono");
  const iCorreo = encabezados.indexOf("Correo");
  const iSalario = encabezados.indexOf("Salario");
  const iFechaFirma = encabezados.indexOf("Fecha firma");
  const iFunciones = encabezados.indexOf("Funciones Especiales");
  const iCreado = encabezados.indexOf("Creado");
  const iLink = encabezados.indexOf("Link");

  const plantillaId = '1pBCfXUOR-bvDe9a9-NKidd_UZajNjaKs13ff9ynn770';
  const carpetaDestino = DriveApp.getFolderById('1wWsr-hEDMrUl5r2nm1urp4Fj7UxJOsy9');

  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];

    const creado = fila[iCreado];
    if (!fila[iNombre] || fila.every(celda => celda === "") || creado === true) continue;

    const nombre = fila[iNombre];
    const salarioNum = parseInt(String(fila[iSalario]).replace(/[^\d]/g, '')) || 0;
    const salarioTexto = convertirNumeroALetras(salarioNum).toUpperCase();
    const salarioFinal = `${salarioTexto} ($${salarioNum.toLocaleString("es-CO")})`;

    const fechaInicio = formatearFecha(fila[iFechaInicio]);
    const fechaFin = formatearFecha(fila[iFechaFin]);
    const fechaFirma = formatearFecha(fila[iFechaFirma]);
    const fechaCompacta = fechaFormatoCompacto(fila[iFechaFirma]);

    const nombreArchivo = `${nombre} - ${fechaCompacta}`;

    // 🔄 Eliminar archivo anterior si ya existe
    const archivosAntiguos = carpetaDestino.getFiles();
    while (archivosAntiguos.hasNext()) {
      const archivo = archivosAntiguos.next();
      const nombreArchivoExistente = archivo.getName().toLowerCase();
      if (nombreArchivoExistente.includes(nombre.toLowerCase())) {
        archivo.setTrashed(true);
      }
    }

    // Crear nuevo contrato
    const copiaDoc = DriveApp.getFileById(plantillaId).makeCopy(nombreArchivo);
    const documento = DocumentApp.openById(copiaDoc.getId());
    const body = documento.getBody();

    body.replaceText('{{NOMBRE_TRABAJADOR}}', nombre);
    body.replaceText('{{CEDULA}}', fila[iCedula]);
    body.replaceText('{{FECHA_INICIO}}', fechaInicio);
    body.replaceText('{{FECHA_FIN}}', fechaFin);
    body.replaceText('{{DIRECCION}}', fila[iDireccion]);
    body.replaceText('{{TELEFONO}}', fila[iTelefono]);
    body.replaceText('{{CORREO}}', fila[iCorreo]);
    body.replaceText('{{SALARIO}}', salarioFinal);
    body.replaceText('{{FECHA_FIRMA}}', fechaFirma);
    body.replaceText('{{FUNCIONES_ESPECIFICAS}}', fila[iFunciones]);

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
  const f = new Date(fecha);
  const año = String(f.getFullYear()).slice(-2);
  const mes = String(f.getMonth() + 1).padStart(2, '0');
  const dia = String(f.getDate()).padStart(2, '0');
  return `${año}${mes}${dia}`;
}
