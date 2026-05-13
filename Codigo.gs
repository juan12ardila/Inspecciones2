const CONFIG = {
  SHEET_INSPECCIONES:    "Inspecciones",
  SHEET_FLOTA:           "VIGENTES",
  ID_LIBRO_FLOTA:        "1rGSNV5PlQk6m-Htzw2mBdLf9Qg4It_c37qWXk2IV6qM",
  GID_CONDUCTORES:       "430651520",
  COL_PLACA_COND:        "Reg Plate",
  COL_DRIVER_NAME:       "Driver Fullname",
  COL_DRIVER_ID:         "Driver National Id Number",
  ID_LIBRO_INSPECCIONES: "1otoY1zTv2nMyqBbfDx2-PdAgHu17eU8fubYwaxXjXyA",
  GID_INSPECCIONES:      "1691322766",
  FOLDER_FOTOS:          "1VwJzrBk4oOcgg2xJ3WYuv_cfpkXupEv7",
  MAX_CELL_CHARS:        45000
};

function doGet(e) {
  return HtmlService
    .createHtmlOutputFromFile("Index")
    .setTitle("Inspección Preoperacional · Cabify")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag("viewport","width=device-width,initial-scale=1.0,maximum-scale=1.0,user-scalable=no");
}

function getSheetByGid(ss, gid) {
  for (const s of ss.getSheets()) {
    if (String(s.getSheetId()) === String(gid)) return s;
  }
  return null;
}

function getInspeccionesSheet(crear) {
  const ss    = SpreadsheetApp.openById(CONFIG.ID_LIBRO_INSPECCIONES);
  let   sheet = getSheetByGid(ss, CONFIG.GID_INSPECCIONES);
  if (!sheet && crear) {
    sheet = ss.insertSheet(CONFIG.SHEET_INSPECCIONES);
    const headers = [
      "ID_Inspeccion","Timestamp","Fecha","Hora",
      "Placa","Marca","Modelo","Año","Color","Tipo Servicio",
      "Nombre Inspector","Cargo","Conductor del Día","Conductor Nuevo",
      "Ciudad","Kilometraje actual",
      "Luces Delanteras","Luces Traseras","Luces Stop","Direccionales",
      "Frenos Principales","Freno de Mano","Nivel Líquido Frenos",
      "Llanta Del. Izq","Llanta Del. Der","Llanta Tras. Izq","Llanta Tras. Der","Llanta Repuesto",
      "Espejos Laterales","Espejo Retrovisor","Limpiabrisas","Bocina",
      "Cinturones Seguridad","Extintor","Botiquín","Chaleco Reflectivo",
      "Conos/Triángulos","Kit de Carretera",
      "Motor (Fugas)","Nivel Aceite","Nivel Refrigerante",
      "Estado Carrocería","Vidrios","Puertas","Tapicería Interior",
      "SOAT Estado","RCC Estado","RCE Estado","RTM Vigente",
      "Observaciones Generales","Zonas con Daño","Fecha Venc. Extintor",
      "Foto_Evidencia_URL","Foto_Evidencia_Nombre","Resultado Final",
      "Es_Electrico","Nivel_Carga_Pct",
      "check_bateria_estado","check_puerto_carga","check_cable_carga",
      "check_refrig_bateria","check_freno_regen","check_dashboard_warn",
      "Firma_URL","Firma_Conductor_URL","Firma_Base64","Firma_Conductor_Base64"
    ];
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    sheet.getRange(1,1,1,headers.length)
      .setBackground("#7B2FF7").setFontColor("#FFFFFF").setFontWeight("bold");
  }
  return sheet;
}

function guardarFirmaEnDrive(base64Data, nombreArchivo) {
  try {
    if (!base64Data || base64Data === '' || !base64Data.startsWith('data:')) return '';
    const parts    = base64Data.split(',');
    const mimeType = parts[0].match(/:(.*?);/)[1];
    const ext      = mimeType.includes('png') ? 'png' : 'jpg';
    const bytes    = Utilities.base64Decode(parts[1]);
    const blob     = Utilities.newBlob(bytes, mimeType, nombreArchivo + '.' + ext);
    const folder   = DriveApp.getFolderById(CONFIG.FOLDER_FOTOS);
    const file     = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return 'https://drive.google.com/uc?export=view&id=' + file.getId();
  } catch(e) {
    Logger.log('guardarFirmaEnDrive ERROR [' + nombreArchivo + ']: ' + e.message);
    return '';
  }
}

function guardarInspeccion(data) {
  const sheet    = getInspeccionesSheet(true);
  const now      = new Date();
  const rand     = Math.random().toString(36).substring(2, 6).toUpperCase();
  const idInsp   = "INS-" + Utilities.formatDate(now, "America/Bogota", "yyyyMMdd-HHmmss") + "-" + rand;

  let fotoUrl = '', fotoNombre = '';
  if (data.fotoEvidencia && data.fotoEvidencia !== '') {
    const fotoRes = guardarFotoEvidencia(data.fotoEvidencia, idInsp, data.placa || '');
    fotoUrl    = fotoRes.url    || '';
    fotoNombre = fotoRes.nombre || '';
  }

  let firmaUrl = '', firmaConductorUrl = '';
  let firmaBase64Truncado = '', firmaConductorBase64Truncado = '';

  if (data.firma && data.firma !== '') {
    firmaUrl = guardarFirmaEnDrive(data.firma, idInsp + '_firma_inspector');
    firmaBase64Truncado = data.firma.length <= CONFIG.MAX_CELL_CHARS
      ? data.firma : data.firma.substring(0, CONFIG.MAX_CELL_CHARS);
  }

  if (data.firmaConductor && data.firmaConductor !== '') {
    firmaConductorUrl = guardarFirmaEnDrive(data.firmaConductor, idInsp + '_firma_conductor');
    firmaConductorBase64Truncado = data.firmaConductor.length <= CONFIG.MAX_CELL_CHARS
      ? data.firmaConductor : data.firmaConductor.substring(0, CONFIG.MAX_CELL_CHARS);
  }

  const fila = [
    idInsp, now,
    Utilities.formatDate(now, "America/Bogota", "dd/MM/yyyy"),
    Utilities.formatDate(now, "America/Bogota", "HH:mm:ss"),
    data.placa||"", data.marca||"", data.modelo||"", data.anio||"",
    data.color||"", data.tipoServicio||"",
    data.nombreInspector||"", data.cargo||"",
    data.conductorDia||"", data.conductorNuevo||"",
    data.ciudad||"",
    data.kilometraje ? Number(data.kilometraje) : "",
    data.check_luces_del||"N/A",   data.check_luces_tra||"N/A",
    data.check_luces_stop||"N/A",  data.check_direccionales||"N/A",
    data.check_frenos||"N/A",      data.check_freno_mano||"N/A",
    data.check_liq_frenos||"N/A",
    data.check_llanta_di||"N/A",   data.check_llanta_dd||"N/A",
    data.check_llanta_ti||"N/A",   data.check_llanta_td||"N/A",
    data.check_llanta_rep||"N/A",
    data.check_espejos_lat||"N/A", data.check_espejo_ret||"N/A",
    data.check_limpiabrisas||"N/A",data.check_bocina||"N/A",
    data.check_cinturones||"N/A",  data.check_extinguidor||"N/A",
    data.check_botiquin||"N/A",    data.check_chaleco||"N/A",
    data.check_conos||"N/A",       data.check_kit_carretera||"N/A",
    data.check_motor||"N/A",       data.check_aceite||"N/A",
    data.check_refrigerante||"N/A",
    data.check_carroceria||"N/A",  data.check_vidrios||"N/A",
    data.check_puertas||"N/A",     data.check_tapiceria||"N/A",
    data.soat_estado||"", data.rcc_estado||"",
    data.rce_estado||"",  data.rtm_vigente||"",
    data.observaciones||"", data.zonasDanio||"",
    data.fechaVencExtinguidor||"",
    fotoUrl, fotoNombre, data.resultado||"",
    data.esElectrico||"NO", data.nivelCarga||"",
    data.check_bateria_estado||"N/A", data.check_puerto_carga||"N/A",
    data.check_cable_carga||"N/A",    data.check_refrig_bateria||"N/A",
    data.check_freno_regen||"N/A",    data.check_dashboard_warn||"N/A",
    firmaUrl, firmaConductorUrl,
    firmaBase64Truncado, firmaConductorBase64Truncado
  ];

  sheet.appendRow(fila);

  const lastRow = sheet.getLastRow();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const colRes  = headers.indexOf("Resultado Final") + 1;
  if (colRes > 0) {
    const cell = sheet.getRange(lastRow, colRes);
    if      (data.resultado === "APROBADO")  cell.setBackground("#d4edda").setFontColor("#155724").setFontWeight("bold");
    else if (data.resultado === "RECHAZADO") cell.setBackground("#f8d7da").setFontColor("#721c24").setFontWeight("bold");
    else                                     cell.setBackground("#fff3cd").setFontColor("#856404").setFontWeight("bold");
  }

  return { ok: true, id: idInsp, fotoUrl: fotoUrl, fotoNombre: fotoNombre };
}

function resolverFirma(urlDrive, base64Fallback) {
  const u = String(urlDrive || '').trim();
  if (u && u.startsWith('https://')) return u;
  const b = String(base64Fallback || '').trim();
  if (b && b.startsWith('data:image')) return b;
  return '';
}

function getInspeccionesByPlaca(placa) {
  try {
    const ss    = SpreadsheetApp.openById(CONFIG.ID_LIBRO_INSPECCIONES);
    const sheet = getSheetByGid(ss, CONFIG.GID_INSPECCIONES);
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];

    const headers      = data[0].map(h => String(h).trim());
    const colPlaca     = headers.indexOf("Placa");
    if (colPlaca < 0) return [];

    const colFirmaUrl    = headers.indexOf("Firma_URL");
    const colFirmaCndUrl = headers.indexOf("Firma_Conductor_URL");
    const colFirmaB64    = headers.indexOf("Firma_Base64");
    const colFirmaCndB64 = headers.indexOf("Firma_Conductor_Base64");

    const placaBuscar = placa.trim().toUpperCase();
    const resultados  = [];

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][colPlaca]).trim().toUpperCase() !== placaBuscar) continue;
      const obj = {};
      headers.forEach((h, j) => {
        const v = data[i][j];
        if (v instanceof Date) {
          if      (h === 'Timestamp') obj[h] = Utilities.formatDate(v, "America/Bogota", "dd/MM/yyyy HH:mm:ss");
          else if (h === 'Hora')      obj[h] = Utilities.formatDate(v, "America/Bogota", "HH:mm:ss");
          else if (h === 'Fecha')     obj[h] = Utilities.formatDate(v, "America/Bogota", "dd/MM/yyyy");
          else                        obj[h] = Utilities.formatDate(v, "America/Bogota", "dd/MM/yyyy HH:mm");
        } else {
          obj[h] = String(v === null || v === undefined ? "" : v);
        }
      });
      obj['Firma_Base64'] = resolverFirma(
        colFirmaUrl    >= 0 ? String(data[i][colFirmaUrl]    || '') : '',
        colFirmaB64    >= 0 ? String(data[i][colFirmaB64]    || '') : ''
      );
      obj['Firma_Conductor_Base64'] = resolverFirma(
        colFirmaCndUrl >= 0 ? String(data[i][colFirmaCndUrl] || '') : '',
        colFirmaCndB64 >= 0 ? String(data[i][colFirmaCndB64] || '') : ''
      );
      resultados.push(obj);
    }
    return resultados.reverse();
  } catch(e) {
    Logger.log('getInspeccionesByPlaca ERROR: ' + e.message);
    return [];
  }
}

function buscarVehiculo(placa) {
  const ss    = SpreadsheetApp.openById(CONFIG.ID_LIBRO_FLOTA);
  const sheet = ss.getSheetByName(CONFIG.SHEET_FLOTA);
  if (!sheet) return null;

  const data        = sheet.getDataRange().getValues();
  const headers     = data[0].map(h => String(h).trim().toLowerCase());
  const placaBuscar = placa.trim().toUpperCase().replace(/\s/g, "");

  for (let i = 1; i < data.length; i++) {
    const placaSheet = String(data[i][0]).trim().toUpperCase().replace(/\s/g, "");
    if (placaSheet !== placaBuscar) continue;

    const row = {};
    headers.forEach((h, j) => { row[h] = data[i][j]; });

    function fmt(val) {
      if (!val || val === "") return "—";
      if (val instanceof Date) return Utilities.formatDate(val, "America/Bogota", "dd/MM/yyyy");
      return String(val);
    }

    return {
      placa:               placaBuscar,
      marca:               row["marca"]         || row["brand"]        || "",
      modelo:              row["linea"]         || row["model"]        || row["line"] || "",
      anio:                row["modelo"]        || row["year"]         || row["año"]  || "",
      color:               row["color"]         || "",
      tipoServicio:        row["tipo_servicio"] || row["service_type"] || row["tipo servicio"] || "",
      clase:               row["clase"]         || row["class"]        || "",
      soat_estado:         row["soat_estado"]         || "—",
      soat_estado_detalle: row["soat_estado_detalle"] || "—",
      soat_vencimiento:    fmt(row["soat_vencimiento"]),
      soat_expedicion:     fmt(row["soat_fecha_expedicion"]),
      soat_aseguradora:    row["soat_aseguradora"]    || "—",
      soat_poliza:         row["soat_poliza"]          || "—",
      rtm_vigente:         row["rtm_vigente"]          || "NO",
      rtm_vencimiento:     fmt(row["rtm_vencimiento"]),
      rtm_expedicion:      fmt(row["rtm_fecha_expedicion"]),
      rtm_cda:             row["rtm_cda"]              || "—",
      rtm_certificado:     row["rtm_certificado"]      || "—",
      rcc_estado:          row["rcc_estado"]           || "—",
      rcc_vencimiento:     fmt(row["rcc_vencimiento"]),
      rcc_expedicion:      fmt(row["rcc_fecha_expedicion"]),
      rcc_aseguradora:     row["rcc_aseguradora"]      || "—",
      rcc_poliza:          row["rcc_poliza"]            || "—",
      rce_estado:          row["rce_estado"]            || "—",
      rce_vencimiento:     fmt(row["rce_vencimiento"]),
      rce_expedicion:      fmt(row["rce_fecha_expedicion"]),
      rce_aseguradora:     row["rce_aseguradora"]      || "—",
      rce_poliza:          row["rce_poliza"]            || "—",
      to_estado:           row["to_estado"]             || "—",
      to_vencimiento:      fmt(row["to_vencimiento"]),
      to_entidad:          row["to_entidad"]            || "—",
      to_modalidad:        row["to_modalidad"]          || "—",
      to_numero:           row["to_numero"]             || "—",
      to_servicio:         row["to_servicio"]           || "—",
      ciudad:              getCiudadByPlaca(placaBuscar),
      combustible:         row["combustible"]           || "",
      fecha_matricula:     fmt(row["fecha_matricula"])  || "",
    };
  }
  return null;
}

function getCiudadByPlaca(placa) {
  try {
    const ss      = SpreadsheetApp.openById(CONFIG.ID_LIBRO_FLOTA);
    const sheet   = getSheetByGid(ss, CONFIG.GID_CONDUCTORES);
    if (!sheet) return "";
    const data    = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim());

    const placaVariants  = [CONFIG.COL_PLACA_COND,"Reg Plate","reg plate","plate","Placa","PLACA","Plate"];
    const regionVariants = ["Region","region","REGION","Ciudad","ciudad","City","city"];

    let idxPlaca = -1, idxRegion = -1;
    for (const v of placaVariants)  { idxPlaca  = headers.findIndex(h => h.toLowerCase() === v.toLowerCase()); if (idxPlaca  >= 0) break; }
    for (const v of regionVariants) { idxRegion = headers.findIndex(h => h.toLowerCase() === v.toLowerCase()); if (idxRegion >= 0) break; }
    if (idxPlaca < 0 || idxRegion < 0) return "";

    const placaBuscar = placa.trim().toUpperCase().replace(/\s/g, "");
    for (let i = data.length - 1; i >= 1; i--) {
      const pr = String(data[i][idxPlaca]).trim().toUpperCase().replace(/\s/g, "");
      if (pr === placaBuscar) {
        const r = String(data[i][idxRegion]).trim();
        if (r && r.toLowerCase() !== "undefined") return r;
      }
    }
    return "";
  } catch(e) { return ""; }
}

// Normaliza una cédula que puede venir como número (13498732), texto ("13498732"),
// número con decimales ("13498732.0") o notación científica ("1.3498732E7")
function normalizarCedula(raw) {
  if (raw === null || raw === undefined) return '';
  // Si es número, convertirlo evitando notación científica
  if (typeof raw === 'number') {
    if (!isFinite(raw)) return '';
    // toFixed(0) elimina decimales; redondea si los hubiera
    return String(Math.round(raw));
  }
  let s = String(raw).trim();
  if (!s || s.toLowerCase() === 'undefined' || s.toLowerCase() === 'null') return '';
  // Quitar .0 final si vino como string desde un número decimal
  s = s.replace(/\.0+$/, '');
  // Si tiene notación científica intentar parsearlo
  if (/^-?\d+(\.\d+)?[eE][+-]?\d+$/.test(s)) {
    const n = Number(s);
    if (isFinite(n)) return String(Math.round(n));
  }
  return s;
}

function getConductoresByPlaca(placa) {
  try {
    const ss    = SpreadsheetApp.openById(CONFIG.ID_LIBRO_FLOTA);
    const sheet = getSheetByGid(ss, CONFIG.GID_CONDUCTORES);
    if (!sheet) return [];
    const data    = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim());

    // Placa: case-insensitive con variantes
    const placaVariants = [CONFIG.COL_PLACA_COND, "Reg Plate", "Plate", "Placa", "PLACA"];
    let idxPlaca = -1;
    for (const v of placaVariants) {
      idxPlaca = headers.findIndex(h => h.toLowerCase() === v.toLowerCase());
      if (idxPlaca >= 0) break;
    }

    // Nombre: case-insensitive con variantes
    const nameVariants = [CONFIG.COL_DRIVER_NAME, "Driver Fullname", "Driver Full Name", "Driver Name", "Conductor", "Nombre Conductor"];
    let idxDriver = -1;
    for (const v of nameVariants) {
      idxDriver = headers.findIndex(h => h.toLowerCase() === v.toLowerCase());
      if (idxDriver >= 0) break;
    }

    // Cédula: case-insensitive con variantes amplias
    const idVariants = [
      CONFIG.COL_DRIVER_ID,
      "Driver National Id Number", "Driver National ID Number",
      "National Id Number", "National ID Number",
      "Driver Id", "Driver ID", "Driver Document",
      "Cedula", "Cédula", "Documento", "Identificacion", "Identificación"
    ];
    let idxCedula = -1;
    for (const v of idVariants) {
      idxCedula = headers.findIndex(h => h.toLowerCase() === v.toLowerCase());
      if (idxCedula >= 0) break;
    }

    Logger.log('getConductoresByPlaca: idxPlaca=' + idxPlaca + ' idxDriver=' + idxDriver + ' idxCedula=' + idxCedula);
    Logger.log('headers: ' + JSON.stringify(headers));

    if (idxPlaca < 0 || idxDriver < 0) return [];

    const placaBuscar = placa.trim().toUpperCase().replace(/\s/g, "");
    const conductores = new Map(); // clave nombre+cedula → display

    for (let i = 1; i < data.length; i++) {
      const pr = String(data[i][idxPlaca]).trim().toUpperCase().replace(/\s/g, "");
      if (pr !== placaBuscar) continue;

      const nombre = String(data[i][idxDriver] || '').trim();
      if (!nombre || nombre.toLowerCase() === 'undefined') continue;

      const cedula = idxCedula >= 0 ? normalizarCedula(data[i][idxCedula]) : '';
      const display = cedula ? (nombre + ' - ' + cedula) : nombre;
      const key = (nombre + '|' + cedula).toLowerCase();
      if (!conductores.has(key)) conductores.set(key, display);
    }
    return Array.from(conductores.values()).sort();
  } catch(e) {
    Logger.log('getConductoresByPlaca ERROR: ' + e.message + ' | ' + e.stack);
    return [];
  }
}

function getUltimoKm(placa) {
  const sheet = getInspeccionesSheet(false);
  if (!sheet) return null;
  const data    = sheet.getDataRange().getValues();
  if (data.length < 2) return null;
  const headers  = data[0].map(h => String(h).trim());
  const colPlaca = headers.indexOf("Placa");
  const colKm    = headers.indexOf("Kilometraje actual");
  const colFecha = headers.indexOf("Fecha");
  if (colPlaca < 0 || colKm < 0) return null;
  const placaBuscar = placa.trim().toUpperCase().replace(/\s/g, "");
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][colPlaca]).trim().toUpperCase().replace(/\s/g, "") === placaBuscar) {
      const km = Number(data[i][colKm]);
      if (!isNaN(km) && km > 0) return { km, fecha: colFecha >= 0 ? String(data[i][colFecha]) : "" };
    }
  }
  return null;
}

// Devuelve { url, nombre } o { url:'', nombre:'' } si falla.
// Importante: NO traga el error silenciosamente — lo registra y devuelve un mensaje
// trazable en el campo nombre para que el usuario pueda diagnosticar.
function guardarFotoEvidencia(base64Data, idInsp, placa) {
  try {
    if (!base64Data || base64Data === '') return { url: '', nombre: '' };
    const dataStr = String(base64Data);
    if (!dataStr.startsWith('data:')) {
      Logger.log('guardarFotoEvidencia: no inicia con data: → ' + dataStr.substring(0, 40));
      return { url: '', nombre: '' };
    }
    const commaIdx = dataStr.indexOf(',');
    if (commaIdx < 0) {
      Logger.log('guardarFotoEvidencia: sin coma separadora');
      return { url: '', nombre: '' };
    }
    const header = dataStr.substring(0, commaIdx);
    const payload = dataStr.substring(commaIdx + 1);
    const mimeMatch = header.match(/:(.*?);/);
    const mimeType = (mimeMatch && mimeMatch[1]) ? mimeMatch[1] : 'image/jpeg';
    const ext = mimeType.indexOf('png') >= 0 ? 'png' : 'jpg';
    const placaLimpia = String(placa || 'SIN_PLACA').replace(/[^A-Za-z0-9]/g, '');
    const nombreArchivo = idInsp + '_' + placaLimpia + '_evidencia.' + ext;
    const bytes = Utilities.base64Decode(payload);
    const blob = Utilities.newBlob(bytes, mimeType, nombreArchivo);
    const folder = DriveApp.getFolderById(CONFIG.FOLDER_FOTOS);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const url = 'https://drive.google.com/uc?export=view&id=' + file.getId();
    Logger.log('guardarFotoEvidencia OK: ' + nombreArchivo + ' → ' + url);
    return { url: url, nombre: nombreArchivo };
  } catch(e) {
    Logger.log('guardarFotoEvidencia ERROR: ' + e.message + ' | stack: ' + (e.stack || ''));
    // Devolver mensaje de error visible en la hoja para diagnóstico
    return { url: '', nombre: 'ERROR: ' + e.message };
  }
}

function getImagenBase64(fileId) {
  try {
    const file     = DriveApp.getFileById(fileId);
    const blob     = file.getBlob();
    const base64   = Utilities.base64Encode(blob.getBytes());
    const mimeType = blob.getContentType();
    return 'data:' + mimeType + ';base64,' + base64;
  } catch(e) {
    Logger.log('getImagenBase64 ERROR: ' + e.message);
    return '';
  }
}

function probarConductores() {
  const resultado = getConductoresByPlaca("NOK261");
  Logger.log('Cantidad: ' + resultado.length);
  Logger.log('Resultado: ' + JSON.stringify(resultado));
}
