// FONCTION POUR OUVRIR LA PAGE HTML SUR MON TÉLÉPHONE





function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index');
}

function getPage(pageName) {
  return HtmlService.createHtmlOutputFromFile(pageName).getContent();
}

function listarSectores() {
  try {
    var sectores = leerSectores_();
    return { ok: true, sectores: sectores };
  } catch (e) {
    return {
      ok: false,
      error: e && e.message ? e.message : 'Error al leer sectores.'
    };
  }
}






// FONCTIONS POUR CRÉER DES DONNÉES TEMPORAIRES





function test_crearPedidos() {
  var sheet = getSheet_('PEDIDOS DIARIOS');
  
  // données de test sans la colonne F
  var pedidos = [
    ['2025-11-03', 'Luis García', '+593987654321', 'Cerca del puente', 2, 'Sí', '12:30', 'Cliente habitual', 'Pendiente'],
    ['2025-11-03', 'Marta López', '+593999888777', 'Junto al coliseo', 1, 'No', '12:45', 'Prefiere sin ensalada', 'Pendiente'],
    ['2025-11-03', 'José Torres', '+593983333222', 'Frente a la iglesia', 3, 'Sí', '12:15', '', 'Pendiente'],
    ['2025-11-03', 'Patricia Ramos', '+593982222111', 'Calle Bolívar', 2, 'No', '13:00', 'Pagar al recibir', 'Pendiente'],
    ['2025-11-03', 'Carlos Méndez', '+593981111000', 'Barrio Central', 4, 'Sí', '12:20', '', 'Pendiente']
  ];

  var startRow = getFirstEmptyRowInColumn_(sheet, 1);

  // Colonnes ciblées: A–E puis G–J
  var ranges = [
    sheet.getRange(startRow, 1, pedidos.length, 5), // A–E
    sheet.getRange(startRow, 7, pedidos.length, 4)  // G–J
  ];

  // On sépare les blocs de données pour ne pas toucher F
  var bloqueAE = pedidos.map(function (p) { return p.slice(0, 5); });
  var bloqueGJ = pedidos.map(function (p) { return p.slice(5); });

  ranges[0].setValues(bloqueAE);
  ranges[1].setValues(bloqueGJ);
}

function test_actualizarMenu() {
  actualizarMenu({
    nombre: 'Seco de pollo',
    descripcion: 'Receta mejorada',
    ingredientes: 'pollo; arroz; ensalada; ají',
    activo: 'Sí'
  });
}

function test_baseClientes() {
  // 1er passage: crée le client
  actualizarUltimoPedidoCliente({
    nombre: 'Luis García',
    telefono: '+593987654321',
    fecha: '2025-11-03',
    direccion: 'Cerca del puente',
    notas: 'Primera compra'
  });

  // 2e passage: même client, autre date, autre adresse, autre note
  actualizarUltimoPedidoCliente({
    nombre: 'Luis García',
    telefono: '+593987654321',
    fecha: '2025-11-04',
    direccion: 'Oficina municipal',
    notas: 'Entregar 12:30'
  });
}





// FINCTIONS DE LA PHASE A




// GUARDAR DATOS





/**
 * Guarda un pedido recibido (normalmente desde un formulario HTML).
 * data debe contener, mínimo:
 * {
 *   fecha: '2025-11-01',
 *   cliente: 'Luis',
 *   telefono: '+593...',
 *   direccion: '...'
 *   cantidad: 2,
 *   horaEntrega: '12:30',
 *   notas: '...'
 * }
 */
function guardarPedido(data) {
  var sheet = getSheet_('PEDIDOS DIARIOS');
  // on cherche la première ligne VRAIMENT vide en colonne A
  var row = getFirstEmptyRowInColumn_(sheet, 1);

  // columnas según la hoja:
  // A Fecha (día del menú)
  // B Cliente
  // C Teléfono / WhatsApp
  // D Dirección / Referencia
  // E Cantidad
  // F Total ($) -> lo calcula la hoja
  // G Pago recibido
  // H Hora de entrega
  // I Notas / Comentarios
  // J Estado

  sheet.getRange(row, 1).setValue(data && data.fecha ? data.fecha : new Date());
  sheet.getRange(row, 2).setValue(data && data.cliente ? data.cliente : '');
  sheet.getRange(row, 3).setValue(data && data.telefono ? data.telefono : '');
  sheet.getRange(row, 4).setValue(data && data.direccion ? data.direccion : '');
  sheet.getRange(row, 5).setValue(data && data.cantidad ? data.cantidad : 1);
  // F = fórmula
  sheet.getRange(row, 7).setValue(data && data.pagoRecibido ? data.pagoRecibido : 'No');
  sheet.getRange(row, 8).setValue(data && data.horaEntrega ? data.horaEntrega : '');
  sheet.getRange(row, 9).setValue(data && data.notas ? data.notas : '');
  sheet.getRange(row, 10).setValue(data && data.estado ? data.estado : 'Pendiente');
}

/**
 * Guarda un gasto en la hoja GASTOS con ID_GASTO e ID_PRODUCTO.
 *
 * data:
 * {
 *   fecha: '2025-11-01' | Date | null,
 *   categoria: 'Ingredientes',
 *   producto: 'Pollo',
 *   cantidad: 6,
 *   unidad: 'kg',
 *   precioUnidad: 2,
 *   proveedor: 'Mercado Misahualli',
 *   observaciones: '...',
 *   idProducto: 'PROD-00001' | '' | null
 * }
 *
 * No devuelve JSON (función utilitaria interna).
 */
function guardarGasto(data) {
  var sheet = getSheet_('GASTOS');
  var row = getFirstEmptyRowInColumn_(sheet, 1);

  // Sécurité minimale
  data = data || {};

  // Fecha
  var fechaObj;
  if (data.fecha) {
    fechaObj = new Date(data.fecha);
    if (isNaN(fechaObj.getTime())) {
      fechaObj = new Date();
    }
  } else {
    fechaObj = new Date();
  }

  // Numériques
  var cantidad = Number(data.cantidad || 0);
  var precioUnidad = Number(data.precioUnidad || 0);
  var monto = cantidad * precioUnidad;

  // ID_PRODUCTO optionnel
  var idProducto = (data.idProducto || '').toString().trim();

  // Generar ID_GASTO
  var idGasto = generarIdGasto_();

  // Estructura actual de GASTOS:
  // A Fecha
  // B Mes (fórmula en la hoja)
  // C Categoría de gasto
  // D Producto
  // E Cantidad
  // F Unidad
  // G Precio por unidad
  // H Monto ($)
  // I Proveedor
  // J Observaciones
  // K ID_GASTO
  // L ID_PRODUCTO

  sheet.getRange(row, 1).setValue(fechaObj);                  // A Fecha
  sheet.getRange(row, 3).setValue(data.categoria || '');      // C
  sheet.getRange(row, 4).setValue(data.producto || '');       // D
  sheet.getRange(row, 5).setValue(cantidad);                  // E
  sheet.getRange(row, 6).setValue(data.unidad || '');         // F
  sheet.getRange(row, 7).setValue(precioUnidad);              // G
  sheet.getRange(row, 8).setValue(monto);                     // H
  sheet.getRange(row, 9).setValue(data.proveedor || '');      // I
  sheet.getRange(row, 10).setValue(data.observaciones || ''); // J
  sheet.getRange(row, 11).setValue(idGasto);                  // K ID_GASTO
  sheet.getRange(row, 12).setValue(idProducto);               // L ID_PRODUCTO

  // Actualiser le résumé mensuel comme avant
  actualizarResumenMensual();
}

/**
 * Guarda un ingreso en la hoja INGRESOS.
 * data:
 * {
 *   fecha: '2025-11-01',
 *   totalComidas: 10,
 *   precio: 1
 * }
 */
function guardarIngreso(data) {
  var sheet = getSheet_('INGRESOS');
  var row = getFirstEmptyRowInColumn_(sheet, 1);

  // A Fecha
  // B Mes (fórmula)
  // C Total comidas vendidas
  // D Precio por comida
  // E Ingreso total ($) -> fórmula
  // F Observaciones

  sheet.getRange(row, 1).setValue(data && data.fecha ? data.fecha : new Date());
  sheet.getRange(row, 3).setValue(data && data.totalComidas ? data.totalComidas : 0);
  sheet.getRange(row, 4).setValue(data && data.precio ? data.precio : 1);
  sheet.getRange(row, 6).setValue(data && data.observaciones ? data.observaciones : '');
}





// FERMER LA JOURNÉE





/**
 * Cierra el día: pasa PEDIDOS DIARIOS → HISTORIAL DE PEDIDOS,
 * registra el ingreso del día en INGRESOS, limpia PEDIDOS DIARIOS
 * y actualiza RESUMEN MENSUAL.
 *
 * Ahora devuelve también un ResumenCierreFront según el PAQUETE 6:
 *
 * Respuesta:
 * {
 *   ok: true|false,
 *   error: string|null,
 *   data: {
 *     resumen: {
 *       fecha: string,
 *       totalPedidos: number,
 *       totalSecos: number,
 *       totalIngresos: number|null
 *     }
 *   } | null
 * }
 */
function cerrarDia() {
  try {
    var sheetPedidos = getSheet_('PEDIDOS DIARIOS');
    var sheetHist = getSheet_('HISTORIAL DE PEDIDOS');

    // 1. Lire les pedidos du jour
    var lastRowPedidos = sheetPedidos.getLastRow();
    if (lastRowPedidos < 2) {
      return {
        ok: false,
        error: 'No hay pedidos para cerrar este día.',
        data: null
      };
    }

    var pedidos = sheetPedidos
      .getRange(2, 1, lastRowPedidos - 1, 10)
      .getValues()
      .filter(function (r) {
        // on garde seulement les lignes qui ont au moins une info
        return r[0] || r[1] || r[2] || r[3] || r[4];
      });

    if (pedidos.length === 0) {
      return {
        ok: false,
        error: 'No hay pedidos válidos para cerrar este día.',
        data: null
      };
    }

    // Construir el resumen ANTES de mover nada
    var resumen = construirResumenCierreDesdePedidos_(pedidos);
    if (!resumen) {
      return {
        ok: false,
        error: 'No se pudo construir el resumen de cierre.',
        data: null
      };
    }

    // 2. Enregistrer l’ingreso du jour dans la feuille INGRESOS
    registrarIngresoDesdePedidos_(pedidos);

    // 3. Sauvegarder l’ancienne ligne 2 de l’historique (si elle existe)
    var hadRow2 = sheetHist.getLastRow() >= 2;
    var row2Values = null;
    var row2Formulas = null;
    if (hadRow2) {
      row2Values = sheetHist.getRange(2, 1, 1, 12).getValues();
      row2Formulas = sheetHist.getRange(2, 1, 1, 12).getFormulasR1C1();
    }

    // 4. Insérer autant de lignes que de pedidos juste sous l’entête
    sheetHist.insertRowsAfter(1, pedidos.length);

    // 5. Préparer les nouvelles lignes à écrire
    var out = [];
    pedidos.forEach(function (row) {
      var fechaMenu   = row[0]; // A PEDIDOS DIARIOS
      var cliente     = row[1]; // B
      var telefono    = row[2]; // C
      var direccion   = row[3]; // D
      var cantidad    = row[4]; // E
      var total       = row[5]; // F
      var notas       = row[8]; // I
      var estado      = row[9]; // J

      // ID
      var idPedido = generarIdPedido_(fechaMenu, cliente);

      // Menu du jour depuis PROGRAMACIÓN MENÚS
      var menuDelDia = getMenuProgramadoPorFecha_(fechaMenu);

      // Mise à jour de la base clients (avec adresse et notes)
      actualizarUltimoPedidoCliente({
        nombre: cliente,
        telefono: telefono,
        fecha: fechaMenu || new Date(),
        direccion: direccion || '',
        notas: notas || ''
      });

      out.push([
        idPedido,                // A ID pedido
        new Date(),              // B Fecha del pedido (aujourd’hui)
        fechaMenu || '',         // C Fecha de entrega
        cliente || '',           // D Cliente
        telefono || '',          // E Teléfono
        direccion || '',         // F Dirección
        menuDelDia || '',        // G Menú del día
        cantidad || '',          // H Cantidad
        1,                       // I Precio unitario ($)
        total || '',             // J Total ($)
        estado || '',            // K Estado
        notas || ''              // L Observaciones
      ]);
    });

    // 6. Écrire les nouvelles lignes dans l’historique
    sheetHist.getRange(2, 1, out.length, out[0].length).setValues(out);
    // enlever le gras sur les lignes qu’on vient d’ajouter
    sheetHist.getRange(2, 1, out.length, 12).setFontWeight('normal');

    // 7. Rétablir l’ancienne ligne 2 (qui a été décalée)
    if (hadRow2) {
      var targetRow = 2 + out.length;
      sheetHist.getRange(targetRow, 1, 1, 12).setValues(row2Values);
      for (var c = 0; c < 12; c++) {
        if (row2Formulas[0][c]) {
          sheetHist.getRange(targetRow, c + 1).setFormulaR1C1(row2Formulas[0][c]);
        }
      }
    }

    // 8. Nettoyer PEDIDOS DIARIOS sans effacer les formules
    limpiarPedidosDelDia();
    // Mettre à jour le résumé mensuel
    actualizarResumenMensual();

    // 9. Retourner le resumen conforme au PAQUETE 6
    return {
      ok: true,
      error: null,
      data: {
        resumen: resumen
      }
    };

  } catch (e) {
    return {
      ok: false,
      error: e && e.message ? e.message : 'Error desconocido al cerrar el día.',
      data: null
    };
  }
}

/**
 * Previsualiza el cierre del día SIN mover datos ni limpiar hojas.
 *
 * Petición:
 *   previsualizarCierreDia(fechaString|null)
 *   - fecha (string 'YYYY-MM-DD') es opcional y solo se usa
 *     como referencia; la lógica sigue la misma que cerrarDia():
 *     toma todos los pedidos presentes en PEDIDOS DIARIOS.
 *
 * Respuesta:
 * {
 *   ok: true|false,
 *   error: string|null,
 *   data: {
 *     resumen: ResumenCierreFront
 *   } | null
 * }
 */
function previsualizarCierreDia(fecha) {
  try {
    var sheetPedidos = getSheet_('PEDIDOS DIARIOS');
    var lastRowPedidos = sheetPedidos.getLastRow();

    if (lastRowPedidos < 2) {
      return {
        ok: false,
        error: 'No hay pedidos para previsualizar el cierre del día.',
        data: null
      };
    }

    var pedidos = sheetPedidos
      .getRange(2, 1, lastRowPedidos - 1, 10)
      .getValues()
      .filter(function (r) {
        return r[0] || r[1] || r[2] || r[3] || r[4];
      });

    if (pedidos.length === 0) {
      return {
        ok: false,
        error: 'No hay pedidos válidos para previsualizar el cierre del día.',
        data: null
      };
    }

    var resumen = construirResumenCierreDesdePedidos_(pedidos);
    if (!resumen) {
      return {
        ok: false,
        error: 'No se pudo construir el resumen de previsualización.',
        data: null
      };
    }

    return {
      ok: true,
      error: null,
      data: {
        resumen: resumen
      }
    };

  } catch (e) {
    return {
      ok: false,
      error: e && e.message ? e.message : 'Error desconocido al previsualizar el cierre del día.',
      data: null
    };
  }
}

/**
 * Construye el resumen de cierre del día a partir de la matriz de pedidos
 * tal como se usa en cerrarDia() y registrarIngresoDesdePedidos_.
 *
 * @param {Array[]} pedidosDelDia - matriz A:J desde PEDIDOS DIARIOS
 * @returns {Object|null} ResumenCierreFront o null si no se puede calcular
 *
 * ResumenCierreFront:
 * {
 *   fecha: string,          // 'yyyy-MM-dd'
 *   totalPedidos: number,
 *   totalSecos: number,
 *   totalIngresos: number
 * }
 */
function construirResumenCierreDesdePedidos_(pedidosDelDia) {
  if (!pedidosDelDia || pedidosDelDia.length === 0) {
    return null;
  }

  var fechaMenu = pedidosDelDia[0][0]; // Columna A
  if (!fechaMenu) {
    return null;
  }

  var tz = Session.getScriptTimeZone();
  var d = new Date(fechaMenu);
  var fechaStr = Utilities.formatDate(d, tz, 'yyyy-MM-dd');

  var totalPedidos = pedidosDelDia.length;
  var totalSecos = 0;

  pedidosDelDia.forEach(function (row) {
    var cant = Number(row[4] || 0); // Columna E = Cantidad
    totalSecos += cant;
  });

  // Debe ser coherente con registrarIngresoDesdePedidos_ (precioUnitario = 1)
  var precioUnitario = 1;
  var totalIngresos = totalSecos * precioUnitario;

  return {
    fecha: fechaStr,
    totalPedidos: totalPedidos,
    totalSecos: totalSecos,
    totalIngresos: totalIngresos
  };
}

function actualizarResumenMensual() {
  var shIng = getSheet_('INGRESOS');
  var shGas = getSheet_('GASTOS');
  var shRes = getSheet_('RESUMEN MENSUAL');

  // 1. Lire INGRESOS
  var lastIng = shIng.getLastRow();
  var ingresosData = [];
  if (lastIng > 1) {
    ingresosData = shIng.getRange(2, 1, lastIng - 1, 6).getValues();
    // A Fecha, B Mes, C Total comidas, D Precio x1, E Ingreso total, F Obs.
  }

  // 2. Lire GASTOS
  var lastGas = shGas.getLastRow();
  var gastosData = [];
  if (lastGas > 1) {
    gastosData = shGas.getRange(2, 1, lastGas - 1, 10).getValues();
    // A Fecha, B Mes, C Categoría, D Producto, E Cant, F Unidad, G Precio, H Monto, I Proveedor, J Obs.
  }

  // 3. On regroupe par AAAA-MM
  var mapa = {}; // { '2025-11': {ingresos:..., gastos:..., secos:...} }

  // Traiter INGRESOS
  ingresosData.forEach(function (row) {
    var fecha = row[0];
    if (!fecha) return;
    var d = new Date(fecha);
    var anio = d.getFullYear();
    var mes = ('0' + (d.getMonth() + 1)).slice(-2);
    var clave = anio + '-' + mes;

    if (!mapa[clave]) {
      mapa[clave] = { ingresos: 0, gastos: 0, secos: 0, anio: anio, mes: mes };
    }

    var totalComidas = Number(row[2] || 0); // C
    var ingresoTotal = Number(row[4] || 0); // E

    mapa[clave].ingresos += ingresoTotal;
    mapa[clave].secos += totalComidas;
  });

  // Traiter GASTOS
  gastosData.forEach(function (row) {
    var fecha = row[0];
    if (!fecha) return;
    var d = new Date(fecha);
    var anio = d.getFullYear();
    var mes = ('0' + (d.getMonth() + 1)).slice(-2);
    var clave = anio + '-' + mes;

    if (!mapa[clave]) {
      mapa[clave] = { ingresos: 0, gastos: 0, secos: 0, anio: anio, mes: mes };
    }

    var monto = Number(row[7] || 0); // H Monto ($)
    mapa[clave].gastos += monto;
  });

  // 4. Transformer en tableau et trier (année desc, mois desc)
  var claves = Object.keys(mapa);
  if (claves.length === 0) {
    // on efface juste la feuille (sauf en-têtes)
    var lastRes = shRes.getLastRow();
    if (lastRes > 1) {
      shRes.getRange(2, 1, lastRes - 1, 8).clearContent();
    }
    return;
  }

  claves.sort(function (a, b) {
    // tri décroissant
    return a < b ? 1 : -1;
  });

  var salida = [];
  claves.forEach(function (k) {
    var obj = mapa[k];
    var ingresos = obj.ingresos;
    var gastos = obj.gastos;
    var secos = obj.secos;
    var beneficio = ingresos - gastos;

    var costoX1 = '';
    var precioX1 = '';
    if (secos > 0) {
      costoX1 = gastos / secos;
      precioX1 = ingresos / secos;
    }

    salida.push([
      obj.anio,          // Año
      obj.mes,           // Mes
      ingresos,          // Ingresos
      gastos,            // Gastos
      beneficio,         // Beneficio
      secos,             // Secos
      costoX1,           // Costo x1
      precioX1           // Precio x1
    ]);
  });

  // 5. Effacer l’ancien contenu
  var lastRes = shRes.getLastRow();
  if (lastRes > 1) {
    shRes.getRange(2, 1, lastRes - 1, 8).clearContent();
  }

  // 6. Écrire le nouveau
  shRes.getRange(2, 1, salida.length, 8).setValues(salida);
}





// ANALYSER LES DONNÉES





/**
 * Busca la primera fila realmente vacía en una columna,
 * sin tenir compte des validations ou du format.
 * Empieza à partir de la fila 2 (fila 1 = encabezados).
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} colIndex - 1 = col A, 2 = col B, etc.
 * @returns {number} número de fila libre
 */
function getFirstEmptyRowInColumn_(sheet, colIndex) {
  // on lit depuis la ligne 2 jusqu’au nombre MAX de lignes
  var max = sheet.getMaxRows();
  var values = sheet.getRange(2, colIndex, max - 1, 1).getValues();
  for (var i = 0; i < values.length; i++) {
    if (!values[i][0]) {
      // ligne vide trouvée
      return i + 2; // +2 car on a commencé à la ligne 2
    }
  }
  // si vraiment rien de vide, on ajoute à la fin
  return sheet.getLastRow() + 1;
}

function sortPedidosDiarios_() {
  var sheet = getSheet_('PEDIDOS DIARIOS');
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow <= 1) return; // aucune donnée

  var range = sheet.getRange(2, 1, lastRow - 1, lastCol);
  range.sort([
    { column: 1, ascending: true }, // Fecha
    { column: 8, ascending: true }  // Hora
  ]);
}





// MENU





function getMenuProgramadoPorFecha_(fecha) {
  if (!fecha) return '';
  var sheet = getSheet_('PROGRAMACIÓN MENÚS');
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return '';

  // normaliser la date reçue
  var f = new Date(fecha);
  var fStr = Utilities.formatDate(f, Session.getScriptTimeZone(), 'yyyy-MM-dd');

  var data = sheet.getRange(2, 1, lastRow - 1, 2).getValues(); // A: fecha, B: menú
  for (var i = 0; i < data.length; i++) {
    var fechaCell = data[i][0];
    var menuCell  = data[i][1];

    if (!fechaCell) continue;

    var fechaCellStr = Utilities.formatDate(new Date(fechaCell), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    if (fechaCellStr === fStr) {
      return menuCell || '';
    }
  }
  return '';
}

/**
 * Guarda un menú en la hoja MENÚS y asegura
 * que LISTAS!G esté también actualizado.
 */
function guardarMenu(data) {
  // data attendu:
  // {
  //   id: 'M-003'        (optionnel, sinon on génère)
  //   nombre: 'Seco de pollo',
  //   descripcion: '...',
  //   ingredientes: 'pollo; arroz; ensalada',
  //   foto: 'https://...',
  //   activo: 'Sí'       // ou 'No'
  // }

  var sheet = getSheet_('MENÚS');
  var row = getFirstEmptyRowInColumn_(sheet, 1);

  // Générer un ID si aucun n’est fourni
  var idMenu = data && data.id ? data.id : generarIdMenu_();

  sheet.getRange(row, 1).setValue(idMenu);
  sheet.getRange(row, 2).setValue(data && data.nombre ? data.nombre : '');
  sheet.getRange(row, 3).setValue(data && data.descripcion ? data.descripcion : '');
  sheet.getRange(row, 4).setValue(data && data.ingredientes ? data.ingredientes : '');
  sheet.getRange(row, 5).setValue(data && data.foto ? data.foto : '');
  sheet.getRange(row, 6).setValue(data && data.activo ? data.activo : 'Sí');

  // très important : mettre à jour LISTAS!G
  sincronizarMenusEnListas();
}

/**
 * Modifica un menú existente en la hoja MENÚS.
 * Se identifica el menú por su ID (columna A) o, si no hay ID,
 * por el nombre (columna B).
 *
 * data esperado:
 * {
 *   id: 'M-003',            // preferible
 *   nombre: 'Seco de pollo',
 *   descripcion: '...',
 *   ingredientes: 'pollo; arroz; ensalada',
 *   foto: 'https://...',
 *   activo: 'Sí'            // o 'No'
 * }
 */
function actualizarMenu(data) {
  var sheet = getSheet_('MENÚS');
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    throw new Error('No hay menús registrados.');
  }

  var values = sheet.getRange(2, 1, lastRow - 1, 6).getValues(); // A:F
  var filaObjetivo = -1;

  // 1. Buscar por ID si viene data.id
  if (data.id) {
    for (var i = 0; i < values.length; i++) {
      var idFila = (values[i][0] || '').toString().trim();
      if (idFila === data.id) {
        filaObjetivo = i + 2; // +2 porque empezamos en la fila 2
        break;
      }
    }
  }

  // 2. Si no hay id o no lo encontró, intentar por nombre
  if (filaObjetivo === -1 && data.nombre) {
    for (var j = 0; j < values.length; j++) {
      var nombreFila = (values[j][1] || '').toString().trim().toLowerCase();
      if (nombreFila === data.nombre.toLowerCase()) {
        filaObjetivo = j + 2;
        break;
      }
    }
  }

  if (filaObjetivo === -1) {
    throw new Error('No se encontró el menú a modificar.');
  }

  // Actualizar solo los campos recibidos
  if (data.nombre !== undefined) {
    sheet.getRange(filaObjetivo, 2).setValue(data.nombre);
  }
  if (data.descripcion !== undefined) {
    sheet.getRange(filaObjetivo, 3).setValue(data.descripcion);
  }
  if (data.ingredientes !== undefined) {
    sheet.getRange(filaObjetivo, 4).setValue(data.ingredientes);
  }
  if (data.foto !== undefined) {
    sheet.getRange(filaObjetivo, 5).setValue(data.foto);
  }
  if (data.activo !== undefined) {
    sheet.getRange(filaObjetivo, 6).setValue(data.activo);
  }

  // Muy importante: refrescar la lista de menús
  sincronizarMenusEnListas();
}

/**
 * Sincroniza la lista de menús en LISTAS!G
 */
function sincronizarMenusEnListas() {
  var sheetMenus = getSheet_('MENÚS');
  var sheetListas = getSheet_('LISTAS');

  var lastRowMenus = sheetMenus.getLastRow();
  if (lastRowMenus < 2) return;

  var menuNames = sheetMenus.getRange('B2:B' + lastRowMenus).getValues().flat().filter(String);

  sheetListas.getRange('G2:G200').clearContent();

  var out = menuNames.map(function (m) { return [m]; });
  out.push(['Otro']);

  sheetListas.getRange(2, 7, out.length, 1).setValues(out);
}

/**
 * Devuelve la configuración del "menú de mañana".
 *
 * Contrato (PAQUETE 5):
 * data:
 * {
 *   config: {
 *     fechaManana: string,        // 'YYYY-MM-DD'
 *     menuActual: string,         // nombre del menú o ''
 *     menusDisponibles: string[]  // desde LISTAS!G
 *   }
 * }
 */
function getMenuManana() {
  try {
    var tz = Session.getScriptTimeZone();
    var hoy = new Date();
    // "Mañana" = hoy + 1 día
    var manana = new Date(hoy);
    manana.setDate(manana.getDate() + 1);
    var fechaMananaStr = Utilities.formatDate(manana, tz, 'yyyy-MM-dd');

    // 1) Leer la programación de menús para encontrar el menú de esa fecha
    var sheetProg = getSheet_('PROGRAMACIÓN MENÚS');
    var lastRow = sheetProg.getLastRow();
    var menuActual = '';

    if (lastRow >= 2) {
      var data = sheetProg.getRange(2, 1, lastRow - 1, 2).getValues(); // A: fecha, B: menú
      for (var i = 0; i < data.length; i++) {
        var fechaCell = data[i][0];
        var menuCell  = data[i][1];

        if (!fechaCell) continue;

        var fechaCellStr = Utilities.formatDate(new Date(fechaCell), tz, 'yyyy-MM-dd');
        if (fechaCellStr === fechaMananaStr) {
          menuActual = (menuCell || '').toString();
          break;
        }
      }
    }

    // 2) Menús disponibles = LISTAS!G (ya sincronizado por sincronizarMenusEnListas)
    var listas = getLists();
    var menusDisponibles = listas.menus || [];

    return {
      ok: true,
      error: null,
      data: {
        config: {
          fechaManana: fechaMananaStr,
          menuActual: menuActual,
          menusDisponibles: menusDisponibles
        }
      }
    };

  } catch (e) {
    return {
      ok: false,
      error: e && e.message ? e.message : 'Error desconocido al obtener el menú de mañana.',
      data: null
    };
  }
}

/**
 * Define / cambia el menú programado para "mañana"
 * (o para la fecha indicada) en PROGRAMACIÓN MENÚS.
 *
 * Contrato (PAQUETE 5):
 * Petición:
 * {
 *   fechaManana: string,        // 'YYYY-MM-DD'
 *   menuSeleccionado: string    // nombre exacto del menú
 * }
 *
 * Respuesta:
 * {
 *   config: MenuConfigFront
 * }
 */
function setMenuManana(data) {
  try {
    if (!data || !data.fechaManana || !data.menuSeleccionado) {
      return {
        ok: false,
        error: 'Faltan fechaManana o menuSeleccionado.',
        data: null
      };
    }

    var tz = Session.getScriptTimeZone();
    var fechaObj = new Date(data.fechaManana);
    if (isNaN(fechaObj.getTime())) {
      return {
        ok: false,
        error: 'fechaManana inválida.',
        data: null
      };
    }

    var fechaStr = Utilities.formatDate(fechaObj, tz, 'yyyy-MM-dd');
    var menuSeleccionado = data.menuSeleccionado.toString().trim();
    if (!menuSeleccionado) {
      return {
        ok: false,
        error: 'El nombre del menú seleccionado no puede estar vacío.',
        data: null
      };
    }

    var sheetProg = getSheet_('PROGRAMACIÓN MENÚS');
    var lastRow = sheetProg.getLastRow();
    var filaEncontrada = -1;

    if (lastRow >= 2) {
      var dataProg = sheetProg.getRange(2, 1, lastRow - 1, 2).getValues(); // A: fecha, B: menú
      for (var i = 0; i < dataProg.length; i++) {
        var fechaCell = dataProg[i][0];
        if (!fechaCell) continue;

        var fechaCellStr = Utilities.formatDate(new Date(fechaCell), tz, 'yyyy-MM-dd');
        if (fechaCellStr === fechaStr) {
          filaEncontrada = i + 2; // fila real
          break;
        }
      }
    }

    if (filaEncontrada === -1) {
      // No había programación para esa fecha → nueva línea
      var newRow = getFirstEmptyRowInColumn_(sheetProg, 1); // columna A
      sheetProg.getRange(newRow, 1).setValue(fechaObj);        // A Fecha
      sheetProg.getRange(newRow, 2).setValue(menuSeleccionado); // B Menú
    } else {
      // Ya existía → solo actualizar el nombre del menú
      sheetProg.getRange(filaEncontrada, 2).setValue(menuSeleccionado);
    }

    // Menús disponibles = LISTAS!G (sin inventar nada)
    var listas = getLists();
    var menusDisponibles = listas.menus || [];

    // Construir MenuConfigFront de respuesta
    var config = {
      fechaManana: fechaStr,
      menuActual: menuSeleccionado,
      menusDisponibles: menusDisponibles
    };

    return {
      ok: true,
      error: null,
      data: {
        config: config
      }
    };

  } catch (e) {
    return {
      ok: false,
      error: e && e.message ? e.message : 'Error desconocido al configurar el menú de mañana.',
      data: null
    };
  }
}




// COMPRAS





/**
 * Genera la lista de compras del día en la hoja LISTA DE COMPRAS
 */
function generarListaCompras(fecha) {
  var sheetMenus = getSheet_('MENÚS');
  var sheetCompras = getSheet_('LISTA DE COMPRAS');

  var lastRowMenus = sheetMenus.getLastRow();
  var menus = sheetMenus.getRange(2, 1, lastRowMenus - 1, 6).getValues();
  var menuElegido = null;
  for (var i = 0; i < menus.length; i++) {
    if (menus[i][5] === 'Sí') {
      menuElegido = menus[i];
      break;
    }
  }
  if (!menuElegido) {
    throw new Error('No hay menú activo para generar la lista de compras.');
  }

  var nombreMenu = menuElegido[1];
  var ingredientesBase = (menuElegido[3] || '').toString();
  if (!ingredientesBase) {
    throw new Error('El menú "' + nombreMenu + '" no tiene ingredientes base definidos.');
  }

  var ingredientes = ingredientesBase.split(';').map(function (s) { return s.trim(); });

  var targetRow = getFirstEmptyRowInColumn_(sheetCompras, 1);
  var out = [];

  ingredientes.forEach(function (ing) {
    out.push([
      fecha || new Date(),
      nombreMenu,
      ing,
      '',
      '',
      '',
      ''
    ]);
  });

  sheetCompras.getRange(targetRow, 1, out.length, out[0].length).setValues(out);
}

/**
 * Genera el mensaje de WhatsApp con la lista de compras del día.
 */
function generarTextoComprasWhatsApp(fecha) {
  var sheetCompras = getSheet_('LISTA DE COMPRAS');
  var lastRow = sheetCompras.getLastRow();
  if (lastRow < 2) {
    return 'No hay compras registradas para hoy.';
  }

  var data = sheetCompras.getRange(2, 1, lastRow - 1, 7).getValues();
  var texto = 'Lista de compras SecoMóvil\n';
  texto += 'Fecha: ' + (fecha || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd')) + '\n\n';

  data.forEach(function (row) {
    var ing = row[2];
    var cant = row[3];
    var uni = row[4];

    if (!ing) return;

    texto += '- ' + ing;
    if (cant) {
      texto += ': ' + cant;
      if (uni) {
        texto += ' ' + uni;
      }
    }
    texto += '\n';
  });

  return texto;
}





// BLOC ID





/**
 * Genera un ID simple para los menús.
 */
function generarIdMenu_() {
  var sheet = getSheet_('MENÚS');
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return 'M-001';
  }
  var lastId = sheet.getRange(lastRow, 1).getValue();
  var num = 1;
  if (lastId && typeof lastId === 'string' && lastId.indexOf('M-') === 0) {
    num = parseInt(lastId.replace('M-', ''), 10) + 1;
  }
  return Utilities.formatString('M-%03d', num);
}

/**
 * Genera un ID simple para los gastos.
 * Formato: G-00001, G-00002, ...
 */
function generarIdGasto_() {
  var sheet = getSheet_('GASTOS');
  var lastRow = sheet.getLastRow();

  // Si no hay filas de datos
  if (lastRow < 2) {
    return 'G-00001';
  }

  // Leemos la columna K (ID_GASTO) desde la fila 2
  var idValues = sheet.getRange(2, 11, lastRow - 1, 1).getValues()
    .map(function (r) { return r[0]; })
    .filter(function (v) { return v; });

  if (idValues.length === 0) {
    return 'G-00001';
  }

  var ultimoId = idValues[idValues.length - 1].toString();
  var num = 1;

  if (ultimoId.indexOf('G-') === 0) {
    num = parseInt(ultimoId.replace('G-', ''), 10);
    if (isNaN(num)) {
      num = 0;
    }
  }

  num++;
  return Utilities.formatString('G-%05d', num);
}

/**
 * Genera un ID simple y secuencial para los pedidos.
 * Formato: P-00001, P-00002, ...
 *
 * Busca los IDs existentes en:
 * - PEDIDOS DIARIOS (columna K = ID_Pedido)
 * - HISTORIAL DE PEDIDOS (columna A = ID_Pedido)
 * y toma el máximo numérico que siga el patrón P-xxxxx.
 *
 * Los antiguos IDs con formato diferente (por ejemplo P-20251115-123000)
 * se ignoran para la secuencia.
 */
function generarIdPedido_(fecha, cliente) {
  var maxNum = 0;

  // 1) Buscar en PEDIDOS DIARIOS (columna K)
  var sheetPedidos = getSheet_('PEDIDOS DIARIOS');
  var lastRowPedidos = sheetPedidos.getLastRow();
  if (lastRowPedidos > 1) {
    var idsPedidos = sheetPedidos.getRange(2, 11, lastRowPedidos - 1, 1).getValues(); // K2:K
    idsPedidos.forEach(function (row) {
      var id = (row[0] || '').toString().trim();
      if (!id) return;
      var match = id.match(/^P-(\d+)$/);
      if (match) {
        var n = parseInt(match[1], 10);
        if (!isNaN(n) && n > maxNum) {
          maxNum = n;
        }
      }
    });
  }

  // 2) Buscar también en HISTORIAL DE PEDIDOS (columna A)
  var sheetHist = null;
  try {
    sheetHist = getSheet_('HISTORIAL DE PEDIDOS');
  } catch (e) {
    sheetHist = null; // si la hoja no existe aún, no pasa nada
  }

  if (sheetHist) {
    var lastRowHist = sheetHist.getLastRow();
    if (lastRowHist > 1) {
      var idsHist = sheetHist.getRange(2, 1, lastRowHist - 1, 1).getValues(); // A2:A
      idsHist.forEach(function (row) {
        var id = (row[0] || '').toString().trim();
        if (!id) return;
        var match = id.match(/^P-(\d+)$/);
        if (match) {
          var n = parseInt(match[1], 10);
          if (!isNaN(n) && n > maxNum) {
            maxNum = n;
          }
        }
      });
    }
  }

  // 3) Siguiente número
  var nextNum = maxNum + 1;
  return Utilities.formatString('P-%05d', nextNum); // P-00001, P-00002, ...
}

/**
 * Expone la generación de IDs de pedido para las páginas HTML.
 */
function generarIdPedido() {
  return generarIdPedido_();
}

/**
 * Genera el próximo ID secuencial para VENTAS DIRECTAS.
 * Formato: V-00001, V-00002, ...
 */
function generarIdVentaDirecta_() {
  var sheet = getSheet_('VENTAS DIRECTAS');
  var lastRow = sheet.getLastRow();
  var maxNum = 0;

  if (lastRow > 1) {
    var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    ids.forEach(function (row) {
      var id = (row[0] || '').toString().trim();
      if (!id) return;
      var match = id.match(/^V-(\d+)$/);
      if (!match) return;
      var n = parseInt(match[1], 10);
      if (!isNaN(n) && n > maxNum) {
        maxNum = n;
      }
    });
  }

  var nextNum = maxNum + 1;
  return Utilities.formatString('V-%05d', nextNum);
}

/**
 * Genera un ID simple para los productos detallados.
 * Formato: "prod-00001", "prod-00002", ...
 */
function generarIdProducto_() {
  var sheet = getSheet_('PRODUCTO');
  var lastRow = sheet.getLastRow();

  // Si no hay datos (solo encabezado o hoja vacía)
  if (lastRow < 2) {
    return 'prod-00001';
  }

  var lastId = sheet.getRange(lastRow, 1).getValue(); // Columna A = PRODUCTO ID
  var num = 0;

  if (lastId && typeof lastId === 'string' && lastId.indexOf('prod-') === 0) {
    num = parseInt(lastId.replace('prod-', ''), 10);
    if (isNaN(num)) {
      num = 0;
    }
  }

  num++;
  // 5 dígitos: 00001, 00002, ...
  return Utilities.formatString('prod-%05d', num);
}

/**
 * Genera un ID simple y secuencial para los clientes.
 * Formato: C-00001, C-00002, ...
 *
 * Lee todos los IDs existentes en BASE DE CLIENTES (columna H)
 * y toma el máximo numérico que siga el patrón C-xxxxx.
 *
 * Soporta tanto IDs antiguos (C-0001) como nuevos: el número interno
 * se interpreta como entero y el resultado se devuelve siempre con 5 dígitos.
 */
function generarIdCliente_() {
  var sheet = getSheet_('BASE DE CLIENTES');
  var lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return 'C-00001';
  }

  var idValues = sheet.getRange(2, 8, lastRow - 1, 1).getValues(); // H2:H
  var maxNum = 0;

  idValues.forEach(function (row) {
    var id = (row[0] || '').toString().trim();
    if (!id) return;

    var match = id.match(/^C-(\d+)$/);
    if (match) {
      var n = parseInt(match[1], 10);
      if (!isNaN(n) && n > maxNum) {
        maxNum = n;
      }
    }
  });

  var nextNum = maxNum + 1;
  return Utilities.formatString('C-%05d', nextNum); // C-00001, C-00002, ...
}

/**
 * Lee los sectores disponibles desde la hoja SECTORES.
 * Devuelve un arreglo de objetos con nombre, columna y precio base.
 */
function leerSectores_() {
  var sheet = getSheet_('SECTORES');
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();

  if (lastRow < 3 || lastColumn < 1) {
    return [];
  }

  var preciosRow = sheet.getRange(2, 1, 1, lastColumn).getValues()[0];
  var etiquetas = sheet.getRange(3, 1, lastRow - 2, lastColumn).getValues();

  var sectores = [];

  for (var col = 0; col < lastColumn; col++) {
    var precioBruto = preciosRow[col];
    var precio = null;
    if (precioBruto !== '' && precioBruto !== null && precioBruto !== undefined) {
      var parsed = Number(precioBruto);
      precio = isNaN(parsed) ? null : parsed;
    }
    var colLetter = columnNumberToLetter_(col + 1);

    for (var row = 0; row < etiquetas.length; row++) {
      var nombre = (etiquetas[row][col] || '').toString().trim();
      if (!nombre) continue;

      sectores.push({
        nombre: nombre,
        columna: colLetter,
        precio: precio
      });
    }
  }

  return sectores;
}

/**
 * Obtiene el precio unitario de un sector leyendo la hoja SECTORES.
 *
 * - La fila 2 contiene los precios unitarios.
 * - La fila 3 en adelante contiene los sectores.
 *
 * @param {string} sectorNombre
 * @returns {number|null}
 */
function obtenerPrecioUnitarioSector_(sectorNombre) {
  var sector = (sectorNombre || '').toString().trim();
  if (!sector) return null;

  var sheet = getSheet_('SECTORES');
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();

  if (lastRow < 3 || lastColumn < 1) return null;

  var sectores = sheet.getRange(3, 1, lastRow - 2, lastColumn).getValues();
  var objetivo = sector.toLowerCase();
  var colEncontrada = -1;

  for (var col = 0; col < lastColumn && colEncontrada === -1; col++) {
    for (var row = 0; row < sectores.length; row++) {
      var nombre = (sectores[row][col] || '').toString().trim();
      if (nombre && nombre.toLowerCase() === objetivo) {
        colEncontrada = col + 1; // 1-based
        break;
      }
    }
  }

  if (colEncontrada === -1) return null;

  var precioBruto = sheet.getRange(2, colEncontrada).getValue();
  var precio = Number(precioBruto);
  return isNaN(precio) ? null : precio;
}

/**
 * Busca el precio y la columna de un sector dado (insensible a mayúsculas/minúsculas).
 * @param {string} sectorNombre
 * @returns {{precio:number|null, columna:string|null}}
 */
function obtenerPrecioYColumnaSector_(sectorNombre) {
  var sector = (sectorNombre || '').toString().trim().toLowerCase();
  if (!sector) {
    return { precio: null, columna: null };
  }

  var lista = leerSectores_();
  for (var i = 0; i < lista.length; i++) {
    if ((lista[i].nombre || '').toString().trim().toLowerCase() === sector) {
      return {
        precio: typeof lista[i].precio === 'number' ? lista[i].precio : null,
        columna: lista[i].columna || null
      };
    }
  }

  return { precio: null, columna: null };
}

/**
 * Devuelve el precio unitario para un sector específico.
 * @param {string} sectorNombre
 * @returns {{ok:boolean, error:string|null, data:{precio:number|null}}}
 */
function obtenerPrecioUnitarioSector(sectorNombre) {
  try {
    var precio = obtenerPrecioUnitarioSector_(sectorNombre);
    return {
      ok: true,
      error: null,
      data: { precio: precio }
    };
  } catch (e) {
    return {
      ok: false,
      error: e && e.message ? e.message : 'Error al obtener el precio unitario del sector.',
      data: { precio: null }
    };
  }
}

/**
 * Convierte un número de columna (1 = A) a su letra correspondiente.
 * @param {number} colNum
 * @returns {string}
 */
function columnNumberToLetter_(colNum) {
  var num = Math.max(1, Math.floor(colNum));
  var letter = '';
  while (num > 0) {
    var rem = (num - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    num = Math.floor((num - 1) / 26);
  }
  return letter;
}

/**
 * Intenta convertir un texto "Lunes, 02/09/2025" o "02/09/2025" en un objeto Date.
 * @param {string} textoFecha
 * @returns {Date|null}
 */
function normalizarFechaDesdeTexto_(textoFecha) {
  if (!textoFecha) return null;

  var m = textoFecha.toString().match(/(\d{2})\/(\d{2})\/(\d{4})/);
  if (!m) return null;

  var dia = parseInt(m[1], 10);
  var mes = parseInt(m[2], 10) - 1; // 0-indexed
  var anio = parseInt(m[3], 10);

  var d = new Date(anio, mes, dia);
  if (isNaN(d.getTime())) return null;
  return d;
}







// BLOC CLIENTES





/**
 * Vide PEDIDOS DIARIOS (garde les en-têtes)
 */
function limpiarPedidosDelDia() {
  var sheet = getSheet_('PEDIDOS DIARIOS');
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  var numRows = lastRow - 1; // nombre de lignes à nettoyer

  // On efface A:E
  sheet.getRange(2, 1, numRows, 5).clearContent();
  // On laisse F tranquille (formule)
  // On efface G:J
  sheet.getRange(2, 7, numRows, 4).clearContent();
}

/**
 * Crea un nuevo cliente en BASE DE CLIENTES.
 *
 * data (desde HTML):
 * {
 *   nombre:   string,
 *   telefono: string,
 *   direccion:string,
 *   sector:   string
 * }
 *
 * Respuesta (contrato JSON):
 * {
 *   ok: true|false,
 *   error: string|null,
 *   data: {
 *     cliente: {
 *       idCliente: string,
 *       nombre: string,
 *       telefono: string,
 *       direccion: string,
 *       sector: string
 *     }
 *   } | null
 * }
 */
function crearCliente(data) {
  try {
    // Validación básica de entrada
    if (!data ||
      !data.nombre ||
      !data.telefono ||
      !data.direccion ||
      !data.sector) {
      return {
        ok: false,
        error: 'Faltan datos obligatorios para crear el cliente.',
        data: null
      };
    }

    var sheet = getSheet_('BASE DE CLIENTES');

    // Buscar la primera fila realmente vacía (columna A = Nombre)
    var newRow = getFirstEmptyRowInColumn_(sheet, 1);

    // Generar nuevo ID cliente secuencial: C-00001, C-00002, ...
    var nuevoId = generarIdCliente_();

    var hoy = new Date();

    // Estructura esperada de BASE DE CLIENTES:
    // A Nombre
    // B Teléfono
    // C Dirección
    // D Fecha último pedido
    // E Frecuencia
    // F Estado
    // G Notas
    // H ID Cliente
    // I Sector
    sheet.getRange(newRow, 1).setValue(data.nombre);
    sheet.getRange(newRow, 2).setValue(data.telefono);
    sheet.getRange(newRow, 3).setValue(data.direccion);
    sheet.getRange(newRow, 4).setValue(hoy);                // Fecha último pedido = creación
    sheet.getRange(newRow, 5).setValue('');                 // Frecuencia (vacío en Fase A)
    sheet.getRange(newRow, 6).setValue('Nuevo');            // Estado
    sheet.getRange(newRow, 7).setValue(data.nota || '');    // Notas (usa "nota" del formulario)
    sheet.getRange(newRow, 8).setValue(nuevoId);            // ID Cliente
    sheet.getRange(newRow, 9).setValue(data.sector);        // Sector

    // Opcional: mantener el orden alfabético por nombre
    var totalRows = sheet.getLastRow();
    if (totalRows > 2) {
      sheet.getRange(2, 1, totalRows - 1, 9).sort({ column: 1, ascending: true });
    }

    // Construir el objeto Cliente para devolver al front
    var cliente = {
      idCliente: nuevoId,
      nombre: data.nombre,
      telefono: data.telefono,
      direccion: data.direccion,
      sector: data.sector,
      nota: data.nota || ''
    };

    return {
      ok: true,
      error: null,
      data: {
        cliente: cliente
      }
    };

  } catch (e) {
    return {
      ok: false,
      error: e && e.message ? e.message : 'Error desconocido al crear el cliente.',
      data: null
    };
  }
}

/**
 * Lista todos los clientes de BASE DE CLIENTES.
 *
 * Contrato JSON (PAQUETE 5):
 *
 * Respuesta:
 * {
 *   ok: true|false,
 *   error: string|null,
 *   data: {
 *     clientes: [
 *       {
 *         idCliente: string,
 *         nombre: string,
 *         telefono: string,
 *         direccion: string,
 *         sector: string
 *       },
 *       ...
 *     ]
 *   } | null
 * }
 */
function listarClientes() {
  try {
    var sheet = getSheet_('BASE DE CLIENTES');
    var lastRow = sheet.getLastRow();

    // Si no hay datos (solo encabezado o ni eso)
    if (lastRow < 2) {
      return {
        ok: true,
        error: null,
        data: { clientes: [] }
      };
    }

    // Estructura esperada:
    // A Nombre
    // B Teléfono
    // C Dirección
    // D Fecha último pedido
    // E Frecuencia
    // F Estado
    // G Notas
    // H ID Cliente
    // I Sector
    var values = sheet.getRange(2, 1, lastRow - 1, 9).getValues();

    var clientes = [];
    values.forEach(function (row) {
      var nombre    = (row[0] || '').toString().trim();
      var telefono  = (row[1] || '').toString().trim();
      var direccion = (row[2] || '').toString().trim();
      var idCliente = (row[7] || '').toString().trim(); // col H
      var sector    = (row[8] || '').toString().trim(); // col I

      // Si no hay nombre ni teléfono, consideramos la fila vacía
      if (!nombre && !telefono) {
        return;
      }

      // Si no hay ID todavía (antiguos registros), devolvemos cadena vacía
      // (no inventamos datos que no existen en la hoja).
      var cliente = {
        idCliente: idCliente || '',
        nombre: nombre,
        telefono: telefono,
        direccion: direccion,
        sector: sector || ''
      };

      clientes.push(cliente);
    });

    // Ordenar por nombre (por seguridad, aunque la hoja ya suele estar ordenada)
    clientes.sort(function (a, b) {
      var na = a.nombre.toLowerCase();
      var nb = b.nombre.toLowerCase();
      if (na < nb) return -1;
      if (na > nb) return 1;
      return 0;
    });

    return {
      ok: true,
      error: null,
      data: {
        clientes: clientes
      }
    };

  } catch (e) {
    return {
      ok: false,
      error: e && e.message ? e.message : 'Error desconocido al listar clientes.',
      data: null
    };
  }
}

/**
 * Actualiza un cliente existente en BASE DE CLIENTES.
 *
 * Contrato JSON (PAQUETE 5):
 *
 * Petición:
 * {
 *   idCliente: string,
 *   nombre: string,
 *   telefono: string,
 *   direccion: string,
 *   sector: string
 * }
 *
 * Respuesta:
 * {
 *   ok: true|false,
 *   error: string|null,
 *   data: {
 *     cliente: {
 *       idCliente, nombre, telefono, direccion, sector
 *     }
 *   } | null
 * }
 */
function actualizarCliente(idCliente, data) {
  try {
    if (!idCliente) {
      return {
        ok: false,
        error: 'ID de cliente no proporcionado.',
        data: null
      };
    }

    var sheet = getSheet_('BASE DE CLIENTES');
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return {
        ok: false,
        error: 'La base de clientes está vacía.',
        data: null
      };
    }

    // Buscar fila por ID cliente (col H)
    var values = sheet.getRange(2, 8, lastRow - 1, 1).getValues(); // Col H
    var rowFound = -1;

    for (var i = 0; i < values.length; i++) {
      var id = (values[i][0] || '').toString().trim();
      if (id === idCliente) {
        rowFound = i + 2; // fila real
        break;
      }
    }

    if (rowFound === -1) {
      return {
        ok: false,
        error: 'No se encontró el cliente con ID: ' + idCliente,
        data: null
      };
    }

    // Columnas:
    // A Nombre
    // B Teléfono
    // C Dirección
    // D Fecha último pedido
    // E Frecuencia
    // F Estado
    // G Notas
    // H ID Cliente
    // I Sector

    if (data.nombre !== undefined) {
      sheet.getRange(rowFound, 1).setValue(data.nombre);
    }
    if (data.telefono !== undefined) {
      sheet.getRange(rowFound, 2).setValue(data.telefono);
    }
    if (data.direccion !== undefined) {
      sheet.getRange(rowFound, 3).setValue(data.direccion);
    }
    if (data.sector !== undefined) {
      sheet.getRange(rowFound, 9).setValue(data.sector); // col I
    }
    if (data.nota !== undefined) {
      sheet.getRange(rowFound, 7).setValue(data.nota); // col G
    }

    // Garantir que el cliente pasa a "Activo"
    sheet.getRange(rowFound, 6).setValue('Activo'); // col F

    // Leer datos actualizados para devolverlos
    var row = sheet.getRange(rowFound, 1, 1, 9).getValues()[0];
    var sectorInfo = obtenerPrecioYColumnaSector_(row[8] || '');
    var cliente = {
      idCliente: idCliente,
      nombre: row[0] || '',
      telefono: row[1] || '',
      direccion: row[2] || '',
      sector: row[8] || '',
      nota: row[6] || '',
      columnaSector: sectorInfo && sectorInfo.columna ? sectorInfo.columna : '',
      precioBase: sectorInfo && sectorInfo.precio !== undefined ? sectorInfo.precio : null
    };

    return {
      ok: true,
      error: null,
      data: { cliente: cliente }
    };

  } catch (e) {
    return {
      ok: false,
      error: e && e.message ? e.message : 'Error desconocido al actualizar el cliente.',
      data: null
    };
  }
}

/**
 * Busca y devuelve un cliente por su ID en BASE DE CLIENTES.
 * @param {string} idCliente
 * @returns {Object|null}
 */
function obtenerClientePorId_(idCliente) {
  var idBuscado = (idCliente || '').toString().trim();
  if (!idBuscado) return null;

  var sheet = getSheet_('BASE DE CLIENTES');
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  var ids = sheet.getRange(2, 8, lastRow - 1, 1).getValues(); // Col H
  for (var i = 0; i < ids.length; i++) {
    var id = (ids[i][0] || '').toString().trim();
    if (id === idBuscado) {
      var row = sheet.getRange(i + 2, 1, 1, 9).getValues()[0];
      var sectorInfo = obtenerPrecioYColumnaSector_(row[8] || '');

      return {
        idCliente: idBuscado,
        nombre: row[0] || '',
        telefono: row[1] || '',
        direccion: row[2] || '',
        nota: row[6] || '',
        sector: row[8] || '',
        columnaSector: sectorInfo && sectorInfo.columna ? sectorInfo.columna : '',
        precioBase: sectorInfo && sectorInfo.precio !== undefined ? sectorInfo.precio : null
      };
    }
  }

  return null;
}

/**
 * Devuelve un cliente (BASE DE CLIENTES) y los sectores disponibles.
 * Reutiliza la lógica existente de lectura para no duplicar hojas.
 *
 * @param {Object} payload
 * @returns {{ok: boolean, error: string|null, data: Object|null}}
 */
function obtenerClientePorId(payload) {
  try {
    payload = payload || {};
    var idContacto = (payload.idContacto || payload.idCliente || payload.id || '').toString().trim();

    if (!idContacto) {
      return { ok: false, error: 'Falta el idContacto para buscar el cliente.', data: null };
    }

    var cliente = obtenerClientePorId_(idContacto);
    if (!cliente) {
      return { ok: false, error: 'No se encontró el contacto solicitado.', data: null };
    }

    var sectores = leerSectores_();

    return {
      ok: true,
      error: null,
      data: {
        cliente: cliente,
        sectores: sectores
      }
    };
  } catch (e) {
    return {
      ok: false,
      error: e && e.message ? e.message : 'Error desconocido al obtener el contacto.',
      data: null
    };
  }
}





// BLOC PEDIDOS





/**
 * Inicializa la página nuevoPedido.html
 *
 * Respuesta:
 * {
 *   ok: true|false,
 *   error: string|null,
 *   data: {
 *     idPedido: string,            // P-00001, P-00002, ...
 *     clientes: [                  // lista real desde BASE DE CLIENTES
 *       {
 *         idCliente: string,
 *         nombre: string,
 *         telefono: string,
 *         direccion: string,
 *         sector: string
 *       },
 *       ...
 *     ],
 *     nextIdCliente: string        // próximo ID cliente sugerido: C-00001, C-00002, ...
 *   } | null
 * }
 */
function initNuevoPedido() {
  try {
    // 1) Generar ID de pedido secuencial
    var idPedido = generarIdPedido_();

    // 2) Listar clientes existentes
    var respClientes = listarClientes();
    if (!respClientes || !respClientes.ok) {
      return {
        ok: false,
        error: (respClientes && respClientes.error) ? respClientes.error : 'Error al listar clientes.',
        data: null
      };
    }

    var clientes = (respClientes.data && respClientes.data.clientes) ? respClientes.data.clientes : [];

    // 3) Calcular el próximo ID cliente (para el caso "Nuevo contacto")
    var nextIdCliente = generarIdCliente_();

    return {
      ok: true,
      error: null,
      data: {
        idPedido: idPedido,
        clientes: clientes,
        nextIdCliente: nextIdCliente
      }
    };

  } catch (e) {
    return {
      ok: false,
      error: e && e.message ? e.message : 'Error desconocido al inicializar nuevo pedido.',
      data: null
    };
  }
}

/**
 * Prepara y muestra registrarPedido.html con el contexto recibido desde
 * nuevoPedido.html o desde nuevoContacto.html.
 *
 * Payload esperado (flexible):
 * {
 *   pedidoId: string,
 *   clienteNombre: string, // opcional si llega un objeto cliente
 *   origen: string,        // ej: 'nuevoPedido' | 'nuevoContacto'
 *   cliente: {             // opcional
 *     nombre, telefono, direccion, sector, nota
 *   }
 * }
 */
function abrirRegistrarPedido(desdeNuevoPedidoPayload) {
  try {
    var clienteObj = (desdeNuevoPedidoPayload && desdeNuevoPedidoPayload.cliente)
      ? desdeNuevoPedidoPayload.cliente
      : null;

    var sectorInfo = obtenerPrecioYColumnaSector_(clienteObj && clienteObj.sector);
    var precioBase = (clienteObj && clienteObj.precioBase !== undefined && clienteObj.precioBase !== null)
      ? Number(clienteObj.precioBase)
      : sectorInfo.precio;

    var ctx = {
      pedidoId: (desdeNuevoPedidoPayload && desdeNuevoPedidoPayload.pedidoId)
        ? desdeNuevoPedidoPayload.pedidoId.toString().trim()
        : '',
      idContacto: (desdeNuevoPedidoPayload && (desdeNuevoPedidoPayload.idContacto || desdeNuevoPedidoPayload.idCliente))
        ? (desdeNuevoPedidoPayload.idContacto || desdeNuevoPedidoPayload.idCliente).toString().trim()
        : (clienteObj && clienteObj.idCliente) ? clienteObj.idCliente.toString().trim() : '',
      clienteNombre: (desdeNuevoPedidoPayload && desdeNuevoPedidoPayload.clienteNombre)
        ? desdeNuevoPedidoPayload.clienteNombre.toString().trim()
        : (clienteObj && clienteObj.nombre) ? clienteObj.nombre.toString().trim() : '',
      origen: (desdeNuevoPedidoPayload && desdeNuevoPedidoPayload.origen)
        ? desdeNuevoPedidoPayload.origen.toString().trim()
        : '',
      cliente: clienteObj,
      precioBase: (typeof precioBase === 'number' && !isNaN(precioBase)) ? precioBase : null,
      columnaSector: (clienteObj && clienteObj.columnaSector)
        ? clienteObj.columnaSector
        : (sectorInfo && sectorInfo.columna) ? sectorInfo.columna : ''
    };

    var userProps = PropertiesService.getUserProperties();
    userProps.setProperty('SECOMOVIL_CTX_NUEVO_PEDIDO', JSON.stringify(ctx));

    return {
      ok: true,
      error: null,
      data: {
        redirectTo: 'registrarPedido.html',
        ctx: ctx
      }
    };

  } catch (e) {
    return {
      ok: false,
      error: e && e.message ? e.message : 'Error desconocido al preparar registrarPedido.',
      data: null
    };
  }
}

/**
 * Prepara y muestra nuevoContacto.html con el ID de pedido y el origen
 * desde donde se abrió (nuevoPedido).
 *
 * Payload esperado (flexible):
 * {
 *   pedidoId: string,
 *   origen: string,
 *   cliente: {...} // opcional
 * }
 */
function abrirNuevoContacto(desdeNuevoPedidoPayload) {
  try {
    var idContacto = generarIdCliente_();
    var sectores = leerSectores_();
    var ctx = {
      pedidoId: (desdeNuevoPedidoPayload && desdeNuevoPedidoPayload.pedidoId)
        ? desdeNuevoPedidoPayload.pedidoId.toString().trim()
        : '',
      origen: (desdeNuevoPedidoPayload && desdeNuevoPedidoPayload.origen)
        ? desdeNuevoPedidoPayload.origen.toString().trim()
        : '',
      idContacto: idContacto,
      cliente: (desdeNuevoPedidoPayload && desdeNuevoPedidoPayload.cliente)
        ? desdeNuevoPedidoPayload.cliente
        : null,
      sectores: sectores
    };

    var userProps = PropertiesService.getUserProperties();
    userProps.setProperty('SECOMOVIL_CTX_NUEVO_PEDIDO', JSON.stringify(ctx));

    return {
      ok: true,
      error: null,
      data: {
        redirectTo: 'nuevoContacto.html',
        ctx: ctx
      }
    };

  } catch (e) {
    return {
      ok: false,
      error: e && e.message ? e.message : 'Error desconocido al preparar nuevoContacto.',
      data: null
    };
  }
}

/**
 * Abre la página editarContacto.html con los datos del contacto existente.
 * @param {Object} payload
 * @returns {{ok:boolean, error:string|null, data:Object|null}}
 */
function abrirEditarContacto(payload) {
  try {
    payload = payload || {};
    var idContacto = (payload.idContacto || payload.idCliente || '').toString().trim();
    var clienteObj = payload.cliente || null;

    if (!clienteObj && idContacto) {
      clienteObj = obtenerClientePorId_(idContacto);
    }

    if (!clienteObj) {
      return {
        ok: false,
        error: 'No se encontró el contacto para editar.',
        data: null
      };
    }

    var sectores = leerSectores_();

    var ctx = {
      pedidoId: (payload && payload.pedidoId) ? payload.pedidoId.toString().trim() : '',
      origen: (payload && payload.origen) ? payload.origen.toString().trim() : '',
      idContacto: clienteObj.idCliente || idContacto,
      cliente: clienteObj,
      sectores: sectores
    };

    var userProps = PropertiesService.getUserProperties();
    userProps.setProperty('SECOMOVIL_CTX_NUEVO_PEDIDO', JSON.stringify(ctx));

    return {
      ok: true,
      error: null,
      data: {
        redirectTo: 'editarContacto.html',
        ctx: ctx
      }
    };

  } catch (e) {
    return {
      ok: false,
      error: e && e.message ? e.message : 'Error desconocido al preparar editarContacto.',
      data: null
    };
  }
}

/**
 * Registra un nuevo contacto en BASE DE CLIENTES utilizando un ID ya generado
 * (idContacto) y devuelve el contexto necesario para continuar con el flujo
 * del pedido.
 *
 * @param {Object} data - Datos del contacto desde el frontend
 * @param {Object} payload - Contexto enviado desde nuevoContacto.html
 * @returns {Object} Respuesta con éxito/fracaso y datos del contacto
 */
function registrarNuevoContacto(data, payload) {
  try {
    data = data || {};
    payload = payload || {};

    var nombre = (data.nombre || '').toString().trim();
    var telefono = (data.telefono || '').toString().trim();
    var direccion = (data.direccion || '').toString().trim();
    var sector = (data.sector || '').toString().trim();
    var nota = (data.nota || '').toString().trim();
    var idContacto = (data.idContacto || payload.idContacto || '').toString().trim();
    if (!idContacto) {
      idContacto = generarIdCliente_();
    }
    var pedidoId = (data.pedidoId || payload.pedidoId || '').toString().trim();

    var sectorInfo = obtenerPrecioYColumnaSector_(sector);

    if (!nombre || !telefono || !direccion || !sector || !idContacto) {
      return {
        ok: false,
        success: false,
        error: 'Faltan datos obligatorios para crear el contacto.',
        data: null
      };
    }

    var sheet = getSheet_('BASE DE CLIENTES');
    var newRow = getFirstEmptyRowInColumn_(sheet, 1);
    var hoy = new Date();

    sheet.getRange(newRow, 1).setValue(nombre);
    sheet.getRange(newRow, 2).setValue(telefono);
    sheet.getRange(newRow, 3).setValue(direccion);
    sheet.getRange(newRow, 4).setValue(hoy);
    sheet.getRange(newRow, 5).setValue('');
    sheet.getRange(newRow, 6).setValue('Nuevo');
    sheet.getRange(newRow, 7).setValue(nota);
    sheet.getRange(newRow, 8).setValue(idContacto);
    sheet.getRange(newRow, 9).setValue(sector);

    var totalRows = sheet.getLastRow();
    if (totalRows > 2) {
      sheet.getRange(2, 1, totalRows - 1, 9).sort({ column: 1, ascending: true });
    }

    var cliente = {
      idCliente: idContacto,
      nombre: nombre,
      telefono: telefono,
      direccion: direccion,
      sector: sector,
      nota: nota,
      columnaSector: sectorInfo && sectorInfo.columna ? sectorInfo.columna : '',
      precioBase: sectorInfo && sectorInfo.precio !== null ? sectorInfo.precio : null
    };

    return {
      ok: true,
      success: true,
      error: null,
      data: {
        cliente: cliente,
        idContacto: idContacto,
        pedidoId: pedidoId,
        clienteNombre: nombre,
        columnaSector: cliente.columnaSector,
        precioBase: cliente.precioBase
      }
    };

  } catch (e) {
    return {
      ok: false,
      success: false,
      error: e && e.message ? e.message : 'Error desconocido al registrar el contacto.',
      data: null
    };
  }
}

/**
 * Actualiza un contacto existente desde editarContacto.html y devuelve el contexto
 * para volver a registrarPedido.html con los datos actualizados.
 * @param {Object} data
 * @param {Object} payload
 * @returns {{ok:boolean, error:string|null, data:Object|null}}
 */
function actualizarContactoDesdeEditar(data, payload) {
  try {
    data = data || {};
    payload = payload || {};

    var idContacto = (data.idContacto || payload.idContacto || '').toString().trim();
    var nombre = (data.nombre || '').toString().trim();
    var telefono = (data.telefono || '').toString().trim();
    var direccion = (data.direccion || '').toString().trim();
    var sector = (data.sector || '').toString().trim();
    var nota = (data.nota || '').toString().trim();
    var pedidoId = (data.pedidoId || payload.pedidoId || '').toString().trim();

    if (!idContacto || !nombre || !telefono || !direccion || !sector) {
      return {
        ok: false,
        error: 'Faltan datos obligatorios para actualizar el contacto.',
        data: null
      };
    }

    var resp = actualizarCliente(idContacto, {
      nombre: nombre,
      telefono: telefono,
      direccion: direccion,
      sector: sector,
      nota: nota
    });

    if (!resp || !resp.ok || !resp.data || !resp.data.cliente) {
      return {
        ok: false,
        error: (resp && resp.error) ? resp.error : 'No se pudo actualizar el contacto.',
        data: null
      };
    }

    var cliente = resp.data.cliente;

    var ctx = {
      origen: payload && payload.origen ? payload.origen : 'editarContacto',
      pedidoId: pedidoId,
      idContacto: cliente.idCliente,
      clienteNombre: cliente.nombre,
      cliente: cliente,
      pedido: payload && payload.pedido ? payload.pedido : null,
      pedidoEnEdicion: payload && payload.pedidoEnEdicion ? payload.pedidoEnEdicion : null,
      sectores: payload && Array.isArray(payload.sectores) ? payload.sectores : []
    };

    return {
      ok: true,
      error: null,
      data: {
        cliente: cliente,
        ctx: ctx
      }
    };

  } catch (e) {
    return {
      ok: false,
      error: e && e.message ? e.message : 'Error desconocido al actualizar el contacto.',
      data: null
    };
  }
}

/**
 * Inserta un script al HTML generado para exponer el contexto (pedido,
 * cliente) y, opcionalmente, aplicar los datos en registrarPedido.html.
 *
 * @param {GoogleAppsScript.HTML.HtmlOutput} htmlOutput
 * @param {Object} ctx
 * @param {Object} [opciones]
 * @param {boolean} [opciones.aplicarContextoRegistrar]
 */
function inyectarContextoEnHtml_(htmlOutput, ctx, opciones) {
  var opts = opciones || {};
  var jsonSeguro = JSON.stringify(ctx || {}).replace(/<\//g, '<\/');

  var script = 'window.SECO_CTX = ' + jsonSeguro + ';';

  if (opts.aplicarContextoRegistrar) {
    script += "(function(){try{"
      + "var ctx=window.SECO_CTX||{};"
      + "var idEl=document.getElementById('pedidoId');if(idEl&&ctx.pedidoId){idEl.textContent=ctx.pedidoId;}"
      + "var card=document.querySelector('.contact-card');"
      + "if(card){"
        + "var nombre=(ctx.cliente&&ctx.cliente.nombre)||ctx.clienteNombre||'';"
        + "var whatsapp=(ctx.cliente&&(ctx.cliente.telefono||ctx.cliente.whatsapp))||ctx.whatsapp||'';"
        + "var subline=nombre?'':'Selecciona o crea un contacto';"
        + "var headerLines=card.querySelectorAll('.contact-top .txt-body div');"
        + "if(headerLines.length>0){headerLines[0].textContent=nombre||'Contacto';}"
        + "if(headerLines.length>1){headerLines[1].textContent=subline;}"
        + "var labels=card.querySelectorAll('.label');"
        + "labels.forEach(function(lbl){var valueEl=lbl.nextElementSibling;if(!valueEl){return;}var key=lbl.textContent.trim().toLowerCase();"
          + "if(key==='sector'&&ctx.cliente&&ctx.cliente.sector){valueEl.textContent=ctx.cliente.sector;}"
          + "else if(key==='direccion'&&ctx.cliente&&ctx.cliente.direccion){valueEl.textContent=ctx.cliente.direccion;}"
          + "else if(key==='nota'&&ctx.cliente&&ctx.cliente.nota){valueEl.textContent=ctx.cliente.nota;}"
          + "else if(key==='whatsapp'&&whatsapp){valueEl.textContent=whatsapp;}"
        + "});"
      + "}"
    + "}catch(e){}})();";
  }

  var tag = '<script>' + script + '</script>';
  var contenido = htmlOutput.getContent();

  if (contenido.indexOf('</body>') !== -1) {
    contenido = contenido.replace('</body>', tag + '</body>');
  } else {
    contenido += tag;
  }

  htmlOutput.setContent(contenido);
}

/**
 * Guarda el pedido recibido desde registrarPedido.html y abre la pantalla
 * de confirmación pedidoRegistrado.html con el contexto cargado.
 *
 * @param {Object} payload
 * @returns {{ok:boolean, error:string|null, data:Object|null}}
 */
function registrarPedidoDesdeRegistrar(payload) {
  try {
    payload = payload || {};

    var cliente = payload.cliente || {};
    var nombre = (cliente.nombre || payload.clienteNombre || '').toString().trim();
    var telefono = (cliente.telefono || cliente.whatsapp || payload.whatsapp || '').toString().trim();
    var direccion = (cliente.direccion || payload.direccion || '').toString().trim();
    var sector = (cliente.sector || payload.sector || '').toString().trim();
    var nota = (payload.nota || payload.notas || cliente.nota || '').toString().trim();

    var fechaTexto = (payload.fechaTexto || payload.fechaEntrega || '').toString();
    var fechaObj = normalizarFechaDesdeTexto_(fechaTexto);
    if (!fechaObj && payload.fechaISO) {
      var fAlt = normalizarFechaDesdeTexto_(payload.fechaISO);
      if (fAlt) fechaObj = fAlt;
      else {
        var dTmp = new Date(payload.fechaISO);
        if (!isNaN(dTmp.getTime())) fechaObj = dTmp;
      }
    }
    if (!fechaObj) {
      fechaObj = new Date();
    }
    var tz = Session.getScriptTimeZone();
    var fechaISO = Utilities.formatDate(fechaObj, tz, 'yyyy-MM-dd');

    var hora = (payload.hora || payload.horaEntrega || '').toString().trim();
    var cantidad = Number(payload.cantidad || 1);
    if (!isFinite(cantidad) || cantidad < 1) cantidad = 1;

    var precioSector = obtenerPrecioYColumnaSector_(sector);
    var unitPrice = (precioSector && typeof precioSector.precio === 'number' && !isNaN(precioSector.precio))
      ? Number(precioSector.precio)
      : null;
    if (unitPrice === null && payload.precioUnitario !== undefined) {
      var tmp = Number(payload.precioUnitario);
      if (!isNaN(tmp)) unitPrice = tmp;
    }
    if (unitPrice === null || !isFinite(unitPrice)) {
      unitPrice = 1;
    }

    var totalCalc = unitPrice * cantidad;
    if (payload.totalCalculado !== undefined) {
      var tTmp = Number(payload.totalCalculado);
      if (!isNaN(tTmp)) totalCalc = tTmp;
    }

    var pedidoId = (payload.pedidoId || '').toString().trim();
    if (!pedidoId) {
      pedidoId = generarIdPedido_(fechaISO, nombre);
    }

    var sheet = getSheet_('PEDIDOS DIARIOS');
    var row = getFirstEmptyRowInColumn_(sheet, 1);

    sheet.getRange(row, 1).setValue(fechaObj);
    sheet.getRange(row, 2).setValue(nombre);
    sheet.getRange(row, 3).setValue(telefono);
    sheet.getRange(row, 4).setValue(direccion);
    sheet.getRange(row, 5).setValue(cantidad);
    sheet.getRange(row, 6).setValue(totalCalc);
    sheet.getRange(row, 7).setValue('No');
    sheet.getRange(row, 8).setValue(hora);
    sheet.getRange(row, 9).setValue(nota);
    sheet.getRange(row, 10).setValue('Pendiente');
    sheet.getRange(row, 11).setValue(pedidoId);

    var ctx = {
      pedidoId: pedidoId,
      clienteNombre: nombre,
      idContacto: payload.idContacto || cliente.idCliente || '',
      origen: payload.origen || '',
      fechaEntrega: fechaISO,
      horaEntrega: hora,
      cantidad: cantidad,
      precioUnitario: unitPrice,
      total: totalCalc,
      sector: sector,
      nota: nota,
      direccion: direccion,
      whatsapp: telefono,
      cliente: {
        idCliente: payload.idContacto || cliente.idCliente || '',
        nombre: nombre,
        telefono: telefono,
        direccion: direccion,
        sector: sector,
        nota: nota,
        whatsapp: telefono
      }
    };

    var userProps = PropertiesService.getUserProperties();
    userProps.setProperty('SECOMOVIL_CTX_NUEVO_PEDIDO', JSON.stringify(ctx));

    sortPedidosDiarios_();

    return {
      ok: true,
      error: null,
      data: { pedido: ctx }
    };

  } catch (e) {
    return {
      ok: false,
      error: e && e.message ? e.message : 'Error desconocido al registrar el pedido.',
      data: null
    };
  }
}

function construirMensajePedidoRegistrado_(pedido) {
  pedido = pedido || {};

  function formatCurrency(num) {
    var n = Number(num);
    if (!isFinite(n)) return '';
    return n.toFixed(2) + ' $';
  }

  var partes = ['Tu pedido SecoMóvil esta registrado :'];
  var datos = [];

  if (pedido.pedidoId) {
    datos.push('ID ' + pedido.pedidoId);
  }

  var nombre = (pedido.clienteNombre || (pedido.cliente && pedido.cliente.nombre) || '').toString().trim();
  if (nombre) {
    datos.push('Cliente: ' + nombre);
  }

  var cantidad = Number(pedido.cantidad || 0);
  if (isFinite(cantidad) && cantidad > 0) {
    var unitTxt = formatCurrency(pedido.precioUnitario);
    var totalTxt = formatCurrency(pedido.total);
    var linea = 'Cantidad: ' + cantidad;
    if (unitTxt) {
      linea += ' (unit ' + unitTxt + ')';
    }
    if (totalTxt) {
      linea += ' total ' + totalTxt;
    }
    datos.push(linea);
  }

  var fechaEntrega = (pedido.fechaEntrega || '').toString().trim();
  var horaEntrega = (pedido.horaEntrega || '').toString().trim();
  if (fechaEntrega || horaEntrega) {
    datos.push('Entrega: ' + fechaEntrega + (horaEntrega ? ' ' + horaEntrega : ''));
  }

  if (pedido.direccion) {
    datos.push('Dirección: ' + pedido.direccion);
  }

  return partes.concat(datos).join(' ');
}

function enviarWhatsappMensaje_(telefono, mensaje) {
  // Punto centralizado para integrar con la API real de WhatsApp si existe.
  Logger.log('[WhatsApp] ' + telefono + ' -> ' + mensaje);
  return true;
}

function enviarWhatsappPedidoRegistrado(payload) {
  try {
    var datos = payload || {};
    var telefono = (datos.whatsapp || (datos.cliente && datos.cliente.telefono) || '').toString().trim();
    if (!telefono) {
      return { ok: false, error: 'No se proporcionó número de WhatsApp.', data: null };
    }

    var mensaje = construirMensajePedidoRegistrado_(datos);
    enviarWhatsappMensaje_(telefono, mensaje);

    return {
      ok: true,
      error: null,
      data: { mensaje: mensaje }
    };
  } catch (e) {
    return {
      ok: false,
      error: e && e.message ? e.message : 'Error al enviar el mensaje de WhatsApp.',
      data: null
    };
  }
}

/**
 * Crea un pedido nuevo en PEDIDOS DIARIOS a partir de los datos
 * del formulario registrarPedido.html.
 *
 * Contrato JSON (PAQUETE 5):
 *
 * Petición:
 * {
 *   cliente: {
 *     idCliente: string,
 *     nombre: string,
 *     telefono: string,
 *     direccion: string,
 *     sector: string
 *   },
 *   fechaEntrega: string,   // 'YYYY-MM-DD'
 *   horaEntrega: string,    // 'HH:MM'
 *   cantidad: number,
 *   notas: string,
 *   enviarWhatsApp: boolean
 * }
 *
 * Respuesta:
 * {
 *   ok: true|false,
 *   error: string|null,
 *   data: {
 *     pedido: {
 *       idPedido: string,
 *       fechaEntrega: string,
 *       horaEntrega: string,
 *       cantidad: number,
 *       total: number|null,
 *       estado: string,
 *       notas: string,
 *       cliente: {
 *         idCliente: string|null,
 *         nombre: string,
 *         telefono: string,
 *         direccion: string,
 *         sector: string
 *       }
 *     }
 *   } | null
 * }
 */
function crearPedido(data) {
  try {
    // 0) Validación básica de entrada
    if (!data || !data.cliente) {
      return {
        ok: false,
        error: 'Datos de cliente no proporcionados para crear el pedido.',
        data: null
      };
    }

    var cliente = data.cliente;

    if (!cliente.nombre || !cliente.telefono || !cliente.direccion || !cliente.sector) {
      return {
        ok: false,
        error: 'Faltan datos obligatorios del cliente para crear el pedido.',
        data: null
      };
    }
    if (!data.fechaEntrega || !data.horaEntrega || !data.cantidad) {
      return {
        ok: false,
        error: 'Faltan fecha, hora o cantidad para crear el pedido.',
        data: null
      };
    }

    var fechaEntregaStr = data.fechaEntrega; // 'YYYY-MM-DD'
    var horaEntregaStr  = data.horaEntrega;  // 'HH:MM'
    var cantidadNum     = Number(data.cantidad) || 1;
    var notasStr        = data.notas || '';

    // 1) Generar el ID de pedido según el formato definido
    //    Usamos la función interna generarIdPedido_(fecha, cliente)
    var idPedido = generarIdPedido_(fechaEntregaStr, cliente.nombre);

    // 2) Guardar el pedido en PEDIDOS DIARIOS
    //    Usamos la función existente guardarPedido() para respetar el plan.
    guardarPedido({
      fecha: fechaEntregaStr,
      cliente: cliente.nombre,
      telefono: cliente.telefono,
      direccion: cliente.direccion,
      cantidad: cantidadNum,
      pagoRecibido: 'No',
      horaEntrega: horaEntregaStr,
      notas: notasStr,
      estado: 'Pendiente'
    });

    // 3) Localizar la fila recién creada para poder escribir el ID en la columna K
    var sheetPedidos = getSheet_('PEDIDOS DIARIOS');
    var lastRowPedidos = sheetPedidos.getLastRow();
    var pedidosValores = [];
    if (lastRowPedidos > 1) {
      pedidosValores = sheetPedidos.getRange(2, 1, lastRowPedidos - 1, 10).getValues();
    }

    var filaEncontrada = -1;
    var tz = Session.getScriptTimeZone();

    for (var i = 0; i < pedidosValores.length; i++) {
      var r = pedidosValores[i];
      var fechaCell   = r[0]; // A
      var clienteCell = (r[1] || '').toString().trim();
      var telCell     = (r[2] || '').toString().trim();
      var dirCell     = (r[3] || '').toString().trim();
      var cantCell    = Number(r[4] || 0);     // E
      var horaCell    = (r[7] || '').toString().trim(); // H
      var notasCell   = (r[8] || '').toString().trim(); // I
      var estadoCell  = (r[9] || '').toString().trim() || 'Pendiente'; // J

      // Normalizar fecha
      var fechaCellStr = '';
      if (fechaCell) {
        var d = new Date(fechaCell);
        fechaCellStr = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
      }

      if (
        fechaCellStr === fechaEntregaStr &&
        clienteCell   === cliente.nombre &&
        telCell       === cliente.telefono &&
        dirCell       === cliente.direccion &&
        cantCell      === cantidadNum &&
        horaCell      === horaEntregaStr &&
        notasCell     === notasStr &&
        estadoCell    === 'Pendiente'
      ) {
        filaEncontrada = i + 2; // +2 porque empezamos a leer en la fila 2
        break;
      }
    }

    var pedidoFront;

    if (filaEncontrada !== -1) {
      // 3.a) Escribimos el ID en la columna K (11)
      sheetPedidos.getRange(filaEncontrada, 11).setValue(idPedido);

      // 3.b) Reconstruimos el PedidoFront desde la fila real
      var rowValues = sheetPedidos.getRange(filaEncontrada, 1, 1, 11).getValues()[0];
      var fechaCell2   = rowValues[0]; // A
      var cantidadCell = rowValues[4]; // E
      var totalCell    = rowValues[5]; // F
      var horaCell2    = rowValues[7]; // H
      var notasCell2   = rowValues[8]; // I
      var estadoCell2  = rowValues[9] || 'Pendiente'; // J
      var idCell       = rowValues[10] || idPedido;   // K

      var fechaEntregaOut = '';
      if (fechaCell2) {
        var d2 = new Date(fechaCell2);
        fechaEntregaOut = Utilities.formatDate(d2, tz, 'yyyy-MM-dd');
      } else {
        fechaEntregaOut = fechaEntregaStr;
      }

      var totalNumber = null;
      if (totalCell !== '' && totalCell !== null && totalCell !== undefined) {
        var t = Number(totalCell);
        if (!isNaN(t)) {
          totalNumber = t;
        }
      }

      pedidoFront = {
        idPedido: (idCell || '').toString(),
        fechaEntrega: fechaEntregaOut,
        horaEntrega: (horaCell2 || '').toString(),
        cantidad: Number(cantidadCell || cantidadNum),
        total: totalNumber,
        estado: estadoCell2.toString(),
        notas: (notasCell2 || '').toString(),
        cliente: {
          idCliente: cliente.idCliente || '',
          nombre: cliente.nombre,
          telefono: cliente.telefono,
          direccion: cliente.direccion,
          sector: cliente.sector
        }
      };

    } else {
      // Cas exceptionnel: si no encontramos la fila, no inventamos nada.
      // Devolvemos lo que sabemos, con el idPedido generado.
      pedidoFront = {
        idPedido: idPedido,
        fechaEntrega: fechaEntregaStr,
        horaEntrega: horaEntregaStr,
        cantidad: cantidadNum,
        total: null,
        estado: 'Pendiente',
        notas: notasStr,
        cliente: {
          idCliente: cliente.idCliente || '',
          nombre: cliente.nombre,
          telefono: cliente.telefono,
          direccion: cliente.direccion,
          sector: cliente.sector
        }
      };
    }

    // 4) Actualizar la BASE DE CLIENTES (último pedido, dirección, notas)
    actualizarUltimoPedidoCliente({
      nombre: cliente.nombre,
      telefono: cliente.telefono,
      fecha: fechaEntregaStr,
      direccion: cliente.direccion,
      notas: notasStr
    });

    // 5) Respuesta estándar
    return {
      ok: true,
      error: null,
      data: {
        pedido: pedidoFront
      }
    };

  } catch (e) {
    return {
      ok: false,
      error: e && e.message ? e.message : 'Error desconocido al crear el pedido.',
      data: null
    };
  }
}

/**
 * Devuelve la lista de pedidos del día (PEDIDOS DIARIOS)
 * para una fecha dada (o para hoy si fecha es null/undefined).
 *
 * Contrato JSON (PAQUETE 5):
 *
 * Petición (conceptual):
 *   getPedidosDelDia(fechaString|null)
 *
 * Respuesta:
 * {
 *   ok: true|false,
 *   error: string|null,
 *   data: {
 *     pedidos: PedidoFront[]
 *   } | null
 * }
 *
 * PedidoFront:
 * {
 *   idPedido: string,
 *   fechaEntrega: string,
 *   horaEntrega: string,
 *   cantidad: number,
 *   total: number | null,
 *   estado: string,
 *   notas: string,
 *   cliente: {
 *     idCliente: string | null,
 *     nombre: string,
 *     telefono: string,
 *     direccion: string,
 *     sector: string
 *   }
 * }
 */
function getPedidosDelDia(fecha) {
  try {
    var tz = Session.getScriptTimeZone();

    function normalizarHoraEntrega(valor) {
      if (valor === null || valor === undefined || valor === '') {
        return '';
      }

      if (Object.prototype.toString.call(valor) === '[object Date]' && !isNaN(valor.getTime())) {
        return Utilities.formatDate(valor, tz, 'H:mm');
      }

      var str = valor.toString().trim();
      if (!str) {
        return '';
      }

      var match = str.match(/^(\d{1,2}):(\d{2})$/);
      if (match) {
        var hora = Number(match[1]);
        var minutos = match[2];
        if (!isNaN(hora)) {
          return hora + ':' + minutos;
        }
      }

      return str;
    }

    // 1) Normalizar la fecha objetivo en formato 'yyyy-MM-dd'
    var fechaObjetivoStr;
    if (fecha) {
      var dReq;

      // a) Si viene como string "yyyy-MM-dd" (caso getProximaEntrega),
      //    crear la fecha en la zona horaria local para evitar desfaces UTC.
      if (typeof fecha === 'string') {
        var isoMatch = fecha.trim().match(/^(\d{4})-(\d{2})-(\d{2})$/);
        if (isoMatch) {
          var y = Number(isoMatch[1]);
          var m = Number(isoMatch[2]) - 1;
          var d = Number(isoMatch[3]);
          dReq = new Date(y, m, d, 0, 0, 0, 0);
        }
      }

      // b) Si no se pudo parsear como ISO, intentar con el constructor normal
      //    (Date o string genérico). Si falla, caeremos al valor por defecto.
      if (!dReq) {
        dReq = new Date(fecha);
      }

      // c) Si sigue siendo inválida, usar hoy.
      if (isNaN(dReq.getTime())) {
        dReq = new Date();
      }

      fechaObjetivoStr = Utilities.formatDate(dReq, tz, 'yyyy-MM-dd');
    } else {
      var hoy = new Date();
      fechaObjetivoStr = Utilities.formatDate(hoy, tz, 'yyyy-MM-dd');
    }

    var sheetPedidos = getSheet_('PEDIDOS DIARIOS');
    var lastRowPedidos = sheetPedidos.getLastRow();
    if (lastRowPedidos < 2) {
      // No hay pedidos
      return {
        ok: true,
        error: null,
        data: { pedidos: [] }
      };
    }

    // Leer todos los pedidos (A:K = 1:11)
    var dataPedidos = sheetPedidos.getRange(2, 1, lastRowPedidos - 1, 11).getValues();

    // 2) Cargar la base de clientes para obtener idCliente y sector
    var sheetClientes = getSheet_('BASE DE CLIENTES');
    var lastRowClientes = sheetClientes.getLastRow();
    var mapaClientes = []; // lista simple, como en actualizarUltimoPedidoCliente

    if (lastRowClientes > 1) {
      // A Nombre, B Teléfono, ..., H ID Cliente, I Sector
      var dataClientes = sheetClientes.getRange(2, 1, lastRowClientes - 1, 9).getValues();
      dataClientes.forEach(function (row) {
        var nombre  = (row[0] || '').toString().trim();
        var tel     = (row[1] || '').toString().trim();
        var idCli   = (row[7] || '').toString().trim();
        var sector  = (row[8] || '').toString().trim();
        if (nombre || tel) {
          mapaClientes.push({
            nombre: nombre,
            telefono: tel,
            idCliente: idCli,
            sector: sector
          });
        }
      });
    }

    // Función de ayuda para encontrar cliente por teléfono o nombre
    function encontrarCliente(nombrePedido, telPedido) {
      var nombreP = (nombrePedido || '').toString().trim();
      var telP    = (telPedido    || '').toString().trim();

      // 1. Intentar por teléfono (coincidencia exacta)
      if (telP) {
        for (var i = 0; i < mapaClientes.length; i++) {
          if (mapaClientes[i].telefono === telP) {
            return mapaClientes[i];
          }
        }
      }

      // 2. Si no, intentar por nombre (sin mayúsculas/minúsculas)
      if (nombreP) {
        var nombrePLower = nombreP.toLowerCase();
        for (var j = 0; j < mapaClientes.length; j++) {
          if (mapaClientes[j].nombre.toLowerCase() === nombrePLower) {
            return mapaClientes[j];
          }
        }
      }

      // 3. Si no, devolvemos algo coherente pero sin inventar ID/sector
      return {
        nombre: nombreP,
        telefono: telP,
        idCliente: '',
        sector: ''
      };
    }

    // 3) Filtrar pedidos por fecha y mapear a PedidoFront
    var pedidosFront = [];

    dataPedidos.forEach(function (row) {
      var fechaCell   = row[0];  // A Fecha menú
      var clienteNom  = row[1];  // B
      var telefono    = row[2];  // C
      var direccion   = row[3];  // D
      var cantidad    = row[4];  // E
      var totalCell   = row[5];  // F
      var horaEntrega = row[7];  // H
      var notas       = row[8];  // I
      var estado      = row[9];  // J
      var idPedido    = row[10]; // K

      // Ignorer les lignes totalement vides
      if (!fechaCell && !clienteNom && !telefono && !direccion && !cantidad) {
        return;
      }

      // Normalizar la fecha de la línea
      var fechaCellStr = '';
      if (fechaCell) {
        var d = new Date(fechaCell);
        fechaCellStr = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
      }

      // Mantener solo los pedidos de la fecha objetivo
      if (fechaCellStr !== fechaObjetivoStr) {
        return;
      }

      // Convertir total
      var totalNum = null;
      if (totalCell !== '' && totalCell !== null && totalCell !== undefined) {
        var t = Number(totalCell);
        if (!isNaN(t)) {
          totalNum = t;
        }
      }

      var estadoStr = (estado || '').toString().trim();
      if (!estadoStr) {
        estadoStr = 'Pendiente';
      }

      var idPedidoStr = (idPedido || '').toString().trim();

      // Récupérer info client (id + sector)
      var cli = encontrarCliente(clienteNom, telefono);
      var infoPrecioSector = obtenerPrecioYColumnaSector_(cli.sector || '');
      var precioUnitarioSector = typeof infoPrecioSector.precio === 'number' && isFinite(infoPrecioSector.precio)
        ? infoPrecioSector.precio
        : null;

      var pedidoFront = {
        idPedido: idPedidoStr,
        fechaEntrega: fechaCellStr,
        horaEntrega: normalizarHoraEntrega(horaEntrega),
        cantidad: Number(cantidad || 0),
        total: totalNum,
        precioUnitarioSector: precioUnitarioSector,
        estado: estadoStr,
        notas: (notas || '').toString(),
        cliente: {
          idCliente: cli.idCliente || '',
          nombre: (clienteNom || '').toString(),
          telefono: (telefono || '').toString(),
          direccion: (direccion || '').toString(),
          sector: cli.sector || ''
        }
      };

      pedidosFront.push(pedidoFront);
    });

    return {
      ok: true,
      error: null,
      data: {
        pedidos: pedidosFront
      }
    };

  } catch (e) {
    return {
      ok: false,
      error: e && e.message ? e.message : 'Error desconocido al obtener los pedidos del día.',
      data: null
    };
  }
}

/**
 * Devuelve el próximo pedido pendiente para mostrar en INICIO.HTML.
 *
 * Contrato JSON (PAQUETE 5):
 *
 * Respuesta:
 * {
 *   ok: true|false,
 *   error: string|null,
 *   data: {
 *     pedido: PedidoFront | null
 *   }
 * }
 */
function getProximaEntrega() {
  try {
    var tz = Session.getScriptTimeZone();
    var hoyStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
    var resp = getPedidosDelDia(hoyStr);

    if (!resp.ok) {
      return {
        ok: false,
        error: resp.error || 'Error al obtener los pedidos del día.',
        data: null
      };
    }

    var pedidos = (resp.data && resp.data.pedidos) || [];

    var pendientesDeHoy = pedidos.filter(function (pedido) {
      var fechaEntrega = (pedido.fechaEntrega || '').toString();
      var estado = (pedido.estado || '').toString().toLowerCase();
      return fechaEntrega === hoyStr && estado === 'pendiente';
    });

    function horaToMinutes(horaStr) {
      var texto = (horaStr || '').toString().trim();
      var partes = texto.split(':');
      if (partes.length !== 2) return null;
      var h = Number(partes[0]);
      var m = Number(partes[1]);
      if (!isFinite(h) || !isFinite(m)) return null;
      if (h < 0 || h > 23 || m < 0 || m > 59) return null;
      return h * 60 + m;
    }

    pendientesDeHoy.sort(function (a, b) {
      var fechaA = (a.fechaEntrega || '').toString();
      var fechaB = (b.fechaEntrega || '').toString();
      if (fechaA !== fechaB) {
        return fechaA < fechaB ? -1 : 1;
      }

      var horaA = horaToMinutes(a.horaEntrega);
      var horaB = horaToMinutes(b.horaEntrega);

      if (horaA === horaB) return 0;
      if (horaA === null) return 1;
      if (horaB === null) return -1;
      return horaA - horaB;
    });

    var siguiente = pendientesDeHoy.length > 0 ? pendientesDeHoy[0] : null;

    return {
      ok: true,
      error: null,
      data: { pedido: siguiente }
    };

  } catch (e) {
    return {
      ok: false,
      error: e && e.message ? e.message : 'Error desconocido al obtener la próxima entrega.',
      data: null
    };
  }
}

/**
 * Edita un pedido existente en PEDIDOS DIARIOS identificado por su ID.
 *
 * Petición:
 * {
 *   idPedido: string,
 *   fechaEntrega: string,  // 'YYYY-MM-DD'
 *   horaEntrega: string,   // 'HH:MM'
 *   cantidad: number,
 *   notas: string
 * }
 *
 * Respuesta:
 * {
 *   ok: true|false,
 *   error: string|null,
 *   data: { pedido: PedidoFront } | null
 * }
 */
function editarPedidoPorId(data) {
  try {
    if (!data || !data.idPedido) {
      return {
        ok: false,
        error: 'No se proporcionó idPedido para editar.',
        data: null
      };
    }

    var idObjetivo = data.idPedido.toString().trim();
    var tz = Session.getScriptTimeZone();

    var sheetPedidos = getSheet_('PEDIDOS DIARIOS');
    var lastRowPedidos = sheetPedidos.getLastRow();
    if (lastRowPedidos < 2) {
      return {
        ok: false,
        error: 'No hay pedidos registrados.',
        data: null
      };
    }

    // Buscar la fila por ID Pedido (columna K = 11)
    var dataPedidos = sheetPedidos.getRange(2, 1, lastRowPedidos - 1, 11).getValues();
    var filaEncontrada = -1;

    for (var i = 0; i < dataPedidos.length; i++) {
      var row = dataPedidos[i];
      var idCell = (row[10] || '').toString().trim(); // K
      if (idCell === idObjetivo) {
        filaEncontrada = i + 2; // +2 porque empezamos en la fila 2
        break;
      }
    }

    if (filaEncontrada === -1) {
      return {
        ok: false,
        error: 'No se encontró el pedido con el ID especificado.',
        data: null
      };
    }

    // Actualizar columnas modificables:
    // A Fecha, E Cantidad, H Hora, I Notas
    if (Object.prototype.hasOwnProperty.call(data, 'fechaEntrega')) {
      sheetPedidos.getRange(filaEncontrada, 1).setValue(data.fechaEntrega || ''); // A
    }
    if (Object.prototype.hasOwnProperty.call(data, 'cantidad')) {
      sheetPedidos.getRange(filaEncontrada, 5).setValue(Number(data.cantidad) || 1); // E
    }
    if (Object.prototype.hasOwnProperty.call(data, 'horaEntrega')) {
      sheetPedidos.getRange(filaEncontrada, 8).setValue(data.horaEntrega || ''); // H
    }
    if (Object.prototype.hasOwnProperty.call(data, 'notas')) {
      sheetPedidos.getRange(filaEncontrada, 9).setValue(data.notas || ''); // I
    }

    sortPedidosDiarios_();

    // Volver a ubicar la fila tras el ordenamiento
    var rowValues = null;
    var filaOrdenada = null;
    var lastRowOrdenado = sheetPedidos.getLastRow();
    if (lastRowOrdenado > 1) {
      var dataPedidosOrdenados = sheetPedidos.getRange(2, 1, lastRowOrdenado - 1, 11).getValues();
      for (var idx = 0; idx < dataPedidosOrdenados.length; idx++) {
        var idFila = (dataPedidosOrdenados[idx][10] || '').toString().trim();
        if (idFila === idObjetivo) {
          rowValues = dataPedidosOrdenados[idx];
          filaOrdenada = idx + 2; // fila real en la hoja (comienza en 2)
          break;
        }
      }
    }

    if (!rowValues) {
      return {
        ok: false,
        error: 'No se encontró el pedido con el ID especificado.',
        data: null
      };
    }

    var fechaCell   = rowValues[0];  // A
    var clienteNom  = rowValues[1];  // B
    var telefono    = rowValues[2];  // C
    var direccion   = rowValues[3];  // D
    var cantidad    = rowValues[4];  // E
    var totalCell   = rowValues[5];  // F
    var horaEntrega = rowValues[7];  // H
    var notas       = rowValues[8];  // I
    var estado      = rowValues[9];  // J
    var idPedido    = rowValues[10]; // K

    var fechaStr = '';
    if (fechaCell) {
      var d = new Date(fechaCell);
      fechaStr = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
    }

    var totalNum = null;
    if (totalCell !== '' && totalCell !== null && totalCell !== undefined) {
      var t = Number(totalCell);
      if (!isNaN(t)) {
        totalNum = t;
      }
    }

    var estadoStr = (estado || '').toString().trim() || 'Pendiente';
    var idPedidoStr = (idPedido || '').toString().trim();

    // Cargar la base de clientes para idCliente/sector
    var sheetClientes = getSheet_('BASE DE CLIENTES');
    var lastRowClientes = sheetClientes.getLastRow();
    var mapaClientes = [];

    if (lastRowClientes > 1) {
      var dataClientes = sheetClientes.getRange(2, 1, lastRowClientes - 1, 9).getValues();
      dataClientes.forEach(function (r) {
        var nom = (r[0] || '').toString().trim();
        var tel = (r[1] || '').toString().trim();
        var idC = (r[7] || '').toString().trim();
        var sec = (r[8] || '').toString().trim();
        if (nom || tel) {
          mapaClientes.push({
            nombre: nom,
            telefono: tel,
            idCliente: idC,
            sector: sec
          });
        }
      });
    }

    function encontrarCliente(nombrePedido, telPedido) {
      var nombreP = (nombrePedido || '').toString().trim();
      var telP    = (telPedido    || '').toString().trim();

      if (telP) {
        for (var i = 0; i < mapaClientes.length; i++) {
          if (mapaClientes[i].telefono === telP) {
            return mapaClientes[i];
          }
        }
      }

      if (nombreP) {
        var nombrePLower = nombreP.toLowerCase();
        for (var j = 0; j < mapaClientes.length; j++) {
          if (mapaClientes[j].nombre.toLowerCase() === nombrePLower) {
            return mapaClientes[j];
          }
        }
      }

      return {
        nombre: nombreP,
        telefono: telP,
        idCliente: '',
        sector: ''
      };
    }

    var cli = encontrarCliente(clienteNom, telefono);
    var infoPrecioSector = obtenerPrecioYColumnaSector_(cli.sector || '');
    var precioUnitario = typeof infoPrecioSector.precio === 'number' ? infoPrecioSector.precio : null;

    var totalCalculado = null;
    if (typeof precioUnitario === 'number') {
      totalCalculado = (Number(cantidad) || 0) * precioUnitario;
      if (filaOrdenada !== null) {
        sheetPedidos.getRange(filaOrdenada, 6).setValue(totalCalculado); // F
      }
      totalNum = totalCalculado;
    }

    var pedidoFront = {
      idPedido: idPedidoStr,
      fechaEntrega: fechaStr,
      horaEntrega: (horaEntrega || '').toString(),
      cantidad: Number(cantidad || 0),
      total: totalNum,
      precioUnitario: typeof precioUnitario === 'number' ? precioUnitario : null,
      estado: estadoStr,
      notas: (notas || '').toString(),
      cliente: {
        idCliente: cli.idCliente || '',
        nombre: (clienteNom || '').toString(),
        telefono: (telefono || '').toString(),
        direccion: (direccion || '').toString(),
        sector: cli.sector || ''
      }
    };

    return {
      ok: true,
      error: null,
      data: { pedido: pedidoFront }
    };

  } catch (e) {
    return {
      ok: false,
      error: e && e.message ? e.message : 'Error desconocido al editar el pedido.',
      data: null
    };
  }
}

/**
 * Cambia la hora de entrega de un pedido identificado por su ID.
 *
 * Petición:
 * {
 *   idPedido: string,
 *   nuevaHora: string  // 'HH:MM'
 * }
 *
 * Respuesta:
 * {
 *   ok: true|false,
 *   error: string|null,
 *   data: { pedido: PedidoFront } | null
 * }
 */
function cambiarHoraPedido(data) {
  try {
    if (!data || !data.idPedido || !data.nuevaHora) {
      return {
        ok: false,
        error: 'Faltan idPedido o nuevaHora.',
        data: null
      };
    }

    var idObjetivo = data.idPedido.toString().trim();
    var nuevaHora  = data.nuevaHora.toString().trim();
    var tz = Session.getScriptTimeZone();

    var sheetPedidos = getSheet_('PEDIDOS DIARIOS');
    var lastRowPedidos = sheetPedidos.getLastRow();
    if (lastRowPedidos < 2) {
      return {
        ok: false,
        error: 'No hay pedidos registrados.',
        data: null
      };
    }

    var dataPedidos = sheetPedidos.getRange(2, 1, lastRowPedidos - 1, 11).getValues();
    var filaEncontrada = -1;

    for (var i = 0; i < dataPedidos.length; i++) {
      var row = dataPedidos[i];
      var idCell = (row[10] || '').toString().trim(); // K
      if (idCell === idObjetivo) {
        filaEncontrada = i + 2;
        break;
      }
    }

    if (filaEncontrada === -1) {
      return {
        ok: false,
        error: 'No se encontró el pedido con el ID especificado.',
        data: null
      };
    }

    // Actualiser la hora (col H = 8)
    sheetPedidos.getRange(filaEncontrada, 8).setValue(nuevaHora);

    sortPedidosDiarios_();

    // Reconstruire PedidoFront comme dans editarPedidoPorId
    var rowValues = null;
    var lastRowOrdenado = sheetPedidos.getLastRow();
    if (lastRowOrdenado > 1) {
      var dataPedidosOrdenados = sheetPedidos.getRange(2, 1, lastRowOrdenado - 1, 11).getValues();
      for (var idx = 0; idx < dataPedidosOrdenados.length; idx++) {
        var idFila = (dataPedidosOrdenados[idx][10] || '').toString().trim();
        if (idFila === idObjetivo) {
          rowValues = dataPedidosOrdenados[idx];
          break;
        }
      }
    }

    if (!rowValues) {
      return {
        ok: false,
        error: 'No se encontró el pedido con el ID especificado.',
        data: null
      };
    }

    var fechaCell   = rowValues[0];
    var clienteNom  = rowValues[1];
    var telefono    = rowValues[2];
    var direccion   = rowValues[3];
    var cantidad    = rowValues[4];
    var totalCell   = rowValues[5];
    var horaEntrega = rowValues[7];
    var notas       = rowValues[8];
    var estado      = rowValues[9];
    var idPedido    = rowValues[10];

    var fechaStr = '';
    if (fechaCell) {
      var d = new Date(fechaCell);
      fechaStr = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
    }

    var totalNum = null;
    if (totalCell !== '' && totalCell !== null && totalCell !== undefined) {
      var t = Number(totalCell);
      if (!isNaN(t)) {
        totalNum = t;
      }
    }

    var estadoStr = (estado || '').toString().trim() || 'Pendiente';
    var idPedidoStr = (idPedido || '').toString().trim();

    // Base de clientes
    var sheetClientes = getSheet_('BASE DE CLIENTES');
    var lastRowClientes = sheetClientes.getLastRow();
    var mapaClientes = [];

    if (lastRowClientes > 1) {
      var dataClientes = sheetClientes.getRange(2, 1, lastRowClientes - 1, 9).getValues();
      dataClientes.forEach(function (r) {
        var nom = (r[0] || '').toString().trim();
        var tel = (r[1] || '').toString().trim();
        var idC = (r[7] || '').toString().trim();
        var sec = (r[8] || '').toString().trim();
        if (nom || tel) {
          mapaClientes.push({
            nombre: nom,
            telefono: tel,
            idCliente: idC,
            sector: sec
          });
        }
      });
    }

    function encontrarCliente(nombrePedido, telPedido) {
      var nombreP = (nombrePedido || '').toString().trim();
      var telP    = (telPedido    || '').toString().trim();

      if (telP) {
        for (var i = 0; i < mapaClientes.length; i++) {
          if (mapaClientes[i].telefono === telP) {
            return mapaClientes[i];
          }
        }
      }

      if (nombreP) {
        var nombrePLower = nombreP.toLowerCase();
        for (var j = 0; j < mapaClientes.length; j++) {
          if (mapaClientes[j].nombre.toLowerCase() === nombrePLower) {
            return mapaClientes[j];
          }
        }
      }

      return {
        nombre: nombreP,
        telefono: telP,
        idCliente: '',
        sector: ''
      };
    }

    var cli = encontrarCliente(clienteNom, telefono);

    var pedidoFront = {
      idPedido: idPedidoStr,
      fechaEntrega: fechaStr,
      horaEntrega: (horaEntrega || '').toString(),
      cantidad: Number(cantidad || 0),
      total: totalNum,
      estado: estadoStr,
      notas: (notas || '').toString(),
      cliente: {
        idCliente: cli.idCliente || '',
        nombre: (clienteNom || '').toString(),
        telefono: (telefono || '').toString(),
        direccion: (direccion || '').toString(),
        sector: cli.sector || ''
      }
    };

    return {
      ok: true,
      error: null,
      data: { pedido: pedidoFront }
    };

  } catch (e) {
    return {
      ok: false,
      error: e && e.message ? e.message : 'Error desconocido al cambiar la hora del pedido.',
      data: null
    };
  }
}

/**
 * Marca un pedido como ENTREGADO.
 *
 * Petición:
 * {
 *   idPedido: string
 * }
 *
 * Respuesta:
 * {
 *   ok: true|false,
 *   error: string|null,
 *   data: { pedido: PedidoFront } | null
 * }
 */
function marcarPedidoEntregado(data) {
  try {
    if (!data || !data.idPedido) {
      return {
        ok: false,
        error: 'No se proporcionó idPedido.',
        data: null
      };
    }

    var idObjetivo = data.idPedido.toString().trim();
    var tz = Session.getScriptTimeZone();

    var sheetPedidos = getSheet_('PEDIDOS DIARIOS');
    var lastRowPedidos = sheetPedidos.getLastRow();
    if (lastRowPedidos < 2) {
      return {
        ok: false,
        error: 'No hay pedidos registrados.',
        data: null
      };
    }

    var dataPedidos = sheetPedidos.getRange(2, 1, lastRowPedidos - 1, 11).getValues();
    var filaEncontrada = -1;

    for (var i = 0; i < dataPedidos.length; i++) {
      var row = dataPedidos[i];
      var idCell = (row[10] || '').toString().trim();
      if (idCell === idObjetivo) {
        filaEncontrada = i + 2;
        break;
      }
    }

    if (filaEncontrada === -1) {
      return {
        ok: false,
        error: 'No se encontró el pedido con el ID especificado.',
        data: null
      };
    }

    // J Estado = 'Entregado'
    sheetPedidos.getRange(filaEncontrada, 10).setValue('Entregado');

    // Reconstruire le PedidoFront
    var rowValues = sheetPedidos.getRange(filaEncontrada, 1, 1, 11).getValues()[0];

    var fechaCell   = rowValues[0];
    var clienteNom  = rowValues[1];
    var telefono    = rowValues[2];
    var direccion   = rowValues[3];
    var cantidad    = rowValues[4];
    var totalCell   = rowValues[5];
    var horaEntrega = rowValues[7];
    var notas       = rowValues[8];
    var estado      = rowValues[9];
    var idPedido    = rowValues[10];

    var fechaStr = '';
    if (fechaCell) {
      var d = new Date(fechaCell);
      fechaStr = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
    }

    var totalNum = null;
    if (totalCell !== '' && totalCell !== null && totalCell !== undefined) {
      var t = Number(totalCell);
      if (!isNaN(t)) {
        totalNum = t;
      }
    }

    var estadoStr = (estado || '').toString().trim() || 'Entregado';
    var idPedidoStr = (idPedido || '').toString().trim();

    // Base de clientes
    var sheetClientes = getSheet_('BASE DE CLIENTES');
    var lastRowClientes = sheetClientes.getLastRow();
    var mapaClientes = [];

    if (lastRowClientes > 1) {
      var dataClientes = sheetClientes.getRange(2, 1, lastRowClientes - 1, 9).getValues();
      dataClientes.forEach(function (r) {
        var nom = (r[0] || '').toString().trim();
        var tel = (r[1] || '').toString().trim();
        var idC = (r[7] || '').toString().trim();
        var sec = (r[8] || '').toString().trim();
        if (nom || tel) {
          mapaClientes.push({
            nombre: nom,
            telefono: tel,
            idCliente: idC,
            sector: sec
          });
        }
      });
    }

    function encontrarCliente(nombrePedido, telPedido) {
      var nombreP = (nombrePedido || '').toString().trim();
      var telP    = (telPedido    || '').toString().trim();

      if (telP) {
        for (var i = 0; i < mapaClientes.length; i++) {
          if (mapaClientes[i].telefono === telP) {
            return mapaClientes[i];
          }
        }
      }

      if (nombreP) {
        var nombrePLower = nombreP.toLowerCase();
        for (var j = 0; j < mapaClientes.length; j++) {
          if (mapaClientes[j].nombre.toLowerCase() === nombrePLower) {
            return mapaClientes[j];
          }
        }
      }

      return {
        nombre: nombreP,
        telefono: telP,
        idCliente: '',
        sector: ''
      };
    }

    var cli = encontrarCliente(clienteNom, telefono);

    var pedidoFront = {
      idPedido: idPedidoStr,
      fechaEntrega: fechaStr,
      horaEntrega: (horaEntrega || '').toString(),
      cantidad: Number(cantidad || 0),
      total: totalNum,
      estado: estadoStr,
      notas: (notas || '').toString(),
      cliente: {
        idCliente: cli.idCliente || '',
        nombre: (clienteNom || '').toString(),
        telefono: (telefono || '').toString(),
        direccion: (direccion || '').toString(),
        sector: cli.sector || ''
      }
    };

    return {
      ok: true,
      error: null,
      data: { pedido: pedidoFront }
    };

  } catch (e) {
    return {
      ok: false,
      error: e && e.message ? e.message : 'Error desconocido al marcar el pedido como entregado.',
      data: null
    };
  }
}

/**
 * Actualiza el estado (columna J) de un pedido identificado por su ID.
 *
 * @param {{idPedido:string, estado:string}} data
 */
function actualizarEstadoPedido(data) {
  try {
    if (!data || !data.idPedido || !data.estado) {
      return {
        ok: false,
        error: 'Faltan idPedido o estado.',
        data: null
      };
    }

    var idObjetivo = data.idPedido.toString().trim();
    var nuevoEstado = data.estado.toString().trim();
    var tz = Session.getScriptTimeZone();

    var sheetPedidos = getSheet_('PEDIDOS DIARIOS');
    var lastRowPedidos = sheetPedidos.getLastRow();
    if (lastRowPedidos < 2) {
      return {
        ok: false,
        error: 'No hay pedidos registrados.',
        data: null
      };
    }

    var dataPedidos = sheetPedidos.getRange(2, 1, lastRowPedidos - 1, 11).getValues();
    var filaEncontrada = -1;

    for (var i = 0; i < dataPedidos.length; i++) {
      var row = dataPedidos[i];
      var idCell = (row[10] || '').toString().trim();
      if (idCell === idObjetivo) {
        filaEncontrada = i + 2;
        break;
      }
    }

    if (filaEncontrada === -1) {
      return {
        ok: false,
        error: 'No se encontró el pedido con el ID especificado.',
        data: null
      };
    }

    sheetPedidos.getRange(filaEncontrada, 10).setValue(nuevoEstado);

    var rowValues = sheetPedidos.getRange(filaEncontrada, 1, 1, 11).getValues()[0];

    var fechaCell   = rowValues[0];
    var clienteNom  = rowValues[1];
    var telefono    = rowValues[2];
    var direccion   = rowValues[3];
    var cantidad    = rowValues[4];
    var totalCell   = rowValues[5];
    var horaEntrega = rowValues[7];
    var notas       = rowValues[8];
    var estado      = rowValues[9];
    var idPedido    = rowValues[10];

    var fechaStr = '';
    if (fechaCell) {
      var d = new Date(fechaCell);
      fechaStr = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
    }

    var totalNum = null;
    if (totalCell !== '' && totalCell !== null && totalCell !== undefined) {
      var t = Number(totalCell);
      if (!isNaN(t)) {
        totalNum = t;
      }
    }

    var estadoStr = (estado || '').toString().trim() || nuevoEstado;
    var idPedidoStr = (idPedido || '').toString().trim();

    var sheetClientes = getSheet_('BASE DE CLIENTES');
    var lastRowClientes = sheetClientes.getLastRow();
    var mapaClientes = [];

    if (lastRowClientes > 1) {
      var dataClientes = sheetClientes.getRange(2, 1, lastRowClientes - 1, 9).getValues();
      dataClientes.forEach(function (r) {
        var nom = (r[0] || '').toString().trim();
        var tel = (r[1] || '').toString().trim();
        var idC = (r[7] || '').toString().trim();
        var sec = (r[8] || '').toString().trim();
        if (nom || tel) {
          mapaClientes.push({
            nombre: nom,
            telefono: tel,
            idCliente: idC,
            sector: sec
          });
        }
      });
    }

    function encontrarCliente(nombrePedido, telPedido) {
      var nombreP = (nombrePedido || '').toString().trim();
      var telP    = (telPedido    || '').toString().trim();

      if (telP) {
        for (var i = 0; i < mapaClientes.length; i++) {
          if (mapaClientes[i].telefono === telP) {
            return mapaClientes[i];
          }
        }
      }

      if (nombreP) {
        var nombrePLower = nombreP.toLowerCase();
        for (var j = 0; j < mapaClientes.length; j++) {
          if (mapaClientes[j].nombre.toLowerCase() === nombrePLower) {
            return mapaClientes[j];
          }
        }
      }

      return {
        nombre: nombreP,
        telefono: telP,
        idCliente: '',
        sector: ''
      };
    }

    var cli = encontrarCliente(clienteNom, telefono);

    var pedidoFront = {
      idPedido: idPedidoStr,
      fechaEntrega: fechaStr,
      horaEntrega: (horaEntrega || '').toString(),
      cantidad: Number(cantidad || 0),
      total: totalNum,
      estado: estadoStr,
      notas: (notas || '').toString(),
      cliente: {
        idCliente: cli.idCliente || '',
        nombre: (clienteNom || '').toString(),
        telefono: (telefono || '').toString(),
        direccion: (direccion || '').toString(),
        sector: cli.sector || ''
      }
    };

    return {
      ok: true,
      error: null,
      data: { pedido: pedidoFront }
    };

  } catch (e) {
    return {
      ok: false,
      error: e && e.message ? e.message : 'Error desconocido al actualizar el estado del pedido.',
      data: null
    };
  }
}

/**
 * Marca un pedido como ELIMINADO (no borra la fila).
 *
 * Petición:
 * {
 *   idPedido: string
 * }
 *
 * Respuesta:
 * {
 *   ok: true|false,
 *   error: string|null,
 *   data: { pedido: PedidoFront } | null
 * }
 */
function eliminarPedido(data) {
  try {
    if (!data || !data.idPedido) {
      return {
        ok: false,
        error: 'No se proporcionó idPedido.',
        data: null
      };
    }

    var idObjetivo = data.idPedido.toString().trim();
    var tz = Session.getScriptTimeZone();

    var sheetPedidos = getSheet_('PEDIDOS DIARIOS');
    var lastRowPedidos = sheetPedidos.getLastRow();
    if (lastRowPedidos < 2) {
      return {
        ok: false,
        error: 'No hay pedidos registrados.',
        data: null
      };
    }

    var dataPedidos = sheetPedidos.getRange(2, 1, lastRowPedidos - 1, 11).getValues();
    var filaEncontrada = -1;

    for (var i = 0; i < dataPedidos.length; i++) {
      var row = dataPedidos[i];
      var idCell = (row[10] || '').toString().trim();
      if (idCell === idObjetivo) {
        filaEncontrada = i + 2;
        break;
      }
    }

    if (filaEncontrada === -1) {
      return {
        ok: false,
        error: 'No se encontró el pedido con el ID especificado.',
        data: null
      };
    }

    // J Estado = 'Eliminado'
    sheetPedidos.getRange(filaEncontrada, 10).setValue('Eliminado');

    var rowValues = sheetPedidos.getRange(filaEncontrada, 1, 1, 11).getValues()[0];

    var fechaCell   = rowValues[0];
    var clienteNom  = rowValues[1];
    var telefono    = rowValues[2];
    var direccion   = rowValues[3];
    var cantidad    = rowValues[4];
    var totalCell   = rowValues[5];
    var horaEntrega = rowValues[7];
    var notas       = rowValues[8];
    var estado      = rowValues[9];
    var idPedido    = rowValues[10];

    var fechaStr = '';
    if (fechaCell) {
      var d = new Date(fechaCell);
      fechaStr = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
    }

    var totalNum = null;
    if (totalCell !== '' && totalCell !== null && totalCell !== undefined) {
      var t = Number(totalCell);
      if (!isNaN(t)) {
        totalNum = t;
      }
    }

    var estadoStr = (estado || '').toString().trim() || 'Eliminado';
    var idPedidoStr = (idPedido || '').toString().trim();

    // Base de clientes
    var sheetClientes = getSheet_('BASE DE CLIENTES');
    var lastRowClientes = sheetClientes.getLastRow();
    var mapaClientes = [];

    if (lastRowClientes > 1) {
      var dataClientes = sheetClientes.getRange(2, 1, lastRowClientes - 1, 9).getValues();
      dataClientes.forEach(function (r) {
        var nom = (r[0] || '').toString().trim();
        var tel = (r[1] || '').toString().trim();
        var idC = (r[7] || '').toString().trim();
        var sec = (r[8] || '').toString().trim();
        if (nom || tel) {
          mapaClientes.push({
            nombre: nom,
            telefono: tel,
            idCliente: idC,
            sector: sec
          });
        }
      });
    }

    function encontrarCliente(nombrePedido, telPedido) {
      var nombreP = (nombrePedido || '').toString().trim();
      var telP    = (telPedido    || '').toString().trim();

      if (telP) {
        for (var i = 0; i < mapaClientes.length; i++) {
          if (mapaClientes[i].telefono === telP) {
            return mapaClientes[i];
          }
        }
      }

      if (nombreP) {
        var nombrePLower = nombreP.toLowerCase();
        for (var j = 0; j < mapaClientes.length; j++) {
          if (mapaClientes[j].nombre.toLowerCase() === nombrePLower) {
            return mapaClientes[j];
          }
        }
      }

      return {
        nombre: nombreP,
        telefono: telP,
        idCliente: '',
        sector: ''
      };
    }

    var cli = encontrarCliente(clienteNom, telefono);

    var pedidoFront = {
      idPedido: idPedidoStr,
      fechaEntrega: fechaStr,
      horaEntrega: (horaEntrega || '').toString(),
      cantidad: Number(cantidad || 0),
      total: totalNum,
      estado: estadoStr,
      notas: (notas || '').toString(),
      cliente: {
        idCliente: cli.idCliente || '',
        nombre: (clienteNom || '').toString(),
        telefono: (telefono || '').toString(),
        direccion: (direccion || '').toString(),
        sector: cli.sector || ''
      }
    };

    return {
      ok: true,
      error: null,
      data: { pedido: pedidoFront }
    };

  } catch (e) {
    return {
      ok: false,
      error: e && e.message ? e.message : 'Error desconocido al eliminar el pedido.',
      data: null
    };
  }
}

/**
 * Reactiva un pedido (pasa su estado a PENDIENTE).
 *
 * Petición:
 * {
 *   idPedido: string
 * }
 *
 * Respuesta:
 * {
 *   ok: true|false,
 *   error: string|null,
 *   data: { pedido: PedidoFront } | null
 * }
 */
function reactivarPedido(data) {
  try {
    if (!data || !data.idPedido) {
      return {
        ok: false,
        error: 'No se proporcionó idPedido.',
        data: null
      };
    }

    var idObjetivo = data.idPedido.toString().trim();
    var tz = Session.getScriptTimeZone();

    var sheetPedidos = getSheet_('PEDIDOS DIARIOS');
    var lastRowPedidos = sheetPedidos.getLastRow();
    if (lastRowPedidos < 2) {
      return {
        ok: false,
        error: 'No hay pedidos registrados.',
        data: null
      };
    }

    var dataPedidos = sheetPedidos.getRange(2, 1, lastRowPedidos - 1, 11).getValues();
    var filaEncontrada = -1;

    for (var i = 0; i < dataPedidos.length; i++) {
      var row = dataPedidos[i];
      var idCell = (row[10] || '').toString().trim();
      if (idCell === idObjetivo) {
        filaEncontrada = i + 2;
        break;
      }
    }

    if (filaEncontrada === -1) {
      return {
        ok: false,
        error: 'No se encontró el pedido con el ID especificado.',
        data: null
      };
    }

    // J Estado = 'Pendiente'
    sheetPedidos.getRange(filaEncontrada, 10).setValue('Pendiente');

    var rowValues = sheetPedidos.getRange(filaEncontrada, 1, 1, 11).getValues()[0];

    var fechaCell   = rowValues[0];
    var clienteNom  = rowValues[1];
    var telefono    = rowValues[2];
    var direccion   = rowValues[3];
    var cantidad    = rowValues[4];
    var totalCell   = rowValues[5];
    var horaEntrega = rowValues[7];
    var notas       = rowValues[8];
    var estado      = rowValues[9];
    var idPedido    = rowValues[10];

    var fechaStr = '';
    if (fechaCell) {
      var d = new Date(fechaCell);
      fechaStr = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
    }

    var totalNum = null;
    if (totalCell !== '' && totalCell !== null && totalCell !== undefined) {
      var t = Number(totalCell);
      if (!isNaN(t)) {
        totalNum = t;
      }
    }

    var estadoStr = (estado || '').toString().trim() || 'Pendiente';
    var idPedidoStr = (idPedido || '').toString().trim();

    // Base de clientes
    var sheetClientes = getSheet_('BASE DE CLIENTES');
    var lastRowClientes = sheetClientes.getLastRow();
    var mapaClientes = [];

    if (lastRowClientes > 1) {
      var dataClientes = sheetClientes.getRange(2, 1, lastRowClientes - 1, 9).getValues();
      dataClientes.forEach(function (r) {
        var nom = (r[0] || '').toString().trim();
        var tel = (r[1] || '').toString().trim();
        var idC = (r[7] || '').toString().trim();
        var sec = (r[8] || '').toString().trim();
        if (nom || tel) {
          mapaClientes.push({
            nombre: nom,
            telefono: tel,
            idCliente: idC,
            sector: sec
          });
        }
      });
    }

    function encontrarCliente(nombrePedido, telPedido) {
      var nombreP = (nombrePedido || '').toString().trim();
      var telP    = (telPedido    || '').toString().trim();

      if (telP) {
        for (var i = 0; i < mapaClientes.length; i++) {
          if (mapaClientes[i].telefono === telP) {
            return mapaClientes[i];
          }
        }
      }

      if (nombreP) {
        var nombrePLower = nombreP.toLowerCase();
        for (var j = 0; j < mapaClientes.length; j++) {
          if (mapaClientes[j].nombre.toLowerCase() === nombrePLower) {
            return mapaClientes[j];
          }
        }
      }

      return {
        nombre: nombreP,
        telefono: telP,
        idCliente: '',
        sector: ''
      };
    }

    var cli = encontrarCliente(clienteNom, telefono);

    var pedidoFront = {
      idPedido: idPedidoStr,
      fechaEntrega: fechaStr,
      horaEntrega: (horaEntrega || '').toString(),
      cantidad: Number(cantidad || 0),
      total: totalNum,
      estado: estadoStr,
      notas: (notas || '').toString(),
      cliente: {
        idCliente: cli.idCliente || '',
        nombre: (clienteNom || '').toString(),
        telefono: (telefono || '').toString(),
        direccion: (direccion || '').toString(),
        sector: cli.sector || ''
      }
    };

    return {
      ok: true,
      error: null,
      data: { pedido: pedidoFront }
    };

  } catch (e) {
    return {
      ok: false,
      error: e && e.message ? e.message : 'Error desconocido al reactivar el pedido.',
      data: null
    };
  }
}





// BLOC VENTAS




function initVentaDirecta() {
  try {
    var idVenta = generarIdVentaDirecta_();
    return {
      ok: true,
      error: null,
      data: { idVenta: idVenta }
    };
  } catch (e) {
    return {
      ok: false,
      error: e && e.message ? e.message : 'Error al inicializar Venta directa.',
      data: null
    };
  }
}

function abrirVentaRegistrada(ctxVenta) {
  var ctx = ctxVenta || {};

  return {
    ok: true,
    error: null,
    data: ctx
  };
}


/**
 * Registra una venta directa en la hoja VENTAS DIRECTAS.
 *
 * Contrato JSON (PAQUETE 5):
 *
 * Petición:
 * {
 *   fecha: string | null,   // 'YYYY-MM-DD' o null => hoy
 *   cantidad: number,
 *   precio: number,
 *   notas: string
 * }
 *
 * Respuesta data:
 * {
 *   venta: {
 *     idVenta: string | null,
 *     fecha: string,
 *     cantidad: number,
 *     precioUnitario: number,
 *     total: number,
 *     notas: string
 *   }
 * }
 */
function registrarVentaDirecta(data) {
  try {
    // 1) Validación básica
    if (!data) {
      return {
        ok: false,
        error: 'No se proporcionaron datos para registrar la venta directa.',
        data: null
      };
    }

    var cantidad = Number(data.cantidad);
    var precio   = Number(data.precio);
    var notas    = (data.notas || '').toString();

    if (!cantidad || isNaN(cantidad) || cantidad <= 0) {
      return {
        ok: false,
        error: 'Cantidad inválida para la venta directa.',
        data: null
      };
    }

    if (!precio || isNaN(precio) || precio <= 0) {
      return {
        ok: false,
        error: 'Precio inválido para la venta directa.',
        data: null
      };
    }

    // 2) Determinar la fecha de la venta
    var tz = Session.getScriptTimeZone();
    var fechaObj;

    if (data.fecha) {
      // Formato esperado 'YYYY-MM-DD'
      fechaObj = new Date(data.fecha);
      if (isNaN(fechaObj.getTime())) {
        return {
          ok: false,
          error: 'Fecha inválida para la venta directa.',
          data: null
        };
      }
    } else {
      // Si no se proporciona fecha, se usa hoy
      fechaObj = new Date();
    }

    var fechaStr = Utilities.formatDate(fechaObj, tz, 'yyyy-MM-dd');

    // 3) Calcular total
    var total = cantidad * precio;

    // 4) Escribir en la hoja VENTAS DIRECTAS
    // Estructura:
    // A ID Venta
    // B Fecha
    // C Cantidad
    // D Precio unitario
    // E Total
    // F Notas
    var sheet = getSheet_('VENTAS DIRECTAS');
    var newRow = getFirstEmptyRowInColumn_(sheet, 1);

    // 4.a) Generar nuevo ID de venta: V-00001, V-00002, ...
    var nuevoId = generarIdVentaDirecta_();

    // 4.b) Escritura de la fila
    sheet.getRange(newRow, 1).setValue(nuevoId);   // A ID Venta
    sheet.getRange(newRow, 2).setValue(fechaObj);  // B Fecha (Date)
    sheet.getRange(newRow, 3).setValue(cantidad);  // C Cantidad
    sheet.getRange(newRow, 4).setValue(precio);    // D Precio unitario
    sheet.getRange(newRow, 5).setValue(total);     // E Total
    sheet.getRange(newRow, 6).setValue(notas);     // F Notas

    // 5) Construir el objeto VentaDirectaFront para el front
    var ventaFront = {
      idVenta: nuevoId,
      fecha: fechaStr,          // 'YYYY-MM-DD'
      cantidad: cantidad,
      precioUnitario: precio,
      total: total,
      notas: notas
    };

    return {
      ok: true,
      error: null,
      data: {
        venta: ventaFront
      }
    };

  } catch (e) {
    return {
      ok: false,
      error: e && e.message ? e.message : 'Error desconocido al registrar la venta directa.',
      data: null
    };
  }
}

function registrarVentaDirectaDesdeUI(payload) {
  var resp = registrarVentaDirecta(payload);
  if (!resp || !resp.ok || !resp.data || !resp.data.venta) {
    return resp || { ok: false, error: 'No se pudo registrar la venta.', data: null };
  }

  abrirVentaRegistrada(resp.data.venta);
  return resp;
}





// BLOC GASTOS





/**
 * Registra un gasto desde el formulario REGISTRAR GASTO.
 *
 * data:
 * {
 *   fecha: string | null,        // 'YYYY-MM-DD' o null -> hoy
 *   categoria: string,
 *   producto: string,
 *   proveedor: string,
 *   unidad: string,
 *   precioUnidad: number,
 *   cantidad: number,
 *   observaciones: string,
 *   idProducto: string | null    // 'PROD-00001' (OFERTA) o null
 * }
 *
 * Respuesta:
 * {
 *   ok: true|false,
 *   error: string|null,
 *   data: {
 *     gasto: {
 *       idGasto: string,
 *       fecha: string,
 *       categoria: string,
 *       producto: string,
 *       proveedor: string,
 *       unidad: string,
 *       precioUnidad: number,
 *       cantidad: number,
 *       monto: number | null,
 *       observaciones: string,
 *       idProducto: string
 *     }
 *   } | null
 * }
 */
function registrarGastoDesdeFormulario(data) {
  try {
    if (!data) {
      return {
        ok: false,
        error: 'No se recibieron datos para registrar el gasto.',
        data: null
      };
    }

    var tz = Session.getScriptTimeZone();

    // 1) Fecha
    var fechaObj;
    if (data.fecha) {
      fechaObj = new Date(data.fecha);
      if (isNaN(fechaObj.getTime())) {
        return {
          ok: false,
          error: 'Fecha inválida.',
          data: null
        };
      }
    } else {
      fechaObj = new Date();
    }
    var fechaStr = Utilities.formatDate(fechaObj, tz, 'yyyy-MM-dd');

    // 2) Cantidad
    var cantidad = Number(data.cantidad || 0);
    if (isNaN(cantidad) || cantidad <= 0) {
      return {
        ok: false,
        error: 'Cantidad inválida para el gasto.',
        data: null
      };
    }

    var idProducto = (data.idProducto || '').toString().trim();

    var categoria;
    var producto;
    var proveedor;
    var unidad;
    var precioUnidad;

    // 3) Si viene idProducto (OFERTA), tomamos todo desde la hoja PRODUCTO
    if (idProducto) {
      var sheetProd = getSheet_('PRODUCTO');
      var lastRowProd = sheetProd.getLastRow();
      if (lastRowProd < 2) {
        return {
          ok: false,
          error: 'No hay ofertas registradas en la hoja PRODUCTO.',
          data: null
        };
      }

      var dataProd = sheetProd.getRange(2, 1, lastRowProd - 1, 6).getValues();
      var filaOferta = -1;

      for (var i = 0; i < dataProd.length; i++) {
        var idFila = (dataProd[i][0] || '').toString().trim(); // A = PRODUCTO ID
        if (idFila === idProducto) {
          filaOferta = i + 2;
          break;
        }
      }

      if (filaOferta === -1) {
        return {
          ok: false,
          error: 'No se encontró la oferta (PRODUCTO ID) especificada: ' + idProducto,
          data: null
        };
      }

      var rowOferta = sheetProd.getRange(filaOferta, 1, 1, 6).getValues()[0];
      // A PRODUCTO ID
      categoria    = (rowOferta[1] || '').toString(); // B CATEGORIA
      producto     = (rowOferta[2] || '').toString(); // C PRODUCTO
      proveedor    = (rowOferta[3] || '').toString(); // D PROVEEDOR
      unidad       = (rowOferta[4] || '').toString(); // E UNIDAD
      precioUnidad = Number(rowOferta[5] || 0);        // F PRECIO UNIDAD

      if (isNaN(precioUnidad) || precioUnidad <= 0) {
        return {
          ok: false,
          error: 'La oferta seleccionada no tiene un precio de unidad válido.',
          data: null
        };
      }

    } else {
      // 4) Sin idProducto -> usamos los datos enviados (modo antiguo)
      categoria    = (data.categoria || '').toString().trim();
      producto     = (data.producto || '').toString().trim();
      proveedor    = (data.proveedor || '').toString().trim();
      unidad       = (data.unidad || '').toString().trim();
      precioUnidad = Number(data.precioUnidad || 0);

      if (!categoria || !producto) {
        return {
          ok: false,
          error: 'Faltan categoría o producto para registrar el gasto.',
          data: null
        };
      }
      if (isNaN(precioUnidad) || precioUnidad <= 0) {
        return {
          ok: false,
          error: 'Precio por unidad inválido para el gasto.',
          data: null
        };
      }
    }

    var observaciones = (data.observaciones || '').toString();

    // 5) Monto = cantidad * precioUnidad
    var monto = cantidad * precioUnidad;

    // 6) Generar ID_GASTO
    var idGasto = generarIdGasto_();

    // 7) Escribir en la hoja GASTOS
    var sheetGastos = getSheet_('GASTOS');
    var newRow = getFirstEmptyRowInColumn_(sheetGastos, 1);

    // Estructura:
    // A Fecha
    // B Mes (fórmula en la hoja)
    // C Categoría de gasto
    // D Producto
    // E Cantidad
    // F Unidad
    // G Precio por unidad
    // H Monto ($)
    // I Proveedor
    // J Observaciones
    // K ID_GASTO
    // L ID_PRODUCTO

    sheetGastos.getRange(newRow, 1).setValue(fechaObj);     // A Fecha
    sheetGastos.getRange(newRow, 3).setValue(categoria);    // C
    sheetGastos.getRange(newRow, 4).setValue(producto);     // D
    sheetGastos.getRange(newRow, 5).setValue(cantidad);     // E
    sheetGastos.getRange(newRow, 6).setValue(unidad);       // F
    sheetGastos.getRange(newRow, 7).setValue(precioUnidad); // G
    sheetGastos.getRange(newRow, 8).setValue(monto);        // H
    sheetGastos.getRange(newRow, 9).setValue(proveedor);    // I
    sheetGastos.getRange(newRow, 10).setValue(observaciones); // J
    sheetGastos.getRange(newRow, 11).setValue(idGasto);     // K ID_GASTO
    sheetGastos.getRange(newRow, 12).setValue(idProducto);  // L ID_PRODUCTO (puede ser '')

    // 8) Construir el objeto GastoFront para devolver
    var gastoFront = {
      idGasto: idGasto,
      fecha: fechaStr,
      categoria: categoria,
      producto: producto,
      proveedor: proveedor,
      unidad: unidad,
      precioUnidad: precioUnidad,
      cantidad: cantidad,
      monto: monto,
      observaciones: observaciones,
      idProducto: idProducto
    };

    return {
      ok: true,
      error: null,
      data: {
        gasto: gastoFront
      }
    };

  } catch (e) {
    return {
      ok: false,
      error: e && e.message ? e.message : 'Error desconocido al registrar el gasto.',
      data: null
    };
  }
}

/**
 * Mise à jour de la BASE DE CLIENTES
 */
function actualizarUltimoPedidoCliente(data) {
  // data attendu:
  // {
  //   nombre: 'Luis García',
  //   telefono: '+593...',
  //   fecha: Date ou '2025-11-03',
  //   direccion: 'Cerca del puente',
  //   notas: 'Paga en efectivo'
  // }

  if (!data || (!data.nombre && !data.telefono)) return;

  var sheet = getSheet_('BASE DE CLIENTES');
  var lastRow = sheet.getLastRow();
  var foundRow = 0;

  // lire les clients existants
  if (lastRow > 1) {
    var clientes = sheet.getRange(2, 1, lastRow - 1, 2).getValues(); // A: Cliente, B: Teléfono
    for (var i = 0; i < clientes.length; i++) {
      var nombre = (clientes[i][0] || '').toString().trim();
      var tel = (clientes[i][1] || '').toString().trim();
      var matchNombre = data.nombre && nombre.toLowerCase() === data.nombre.toLowerCase();
      var matchTel = data.telefono && tel === data.telefono;
      if (matchNombre || matchTel) {
        foundRow = i + 2;
        break;
      }
    }
  }

  // si trouvé → mise à jour
  if (foundRow > 0) {
    if (data.direccion) sheet.getRange(foundRow, 3).setValue(data.direccion);
    sheet.getRange(foundRow, 4).setValue(data.fecha || new Date());
    sheet.getRange(foundRow, 6).setValue('Activo');
    if (data.notas) {
      var currentNotes = sheet.getRange(foundRow, 7).getValue();
      sheet.getRange(foundRow, 7).setValue(currentNotes ? currentNotes + ' | ' + data.notas : data.notas);
    }
  } else {
    // sinon → nouvelle ligne
    var newRow = getFirstEmptyRowInColumn_(sheet, 1);
    sheet.getRange(newRow, 1).setValue(data.nombre || '');
    sheet.getRange(newRow, 2).setValue(data.telefono || '');
    sheet.getRange(newRow, 3).setValue(data.direccion || '');
    sheet.getRange(newRow, 4).setValue(data.fecha || new Date());
    sheet.getRange(newRow, 5).setValue('');        // Frecuencia
    sheet.getRange(newRow, 6).setValue('Nuevo');   // Estado
    sheet.getRange(newRow, 7).setValue(data.notas || '');
  }

  // **TRI AUTOMATIQUE**
  var totalRows = sheet.getLastRow();
  if (totalRows > 2) {
    sheet.getRange(2, 1, totalRows - 1, 7).sort({ column: 1, ascending: true });
  }
}

/**
 * Registra en la hoja INGRESOS el resumen del día
 * a partir de los pedidos del día.
 *
 * @param {Array} pedidosDelDia matriz de PEDIDOS DIARIOS (A:J) ya filtrada
 */
function registrarIngresoDesdePedidos_(pedidosDelDia) {
  if (!pedidosDelDia || pedidosDelDia.length === 0) return;

  // on prend la date du 1er pedido
  var fechaMenu = pedidosDelDia[0][0]; // columna A
  if (!fechaMenu) {
    // si pas de date dans le premier, on arrête
    return;
  }

  // total de comidas vendidas = somme des colonnes E
  var totalComidas = 0;
  pedidosDelDia.forEach(function (row) {
    var cant = Number(row[4] || 0); // col E
    totalComidas += cant;
  });

  // prix unitaire : pour l’instant, fixe = 1
  var precioUnitario = 1;

  // on enregistre dans INGRESOS
  guardarIngreso({
    fecha: fechaMenu,
    totalComidas: totalComidas,
    precio: precioUnitario,
    observaciones: 'Ingreso generado automáticamente desde PEDIDOS DIARIOS'
  });
}

/**
 * Crea una nueva categoría de gasto en LISTAS!D (si no existe ya).
 *
 * Petición:
 * {
 *   nombreCategoria: string
 * }
 *
 * Respuesta:
 * {
 *   ok: true|false,
 *   error: string|null,
 *   data: {
 *     categoriasGasto: string[]
 *   } | null
 * }
 */
function crearCategoriaGasto(data) {
  try {
    if (!data || !data.nombreCategoria) {
      return {
        ok: false,
        error: 'No se proporcionó el nombre de la categoría de gasto.',
        data: null
      };
    }

    var nombre = data.nombreCategoria.toString().trim();
    if (!nombre) {
      return {
        ok: false,
        error: 'El nombre de la categoría no puede estar vacío.',
        data: null
      };
    }

    // Añadir a LISTAS!D usando el helper genérico
    addToList_('LISTAS', 'D', nombre);

    // Volver a leer las listas para devolver la lista actualizada
    var listas = getLists();

    return {
      ok: true,
      error: null,
      data: {
        categoriasGasto: listas.categoriasGasto || []
      }
    };

  } catch (e) {
    return {
      ok: false,
      error: e && e.message ? e.message : 'Error desconocido al crear la categoría de gasto.',
      data: null
    };
  }
}

/**
 * Crea un nuevo producto (nombre genérico) en LISTAS!H (si no existe ya).
 * Esto sirve para la lista de productos disponibles en los desplegables.
 *
 * Petición:
 * {
 *   categoria: string,   // opcional aquí, pero previsto por el contrato
 *   producto: string
 * }
 *
 * Respuesta:
 * {
 *   ok: true|false,
 *   error: string|null,
 *   data: {
 *     productos: string[]
 *   } | null
 * }
 */
function crearProducto(data) {
  try {
    if (!data || !data.producto) {
      return {
        ok: false,
        error: 'No se proporcionó el nombre del producto.',
        data: null
      };
    }

    var nombreProducto = data.producto.toString().trim();
    if (!nombreProducto) {
      return {
        ok: false,
        error: 'El nombre del producto no puede estar vacío.',
        data: null
      };
    }

    // Añadir a LISTAS!H (lista genérica de productos)
    addToList_('LISTAS', 'H', nombreProducto);

    // Volver a leer las listas para devolver la lista actualizada
    var listas = getLists();

    return {
      ok: true,
      error: null,
      data: {
        productos: listas.productos || []
      }
    };

  } catch (e) {
    return {
      ok: false,
      error: e && e.message ? e.message : 'Error desconocido al crear el producto.',
      data: null
    };
  }
}

/**
 * Crea un nuevo proveedor en LISTAS!F (si no existe ya).
 *
 * Petición:
 * {
 *   proveedor: string
 * }
 *
 * Respuesta:
 * {
 *   ok: true|false,
 *   error: string|null,
 *   data: {
 *     proveedores: string[]
 *   } | null
 * }
 */
function crearProveedor(data) {
  try {
    if (!data || !data.proveedor) {
      return {
        ok: false,
        error: 'No se proporcionó el nombre del proveedor.',
        data: null
      };
    }

    var nombreProveedor = data.proveedor.toString().trim();
    if (!nombreProveedor) {
      return {
        ok: false,
        error: 'El nombre del proveedor no puede estar vacío.',
        data: null
      };
    }

    // Añadir a LISTAS!F
    addToList_('LISTAS', 'F', nombreProveedor);

    // Volver a leer las listas para devolver la lista actualizada
    var listas = getLists();

    return {
      ok: true,
      error: null,
      data: {
        proveedores: listas.proveedores || []
      }
    };

  } catch (e) {
    return {
      ok: false,
      error: e && e.message ? e.message : 'Error desconocido al crear el proveedor.',
      data: null
    };
  }
}

/**
 * Crea un producto detallado en la hoja PRODUCTO.
 *
 * Esta función NO reemplaza a crearProducto(data):
 * - crearProducto(data) mantiene la lista genérica de nombres en LISTAS!H
 * - crearProductoDetallado(data) crea una línea concreta en PRODUCTO
 *   con ID, categoría, proveedor, unidad y precio unidad.
 *
 * Petición:
 * {
 *   categoria: string,
 *   producto: string,
 *   proveedor: string,
 *   unidad: string,
 *   precioUnidad: number
 * }
 *
 * Respuesta:
 * {
 *   ok: true|false,
 *   error: string|null,
 *   data: {
 *     producto: {
 *       idProducto: string,
 *       categoria: string,
 *       producto: string,
 *       proveedor: string,
 *       unidad: string,
 *       precioUnidad: number
 *     }
 *   } | null
 * }
 */
function crearProductoDetallado(data) {
  try {
    if (!data) {
      return {
        ok: false,
        error: 'No se proporcionaron datos para crear el producto detallado.',
        data: null
      };
    }

    var categoria = (data.categoria || '').toString().trim();
    var producto  = (data.producto  || '').toString().trim();
    var proveedor = (data.proveedor || '').toString().trim();
    var unidad    = (data.unidad    || '').toString().trim();
    var precioRaw = data.precioUnidad;

    if (!categoria || !producto || !proveedor || !unidad) {
      return {
        ok: false,
        error: 'Faltan categoría, producto, proveedor o unidad.',
        data: null
      };
    }

    var precioNum = Number(precioRaw);
    if (isNaN(precioNum) || precioNum <= 0) {
      return {
        ok: false,
        error: 'El precio por unidad debe ser un número mayor que cero.',
        data: null
      };
    }

    var sheet = getSheet_('PRODUCTO');

    // Generar el ID de producto detallado
    var idProducto = generarIdProducto_();

    // Buscar primera fila realmente vacía (columna A)
    var newRow = getFirstEmptyRowInColumn_(sheet, 1);

    // Estructura de PRODUCTO:
    // A PRODUCTO ID
    // B CATEGORIA
    // C PRODUCTO
    // D PROVEEDOR
    // E UNIDAD
    // F PRECIO UNIDAD
    sheet.getRange(newRow, 1).setValue(idProducto);
    sheet.getRange(newRow, 2).setValue(categoria);
    sheet.getRange(newRow, 3).setValue(producto);
    sheet.getRange(newRow, 4).setValue(proveedor);
    sheet.getRange(newRow, 5).setValue(unidad);
    sheet.getRange(newRow, 6).setValue(precioNum);

    // Mantener LISTAS en coherencia (sin inventar nada nuevo):
    // D = categoriasGasto
    // E = unidades
    // F = proveedores
    // H = productos
    addToList_('LISTAS', 'D', categoria);
    addToList_('LISTAS', 'E', unidad);
    addToList_('LISTAS', 'F', proveedor);
    addToList_('LISTAS', 'H', producto);

    var productoFront = {
      idProducto: idProducto,
      categoria: categoria,
      producto: producto,
      proveedor: proveedor,
      unidad: unidad,
      precioUnidad: precioNum
    };

    return {
      ok: true,
      error: null,
      data: {
        producto: productoFront
      }
    };

  } catch (e) {
    return {
      ok: false,
      error: e && e.message ? e.message : 'Error desconocido al crear el producto detallado.',
      data: null
    };
  }
}






