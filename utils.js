/**
 * utils.gs
 * Fonctions utilitaires pour le projet SecoMóvil
 * Ce fichier ne contient pas de logique métier.
 * Il sert uniquement à lire les feuilles, récupérer les listes
 * et gérer l’ajout d’éléments dans LISTAS.
 */

/**
 * Retourne une feuille par son nom.
 * Lance une erreur claire si la feuille n’existe pas.
 *
 * @param {string} name - Nom de la feuille (ex.: "PEDIDOS DIARIOS")
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getSheet_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    throw new Error('La feuille "' + name + '" est introuvable.');
  }
  return sheet;
}

/**
 * Retourne toutes les listes de la feuille LISTAS
 * dans un objet JSON.
 * Ce sera utilisé par les formulaires HTML.
 *
 * Structure retournée :
 * {
 *   estadosPedido: [...],
 *   pagoRecibido: [...],
 *   estadosCliente: [...],
 *   categoriasGasto: [...],
 *   unidades: [...],
 *   proveedores: [...],
 *   menus: [...],
 *   productos: [...]
 * }
 *
 * On lit seulement les lignes non vides.
 */
function getLists() {
  var sheet = getSheet_('LISTAS');
  var lastRow = sheet.getLastRow();

  // Colonnes fixes selon ce qu’on a défini ensemble
  var estadosPedido   = sheet.getRange('A2:A' + lastRow).getValues().flat().filter(String);
  var pagoRecibido    = sheet.getRange('B2:B' + lastRow).getValues().flat().filter(String);
  var estadosCliente  = sheet.getRange('C2:C' + lastRow).getValues().flat().filter(String);
  var categoriasGasto = sheet.getRange('D2:D' + lastRow).getValues().flat().filter(String);
  var unidades        = sheet.getRange('E2:E' + lastRow).getValues().flat().filter(String);
  var proveedores     = sheet.getRange('F2:F' + lastRow).getValues().flat().filter(String);
  var menus           = sheet.getRange('G2:G' + lastRow).getValues().flat().filter(String);
  var productos       = sheet.getRange('H2:H' + lastRow).getValues().flat().filter(String);

  return {
    estadosPedido: estadosPedido,
    pagoRecibido: pagoRecibido,
    estadosCliente: estadosCliente,
    categoriasGasto: categoriasGasto,
    unidades: unidades,
    proveedores: proveedores,
    menus: menus,
    productos: productos
  };
}

/**
 * Ajoute une valeur dans une colonne de LISTAS si elle n'existe pas déjà.
 * Exemple d’appel:
 *   addToList_('LISTAS', 'H', 'Ajo')
 *
 * @param {string} sheetName - en général "LISTAS"
 * @param {string} colLetter - lettre de la colonne (A, B, C, ...)
 * @param {string} value - valeur à ajouter
 * @returns {boolean} true si ajouté, false si déjà présent
 */
function addToList_(sheetName, colLetter, value) {
  if (!value) return false;

  var sheet = getSheet_(sheetName);
  var lastRow = sheet.getLastRow();
  var range = sheet.getRange(colLetter + '2:' + colLetter + lastRow);
  var values = range.getValues().flat();

  // Vérifier si la valeur existe déjà (sans tenir compte des majuscules)
  var exists = values.some(function (v) {
    return (v || '').toString().trim().toLowerCase() === value.toString().trim().toLowerCase();
  });

  if (exists) {
    return false;
  }

  // Chercher la première ligne vide sous la liste
  var insertRow = lastRow + 1;
  // mais on peut aussi chercher le premier vide dans la colonne
  for (var i = 0; i < values.length; i++) {
    if (!values[i]) {
      insertRow = i + 2; // +2 car démarré à la ligne 2
      break;
    }
  }

  sheet.getRange(colLetter + insertRow).setValue(value);

  // On peut trier la liste après ajout
  ordenarListaColumna_(sheet, colLetter);

  return true;
}

/**
 * Trie une colonne précise de la feuille LISTAS
 * et remet "Otro" en dernier si présent.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {string} colLetter
 */
function ordenarListaColumna_(sheet, colLetter) {
  var lastRow = sheet.getLastRow();
  var range = sheet.getRange(colLetter + '2:' + colLetter + lastRow);
  var values = range.getValues().flat();

  // Filtrer les vides
  var filled = values.filter(function (v) { return v && v.toString().trim() !== ''; });

  if (filled.length === 0) {
    return;
  }

  // Séparer "Otro" s'il existe
  var otros = filled.filter(function (v) {
    return v.toString().trim().toLowerCase() === 'otro';
  });
  var normales = filled.filter(function (v) {
    return v.toString().trim().toLowerCase() !== 'otro';
  });

  // Trier les valeurs normales
  normales.sort(function (a, b) {
    a = a.toString().toLowerCase();
    b = b.toString().toLowerCase();
    if (a < b) return 1 * -1; // ordre alphabétique ascendant
    if (a > b) return 1 * 1;
    return 0;
  });

  // Recomposer la liste: d’abord les normales, puis "Otro"
  var finalList = normales.concat(otros);

  // Réécrire la colonne proprement
  var out = finalList.map(function (v) { return [v]; });
  range.clearContent();
  sheet.getRange(colLetter + '2:' + colLetter + (finalList.length + 1)).setValues(out);
}

/**
 * Trie rapidement toutes les colonnes de LISTAS que nous utilisons.
 * À appeler ponctuellement si beaucoup d'éléments ont été ajoutés.
 */
function ordenarListas() {
  var sheet = getSheet_('LISTAS');
  // A: estados pedido
  ordenarListaColumna_(sheet, 'A');
  // B: pago recibido
  ordenarListaColumna_(sheet, 'B');
  // C: estados cliente
  ordenarListaColumna_(sheet, 'C');
  // D: categorías gasto
  ordenarListaColumna_(sheet, 'D');
  // E: unidades
  ordenarListaColumna_(sheet, 'E');
  // F: proveedores
  ordenarListaColumna_(sheet, 'F');
  // G: menús
  ordenarListaColumna_(sheet, 'G');
  // H: productos
  ordenarListaColumna_(sheet, 'H');
}
