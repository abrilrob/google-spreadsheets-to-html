// Includes functions for exporting spreadsheet to HTML Voluntaries page.

var STRUCTURE_LIST = 'List';
var STRUCTURE_HASH = 'Hash (keyed by "id" column)';

var DEFAULT_LANGUAGE = 'JavaScript';
var DEFAULT_FORMAT = 'Pretty';
var DEFAULT_STRUCTURE = STRUCTURE_LIST;

// Here we add the option in the toolbar menu.
function onOpen() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var menuEntries = [
        {name: "Export Otros", functionName: "exportOtros"},
        {name: "Export Colaboradores", functionName: "exportColaboradores"}
    ];
    ss.addMenu("Export Web", menuEntries);
}
// Method to export.
function exportOtros(e) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();
    var rowsData = getRowsData_(sheet, getExportOptions_(e));
    var html = '<ul>';
    for(var x = 0; x < rowsData.length; x++){
        if(rowsData[x].tipo == 'Otros'){
            html+='<li><a href="'+rowsData[x].link+'" target="_blank">'+rowsData[x].nombre+'</a></li>';
        }
    }
    html += '</ul>';
    displayText_(html);
}
// Method to export.
function exportColaboradores(e) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();
    var rowsData = getRowsData_(sheet, getExportOptions_(e));
    var html = '<ul>';
    for(var x = 0; x < rowsData.length; x++){
        if(rowsData[x].tipo == 'Colaborador'){
            html+='<li><a href="'+rowsData[x].link+'" target="_blank">'+rowsData[x].nombre+'</a></li>';
        }
    }
    html += '</ul>';
    displayText_(html);
}

/*
    The next functions are used to export, you can go in more detail, but the important to do the export are the previous functions.
*/

function getExportOptions_(e) {
    var options = {};

    options.language = e && e.parameter.language || DEFAULT_LANGUAGE;
    options.format   = e && e.parameter.format || DEFAULT_FORMAT;
    options.structure = e && e.parameter.structure || DEFAULT_STRUCTURE;

    var cache = CacheService.getPublicCache();
    cache.put('language', options.language);
    cache.put('format',   options.format);
    cache.put('structure',   options.structure);

    Logger.log(options);
    return options;
}

function displayText_(text) {
    var output = HtmlService.createHtmlOutput("<textarea style='width:100%;' rows='20'>" + text + "</textarea>");
    output.setWidth(400)
    output.setHeight(300);
    SpreadsheetApp.getUi().showModalDialog(output, 'Exported HTML');
}

function getRowsData_(sheet, options) {
    var headersRange = sheet.getRange(1, 1, sheet.getFrozenRows(), sheet.getMaxColumns());
    var headers = headersRange.getValues()[0];
    var dataRange = sheet.getRange(sheet.getFrozenRows()+1, 1, sheet.getMaxRows(), sheet.getMaxColumns());
    var objects = getObjects_(dataRange.getValues(), normalizeHeaders_(headers));
    if (options.structure == STRUCTURE_HASH) {
        var objectsById = {};
        objects.forEach(function(object) {
            objectsById[object.id] = object;
        });
        return objectsById;
    } else {
        return objects;
    }
}

function getColumnsData_(sheet, range, rowHeadersColumnIndex) {
    rowHeadersColumnIndex = rowHeadersColumnIndex || range.getColumnIndex() - 1;
    var headersTmp = sheet.getRange(range.getRow(), rowHeadersColumnIndex, range.getNumRows(), 1).getValues();
    var headers = normalizeHeaders_(arrayTranspose_(headersTmp)[0]);
    return getObjects(arrayTranspose_(range.getValues()), headers);
}

function getObjects_(data, keys) {
    var objects = [];
    for (var i = 0; i < data.length; ++i) {
        var object = {};
        var hasData = false;
        for (var j = 0; j < data[i].length; ++j) {
            var cellData = data[i][j];
            if (isCellEmpty_(cellData)) {
                continue;
            }
            object[keys[j]] = cellData;
            hasData = true;
        }
        if (hasData) {
            objects.push(object);
        }
    }
    return objects;
}

function normalizeHeaders_(headers) {
    var keys = [];
    for (var i = 0; i < headers.length; ++i) {
        var key = normalizeHeader_(headers[i]);
        if (key.length > 0) {
            keys.push(key);
        }
    }
    return keys;
}

function normalizeHeader_(header) {
    var key = "";
    var upperCase = false;
    for (var i = 0; i < header.length; ++i) {
        var letter = header[i];
        if (letter == " " && key.length > 0) {
            upperCase = true;
            continue;
        }
        if (!isAlnum_(letter)) {
            continue;
        }
        if (key.length == 0 && isDigit_(letter)) {
            continue; // first character must be a letter
        }
        if (upperCase) {
            upperCase = false;
            key += letter.toUpperCase();
        } else {
            key += letter.toLowerCase();
        }
    }
    return key;
}

function isCellEmpty_(cellData) {
    return typeof(cellData) == "string" && cellData == "";
}

function isAlnum_(char) {
    return char >= 'A' && char <= 'Z' ||
        char >= 'a' && char <= 'z' ||
        isDigit_(char);
}

function isDigit_(char) {
    return char >= '0' && char <= '9';
}

function arrayTranspose_(data) {
    if (data.length == 0 || data[0].length == 0) {
        return null;
    }

    var ret = [];
    for (var i = 0; i < data[0].length; ++i) {
        ret.push([]);
    }

    for (var i = 0; i < data.length; ++i) {
        for (var j = 0; j < data[i].length; ++j) {
            ret[j][i] = data[i][j];
        }
    }

    return ret;
}