/**
 * SUMFONT — сумма числовых ячеек по цвету шрифта.
 *
 * Пример:
 * =SUMFONT(M5:M44;SUMFONT_GREEN;SUMFONT_TICK)
 *
 * @param {Range} values Диапазон, который нужно суммировать.
 * @param {Range} colorSample Ячейка-образец нужного цвета шрифта.
 * @param {*} tick Служебный аргумент для принудительного пересчёта.
 * @return {number} Сумма числовых ячеек с нужным цветом шрифта.
 * @customfunction
 */
function SUMFONT(values, colorSample, tick) {
  // Аргументы values, colorSample и tick нужны Google Таблицам
  // как зависимости формулы. Ссылки на диапазоны извлекаются
  // из текста самой формулы, чтобы получить форматирование.
  void values;
  void colorSample;
  void tick;

  const formulaRange = SpreadsheetApp.getActiveRange();

  if (!formulaRange) {
    throw new Error(
      'Не удалось определить ячейку с формулой SUMFONT.'
    );
  }

  const formulaCell = formulaRange.getCell(1, 1);
  const formula = formulaCell.getFormula();
  const args = parseSUMFONTArgs_(formula);

  if (args.length < 2) {
    throw new Error(
      'Используйте формулу вида ' +
      '=SUMFONT(M5:M44;SUMFONT_GREEN;SUMFONT_TICK)'
    );
  }

  const spreadsheet = formulaCell
    .getSheet()
    .getParent();

  const sourceRange = getRangeFromReference_(
    args[0],
    formulaCell.getSheet()
  );

  const sampleRange = getRangeFromReference_(
    args[1],
    formulaCell.getSheet()
  );

  if (
    sampleRange.getNumRows() !== 1 ||
    sampleRange.getNumColumns() !== 1
  ) {
    throw new Error(
      'Второй аргумент SUMFONT должен быть одной ячейкой.'
    );
  }

  const sourceValues = sourceRange.getValues();
  const sourceColors = sourceRange.getFontColorObjects();

  const targetColor = colorKey_(
    sampleRange.getFontColorObject(),
    spreadsheet
  );

  let total = 0;

  for (
    let rowIndex = 0;
    rowIndex < sourceValues.length;
    rowIndex += 1
  ) {
    for (
      let columnIndex = 0;
      columnIndex < sourceValues[rowIndex].length;
      columnIndex += 1
    ) {
      const value =
        sourceValues[rowIndex][columnIndex];

      const color = colorKey_(
        sourceColors[rowIndex][columnIndex],
        spreadsheet
      );

      if (
        color === targetColor &&
        typeof value === 'number' &&
        Number.isFinite(value)
      ) {
        total += value;
      }
    }
  }

  return total;
}


/**
 * Устанавливает SUMFONT с нуля:
 * создаёт служебный лист и именованные диапазоны,
 * удаляет старые триггеры SUMFONT и создаёт один новый.
 *
 * Эту функцию нужно запустить вручную один раз.
 */
function setupSUMFONT() {
  const spreadsheet =
    SpreadsheetApp.getActiveSpreadsheet();

  if (!spreadsheet) {
    throw new Error(
      'Откройте Apps Script из нужной Google Таблицы.'
    );
  }

  let serviceSheet = spreadsheet.getSheetByName(
    SUMFONT_SERVICE_SHEET_
  );

  if (!serviceSheet) {
    serviceSheet = spreadsheet.insertSheet(
      SUMFONT_SERVICE_SHEET_
    );
  }

  serviceSheet.showSheet();
  serviceSheet.clear();

  const tickCell = serviceSheet.getRange('A1');
  const greenCell = serviceSheet.getRange('B1');
  const blueCell = serviceSheet.getRange('B2');

  tickCell
    .setNumberFormat('0')
    .setValue(1)
    .setNote(
      'Служебный счётчик SUMFONT. Не изменять вручную.'
    );

  greenCell
    .setValue('GREEN')
    .setFontColor('#34A853')
    .setNote(
      'Образец зелёного цвета для SUMFONT.'
    );

  blueCell
    .setValue('BLUE')
    .setFontColor('#0000FF')
    .setNote(
      'Образец синего цвета для SUMFONT.'
    );

  replaceNamedRange_(
    spreadsheet,
    SUMFONT_TICK_NAME_,
    tickCell
  );

  replaceNamedRange_(
    spreadsheet,
    SUMFONT_GREEN_NAME_,
    greenCell
  );

  replaceNamedRange_(
    spreadsheet,
    SUMFONT_BLUE_NAME_,
    blueCell
  );

  removeOldSUMFONTTriggers_();

  const properties =
    PropertiesService.getDocumentProperties();

  properties.deleteProperty(
    SUMFONT_SIGNATURE_PROPERTY_
  );

  properties.setProperty(
    SUMFONT_SIGNATURE_PROPERTY_,
    buildColorSignature_(spreadsheet)
  );

  ScriptApp
    .newTrigger('SUMFONT_WATCH')
    .forSpreadsheet(spreadsheet)
    .onChange()
    .create();

  serviceSheet.hideSheet();

  spreadsheet.toast(
    'SUMFONT установлен. Теперь вставьте формулы и проверьте пересчёт.',
    'SUMFONT',
    8
  );
}


/**
 * Устанавливаемый триггер «При изменении».
 *
 * Он сравнивает текущую цветовую сигнатуру с предыдущей.
 * Если значения или цвета в диапазонах SUMFONT изменились,
 * увеличивает SUMFONT_TICK и заставляет формулы пересчитаться.
 */
function SUMFONT_WATCH(e) {
  const spreadsheet =
    e && e.source
      ? e.source
      : SpreadsheetApp.getActiveSpreadsheet();

  if (!spreadsheet) {
    return;
  }

  const lock = LockService.getDocumentLock();

  if (!lock.tryLock(10000)) {
    return;
  }

  try {
    const properties =
      PropertiesService.getDocumentProperties();

    const oldSignature = properties.getProperty(
      SUMFONT_SIGNATURE_PROPERTY_
    );

    const newSignature =
      buildColorSignature_(spreadsheet);

    if (oldSignature === newSignature) {
      return;
    }

    const tickCell = spreadsheet.getRangeByName(
      SUMFONT_TICK_NAME_
    );

    if (!tickCell) {
      throw new Error(
        'Не найден именованный диапазон ' +
        SUMFONT_TICK_NAME_ +
        '. Запустите setupSUMFONT.'
      );
    }

    const currentValue =
      Number(tickCell.getValue()) || 0;

    const nextValue =
      currentValue >= 9000000000000
        ? 1
        : currentValue + 1;

    tickCell.setValue(nextValue);
    SpreadsheetApp.flush();

    properties.setProperty(
      SUMFONT_SIGNATURE_PROPERTY_,
      newSignature
    );
  } finally {
    lock.releaseLock();
  }
}


/**
 * Ручной принудительный пересчёт.
 *
 * Запускайте эту функцию только при необходимости проверки.
 */
function refreshSUMFONT() {
  const spreadsheet =
    SpreadsheetApp.getActiveSpreadsheet();

  const tickCell = spreadsheet.getRangeByName(
    SUMFONT_TICK_NAME_
  );

  if (!tickCell) {
    throw new Error(
      'Не найден SUMFONT_TICK. Сначала запустите setupSUMFONT.'
    );
  }

  tickCell.setValue(
    (Number(tickCell.getValue()) || 0) + 1
  );

  SpreadsheetApp.flush();

  PropertiesService
    .getDocumentProperties()
    .setProperty(
      SUMFONT_SIGNATURE_PROPERTY_,
      buildColorSignature_(spreadsheet)
    );

  spreadsheet.toast(
    'Формулы SUMFONT отправлены на пересчёт.',
    'SUMFONT',
    5
  );
}


/**
 * Проверка настройки.
 */
function checkSUMFONT() {
  const spreadsheet =
    SpreadsheetApp.getActiveSpreadsheet();

  const formulaCells =
    findSUMFONTFormulaCells_(spreadsheet);

  const triggers = ScriptApp
    .getProjectTriggers()
    .filter(function (trigger) {
      return (
        trigger.getHandlerFunction() ===
        'SUMFONT_WATCH'
      );
    });

  const result = {
    tickNamedRange:
      Boolean(
        spreadsheet.getRangeByName(
          SUMFONT_TICK_NAME_
        )
      ),

    greenNamedRange:
      Boolean(
        spreadsheet.getRangeByName(
          SUMFONT_GREEN_NAME_
        )
      ),

    formulasFound:
      formulaCells.map(function (cell) {
        return (
          cell.getSheet().getName() +
          '!' +
          cell.getA1Notation()
        );
      }),

    triggerCount: triggers.length
  };

  console.log(
    JSON.stringify(result, null, 2)
  );

  return result;
}


/**
 * Собирает сигнатуру всех формул SUMFONT:
 * их диапазонов, значений и цветов.
 */
function buildColorSignature_(spreadsheet) {
  const formulaCells =
    findSUMFONTFormulaCells_(spreadsheet);

  const items = formulaCells.map(
    function (formulaCell) {
      const formula = formulaCell.getFormula();
      const args = parseSUMFONTArgs_(formula);

      if (args.length < 2) {
        return {
          formulaCell:
            formulaCell.getSheet().getName() +
            '!' +
            formulaCell.getA1Notation(),
          formula: formula,
          error: 'Недостаточно аргументов'
        };
      }

      const sourceRange =
        getRangeFromReference_(
          args[0],
          formulaCell.getSheet()
        );

      const sampleRange =
        getRangeFromReference_(
          args[1],
          formulaCell.getSheet()
        );

      const values = sourceRange.getValues();
      const colorObjects =
        sourceRange.getFontColorObjects();

      const colors = colorObjects.map(
        function (row) {
          return row.map(function (color) {
            return colorKey_(
              color,
              spreadsheet
            );
          });
        }
      );

      return {
        formulaCell:
          formulaCell.getSheet().getName() +
          '!' +
          formulaCell.getA1Notation(),

        formula: formula,

        source:
          sourceRange.getSheet().getName() +
          '!' +
          sourceRange.getA1Notation(),

        sample:
          sampleRange.getSheet().getName() +
          '!' +
          sampleRange.getA1Notation(),

        sampleColor: colorKey_(
          sampleRange.getFontColorObject(),
          spreadsheet
        ),

        values: values,
        colors: colors
      };
    }
  );

  return hashText_(
    JSON.stringify(items)
  );
}


/**
 * Находит все ячейки с формулами SUMFONT.
 */
function findSUMFONTFormulaCells_(spreadsheet) {
  return spreadsheet
    .createTextFinder('SUMFONT(')
    .matchFormulaText(true)
    .matchCase(false)
    .findAll()
    .filter(function (cell) {
      return (
        cell.getSheet().getName() !==
          SUMFONT_SERVICE_SHEET_ &&
        /^=SUMFONT\s*\(/i.test(
          cell.getFormula()
        )
      );
    })
    .sort(function (first, second) {
      const firstKey =
        first.getSheet().getSheetId() +
        ':' +
        first.getRow() +
        ':' +
        first.getColumn();

      const secondKey =
        second.getSheet().getSheetId() +
        ':' +
        second.getRow() +
        ':' +
        second.getColumn();

      return firstKey.localeCompare(secondKey);
    });
}


/**
 * Разбирает аргументы формулы SUMFONT.
 * Поддерживает как «;», так и «,» в качестве разделителя.
 */
function parseSUMFONTArgs_(formula) {
  const openIndex = formula.indexOf('(');
  const closeIndex = formula.lastIndexOf(')');

  if (
    openIndex === -1 ||
    closeIndex === -1 ||
    closeIndex <= openIndex
  ) {
    return [];
  }

  const body = formula.slice(
    openIndex + 1,
    closeIndex
  );

  const args = [];
  let current = '';
  let depth = 0;
  let quote = '';

  for (
    let index = 0;
    index < body.length;
    index += 1
  ) {
    const char = body[index];

    if (quote) {
      current += char;

      if (char === quote) {
        if (
          quote === "'" &&
          body[index + 1] === "'"
        ) {
          current += body[index + 1];
          index += 1;
        } else {
          quote = '';
        }
      }

      continue;
    }

    if (char === '"' || char === "'") {
      quote = char;
      current += char;
      continue;
    }

    if (char === '(') {
      depth += 1;
      current += char;
      continue;
    }

    if (char === ')') {
      depth -= 1;
      current += char;
      continue;
    }

    if (
      depth === 0 &&
      (char === ';' || char === ',')
    ) {
      args.push(current.trim());
      current = '';
      continue;
    }

    current += char;
  }

  args.push(current.trim());

  return args;
}


/**
 * Превращает ссылку из формулы в объект Range.
 *
 * Поддерживает:
 * M5:M44
 * $AD$5
 * 'Другой лист'!A1:B10
 * именованные диапазоны SUMFONT_GREEN и SUMFONT_TICK
 */
function getRangeFromReference_(
  reference,
  defaultSheet
) {
  const spreadsheet =
    defaultSheet.getParent();

  let text = String(reference || '')
    .trim();

  if (!text) {
    throw new Error(
      'В SUMFONT передана пустая ссылка.'
    );
  }

  while (
    text.startsWith('(') &&
    text.endsWith(')')
  ) {
    text = text.slice(1, -1).trim();
  }

  const namedRange =
    spreadsheet.getRangeByName(
      text.replace(/\$/g, '')
    );

  if (namedRange) {
    return namedRange;
  }

  if (/[()]/.test(text)) {
    throw new Error(
      'SUMFONT принимает прямую ссылку на диапазон, ' +
      'а не другую функцию: ' +
      text
    );
  }

  let sheet = defaultSheet;
  let a1Notation = text;

  const exclamationIndex =
    text.lastIndexOf('!');

  if (exclamationIndex !== -1) {
    const sheetName = unquoteSheetName_(
      text.slice(0, exclamationIndex)
    );

    sheet = spreadsheet.getSheetByName(
      sheetName
    );

    if (!sheet) {
      throw new Error(
        'Не найден лист: ' +
        sheetName
      );
    }

    a1Notation = text.slice(
      exclamationIndex + 1
    );
  }

  a1Notation = a1Notation
    .replace(/\$/g, '')
    .trim();

  return sheet.getRange(a1Notation);
}


/**
 * Возвращает фактический HEX-цвет.
 * Цвета темы переводятся в их конкретный RGB-цвет.
 */
function colorKey_(color, spreadsheet) {
  if (!color) {
    return '';
  }

  const colorType = color.getColorType();

  if (
    colorType ===
    SpreadsheetApp.ColorType.RGB
  ) {
    return color
      .asRgbColor()
      .asHexString()
      .toUpperCase();
  }

  if (
    colorType ===
    SpreadsheetApp.ColorType.THEME
  ) {
    const themeColorType = color
      .asThemeColor()
      .getThemeColorType();

    const theme =
      spreadsheet.getSpreadsheetTheme();

    if (theme) {
      return theme
        .getConcreteColor(themeColorType)
        .asRgbColor()
        .asHexString()
        .toUpperCase();
    }

    return (
      'THEME:' +
      String(themeColorType)
    );
  }

  return '';
}


/**
 * Создаёт или заменяет именованный диапазон.
 */
function replaceNamedRange_(
  spreadsheet,
  name,
  range
) {
  if (spreadsheet.getRangeByName(name)) {
    spreadsheet.removeNamedRange(name);
  }

  spreadsheet.setNamedRange(
    name,
    range
  );
}


/**
 * Удаляет только старые триггеры,
 * связанные с SUMFONT.
 */
function removeOldSUMFONTTriggers_() {
  const oldHandlerNames = [
    'SUMFONT_WATCH',
    'refreshSumFontOnChange',
    'onSumFontChange'
  ];

  ScriptApp
    .getProjectTriggers()
    .forEach(function (trigger) {
      if (
        oldHandlerNames.indexOf(
          trigger.getHandlerFunction()
        ) !== -1
      ) {
        ScriptApp.deleteTrigger(trigger);
      }
    });
}


/**
 * Снимает кавычки с имени листа.
 */
function unquoteSheetName_(value) {
  let text = String(value || '').trim();

  if (
    text.startsWith("'") &&
    text.endsWith("'")
  ) {
    text = text
      .slice(1, -1)
      .replace(/''/g, "'");
  }

  return text;
}


/**
 * MD5-подпись строки.
 */
function hashText_(text) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.MD5,
    text,
    Utilities.Charset.UTF_8
  );

  return bytes
    .map(function (byte) {
      const unsigned =
        byte < 0
          ? byte + 256
          : byte;

      return (
        '0' +
        unsigned.toString(16)
      ).slice(-2);
    })
    .join('');
}


const SUMFONT_SERVICE_SHEET_ =
  '__SUMFONT_SERVICE';

const SUMFONT_TICK_NAME_ =
  'SUMFONT_TICK';

const SUMFONT_GREEN_NAME_ =
  'SUMFONT_GREEN';

const SUMFONT_BLUE_NAME_ =
  'SUMFONT_BLUE';

const SUMFONT_SIGNATURE_PROPERTY_ =
  'SUMFONT_SIGNATURE_V3';
