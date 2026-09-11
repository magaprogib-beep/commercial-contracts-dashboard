/**
 * Commercial Contracts Dashboard
 * Полный серверный слой данных для Apps Script.
 *
 * Этот файл заменяет предыдущий тестовый Data.gs.
 * Он сохраняет совместимость с контрольной страницей:
 * data.receipts, data.contracting и data.overallKpi
 * остаются доступны в прежнем формате.
 */

const DASHBOARD_CACHE_KEY = 'commercial-dashboard-full-v39';
const DASHBOARD_CACHE_SECONDS = 30;

const DASHBOARD_SECTIONS = [
  'Разработка ТС',
  'Разработка ПОДД',
  'Прочие договоры'
];


/**
 * Главная функция, которую вызывает HTML-страница.
 */
function getDashboardData() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(DASHBOARD_CACHE_KEY);

  if (cached) {
    return JSON.parse(cached);
  }

  const data = buildFullDashboardData_();
  const json = JSON.stringify(data);

  /*
   * CacheService ограничивает размер одного значения.
   * Если полный объект помещается — кэшируем.
   * Если нет — просто возвращаем результат без ошибки.
   */
  if (json.length < 90000) {
    try {
      cache.put(
        DASHBOARD_CACHE_KEY,
        json,
        DASHBOARD_CACHE_SECONDS
      );
    } catch (error) {
      console.log(
        'Кэш не записан, но данные рассчитаны: ' +
        error.message
      );
    }
  }

  return data;
}


/**
 * Сбрасывает кэш после изменения таблицы
 * или во время ручной проверки.
 */
function clearDashboardCache() {
  CacheService
    .getScriptCache()
    .remove(DASHBOARD_CACHE_KEY);
}


/**
 * Однократно устанавливает триггеры для сброса кэша
 * после редактирования или изменения структуры таблицы.
 */
function installDashboardTriggers() {
  const spreadsheet = SpreadsheetApp.openById(
    getSpreadsheetId_()
  );

  ScriptApp.getProjectTriggers().forEach(
    function (trigger) {
      if (
        trigger.getHandlerFunction() ===
        'clearDashboardCache'
      ) {
        ScriptApp.deleteTrigger(trigger);
      }
    }
  );

  ScriptApp
    .newTrigger('clearDashboardCache')
    .forSpreadsheet(spreadsheet)
    .onEdit()
    .create();

  ScriptApp
    .newTrigger('clearDashboardCache')
    .forSpreadsheet(spreadsheet)
    .onChange()
    .create();

  return 'Триггеры дашборда установлены.';
}


/**
 * Компактная ручная проверка полного набора данных.
 *
 * Запускается из редактора Apps Script:
 * testDashboardDataFull → Выполнить.
 */
function testDashboardDataFull() {
  clearDashboardCache();

  const data = getDashboardData();

  const check = {
    spreadsheetName: data.spreadsheetName,
    updatedAt: data.updatedAt,

    receipts: {
      quarters: data.receipts.quarters,
      actualTotal: data.receipts.actualTotal,
      planTotal: data.receipts.planTotal,
      attainment: data.receipts.attainment,
      forecastQuarters:
        data.receipts.forecastQuarters,
      blueForecastQuarters:
        data.receipts.blueForecastQuarters,
      currentQuarter:
        data.receipts.currentQuarter,
      currentQuarterForecast:
        data.receipts.currentQuarterForecast
    },

    contracting: {
      quarters: data.contracting.quarters,
      actualTotal: data.contracting.actualTotal,
      planTotal: data.contracting.planTotal,
      attainment: data.contracting.attainment,
      detailedItems:
        data.contracting.items.length,
      unmatchedItems:
        data.contracting.unmatchedItems,
      total2025:
        data.contracting.total2025
    },

    overallKpi: data.overallKpi,

    portfolio: {
      signed: data.portfolio.signedAgg.count,
      opportunities:
        data.portfolio.pipelineAgg.count,
      subsidies:
        data.portfolio.subsidyAgg.count,
      signedValue:
        data.portfolio.signedAgg.value,
      signedBalance:
        data.portfolio.signedAgg.balance
    },

    dataQuality: {
      total: data.dataQuality.summary.total,
      critical: data.dataQuality.summary.critical,
      review: data.dataQuality.summary.review,
      sample: data.dataQuality.issues.slice(0, 5)
    },

    registrySamples: {
      signed: data.registry.signed.slice(0, 2),
      opportunities:
        data.registry.pipeline.slice(0, 2)
    }
  };

  console.log(
    JSON.stringify(check, null, 2)
  );

  return check;
}


/**
 * Сохраняем функцию со старым именем:
 * если она уже использовалась в инструкции,
 * она продолжит работать.
 */
function testDashboardData() {
  return testDashboardDataFull();
}


/**
 * Собирает единый объект данных.
 */
function buildFullDashboardData_() {
  const spreadsheetId = getSpreadsheetId_();

  const spreadsheet = SpreadsheetApp.openById(
    spreadsheetId
  );

  const statusSheet = spreadsheet.getSheetByName(
    'Статус'
  );

  const contractingSheet = spreadsheet.getSheetByName(
    'Контрактация/закрытие 2026'
  );

  if (!statusSheet) {
    throw new Error(
      'Не найден лист «Статус».'
    );
  }

  if (!contractingSheet) {
    throw new Error(
      'Не найден лист «Контрактация/закрытие 2026».'
    );
  }

  const statusData = parseStatusSheet_(
    statusSheet
  );

  const contractingData =
    parseContractingSheet_(
      contractingSheet,
      statusData.signed
    );

  const statusGroups = buildStatusGroups_(
    statusData.signed
  );

  const signedAgg = aggregateRecords_(
    statusData.signed
  );

  const pipelineAgg = aggregateRecords_(
    statusData.pipeline
  );

  const subsidyAgg = aggregateRecords_(
    statusData.subsidy
  );

  const pipelineTypes = countTypes_(
    statusData.pipeline
  );

  const qualityIssues = statusData
    .qualityIssues
    .slice();

  contractingData.unmatchedItems.forEach(
    function (item) {
      qualityIssues.push({
        row: item.row || null,
        severity: 'critical',
        title: 'Строка контрактации не связана с договором',
        description:
          '«' + item.name + '», ' +
          item.quarter + ' кв., ' +
          roundMoney_(item.amount) +
          ' ₽: позиция есть на листе контрактации, ' +
          'но не сопоставилась с договором.',
        query: item.name,
        scope: 'signed'
      });
    }
  );

  qualityIssues.sort(function (first, second) {
    const severityOrder = {
      critical: 0,
      review: 1
    };

    const firstOrder =
      severityOrder[first.severity] ?? 9;

    const secondOrder =
      severityOrder[second.severity] ?? 9;

    if (firstOrder !== secondOrder) {
      return firstOrder - secondOrder;
    }

    return numberOrZero_(first.row) -
      numberOrZero_(second.row);
  });

  const qualitySummary = {
    total: qualityIssues.length,
    critical: qualityIssues.filter(
      function (issue) {
        return issue.severity === 'critical';
      }
    ).length,
    review: qualityIssues.filter(
      function (issue) {
        return issue.severity === 'review';
      }
    ).length
  };

  const core = {
    spreadsheetName: spreadsheet.getName(),

    receipts: {
      quarters:
        contractingData.receipts.quarters,

      actualTotal:
        contractingData.receipts.actualTotal,

      planQuarters:
        contractingData.receipts.planQuarters,

      planTotal:
        contractingData.receipts.planTotal,

      attainment:
        contractingData.receipts.attainment,

      forecastQuarters:
        statusData.chartForecastQuarters,

      blueForecastQuarters:
        statusData.blueForecastQuarters,

      currentQuarter: 4,

      currentQuarterForecast:
        statusData.currentQuarterForecast,

      receipts2025:
        statusData.receipts2025
    },

    contracting: {
      quarters:
        contractingData.contracting.quarters,

      actualTotal:
        contractingData.contracting.actualTotal,

      planQuarters:
        contractingData.contracting.planQuarters,

      planTotal:
        contractingData.contracting.planTotal,

      attainment:
        contractingData.contracting.attainment,

      items:
        contractingData.items,

      unmatchedItems:
        contractingData.unmatchedItems,

      total2025:
        contractingData.contracting2025Total
    },

    overallKpi:
      contractingData.overallKpi,

    portfolio: {
      signedAgg: signedAgg,
      pipelineAgg: pipelineAgg,
      subsidyAgg: subsidyAgg,
      statusGroups: statusGroups,
      pipelineTypes: pipelineTypes
    },

    registry: {
      signed: statusData.signed,
      pipeline: statusData.pipeline,
      subsidy: statusData.subsidy
    },

    dataQuality: {
      summary: qualitySummary,
      issues: qualityIssues
    }
  };

  core.version = createDataVersion_(core);

  core.updatedAt = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    'dd.MM.yyyy HH:mm:ss'
  );

  core.updatedAtIso = new Date().toISOString();

  return core;
}


/**
 * Читает основной лист договоров,
 * включая значения и цвета текста.
 */
function parseStatusSheet_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  const range = sheet.getRange(
    1,
    1,
    lastRow,
    lastColumn
  );

  const values = range.getValues();
  const displayValues = range.getDisplayValues();
  const fontColors = range.getFontColors();

  const firstHeader = displayValues[0];
  const secondHeader = displayValues[1];

  const subjectColumn = findHeaderColumn_(
    firstHeader,
    'Наименование работ / предмет договора'
  );

  const nameColumn = findHeaderColumn_(
    firstHeader,
    'Краткое название'
  );

  const contractGroupColumn = findHeaderColumn_(
    firstHeader,
    'Договор'
  );

  const contractNumberColumn =
    findSubheaderColumn_(
      secondHeader,
      '№',
      contractGroupColumn
    );

  const contractDateColumn =
    findSubheaderColumn_(
      secondHeader,
      'дата',
      contractGroupColumn
    );

  const counterpartyColumn = findHeaderColumn_(
    firstHeader,
    'Контрагент'
  );

  const valueColumn = findHeaderColumn_(
    firstHeader,
    'Цена договора, руб.'
  );

  const paidColumn = findHeaderColumnStartsWith_(
    firstHeader,
    'Работ оплачено на сумму'
  );

  const payment2026StartColumn =
    findHeaderColumn_(
      firstHeader,
      'План платежей в 2026 году'
    );

  const balanceColumn = findHeaderColumn_(
    firstHeader,
    'Остаток по контракту'
  );

  const statusColumn = findHeaderColumn_(
    firstHeader,
    'Статус'
  );

  const signed = [];
  const pipeline = [];
  const subsidy = [];
  const qualityIssues = [];

  const forecastQuarters = [
    0,
    0,
    0,
    0
  ];

  let currentSection = '';

  for (
    let rowIndex = 2;
    rowIndex < values.length;
    rowIndex += 1
  ) {
    const displayRow = displayValues[rowIndex];
    const valueRow = values[rowIndex];
    const colorRow = fontColors[rowIndex];

    /*
     * Названия разделов («Разработка ТС», «Разработка ПОДД»,
     * «Прочие договоры») находятся в колонке A, а предмет
     * договора — в колонке B. Поэтому ищем раздел по всей строке.
     */
    const sectionName = findSectionNameInRow_(
      displayRow
    );

    if (sectionName) {
      currentSection = sectionName;
      continue;
    }

    const subject = cleanText_(
      displayRow[subjectColumn]
    );

    if (!currentSection) {
      continue;
    }

    if (
      !subject ||
      subject === 'ВСЕГО:' ||
      subject.indexOf(
        'Суммарно по всем договорам'
      ) === 0
    ) {
      continue;
    }

    const shortName = cleanText_(
      displayRow[nameColumn]
    );

    const contractRaw =
      valueRow[contractNumberColumn];

    const dateRaw =
      valueRow[contractDateColumn];

    const status = cleanText_(
      displayRow[statusColumn]
    );

    /*
     * Нижние сводные строки не являются договорами:
     * у них нет одновременно предмета, краткого названия,
     * договора или статуса.
     */
    if (
      !shortName &&
      !hasValue_(contractRaw) &&
      !status
    ) {
      continue;
    }

    const recordForecast = sumByFontType_(
      valueRow,
      colorRow,
      payment2026StartColumn,
      4,
      'blue'
    );

    const received2026 = sumByFontType_(
      valueRow,
      colorRow,
      payment2026StartColumn,
      4,
      'green'
    );

    for (
      let quarterIndex = 0;
      quarterIndex < 4;
      quarterIndex += 1
    ) {
      const cellValue = numberOrZero_(
        valueRow[
          payment2026StartColumn +
          quarterIndex
        ]
      );

      const cellColor =
        colorRow[
          payment2026StartColumn +
          quarterIndex
        ];

      if (isBlueFont_(cellColor)) {
        forecastQuarters[quarterIndex] +=
          cellValue;
      }
    }

    const dateIso = dateToIso_(dateRaw);

    const record = {
      row: rowIndex + 1,
      section: currentSection,
      subject: subject,
      name: shortName || subject,
      contract: contractToString_(
        contractRaw
      ),
      date: dateIso,
      year: dateIso
        ? Number(dateIso.slice(0, 4))
        : null,
      counterparty: cleanText_(
        displayRow[counterpartyColumn]
      ),
      value: numberOrZero_(
        valueRow[valueColumn]
      ),
      paid: numberOrZero_(
        valueRow[paidColumn]
      ),
      balance: nullableNumber_(
        valueRow[balanceColumn]
      ),
      status: status,
      scope: '',
      received2026: received2026,
      forecast2026: recordForecast,
      contracting2026: 0
    };

    const contractText = cleanText_(
      record.contract
    ).toLowerCase();

    let destination = null;

    if (contractText === 'субсидия') {
      record.scope = 'subsidy';
      destination = subsidy;
    } else if (
      status.toLowerCase() ===
      'opportunities'
    ) {
      record.scope = 'pipeline';
      destination = pipeline;
    } else if (
      hasValue_(contractRaw) &&
      hasValue_(dateRaw)
    ) {
      record.scope = 'signed';
      destination = signed;
    } else {
      qualityIssues.push({
        row: rowIndex + 1,
        severity: 'critical',
        title: 'Неполные реквизиты договора',
        description:
          'Для строки «' +
          (shortName || subject) +
          '» должны быть одновременно заполнены ' +
          'номер и дата договора.',
        query: shortName || subject,
        scope: 'signed'
      });

      continue;
    }

    collectRecordQualityIssues_(
      qualityIssues,
      {
        record: record,
        shortName: shortName,
        contractRaw: contractRaw,
        dateRaw: dateRaw,
        valueRow: valueRow,
        displayRow: displayRow,
        colorRow: colorRow,
        payment2026StartColumn:
          payment2026StartColumn
      }
    );

    destination.push(record);
  }

  const currentQuarterForecast =
    readCurrentQuarterForecast_(
      sheet,
      values,
      displayValues,
      payment2026StartColumn
    );

  return {
    signed: signed,
    pipeline: pipeline,
    subsidy: subsidy,

    blueForecastQuarters:
      forecastQuarters.map(roundMoney_),

    currentQuarterForecast:
      roundMoney_(currentQuarterForecast),

    chartForecastQuarters: [
      0,
      0,
      0,
      roundMoney_(currentQuarterForecast)
    ],

    qualityIssues: qualityIssues,

    receipts2025:
      parseReceipts2025_(
        values,
        displayValues
      )
  };
}


/**
 * Собирает замечания по одной строке договора.
 */
function collectRecordQualityIssues_(
  issues,
  context
) {
  const record = context.record;
  const row = record.row;
  const query = record.name || record.subject;
  const scope = record.scope || 'signed';

  const knownStatuses = [
    'Действует',
    'Ожидание ИД',
    'Процесс закрытия',
    'Выполнен и закрыт',
    'Расторжение',
    'Расторгнут',
    'Отказ',
    'Приостановлен',
    'opportunities'
  ];

  if (
    typeof context.contractRaw === 'number' &&
    context.contractRaw === 0
  ) {
    issues.push({
      row: row,
      severity: 'critical',
      title: 'Номер договора равен 0',
      description:
        'В строке «' + query +
        '» вместо фактического номера договора стоит 0.',
      query: query,
      scope: scope
    });
  }

  if (
    record.scope === 'signed' &&
    !context.shortName
  ) {
    issues.push({
      row: row,
      severity: 'review',
      title: 'Не заполнено краткое название',
      description:
        'Столбец «Краткое название» пуст. ' +
        'Это ухудшает поиск и отображение договора.',
      query: record.contract || record.subject,
      scope: scope
    });
  }

  if (
    !record.status &&
    record.scope !== 'subsidy'
  ) {
    issues.push({
      row: row,
      severity: 'critical',
      title: 'Не заполнен статус',
      description:
        'У строки «' + query +
        '» отсутствует значение в столбце «Статус».',
      query: query,
      scope: scope
    });
  } else if (
    knownStatuses.indexOf(record.status) === -1 &&
    record.scope !== 'subsidy'
  ) {
    issues.push({
      row: row,
      severity: 'review',
      title: 'Неизвестный статус',
      description:
        'Статус «' + record.status +
        '» не входит в утверждённый перечень.',
      query: query,
      scope: scope
    });
  }

  if (
    record.scope === 'signed' &&
    (
      record.balance === null ||
      record.balance === undefined
    )
  ) {
    issues.push({
      row: row,
      severity: 'review',
      title: 'Не заполнен остаток по договору',
      description:
        'Для договора «' + query +
        '» не заполнен столбец «Остаток по контракту».',
      query: query,
      scope: 'signed'
    });
  }

  if (
    record.scope === 'pipeline' &&
    (
      record.balance === null ||
      record.balance === undefined
    )
  ) {
    issues.push({
      row: row,
      severity: 'review',
      title: 'Не заполнен остаток opportunity',
      description:
        'Для позиции «' + query +
        '» указана стоимость ' +
        roundMoney_(record.value) +
        ' ₽, но остаток по контракту пуст.',
      query: query,
      scope: 'pipeline'
    });
  }

  if (
    record.balance !== null &&
    record.balance !== undefined &&
    record.value > 0
  ) {
    const difference = roundMoney_(
      record.value -
      record.paid -
      record.balance
    );

    const terminatedStatuses = [
      'Расторжение',
      'Расторгнут',
      'Отказ',
      'Приостановлен'
    ];

    if (
      Math.abs(difference) > 1 &&
      terminatedStatuses.indexOf(
        record.status
      ) === -1
    ) {
      const severity =
        record.status === 'Выполнен и закрыт'
          ? 'review'
          : 'critical';

      issues.push({
        row: row,
        severity: severity,
        title: 'Не сходятся стоимость, поступления и остаток',
        description:
          '«' + query + '»: G = ' +
          roundMoney_(record.value) +
          ' ₽, L = ' +
          roundMoney_(record.paid) +
          ' ₽, остаток = ' +
          roundMoney_(record.balance) +
          ' ₽. Разница: ' +
          roundMoney_(difference) +
          ' ₽.',
        query: query,
        scope: scope
      });
    }
  }

  const textQuarters = [];
  const unknownColorQuarters = [];
  const quarterLabels = [
    '1 кв.',
    '2 кв.',
    '3 кв.',
    '4 кв.'
  ];

  for (
    let offset = 0;
    offset < 4;
    offset += 1
  ) {
    const columnIndex =
      context.payment2026StartColumn +
      offset;

    const rawValue =
      context.valueRow[columnIndex];

    const displayValue = cleanText_(
      context.displayRow[columnIndex]
    );

    const color =
      context.colorRow[columnIndex];

    const parsedValue =
      normaliseNumber_(rawValue);

    if (
      displayValue &&
      typeof rawValue === 'string' &&
      Number.isFinite(parsedValue)
    ) {
      textQuarters.push(
        quarterLabels[offset]
      );
    }

    if (
      Number.isFinite(parsedValue) &&
      parsedValue !== 0 &&
      !isGreenFont_(color) &&
      !isBlueFont_(color)
    ) {
      unknownColorQuarters.push(
        quarterLabels[offset]
      );
    }
  }

  if (textQuarters.length) {
    issues.push({
      row: row,
      severity: 'critical',
      title: 'Сумма записана текстом',
      description:
        'В строке «' + query +
        '» числовое значение в ' +
        textQuarters.join(', ') +
        ' сохранено как текст. ' +
        'Его нужно преобразовать в число.',
      query: query,
      scope: scope
    });
  }

  if (unknownColorQuarters.length) {
    issues.push({
      row: row,
      severity: 'critical',
      title: 'Платёж не классифицирован цветом',
      description:
        'В строке «' + query +
        '» значение в ' +
        unknownColorQuarters.join(', ') +
        ' имеет цвет, который не распознан ' +
        'ни как зелёный факт, ни как синий прогноз.',
      query: query,
      scope: scope
    });
  }
}


/**
 * Находит актуальный прогноз текущего 4 квартала.
 *
 * Основной источник — значение под подписью
 * «актуальное ожидание в 4 квартале»
 * в блоке 2026 года. Если структура блока изменится,
 * используется резервная ячейка M143.
 */
function readCurrentQuarterForecast_(
  sheet,
  values,
  displayValues,
  payment2026StartColumn
) {
  const label =
    'актуальное ожидание в 4 квартале';

  for (
    let rowIndex = 0;
    rowIndex < displayValues.length - 1;
    rowIndex += 1
  ) {
    for (
      let columnIndex =
        payment2026StartColumn;
      columnIndex <
        Math.min(
          displayValues[rowIndex].length,
          payment2026StartColumn + 4
        );
      columnIndex += 1
    ) {
      if (
        cleanText_(
          displayValues[rowIndex][columnIndex]
        ).toLowerCase() === label
      ) {
        const candidate =
          normaliseNumber_(
            values[rowIndex + 1][columnIndex]
          );

        if (Number.isFinite(candidate)) {
          return candidate;
        }
      }
    }
  }

  const direct = normaliseNumber_(
    sheet.getRange('M143').getValue()
  );

  return Number.isFinite(direct)
    ? direct
    : 0;
}


/**
 * Надёжно читает итог контрактации за 2025 год.
 *
 * Ожидаемая структура:
 *   2025
 *   Контр 1 кв | Контр 2 кв | Контр 3 кв | Контр 4 кв | Итого
 *   ...        | ...        | ...        | ...        | 91 127 814,59
 *
 * Между строкой года и заголовком допускаются пустые строки.
 */
function readContracting2025Total_(values) {
  for (
    let yearRow = 0;
    yearRow < values.length;
    yearRow += 1
  ) {
    const firstCell = cleanText_(
      values[yearRow] &&
      values[yearRow][0]
    );

    if (firstCell !== '2025') {
      continue;
    }

    /*
     * После строки с годом ищем ближайшую строку
     * заголовков, чтобы вставка пустых строк
     * не ломала чтение.
     */
    const headerLimit = Math.min(
      values.length,
      yearRow + 6
    );

    for (
      let headerRow = yearRow + 1;
      headerRow < headerLimit;
      headerRow += 1
    ) {
      const row = values[headerRow] || [];

      const firstHeader =
        cleanText_(row[0]).toLowerCase();

      if (
        firstHeader.indexOf('контр 1 кв') !== 0
      ) {
        continue;
      }

      let totalColumn = -1;

      for (
        let columnIndex = 0;
        columnIndex < row.length;
        columnIndex += 1
      ) {
        if (
          cleanText_(row[columnIndex])
            .toLowerCase() === 'итого'
        ) {
          totalColumn = columnIndex;
          break;
        }
      }

      if (totalColumn === -1) {
        continue;
      }

      /*
       * Ищем ближайшую следующую строку,
       * где в колонке «Итого» находится число.
       */
      const valueLimit = Math.min(
        values.length,
        headerRow + 5
      );

      for (
        let valueRow = headerRow + 1;
        valueRow < valueLimit;
        valueRow += 1
      ) {
        const candidate =
          normaliseNumber_(
            values[valueRow] &&
            values[valueRow][totalColumn]
          );

        if (
          Number.isFinite(candidate) &&
          candidate > 0
        ) {
          return candidate;
        }
      }
    }
  }

  throw new Error(
    'Не удалось найти итог контрактации 2025. ' +
    'Проверьте блок «2025 / Контр 1 кв ... Итого» ' +
    'на листе «Контрактация/закрытие 2026».'
  );
}


/**
 * Читает лист контрактации:
 * квартальные факты, планы, KPI
 * и перечень фактически заключённых договоров.
 */
function parseContractingSheet_(
  sheet,
  signedRecords
) {
  const values = sheet
    .getDataRange()
    .getValues();

  /*
   * Контрактация 2025 читается по структуре блока,
   * а не по фиксированной ячейке E86.
   *
   * Это защищает дашборд от вставки/удаления строк:
   * ищем строку с годом 2025, затем строку заголовков
   * «Контр 1 кв ... Итого» и берём значение из
   * следующей числовой строки в колонке «Итого».
   */
  const contracting2025Total =
    readContracting2025Total_(values);

  const contractingHeader =
    findRowStartingWith_(
      values,
      'Контрактация факт/план'
    );

  const receiptsHeader =
    findRowStartingWith_(
      values,
      'Приход факт/план'
    );

  const overallHeader =
    findRowStartingWith_(
      values,
      'Сводный KPI.'
    );

  const contractingRows =
    findNumericRowsAfter_(
      values,
      contractingHeader,
      2,
      5
    );

  const receiptsRows =
    findNumericRowsAfter_(
      values,
      receiptsHeader,
      2,
      5
    );

  const overallRows =
    findNumericRowsAfter_(
      values,
      overallHeader,
      1,
      5
    );

  const contractingActual =
    contractingRows[0];

  const contractingPlan =
    contractingRows[1];

  const receiptsActual =
    receiptsRows[0];

  const receiptsPlan =
    receiptsRows[1];

  const overallKpi =
    overallRows[0];

  const items = [];

  for (
    let rowIndex = 0;
    rowIndex < values.length;
    rowIndex += 1
  ) {
    const row = values[rowIndex];

    const quarter = Number(row[0]);
    const name = cleanText_(row[1]);
    const amount = normaliseNumber_(row[2]);
    const note = cleanText_(row[3])
      .toLowerCase();

    if (
      !Number.isInteger(quarter) ||
      quarter < 1 ||
      quarter > 4 ||
      !name ||
      !Number.isFinite(amount) ||
      amount <= 0
    ) {
      continue;
    }

    /*
     * Строки «в процессе» ещё не входят
     * в фактическую контрактацию.
     */
    if (
      note.indexOf('в процессе') !== -1
    ) {
      continue;
    }

    items.push({
      quarter: quarter,
      name: name,
      amount: amount,
      row: rowIndex + 1
    });
  }

  const unmatchedItems =
    assignContractingToRecords_(
      signedRecords,
      items
    );

  return {
    receipts: {
      quarters:
        receiptsActual.slice(0, 4),
      actualTotal:
        receiptsActual[4],
      planQuarters:
        receiptsPlan.slice(0, 4),
      planTotal:
        receiptsPlan[4],
      attainment:
        safeDivide_(
          receiptsActual[4],
          receiptsPlan[4]
        )
    },

    contracting: {
      quarters:
        contractingActual.slice(0, 4),
      actualTotal:
        contractingActual[4],
      planQuarters:
        contractingPlan.slice(0, 4),
      planTotal:
        contractingPlan[4],
      attainment:
        safeDivide_(
          contractingActual[4],
          contractingPlan[4]
        )
    },

    overallKpi: {
      quarters:
        overallKpi.slice(0, 4),
      total:
        overallKpi[4]
    },

    items: items,
    unmatchedItems: unmatchedItems,
    contracting2025Total:
      roundMoney_(contracting2025Total)
  };
}


/**
 * Привязывает строки контрактации
 * к конкретным договорам реестра.
 */
function assignContractingToRecords_(
  records,
  items
) {
  const assignedRecordIndexes = {};
  const unmatched = [];

  items.forEach(function (item) {
    const itemKey = normaliseMatchKey_(
      item.name
    );

    let bestIndex = -1;
    let bestScore = -1;

    records.forEach(function (
      record,
      recordIndex
    ) {
      if (
        assignedRecordIndexes[recordIndex]
      ) {
        return;
      }

      const nameKey = normaliseMatchKey_(
        record.name
      );

      const counterpartyKey =
        normaliseMatchKey_(
          record.counterparty
        );

      let score = -1;

      if (
        itemKey &&
        itemKey === nameKey
      ) {
        score = 100;
      } else if (
        itemKey &&
        itemKey === counterpartyKey
      ) {
        score = 90;
      } else if (
        isUsefulPartialMatch_(
          itemKey,
          nameKey
        )
      ) {
        score = 70;
      } else if (
        isUsefulPartialMatch_(
          itemKey,
          counterpartyKey
        )
      ) {
        score = 60;
      }

      if (score > bestScore) {
        bestScore = score;
        bestIndex = recordIndex;
      }
    });

    if (
      bestIndex === -1 ||
      bestScore < 0
    ) {
      unmatched.push({
        quarter: item.quarter,
        name: item.name,
        amount: item.amount,
        row: item.row
      });
      return;
    }

    records[bestIndex].contracting2026 =
      roundMoney_(
        numberOrZero_(
          records[bestIndex]
            .contracting2026
        ) +
        item.amount
      );

    assignedRecordIndexes[bestIndex] = true;
  });

  return unmatched;
}


/**
 * Формирует управленческие группы статусов.
 */
function buildStatusGroups_(records) {
  const definitions = [
    {
      name: 'В работе',
      statuses: ['Действует']
    },
    {
      name: 'Ожидание ИД',
      statuses: ['Ожидание ИД']
    },
    {
      name: 'На закрытии',
      statuses: ['Процесс закрытия']
    },
    {
      name: 'Закрыты',
      statuses: ['Выполнен и закрыт']
    },
    {
      name: 'Расторгнут',
      statuses: [
        'Расторжение',
        'Расторгнут',
        'Отказ'
      ]
    },
    {
      name: 'Приостановлен',
      statuses: ['Приостановлен']
    }
  ];

  const result = {};

  definitions.forEach(function (definition) {
    const rows = records.filter(
      function (record) {
        return (
          definition.statuses.indexOf(
            record.status
          ) !== -1
        );
      }
    );

    const aggregate = aggregateRecords_(
      rows
    );

    aggregate.statuses =
      definition.statuses.slice();

    aggregate.types = countTypes_(rows);

    result[definition.name] = aggregate;
  });

  return result;
}


/**
 * Итоги набора договоров.
 */
function aggregateRecords_(records) {
  let value = 0;
  let paid = 0;
  let balance = 0;
  let balanceKnown = 0;
  let balanceMissing = 0;

  records.forEach(function (record) {
    value += numberOrZero_(record.value);
    paid += numberOrZero_(record.paid);

    if (
      record.balance === null ||
      record.balance === undefined
    ) {
      balanceMissing += 1;
    } else {
      balance += numberOrZero_(
        record.balance
      );
      balanceKnown += 1;
    }
  });

  return {
    count: records.length,
    value: roundMoney_(value),
    paid: roundMoney_(paid),
    balance: roundMoney_(balance),
    balanceKnown: balanceKnown,
    balanceMissing: balanceMissing
  };
}


/**
 * Разбивка ТС / ПОДД / Другие.
 */
function countTypes_(records) {
  const result = {
    'ТС': 0,
    'ПОДД': 0,
    'Другие': 0
  };

  records.forEach(function (record) {
    const type = workTypeFromSection_(
      record.section
    );

    result[type] += 1;
  });

  return result;
}


/**
 * Читает проверенные квартальные данные 2025 года.
 */
function parseReceipts2025_(
  values,
  displayValues
) {
  let headerRow = -1;

  for (
    let rowIndex = 0;
    rowIndex < displayValues.length;
    rowIndex += 1
  ) {
    const row = displayValues[rowIndex];

    if (
      cleanText_(row[7]) ===
        'Поступило 1кв' &&
      cleanText_(row[12]) ===
        'Поступило 1кв'
    ) {
      headerRow = rowIndex;
      break;
    }
  }

  if (
    headerRow === -1 ||
    !values[headerRow + 1]
  ) {
    throw new Error(
      'Не найдены квартальные поступления 2025/2026 на листе «Статус».'
    );
  }

  const rawQuarters = values[
    headerRow + 1
  ]
    .slice(7, 11)
    .map(numberOrZero_);

  let transferTo2026 = 0;
  let annualPlan = 53000000;

  for (
    let rowIndex = 0;
    rowIndex < displayValues.length;
    rowIndex += 1
  ) {
    const row = displayValues[rowIndex];

    for (
      let columnIndex = 0;
      columnIndex < row.length;
      columnIndex += 1
    ) {
      const text = cleanText_(
        row[columnIndex]
      );

      if (
        text ===
        'Перераспределение в 2026'
      ) {
        transferTo2026 =
          firstNumericAfter_(
            values[rowIndex],
            columnIndex
          );
      }

      if (
        text.indexOf(
          'план на год 53'
        ) !== -1
      ) {
        const parsed =
          numberFromText_(text);

        if (parsed) {
          annualPlan = parsed;
        }
      }
    }
  }

  const adjustedQuarters =
    rawQuarters.slice();

  adjustedQuarters[3] =
    roundMoney_(
      adjustedQuarters[3] -
      transferTo2026
    );

  const adjustedTotal =
    adjustedQuarters.reduce(
      function (sum, value) {
        return sum + value;
      },
      0
    );

  return {
    rawQuarters:
      rawQuarters.map(roundMoney_),

    adjustedQuarters:
      adjustedQuarters.map(roundMoney_),

    transferTo2026:
      roundMoney_(transferTo2026),

    adjustedTotal:
      roundMoney_(adjustedTotal),

    annualPlan: annualPlan
  };
}


/**
 * Ищет колонку по точному заголовку.
 */
function findHeaderColumn_(
  headerRow,
  headerText
) {
  const target = cleanText_(headerText);

  for (
    let columnIndex = 0;
    columnIndex < headerRow.length;
    columnIndex += 1
  ) {
    if (
      cleanText_(
        headerRow[columnIndex]
      ) === target
    ) {
      return columnIndex;
    }
  }

  throw new Error(
    'Не найден заголовок: ' +
    headerText
  );
}


/**
 * Ищет колонку по началу заголовка.
 */
function findHeaderColumnStartsWith_(
  headerRow,
  headerText
) {
  const target = cleanText_(headerText);

  for (
    let columnIndex = 0;
    columnIndex < headerRow.length;
    columnIndex += 1
  ) {
    if (
      cleanText_(
        headerRow[columnIndex]
      ).indexOf(target) === 0
    ) {
      return columnIndex;
    }
  }

  throw new Error(
    'Не найден заголовок, начинающийся с: ' +
    headerText
  );
}


/**
 * Ищет подзаголовок рядом с группой «Договор».
 */
function findSubheaderColumn_(
  secondHeader,
  subheaderText,
  startColumn
) {
  const target = cleanText_(
    subheaderText
  ).toLowerCase();

  const finalColumn = Math.min(
    secondHeader.length,
    startColumn + 4
  );

  for (
    let columnIndex = startColumn;
    columnIndex < finalColumn;
    columnIndex += 1
  ) {
    if (
      cleanText_(
        secondHeader[columnIndex]
      ).toLowerCase() === target
    ) {
      return columnIndex;
    }
  }

  throw new Error(
    'Не найден подзаголовок: ' +
    subheaderText
  );
}


/**
 * Ищет строку, в которой какая-либо ячейка
 * начинается с нужного текста.
 */
function findRowStartingWith_(
  values,
  searchText
) {
  const target = cleanText_(
    searchText
  );

  for (
    let rowIndex = 0;
    rowIndex < values.length;
    rowIndex += 1
  ) {
    const row = values[rowIndex];

    for (
      let columnIndex = 0;
      columnIndex < row.length;
      columnIndex += 1
    ) {
      if (
        cleanText_(
          row[columnIndex]
        ).indexOf(target) === 0
      ) {
        return rowIndex;
      }
    }
  }

  throw new Error(
    'Не найдена строка: ' +
    searchText
  );
}


/**
 * Ищет нужное количество числовых строк
 * после найденного заголовка.
 */
function findNumericRowsAfter_(
  values,
  startRow,
  requiredRows,
  width
) {
  const result = [];

  const finalRow = Math.min(
    values.length,
    startRow + 14
  );

  for (
    let rowIndex = startRow + 1;
    rowIndex < finalRow;
    rowIndex += 1
  ) {
    const candidate = values[rowIndex]
      .slice(0, width)
      .map(normaliseNumber_);

    if (
      candidate.length < width ||
      !candidate.every(Number.isFinite)
    ) {
      continue;
    }

    result.push(candidate);

    if (
      result.length === requiredRows
    ) {
      return result;
    }
  }

  throw new Error(
    'После строки ' +
    (startRow + 1) +
    ' не найдены требуемые числовые данные.'
  );
}


/**
 * Сумма значений заданного цвета текста.
 */
function sumByFontType_(
  valueRow,
  colorRow,
  startColumn,
  width,
  fontType
) {
  let total = 0;

  for (
    let offset = 0;
    offset < width;
    offset += 1
  ) {
    const columnIndex =
      startColumn + offset;

    const color = colorRow[columnIndex];

    const matches =
      fontType === 'green'
        ? isGreenFont_(color)
        : isBlueFont_(color);

    if (matches) {
      total += numberOrZero_(
        valueRow[columnIndex]
      );
    }
  }

  return roundMoney_(total);
}


/**
 * Определяет зелёный фактический цвет.
 */
function isGreenFont_(hexColor) {
  const rgb = hexToRgb_(hexColor);

  if (!rgb) {
    return false;
  }

  return (
    rgb.g >= 120 &&
    rgb.g >= rgb.r + 35 &&
    rgb.g >= rgb.b + 20
  );
}


/**
 * Определяет синий прогнозный цвет.
 */
function isBlueFont_(hexColor) {
  const rgb = hexToRgb_(hexColor);

  if (!rgb) {
    return false;
  }

  return (
    rgb.b >= 150 &&
    rgb.b >= rgb.r + 60 &&
    rgb.b >= rgb.g + 40
  );
}


/**
 * Преобразует #RRGGBB в RGB.
 */
function hexToRgb_(hexColor) {
  const text = cleanText_(hexColor)
    .replace('#', '');

  if (!/^[0-9a-fA-F]{6}$/.test(text)) {
    return null;
  }

  return {
    r: parseInt(text.slice(0, 2), 16),
    g: parseInt(text.slice(2, 4), 16),
    b: parseInt(text.slice(4, 6), 16)
  };
}


/**
 * ID рабочей таблицы.
 */
function getSpreadsheetId_() {
  const spreadsheetId = PropertiesService
    .getScriptProperties()
    .getProperty('SPREADSHEET_ID');

  if (!spreadsheetId) {
    throw new Error(
      'Не заполнено свойство SPREADSHEET_ID.'
    );
  }

  return spreadsheetId.trim();
}


/**
 * Ищет название раздела в любой ячейке строки.
 *
 * В рабочей таблице названия разделов находятся в колонке A,
 * а данные договора начинаются с колонки B.
 */
function findSectionNameInRow_(row) {
  for (
    let columnIndex = 0;
    columnIndex < row.length;
    columnIndex += 1
  ) {
    const value = cleanText_(
      row[columnIndex]
    );

    if (
      DASHBOARD_SECTIONS.indexOf(value) !== -1
    ) {
      return value;
    }
  }

  return '';
}


/**
 * Тип работ для фильтров и карточек.
 */
function workTypeFromSection_(section) {
  if (section === 'Разработка ТС') {
    return 'ТС';
  }

  if (section === 'Разработка ПОДД') {
    return 'ПОДД';
  }

  return 'Другие';
}


/**
 * Подготавливает текст для сопоставления названий.
 */
function normaliseMatchKey_(value) {
  return cleanText_(value)
    .toLowerCase()
    .replace(/ё/g, 'е')
    .replace(/[«»„“”"'№]/g, '')
    .replace(
      /[^a-zа-я0-9]+/gi,
      ' '
    )
    .replace(/\s+/g, ' ')
    .trim();
}


/**
 * Безопасное частичное совпадение.
 */
function isUsefulPartialMatch_(
  first,
  second
) {
  if (
    !first ||
    !second ||
    first.length < 5 ||
    second.length < 5
  ) {
    return false;
  }

  return (
    first.indexOf(second) !== -1 ||
    second.indexOf(first) !== -1
  );
}


/**
 * Дата в формате YYYY-MM-DD.
 */
function dateToIso_(value) {
  if (
    Object.prototype.toString.call(value) ===
      '[object Date]' &&
    !isNaN(value.getTime())
  ) {
    return Utilities.formatDate(
      value,
      Session.getScriptTimeZone(),
      'yyyy-MM-dd'
    );
  }

  return '';
}


/**
 * Номер договора без лишнего .0.
 */
function contractToString_(value) {
  if (!hasValue_(value)) {
    return null;
  }

  if (
    typeof value === 'number' &&
    Number.isInteger(value)
  ) {
    return String(value);
  }

  return cleanText_(value);
}


/**
 * Число или ноль.
 */
function numberOrZero_(value) {
  const number = normaliseNumber_(value);

  return Number.isFinite(number)
    ? number
    : 0;
}


/**
 * Число или null.
 */
function nullableNumber_(value) {
  if (!hasValue_(value)) {
    return null;
  }

  const number = normaliseNumber_(value);

  return Number.isFinite(number)
    ? number
    : null;
}


/**
 * Универсальное преобразование числа.
 */
function normaliseNumber_(value) {
  if (typeof value === 'number') {
    return value;
  }

  if (!hasValue_(value)) {
    return NaN;
  }

  const prepared = String(value)
    .replace(/\s/g, '')
    .replace(',', '.');

  const result = Number(prepared);

  return Number.isFinite(result)
    ? result
    : NaN;
}


/**
 * Проверка непустого значения.
 */
function hasValue_(value) {
  return !(
    value === null ||
    value === undefined ||
    value === ''
  );
}


/**
 * Очищает текст.
 */
function cleanText_(value) {
  if (
    value === null ||
    value === undefined
  ) {
    return '';
  }

  return String(value)
    .replace(/\u00a0/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}


/**
 * Первое число справа от текстовой ячейки.
 */
function firstNumericAfter_(
  row,
  startColumn
) {
  for (
    let columnIndex = startColumn + 1;
    columnIndex < row.length;
    columnIndex += 1
  ) {
    const value = normaliseNumber_(
      row[columnIndex]
    );

    if (Number.isFinite(value)) {
      return value;
    }
  }

  return 0;
}


/**
 * Извлекает число из подписи.
 */
function numberFromText_(text) {
  const digits = cleanText_(text)
    .replace(/[^\d]/g, '');

  if (!digits) {
    return 0;
  }

  return Number(digits);
}


/**
 * Деление без Infinity и NaN.
 */
function safeDivide_(
  numerator,
  denominator
) {
  if (!denominator) {
    return 0;
  }

  return numerator / denominator;
}


/**
 * Округление денежных значений.
 */
function roundMoney_(value) {
  return Math.round(
    Number(value || 0) * 100
  ) / 100;
}


/**
 * Короткий отпечаток набора данных.
 */
function createDataVersion_(data) {
  const source = JSON.stringify(data);

  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.MD5,
    source,
    Utilities.Charset.UTF_8
  );

  return digest
    .map(function (byte) {
      const value = (
        byte < 0
          ? byte + 256
          : byte
      );

      return (
        '0' +
        value.toString(16)
      ).slice(-2);
    })
    .join('');
}
