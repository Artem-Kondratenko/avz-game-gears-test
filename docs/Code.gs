const SHEETS = {
  settings: 'Настройка',
  balance: 'Балансовая таблица',
  viz: 'Визуализация данных',
};

const STAT_LABELS = {
  tower_damage: 'Урон башен',
  ufo_damage: 'Урон НЛО',
  move_speed: 'Скорость НЛО',
  cargo_capacity: 'Вместимость ресурсов',
};

const BASE_STATS = {
  tower_damage: 100,
  ufo_damage: 100,
  move_speed: 100,
  cargo_capacity: 10,
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('UFO Calculator')
    .addItem('Создать демо-таблицу', 'setupUfoCalculator')
    .addItem('Обновить расчет', 'updateCalculator')
    .addToUi();
}

function onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== SHEETS.viz) return;
  const row = e.range.getRow();
  const col = e.range.getColumn();
  if (col === 2 && row >= 11 && row <= 13) {
    updateCalculator();
  }
}

function setupUfoCalculator() {
  const ss = SpreadsheetApp.getActive();
  const settings = recreateSheet_(ss, SHEETS.settings);
  const balance = recreateSheet_(ss, SHEETS.balance);
  const viz = recreateSheet_(ss, SHEETS.viz);

  buildSettingsSheet_(settings);
  buildBalanceSheet_(balance);
  buildVisualizationSheet_(viz, balance);
  updateCalculator();
}

function updateCalculator() {
  const ss = SpreadsheetApp.getActive();
  const viz = ss.getSheetByName(SHEETS.viz);
  const balance = ss.getSheetByName(SHEETS.balance);
  if (!viz || !balance) return;

  const selected = viz.getRange('B11:B13').getValues().flat().filter(Boolean);
  const items = readItems_(balance);
  const totals = {};
  Object.keys(STAT_LABELS).forEach((stat) => {
    totals[stat] = { flat: 0, percent: 0 };
  });

  selected.forEach((itemName) => {
    const item = items.find((row) => row.name === itemName);
    if (!item) return;
    item.stats.forEach((entry) => {
      if (!entry.stat || !totals[entry.stat]) return;
      if (entry.type === 'percent') {
        totals[entry.stat].percent += entry.value;
      } else {
        totals[entry.stat].flat += entry.value;
      }
    });
  });

  const rows = Object.keys(STAT_LABELS).map((stat) => {
    const base = BASE_STATS[stat];
    const flat = totals[stat].flat;
    const percent = totals[stat].percent;
    const finalValue = (base + flat) * (1 + percent);
    const growth = base === 0 ? 0 : (finalValue - base) / base;
    return [
      STAT_LABELS[stat],
      base,
      flat,
      percent,
      finalValue,
      growth,
    ];
  });

  viz.getRange('A18:F21').setValues(rows);
  viz.getRange('B18:C21').setNumberFormat('0.0');
  viz.getRange('D18:D21').setNumberFormat('0.0%');
  viz.getRange('E18:E21').setNumberFormat('0.0');
  viz.getRange('F18:F21').setNumberFormat('0.0%');

  const summary = buildSummary_(selected, items);
  viz.getRange('A24:F27').clearContent();
  if (summary.length) {
    viz.getRange(24, 1, summary.length, 6).setValues(summary);
  }
}

function recreateSheet_(ss, name) {
  const existing = ss.getSheetByName(name);
  if (existing) ss.deleteSheet(existing);
  return ss.insertSheet(name);
}

function buildSettingsSheet_(sheet) {
  sheet.clear();
  sheet.getRange('A1:D1').merge().setValue('Слоты экипировки НЛО');
  sheet.getRange('A2:D2').setValues([['slot_id', 'Название', 'Роль', 'Комментарий']]);
  sheet.getRange('A3:D5').setValues([
    ['engine', 'Двигатель', 'Мобильность НЛО', 'Скорость сбора ресурсов и активного участия в бою'],
    ['steering', 'Руль', 'Боевой контур', 'Стабилизация наведения и боевой эффективности без изменения управления'],
    ['cargo', 'Грузовой модуль', 'Экономика боя', 'Вместимость ресурсов для строительства и апгрейда башен'],
  ]);

  sheet.getRange('F1:I1').merge().setValue('Редкости');
  sheet.getRange('F2:I2').setValues([['rarity_id', 'Название', 'Ранг', 'Множитель бюджета']]);
  sheet.getRange('F3:I6').setValues([
    ['common', 'Обычная', 1, 1.00],
    ['uncommon', 'Необычная', 2, 1.45],
    ['rare', 'Редкая', 3, 2.10],
    ['epic', 'Эпичная', 4, 3.00],
  ]);

  sheet.getRange('A8:D8').setValues([['stat_id', 'Название', 'Типовое отображение', 'Зачем нужен']]);
  sheet.getRange('A9:D12').setValues([
    ['tower_damage', 'Урон башен', '% / flat', 'Усиливает защитную часть core loop'],
    ['ufo_damage', 'Урон НЛО', '% / flat', 'Делает НЛО более активным участником боя'],
    ['move_speed', 'Скорость НЛО', '%', 'Ускоряет сбор ресурсов и реакцию на угрозы'],
    ['cargo_capacity', 'Вместимость ресурсов', 'flat / %', 'Позволяет дольше строить и апгрейдить башни без возврата к сбору'],
  ]);

  styleReferenceSheet_(sheet);
}

function buildBalanceSheet_(sheet) {
  const headers = [
    'item_id', 'Название', 'slot_id', 'Слот', 'rarity_id', 'Редкость',
    'Стат 1', 'Тип 1', 'Значение 1',
    'Стат 2', 'Тип 2', 'Значение 2',
    'Стат 3', 'Тип 3', 'Значение 3',
    'Стат 4', 'Тип 4', 'Значение 4',
    'Комментарий баланса',
  ];

  const rows = [
    ['engine_common', 'Пульсарный двигатель', 'engine', 'Двигатель', 'common', 'Обычная', 'move_speed', 'percent', 0.04, 'ufo_damage', 'flat', 5, '', '', '', '', '', '', 'Базовая мобильность и небольшой вклад НЛО в бой'],
    ['engine_uncommon', 'Пульсарный двигатель+', 'engine', 'Двигатель', 'uncommon', 'Необычная', 'move_speed', 'percent', 0.07, 'ufo_damage', 'flat', 8, '', '', '', '', '', '', 'Ускоряет сбор ресурсов и раннюю активность НЛО'],
    ['engine_rare', 'Пульсарный двигатель Mk II', 'engine', 'Двигатель', 'rare', 'Редкая', 'move_speed', 'percent', 0.10, 'ufo_damage', 'percent', 0.05, 'tower_damage', 'percent', 0.03, '', '', '', 'Гибрид мобильности и боевой эффективности'],
    ['engine_epic', 'Пульсарный двигатель Омега', 'engine', 'Двигатель', 'epic', 'Эпичная', 'move_speed', 'percent', 0.15, 'ufo_damage', 'percent', 0.09, 'tower_damage', 'percent', 0.05, 'cargo_capacity', 'percent', 0.05, 'Сильная мобильность без превращения предмета в чистый DPS'],
    ['steering_common', 'Боевой руль Азимут', 'steering', 'Руль', 'common', 'Обычная', 'tower_damage', 'flat', 6, 'ufo_damage', 'flat', 4, '', '', '', '', '', '', 'Небольшой универсальный прирост боевой эффективности'],
    ['steering_uncommon', 'Боевой руль Азимут+', 'steering', 'Руль', 'uncommon', 'Необычная', 'tower_damage', 'percent', 0.05, 'ufo_damage', 'flat', 7, '', '', '', '', '', '', 'Руль трактуется как стабилизация наведения, а не изменение управления'],
    ['steering_rare', 'Боевой руль Азимут Mk II', 'steering', 'Руль', 'rare', 'Редкая', 'tower_damage', 'percent', 0.08, 'ufo_damage', 'percent', 0.05, '', '', '', '', '', '', 'Хорош для сбалансированной боевой сборки'],
    ['steering_epic', 'Боевой руль Азимут Омега', 'steering', 'Руль', 'epic', 'Эпичная', 'tower_damage', 'percent', 0.12, 'ufo_damage', 'percent', 0.08, 'move_speed', 'percent', 0.04, '', '', '', 'Топовый боевой слот с небольшим бонусом к темпу боя'],
    ['cargo_common', 'Компрессионный трюм', 'cargo', 'Грузовой модуль', 'common', 'Обычная', 'cargo_capacity', 'flat', 2, 'move_speed', 'percent', -0.01, '', '', '', '', '', '', 'Ресурсная емкость с маленькой ценой в мобильности'],
    ['cargo_uncommon', 'Компрессионный трюм+', 'cargo', 'Грузовой модуль', 'uncommon', 'Необычная', 'cargo_capacity', 'flat', 4, 'tower_damage', 'flat', 5, '', '', '', '', '', '', 'Помогает быстрее ставить оборону и слегка усиливает башни'],
    ['cargo_rare', 'Компрессионный трюм Mk II', 'cargo', 'Грузовой модуль', 'rare', 'Редкая', 'cargo_capacity', 'percent', 0.20, 'tower_damage', 'percent', 0.05, 'ufo_damage', 'flat', 6, '', '', '', 'Экономика боя плюс умеренный урон'],
    ['cargo_epic', 'Компрессионный трюм Омега', 'cargo', 'Грузовой модуль', 'epic', 'Эпичная', 'cargo_capacity', 'percent', 0.35, 'tower_damage', 'percent', 0.08, 'ufo_damage', 'percent', 0.05, 'move_speed', 'percent', -0.02, 'Мощная экономика с небольшой компенсацией через вес модуля'],
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  sheet.getRange('I2:I13').setNumberFormat('0.0%');
  sheet.getRange('L2:L13').setNumberFormat('0.0');
  sheet.getRange('O2:O13').setNumberFormat('0.0');
  sheet.getRange('R2:R13').setNumberFormat('0.0');
  styleTableSheet_(sheet, headers.length);
}

function buildVisualizationSheet_(sheet, balance) {
  sheet.clear();
  sheet.getRange('A1:F1').merge().setValue('Калькулятор экипировки НЛО');
  sheet.getRange('A2:F2').merge().setValue('Выберите предметы в трех слотах и оцените итоговый прирост характеристик.');

  sheet.getRange('A4:B4').setValues([['Базовая характеристика', 'Значение']]);
  sheet.getRange('A5:B8').setValues([
    ['Урон башен', BASE_STATS.tower_damage],
    ['Урон НЛО', BASE_STATS.ufo_damage],
    ['Скорость НЛО', BASE_STATS.move_speed],
    ['Вместимость ресурсов', BASE_STATS.cargo_capacity],
  ]);

  sheet.getRange('A10:F10').setValues([['Слот', 'Выбранный предмет', 'Заметка', '', '', '']]);
  sheet.getRange('A11:A13').setValues([['Двигатель'], ['Руль'], ['Грузовой модуль']]);
  sheet.getRange('C11:C13').setValues([
    ['Мобильность и участие НЛО в бою'],
    ['Боевой контур без изменения управления'],
    ['Экономика строительства башен'],
  ]);

  const itemNames = balance.getRange('B2:B13');
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(itemNames, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('B11:B13').setDataValidation(rule);
  sheet.getRange('B11:B13').setValues([
    ['Пульсарный двигатель Mk II'],
    ['Боевой руль Азимут Mk II'],
    ['Компрессионный трюм Mk II'],
  ]);

  sheet.getRange('A17:F17').setValues([['Стат', 'База', 'Flat бонус', '% бонус', 'Итог', 'Прирост']]);
  sheet.getRange('A23:F23').setValues([['Выбранный предмет', 'Слот', 'Редкость', 'Ключевые бонусы', 'Комментарий', '']]);

  styleVisualizationSheet_(sheet);
}

function readItems_(sheet) {
  const values = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 0), 19).getValues();
  return values
    .filter((row) => row[0])
    .map((row) => ({
      id: row[0],
      name: row[1],
      slot: row[3],
      rarity: row[5],
      comment: row[18],
      stats: [
        { stat: row[6], type: row[7], value: Number(row[8]) || 0 },
        { stat: row[9], type: row[10], value: Number(row[11]) || 0 },
        { stat: row[12], type: row[13], value: Number(row[14]) || 0 },
        { stat: row[15], type: row[16], value: Number(row[17]) || 0 },
      ],
    }));
}

function buildSummary_(selected, items) {
  return selected.map((itemName) => {
    const item = items.find((row) => row.name === itemName);
    if (!item) return [itemName, '', '', '', '', ''];
    const bonuses = item.stats
      .filter((entry) => entry.stat)
      .map((entry) => {
        const label = STAT_LABELS[entry.stat] || entry.stat;
        const value = entry.type === 'percent'
          ? Utilities.formatString('%s%%', Math.round(entry.value * 1000) / 10)
          : Utilities.formatString('+%s', entry.value);
        return `${label}: ${value}`;
      })
      .join('; ');
    return [item.name, item.slot, item.rarity, bonuses, item.comment, ''];
  });
}

function styleReferenceSheet_(sheet) {
  sheet.setFrozenRows(2);
  sheet.getRange('A1:D1').setBackground('#0b5cad').setFontColor('#ffffff').setFontWeight('bold');
  sheet.getRange('F1:I1').setBackground('#0b5cad').setFontColor('#ffffff').setFontWeight('bold');
  sheet.getRange('A2:D2').setBackground('#d9ecff').setFontWeight('bold');
  sheet.getRange('F2:I2').setBackground('#d9ecff').setFontWeight('bold');
  sheet.getRange('A8:D8').setBackground('#d9ecff').setFontWeight('bold');
  sheet.autoResizeColumns(1, 9);
}

function styleTableSheet_(sheet, width) {
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, width).setBackground('#0b5cad').setFontColor('#ffffff').setFontWeight('bold');
  sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 1), width).setBackground('#f7fbff');
  sheet.autoResizeColumns(1, width);
  sheet.getRange(1, 1, sheet.getLastRow(), width).createFilter();
}

function styleVisualizationSheet_(sheet) {
  sheet.setColumnWidths(1, 6, 150);
  sheet.setColumnWidth(2, 230);
  sheet.setColumnWidth(4, 260);
  sheet.setColumnWidth(5, 330);
  sheet.getRange('A1:F1').setBackground('#08345f').setFontColor('#ffffff').setFontWeight('bold').setFontSize(16);
  sheet.getRange('A2:F2').setBackground('#d9ecff').setFontColor('#16324f');
  sheet.getRange('A4:B4').setBackground('#0b5cad').setFontColor('#ffffff').setFontWeight('bold');
  sheet.getRange('A10:F10').setBackground('#0b5cad').setFontColor('#ffffff').setFontWeight('bold');
  sheet.getRange('A17:F17').setBackground('#0b5cad').setFontColor('#ffffff').setFontWeight('bold');
  sheet.getRange('A23:F23').setBackground('#0b5cad').setFontColor('#ffffff').setFontWeight('bold');
  sheet.getRange('A5:B8').setBackground('#f7fbff');
  sheet.getRange('A11:C13').setBackground('#f7fbff');
  sheet.getRange('A18:F21').setBackground('#f7fbff');
  sheet.getRange('A24:F27').setBackground('#f7fbff').setWrap(true);
}
