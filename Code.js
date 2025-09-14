/**
 * @file Code.js
 * @description Основной файл скрипта Google Apps Script для проекта gostanalysis_gas.
 * @version 2.1
 * @build 2.2.0 
 * @author ИИ-агент Ядро
 */

// --- Глобальная конфигурация ---
const GEMINI_API_URL = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent';

/**
 * @description Функция для обработки GET-запросов к веб-приложению.
 * Она загружает и отображает пользовательский интерфейс HTML, определенный в файле 'index.html'.
 * @param {GoogleAppsScript.Events.AppsScriptHttpRequestEvent} e - Объект события GET-запроса, содержащий параметры запроса.
 * @returns {GoogleAppsScript.HTML.HtmlOutput} HTML-содержимое для отображения в браузере.
 * @module {UI}
 * @entrypoint
 */
function doGet(e) {
  Logger.log('doGet вызван.');
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Вова-Стандарт: Анализ и Консультации');
}

/**
 * @description Получает имя используемой модели Gemini из глобальной конфигурации URL.
 * Это позволяет динамически отображать текущую модель, используемую для API-взаимодействий.
 * @returns {string} Имя модели Gemini (например, 'gemini-2.5-flash') или 'unknown'/'error' в случае ошибки.
 * @module {API_Integration}
 */
function getModelNameGAS() {
  try {
    const match = GEMINI_API_URL.match(/models\/(.*?):/);
    const modelName = match && match[1] ? match[1] : 'unknown';
    Logger.log('Запрошено имя модели: ' + modelName);
    return modelName;
  } catch (e) {
    Logger.log('Ошибка при получении имени модели: ' + e.message);
    return 'error';
  }
}

/**
 * @description Получает значение долговременной памяти из PropertiesService скрипта.
 * Эта функция используется для извлечения пользовательских инструкций или контекста,
 * который должен быть сохранен между сессиями или запросами.
 * @returns {string} Строка, содержащая долговременную память. Возвращает пустую строку, если память не установлена.
 * @module {DataStorage}
 */
function getMemoryProperty() {
  Logger.log('getMemoryProperty вызван.');
  return PropertiesService.getScriptProperties().getProperty('LONG_TERM_MEMORY') || '';
}

/**
 * @description Устанавливает значение долговременной памяти в PropertiesService скрипта.
 * Эта функция позволяет пользователям сохранять важную информацию или инструкции,
 * которые будут использоваться в последующих взаимодействиях с ИИ-экспертом.
 * @param {string} memory - Строка долговременной памяти для сохранения.
 * @module {DataStorage}
 */
function setMemoryProperty(memory) {
  Logger.log('setMemoryProperty вызван.');
  PropertiesService.getScriptProperties().setProperty('LONG_TERM_MEMORY', memory);
}

/**
 * @description Обрабатывает сообщение чата с экспертом по стандартизации "Вова-Стандарт" через Gemini API.
 * Функция включает в себя системные инструкции, контекст анализа и долговременную память
 * для формирования релевантного и профессионального ответа.
 * @param {Object[]} messageHistory - История диалога в формате, ожидаемом Gemini API. Каждый объект должен содержать поле `parts` с текстом.
 * @param {string} analysisContext - Контекст из последнего анализа стандартов, который ИИ должен учитывать. Может быть пустой строкой.
 * @returns {string} Ответ от модели Gemini в виде текстовой строки.
 * @throws {Error} Если API-ключ Gemini не установлен или произошла ошибка при вызове API.
 * @module {API_Integration}
 */
function chatWithExpertGAS(messageHistory, analysisContext) {
  Logger.log('chatWithExpertGAS вызван. История сообщений: ' + messageHistory.length);
  const GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!GEMINI_API_KEY) throw new Error('API-ключ Gemini не установлен.');

  const longTermMemory = getMemoryProperty();
  const expertSystemInstruction = `
# РОЛЬ И ЗАДАЧА
Ты — «Вова-Стандарт», ИИ-эксперт мирового класса в области международной и национальной стандартизации (ISO, IEC, EN, ГОСТ, ДСТУ и др.). Твоя основная задача — предоставлять пользователю исчерпывающие, точные и профессиональные консультации.

# СТИЛЬ И ТОН
- **Тон:** Вежливый, деловой, но при этом проактивный и готовый помочь.
- **Структура ответа:** Используй Markdown для форматирования. Ключевые моменты выделяй жирным шрифтом, списки — для перечислений. Ответы должны быть хорошо структурированы и легко читаемы.
- **Полнота важнее краткости:** Сначала дай полный и точный ответ, и только потом стремись к сжатости. Не упускай важные детали ради краткости.

# ПРАВИЛА ПОВЕДЕНИЯ
1.  **Точность и ссылки:** Всегда ссылайся на конкретные стандарты по их полному обозначению (например, "ГОСТ Р ИСО 9001-2015"). Если возможно, указывай конкретные пункты или разделы стандарта.
2.  **Обработка неясностей:** Если запрос пользователя неоднозначен (например, "ГОСТ 12345" без года), НЕ ПРЕДПОЛАГАЙ. Задай уточняющий вопрос. Пример: "Уточните, пожалуйста, год стандарта ГОСТ 12345, так как существует несколько версий."
3.  **Признание ограничений:** Если ты не знаешь ответа или не уверен в его точности, честно сообщи об этом. Пример: "Я не могу найти точную информацию по вашему запросу. Рекомендую обратиться к официальному тексту стандарта."

# СПЕЦИАЛЬНЫЕ КОМАНДЫ
- **[ИНСТРУКЦИЯ ПО ЗАПОМИНАНИЮ]:** Если пользователь просит тебя что-то запомнить (используя фразы "запомни", "запиши в память", "нужно помнить" и т.п.), твоим ЕДИНСТВЕННЫМ ответом должен быть JSON-объект. Не добавляй никакого текста до или после него. JSON должен иметь строго следующую структуру: {"action": "propose_memory_update", "data": "сформулированная_суть_для_запоминания"}.

${analysisContext ? `
# АКТУАЛЬНЫЙ КОНТЕКСТ АНАЛИЗА
Пользователь только что проанализировал следующие стандарты. Учитывай эту информацию при ответах.
${analysisContext}` : ''}

${longTermMemory ? `
# ДОЛГОВРЕМЕННАЯ ПАМЯТЬ (ВЫСШИЙ ПРИОРИТЕТ)
Следуй этим инструкциям пользователя в первую очередь.
${longTermMemory}` : ''}
`;

  const payload = {
    contents: messageHistory,
    systemInstruction: { parts: [{ text: expertSystemInstruction }] },
    generationConfig: {
      temperature: 0.1,
      topK: 20,
      topP: 0.8,
      maxOutputTokens: 8192,
    },
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };

  const response = UrlFetchApp.fetch(GEMINI_API_URL + '?key=' + GEMINI_API_KEY, options);
  const responseCode = response.getResponseCode();
  const responseBody = response.getContentText();

  if (responseCode === 200) {
    const jsonResponse = JSON.parse(responseBody);
    if (jsonResponse.candidates && jsonResponse.candidates[0].content && jsonResponse.candidates[0].content.parts && jsonResponse.candidates[0].content.parts[0].text) {
      return jsonResponse.candidates[0].content.parts[0].text;
    }
    // Handle cases where the response might be blocked or empty
    Logger.log('Ответ от Gemini API не содержит текста. Finish reason: ' + (jsonResponse.candidates ? jsonResponse.candidates[0].finishReason : 'N/A'));
    return 'Не удалось получить ответ от ИИ. Попробуйте переформулировать запрос.';
  } else {
    throw new Error(`Ошибка API чата: Код ${responseCode}. Ответ: ${responseBody}`);
  }
}

/**
 * @description Анализирует список стандартов с использованием Gemini API в режиме JSON, предоставляя информацию о их существовании, статусе и полных наименованиях.
 * Результаты анализа могут включать дополнительные столбцы на основе пользовательских настроек.
 * @param {string[]} standardsList - Массив наименований стандартов, которые необходимо проанализировать (например, ['ГОСТ Р 52289-2004']).
 * @param {string} selectedCountry - Выбранная страна, для которой проводится анализ стандартов (например, 'Россия', 'Казахстан').
 * @param {Object} optionalColumns - Объект с флагами, указывающими, какие дополнительные столбцы (например, `replacedBy`, `sources`) должны быть включены в результат.
 * @returns {Object[]} - Массив объектов, где каждый объект представляет результат анализа одного стандарта и соответствует заданной структуре JSON.
 * @throws {Error} Если API-ключ Gemini не установлен или произошла ошибка при вызове API.
 * @module {API_Integration}
 */
function analyzeStandardsGAS(standardsList, selectedCountry, optionalColumns) {
  Logger.log('analyzeStandardsGAS вызван с JSON режимом.');
  const GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!GEMINI_API_KEY) {
    throw new Error('API-ключ Gemini не установлен.');
  }

  const longTermMemory = getMemoryProperty();
  const analysisBaseInstruction = 'Вы — экспертный ИИ-помощник, специализирующийся на стандартизации. Ваша задача — проанализировать список стандартов для указанной страны. Для каждого стандарта определите его существование, полное наименование и текущий статус. Предоставьте точную и краткую информацию. Ваш ответ должен быть ТОЛЬКО JSON-массивом, соответствующим предоставленной схеме. Не добавляйте текст до или после JSON.';
  
  const memoryBlock = longTermMemory.trim()
      ? `[ДОЛГОВРЕМЕННАЯ ПАМЯТЬ ПОЛЬЗОВАТЕЛЯ - ЭТИ ИНСТРУКЦИИ ИМЕЮТ ВЫСШИЙ ПРИОРИТЕТ]\n${longTermMemory}\n[/ДОЛГОВРЕМЕННАЯ ПАМЯТЬ ПОЛЬЗОВАТЕЛЯ]\n\n`
      : '';
  const systemInstruction = `${memoryBlock}${analysisBaseInstruction}`;

  const prompt = `Проанализируй следующие стандарты для страны "${selectedCountry}".
  Список стандартов:\n${standardsList.join('\n')}`;

  // --- Формирование JSON-схемы для ответа ---
  const schemaProperties = {
    requestedDesignation: { type: 'STRING', description: 'Обозначение стандарта, как его запросил пользователь.' },
    exists: { type: 'STRING', description: 'Существует ли стандарт ("Да" или "Нет").' },
    fullName: { type: 'STRING', description: 'Полное официальное наименование стандарта.' },
    status: { type: 'STRING', description: 'Текущий статус (Действующий, Отменен, Заменен и т.д.).' },
    aiNote: { type: 'STRING', description: 'Краткое примечание от ИИ по стандарту (до 100 символов).' }
  };

  if (optionalColumns.replacedBy) {
    schemaProperties.replacedBy = { type: 'STRING', description: 'Обозначение стандарта, на который был произведен замен.' };
  }
  if (optionalColumns.sources) {
    schemaProperties.sources = {
      type: 'ARRAY',
      description: 'Список URL-адресов или названий документов, подтверждающих информацию.',
      items: { type: 'STRING' }
    };
  }
  
  const responseSchema = {
    type: 'ARRAY',
    items: {
      type: 'OBJECT',
      properties: schemaProperties
    }
  };
  // --- Конец формирования схемы ---

  const url = GEMINI_API_URL + '?key=' + GEMINI_API_KEY;
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({
      contents: [{ parts: [{ text: prompt }] }],
      systemInstruction: { parts: [{ text: systemInstruction }] },
      generationConfig: {
        temperature: 0.1,
        topK: 40,
        topP: 0.95,
        maxOutputTokens: 8192,
        responseMimeType: 'application/json',
        responseSchema: responseSchema,
      },
    }),
    muteHttpExceptions: true,
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();
    Logger.log('Gemini API Response Code: ' + responseCode);

    if (responseCode === 200) {
      const jsonResponse = JSON.parse(responseBody);
      try {
        logApiUsage(
          standardsList.length, 
          selectedCountry, 
          jsonResponse.usageMetadata || {}, 
          (jsonResponse.candidates && jsonResponse.candidates[0]) ? jsonResponse.candidates[0].finishReason : 'UNKNOWN'
        );
      } catch (logError) {
          Logger.log('Ошибка при логировании: ' + logError.message);
      }

      if (!jsonResponse.candidates || !jsonResponse.candidates[0].content) {
        throw new Error('Некорректный ответ от API. Возможно, сработали фильтры безопасности. Причина: ' + (jsonResponse.candidates ? jsonResponse.candidates[0].finishReason : 'N/A'));
      }

      const textContent = jsonResponse.candidates[0].content.parts[0].text;
      return JSON.parse(textContent); // Ответ уже должен быть валидным JSON
    } else {
      throw new Error(`Ошибка API анализа: Код ${responseCode}. Ответ: ${responseBody}`);
    }
  } catch (e) {
    Logger.log('Ошибка при вызове Gemini API: ' + e.message);
    throw e;
  }
}

/**
 * @description Записывает данные об использовании Gemini API (количество токенов, стандартов) в Google Sheet для мониторинга и отладки.
 * @param {number} standardsCount - Количество анализируемых стандартов в текущем запросе.
 * @param {string} selectedCountry - Выбранная страна, для которой проводился анализ.
 * @param {object} usageMetadata - Объект с метаданными об использовании токенов от API (promptTokenCount, candidatesTokenCount, totalTokenCount).
 * @param {string} finishReason - Причина завершения генерации ответа моделью (например, 'STOP', 'SAFETY', 'LENGTH').
 * @module {Logging}
 */
function logApiUsage(standardsCount, selectedCountry, usageMetadata, finishReason) {
  try {
    const LOG_SHEET_ID = PropertiesService.getScriptProperties().getProperty('LOG_SHEET_ID');
    if (!LOG_SHEET_ID) {
      Logger.log('ID таблицы для логирования (LOG_SHEET_ID) не установлен. Логирование пропускается.');
      return;
    }

    const spreadsheet = SpreadsheetApp.openById(LOG_SHEET_ID);
    const sheet = spreadsheet.getSheets()[0]; 
    
    if (!sheet) {
        Logger.log('Лист для логирования не найден в таблице с ID: ' + LOG_SHEET_ID);
        return;
    }
    
    const timestamp = new Date();
    const promptTokens = usageMetadata.promptTokenCount || 0;
    const outputTokens = usageMetadata.candidatesTokenCount || 0;
    const totalTokens = usageMetadata.totalTokenCount || 0;

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Timestamp', 'Standards Count', 'Country', 'Prompt Tokens', 'Output Tokens', 'Total Tokens', 'Finish Reason']);
    }

    sheet.appendRow([
      timestamp,
      standardsCount,
      selectedCountry,
      promptTokens,
      outputTokens,
      totalTokens,
      finishReason || 'N/A'
    ]);
    Logger.log('Вызов API успешно залогирован.');

  } catch (e) {
    Logger.log('ОШИБКА при логировании вызова API: ' + e.message);
  }
}