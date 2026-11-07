const APP = {
  TIMEZONE: 'Asia/Tokyo',
  SHEETS: {
    CONFIG: 'config',
    CATALOG: 'catalog',
    PLANS: 'plans',
    LOGS: 'logs',
    WEAKNESS: 'weakness',
    EVENTS: 'line_events',
    USERS: 'users'
  },
  SECRET_KEYS: {
    LINE_TOKEN: 'LINE_CHANNEL_ACCESS_TOKEN',
    OPENAI_KEY: 'OPENAI_API_KEY',
    RELAY_SECRET: 'VERCEL_RELAY_SECRET'
  }
};

let CURRENT_USER_ID_ = '';

// =====================================================
// 0. 初回セットアップ
// =====================================================

function setupV3Schema() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    ensureSheetHeaders_(APP.SHEETS.PLANS, [
      '日付',
      'plan_id',
      '科目',
      '教材',
      '範囲',
      '予定分',
      '目的',
      '完了',
      '完了率%',
      '実績分',
      '理解度1-5',
      'メモ',
      '翌日繰越',
      'catalog_id',
      '作成区分',
      'fixed',
      'report_received',
      'user_id'
    ]);

    ensureSheetHeaders_(APP.SHEETS.LOGS, [
      'timestamp',
      'date',
      'plan_id',
      'catalog_id',
      '科目',
      '教材',
      '範囲',
      '予定分',
      '実績分',
      '完了率%',
      '理解度1-5',
      'つまずき',
      'コメント',
      '調整',
      'user_id'
    ]);

    ensureSheetHeaders_(APP.SHEETS.WEAKNESS, [
      'weakness_id',
      'catalog_id',
      '科目',
      'テーマ',
      '理解度',
      '原因/メモ',
      '登録日',
      '次回復習日',
      '復習間隔日',
      '状態',
      '最終更新',
      'user_id'
    ]);

    ensureSheetHeaders_(APP.SHEETS.EVENTS, [
      'timestamp',
      'event_type',
      'webhook_event_id',
      'user_id',
      'message_text',
      'reply_token',
      'status',
      'parsed_json',
      'note'
    ]);

    ensureSheetHeaders_(APP.SHEETS.USERS, [
      'registration_code',
      'display_name',
      'user_id',
      'role',
      'active',
      'weekday_target_min',
      'weekend_target_min',
      'registered_at',
      'note',
      'last_sent_date'
    ]);

    ensureCatalogHeaders_();

    ensureConfigRows_({
      BOT_TITLE: '簿記3級 AI担任',
      REPORT_DRIVEN_MODE: 'ON',
      RESEND_PENDING_MENU: 'ON',
      WEEKDAY_TARGET_MIN: '60',
      WEEKEND_TARGET_MIN: '120',
      PLAN_RULES_WEEKDAY: '簿記3級:2',
      PLAN_RULES_WEEKEND: '簿記3級:4',
      FILL_SUBJECTS_WEEKEND: '簿記3級',
      OPENAI_MODEL: 'gpt-4o-mini',
      MORNING_HOUR: '7',
      MORNING_MINUTE: '0',
      AI_ADJUSTMENT: 'ON',
      WEAKNESS_REVIEW: 'ON'
    });

    backfillPlanIds_();
    normalizePlanDefaults_();
    ensureV21DisplayFormats_();

    SpreadsheetApp.getActive().setSpreadsheetTimeZone(APP.TIMEZONE);
    clearConfigCache();

    return 'v3 multi-user setup complete';
  } finally {
    lock.releaseLock();
  }
}

function setupV2Schema() {
  return setupV3Schema();
}

function ensureCatalogHeaders_() {
  const sh = getSheet_(APP.SHEETS.CATALOG);
  const required = [
    'id',
    '科目',
    '教材',
    '範囲',
    '目安分',
    '目的',
    '確認方法',
    '優先度',
    '状態'
  ];
  ensureHeadersOnSheet_(sh, required);
}

function ensureSheetHeaders_(sheetName, headers) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(sheetName);
  if (!sh) sh = ss.insertSheet(sheetName);
  ensureHeadersOnSheet_(sh, headers);
}

function ensureHeadersOnSheet_(sh, required) {
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const existing = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0]
    .map(v => String(v || '').trim());

  const missing = required.filter(h => !existing.includes(h));
  if (!missing.length) return;

  const startCol = existing.filter(Boolean).length + 1;
  sh.getRange(1, startCol, 1, missing.length).setValues([missing]);
  sh.setFrozenRows(1);
}

function ensureConfigRows_(defaults) {
  const sh = getSheet_(APP.SHEETS.CONFIG);
  ensureHeadersOnSheet_(sh, ['key', 'value', 'description']);

  const table = readTable_(APP.SHEETS.CONFIG);
  const existing = {};
  table.rows.forEach(r => {
    const key = String(r.values[table.col('key')] || '').trim();
    if (key) existing[key] = true;
  });

  const descriptions = {
    BOT_TITLE: 'LINEに表示する学習Bot名',
    REPORT_DRIVEN_MODE: '進捗報告を基準に翌日計画を作成',
    RESEND_PENDING_MENU: '報告がない場合は同じメニューを再送',
    WEEKDAY_TARGET_MIN: '平日の学習目標時間（分）',
    WEEKEND_TARGET_MIN: '休日の学習目標時間（分）',
    PLAN_RULES_WEEKDAY: '平日の科目:件数ルール',
    PLAN_RULES_WEEKEND: '休日の科目:件数ルール',
    FILL_SUBJECTS_WEEKEND: '休日に余った時間を埋める科目',
    OPENAI_MODEL: '進捗解析に使うモデル',
    MORNING_HOUR: '朝の送信時刻（時）',
    MORNING_MINUTE: '朝の送信時刻（分、±15分程度）',
    AI_ADJUSTMENT: 'AIによる進捗解析をON/OFF',
    WEAKNESS_REVIEW: 'weaknessの復習を計画に再投入'
  };

  const rows = [];
  Object.keys(defaults).forEach(key => {
    if (!existing[key]) rows.push([key, defaults[key], descriptions[key] || '']);
  });

  if (rows.length) {
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, 3).setValues(rows);
  }
}

function migrateOwnerUserIdToScriptProperties_() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty(APP.SECRET_KEYS.OWNER_USER_ID)) return;

  const legacy = getConfig_('LINE_USER_ID');
  if (legacy) props.setProperty(APP.SECRET_KEYS.OWNER_USER_ID, legacy);
}

function backfillPlanIds_() {
  const table = readTable_(APP.SHEETS.PLANS);
  const planIdCol = table.col('plan_id');

  table.rows.forEach((r, idx) => {
    if (String(r.values[planIdCol] || '').trim()) return;

    const date = normalizeDate_(r.values[table.col('日付')]);
    if (!date) return;

    const id = makePlanId_(date, r.rowNumber);
    getSheet_(APP.SHEETS.PLANS).getRange(r.rowNumber, planIdCol + 1).setValue(id);
  });
}

function normalizePlanDefaults_() {
  const table = readTable_(APP.SHEETS.PLANS);
  const sh = getSheet_(APP.SHEETS.PLANS);

  table.rows.forEach(r => {
    const values = r.values.slice();

    if (values[table.col('完了率%')] === '') values[table.col('完了率%')] = 0;
    if (values[table.col('完了')] === '') values[table.col('完了')] = false;
    if (values[table.col('fixed')] === '') values[table.col('fixed')] = false;
    if (values[table.col('report_received')] === '') values[table.col('report_received')] = false;

    sh.getRange(r.rowNumber, 1, 1, table.headers.length).setValues([values]);
  });
}


function ensureV21DisplayFormats_() {
  const eventSh = getSheet_(APP.SHEETS.EVENTS);
  const logSh = getSheet_(APP.SHEETS.LOGS);

  if (eventSh.getMaxRows() >= 2) {
    eventSh.getRange(2, 1, eventSh.getMaxRows() - 1, 1)
      .setNumberFormat('yyyy/MM/dd HH:mm:ss');
  }

  if (logSh.getMaxRows() >= 2) {
    logSh.getRange(2, 1, logSh.getMaxRows() - 1, 1)
      .setNumberFormat('yyyy/MM/dd HH:mm:ss');
  }
}

// =====================================================
// 1. VercelからのWebhook受信
// =====================================================

function doGet() {
  return ContentService.createTextOutput('ok');
}

function doPost(e) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(30000);

    const envelope = JSON.parse((e.postData && e.postData.contents) || '{}');
    const expectedRelaySecret = getSecret_(APP.SECRET_KEYS.RELAY_SECRET);

    if (!expectedRelaySecret) {
      return ContentService.createTextOutput('ERROR: VERCEL_RELAY_SECRET missing');
    }

    if (!safeTextEqual_(String(envelope.relay_secret || ''), expectedRelaySecret)) {
      return ContentService.createTextOutput('UNAUTHORIZED');
    }

    const rawLineBody = String(envelope.line_raw_body || '');
    if (!rawLineBody) return ContentService.createTextOutput('ERROR: body missing');

    const body = JSON.parse(rawLineBody);
    const events = body.events || [];

    for (const event of events) {
      processLineEvent_(event);
    }

    return ContentService.createTextOutput('OK');
  } catch (err) {
    try {
      upsertLineEvent_({
        eventType: 'system_error',
        webhookEventId: '',
        userId: '',
        messageText: '',
        replyToken: '',
        status: 'error',
        parsedJson: '',
        note: String(err)
      });
    } catch (_) {}

    return ContentService.createTextOutput('ERROR');
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

function processLineEvent_(event) {
  const webhookEventId = String(event.webhookEventId || '').trim();

  if (webhookEventId && isWebhookEventProcessed_(webhookEventId)) return;

  const userId = event.source && event.source.userId
    ? String(event.source.userId)
    : '';

  if (
    event.type !== 'message' ||
    !event.message ||
    event.message.type !== 'text'
  ) {
    upsertLineEvent_({
      eventType: event.type || '',
      webhookEventId,
      userId,
      messageText: '',
      replyToken: event.replyToken || '',
      status: 'processed',
      parsedJson: '',
      note: 'text message以外'
    });
    return;
  }

  const text = String(event.message.text || '').trim();
  const replyToken = String(event.replyToken || '');

  upsertLineEvent_({
    eventType: 'message',
    webhookEventId,
    userId,
    messageText: text,
    replyToken,
    status: 'received',
    parsedJson: '',
    note: ''
  });

  try {
    const user = findUserByLineUserId_(userId);

    // 未登録ユーザーは、usersシートのregistration_codeで登録する。
    if (!user) {
      const registered = registerUserByCode_(userId, text);

      if (!registered) {
        upsertLineEvent_({
          eventType: 'message',
          webhookEventId,
          userId,
          messageText: '[UNREGISTERED]',
          replyToken,
          status: 'processed',
          parsedJson: '',
          note: 'unregistered_user'
        });

        sendLineResponse_(
          replyToken,
          userId,
          'この簿記3級AI担任は登録制です。管理者から案内された登録コードを、そのまま送ってください。'
        );
        return;
      }

      upsertLineEvent_({
        eventType: 'message',
        webhookEventId,
        userId,
        messageText: '[REGISTER]',
        replyToken,
        status: 'processed',
        parsedJson: '',
        note: 'user_registered'
      });

      sendLineResponse_(
        replyToken,
        userId,
        '✅ 簿記3級 AI担任への登録が完了したよ。明日から学習メニューを送るね。すぐ試したい場合は「今日のメニュー」と送ってね。'
      );
      return;
    }

    if (!bool_(user.obj['active'])) {
      upsertLineEvent_({
        eventType: 'message',
        webhookEventId,
        userId,
        messageText: '[INACTIVE]',
        replyToken,
        status: 'processed',
        parsedJson: '',
        note: 'inactive_user'
      });

      sendLineResponse_(replyToken, userId, '現在、このアカウントの利用は停止中です。管理者に確認してください。');
      return;
    }

    setCurrentUserId_(userId);

    if (/^(今日のメニュー|メニュー|今日の予定)$/i.test(text)) {
      ensureTodayPlan_(today_());
      const tasks = getPlanRowsByDate_(today_());
      upsertLineEvent_({
        eventType: 'message',
        webhookEventId,
        userId,
        messageText: text,
        replyToken,
        status: 'processed',
        parsedJson: '',
        note: 'menu_requested'
      });
      sendLineResponse_(replyToken, userId, buildPlanMessage_(tasks, ''));
      return;
    }

    upsertLineEvent_({
      eventType: 'message',
      webhookEventId,
      userId,
      messageText: text,
      replyToken,
      status: 'processing',
      parsedJson: '',
      note: ''
    });

    const result = handleStudyMessage_(text);

    upsertLineEvent_({
      eventType: 'message',
      webhookEventId,
      userId,
      messageText: text,
      replyToken,
      status: 'processed',
      parsedJson: JSON.stringify(result.parsed || {}),
      note: result.note || ''
    });
    SpreadsheetApp.flush();

    sendLineResponse_(replyToken, userId, result.replyText);
  } catch (err) {
    upsertLineEvent_({
      eventType: 'message',
      webhookEventId,
      userId,
      messageText: text,
      replyToken,
      status: 'error',
      parsedJson: '',
      note: String(err)
    });

    sendLineResponse_(
      replyToken,
      userId,
      'メッセージは受け取ったよ。ただ、処理中にエラーが出たので計画は変更していないよ。管理者に確認してください。'
    );
  } finally {
    clearCurrentUserId_();
  }
}

function safeTextEqual_(a, b) {
  if (a.length !== b.length) return false;
  let diff = 0;
  for (let i = 0; i < a.length; i++) diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return diff === 0;
}

// =====================================================
// 2. 朝の送信
// =====================================================

function sendTodayPlan() {
  sendTodayPlanInternal_(false);
}

function sendTodayPlanForce() {
  sendTodayPlanInternal_(true);
}

function sendTodayPlanInternal_(force) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(30000);

    const today = today_();
    const users = getActiveRegisteredUsers_();

    users.forEach(user => {
      const userId = String(user.obj['user_id'] || '').trim();
      if (!userId) return;

      try {
        if (!force && normalizeDate_(user.obj['last_sent_date']) === today) return;

        setCurrentUserId_(userId);
        ensureTodayPlan_(today);

        const tasks = getPlanRowsByDate_(today);
        if (!tasks.length) throw new Error('今日のplansを作成できませんでした。');

        const latestPreviousDate = findLatestPlanDateBefore_(today);
        let note = '';

        if (
          latestPreviousDate &&
          !planDateHasReport_(latestPreviousDate) &&
          tasks.some(t => String(t.obj['作成区分'] || '') === '未報告再送')
        ) {
          note = '前回の学習報告がまだないので、同じメニューを再送するね。新しい範囲へはまだ進めていないよ。';
        }

        pushLine_(userId, buildPlanMessage_(tasks, note));
        setUserLastSentDate_(user.rowNumber, today);
      } catch (userErr) {
        logSystemErrorForUser_(userId, '朝のメニュー送信でエラー: ' + userErr);
      } finally {
        clearCurrentUserId_();
      }
    });
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

function ensureTodayPlan_(today) {
  if (getPlanRowsByDate_(today).length) return;

  const latestDate = findLatestPlanDateBefore_(today);

  if (!latestDate) {
    generateInitialPlanForDate_(today);
    return;
  }

  if (!planDateHasReport_(latestDate)) {
    if (getConfig_('RESEND_PENDING_MENU') === 'ON') {
      copyUnreportedPlan_(latestDate, today);
    }
    return;
  }

  generateAdaptivePlanForDate_(today, latestDate);
}

function buildPlanMessage_(tasks, note) {
  const title = getConfig_('BOT_TITLE') || 'AI担任';
  let total = 0;
  const lines = ['📚 ' + title + '｜今日の学習メニュー'];

  if (note) lines.push('', note);
  lines.push('');

  tasks.forEach((t, idx) => {
    const o = t.obj;
    total += number_(o['予定分']);

    lines.push(
      '【' + (idx + 1) + '】' + String(o['科目'] || ''),
      '教材：' + String(o['教材'] || ''),
      '範囲：' + String(o['範囲'] || ''),
      '目安：' + number_(o['予定分']) + '分',
      '目的：' + String(o['目的'] || ''),
      ''
    );
  });

  lines.push(
    '合計：約' + total + '分',
    '',
    '終わったらLINEでそのまま報告してね。',
    '例：1番と2番完了。3番は半分。理解度2。DNA複製が難しかった。',
    '追加で計画外の範囲を進めた場合も、その内容を書けば記録するよ。'
  );

  return lines.join('\n');
}

// =====================================================
// 3. LINE報告 → 累積更新 → 翌日再計画
// =====================================================

function handleStudyMessage_(messageText) {
  if (getConfig_('AI_ADJUSTMENT') !== 'ON') {
    return {
      parsed: {},
      note: 'AI_ADJUSTMENT_OFF',
      replyText: '進捗は受け取ったよ。AI_ADJUSTMENTがOFFなので、まだ翌日計画は変更していないよ。'
    };
  }

  const apiKey = getSecret_(APP.SECRET_KEYS.OPENAI_KEY);
  if (!apiKey) throw new Error('OPENAI_API_KEY がScript Propertiesにありません。');

  const today = today_();
  const planDate = findLatestPlanDateOnOrBefore_(today);

  if (!planDate) {
    return {
      parsed: {},
      note: 'no_plan',
      replyText: '学習報告は受け取ったけど、対応する学習メニューが見つからなかったよ。'
    };
  }

  const tasks = getPlanRowsByDate_(planDate);
  const catalog = readCatalogForPrompt_();
  const parsed = analyzeStudyReport_(messageText, planDate, tasks, catalog);

  if (!parsed.is_study_report) {
    return {
      parsed,
      note: 'not_study_report',
      replyText: 'メッセージは受け取ったよ。学習報告とは判定しなかったので、計画は変更していないよ。'
    };
  }

  const applyResult = applyStudyReport_(planDate, tasks, parsed, messageText);

  markPlanDateReported_(planDate);

  const tomorrow = nextDate_(today, 1);
  generateAdaptivePlanForDate_(tomorrow, planDate);

  const tomorrowTasks = getPlanRowsByDate_(tomorrow);

  return {
    parsed,
    note: 'report_applied',
    replyText: buildReportReply_(parsed, applyResult, tomorrowTasks)
  };
}

function analyzeStudyReport_(messageText, planDate, tasks, catalogText) {
  const schema = {
    type: 'object',
    additionalProperties: false,
    properties: {
      is_study_report: { type: 'boolean' },
      summary: { type: 'string' },
      task_updates: {
        type: 'array',
        items: {
          type: 'object',
          additionalProperties: false,
          properties: {
            plan_id: { type: 'string' },
            task_no: { type: 'integer' },
            reported: { type: 'boolean' },
            evidence: { type: 'string' },
            completion_rate: { type: 'integer' },
            actual_minutes: { type: 'integer' },
            understanding: { type: 'integer' },
            note: { type: 'string' }
          },
          required: [
            'plan_id',
            'task_no',
            'reported',
            'evidence',
            'completion_rate',
            'actual_minutes',
            'understanding',
            'note'
          ]
        }
      },
      extra_progress: {
        type: 'array',
        items: {
          type: 'object',
          additionalProperties: false,
          properties: {
            reported: { type: 'boolean' },
            evidence: { type: 'string' },
            catalog_id: { type: 'string' },
            subject: { type: 'string' },
            material: { type: 'string' },
            range: { type: 'string' },
            completion_rate: { type: 'integer' },
            actual_minutes: { type: 'integer' },
            understanding: { type: 'integer' },
            note: { type: 'string' }
          },
          required: [
            'reported',
            'evidence',
            'catalog_id',
            'subject',
            'material',
            'range',
            'completion_rate',
            'actual_minutes',
            'understanding',
            'note'
          ]
        }
      },
      weaknesses: {
        type: 'array',
        items: {
          type: 'object',
          additionalProperties: false,
          properties: {
            reported: { type: 'boolean' },
            evidence: { type: 'string' },
            task_no: { type: 'integer' },
            catalog_id: { type: 'string' },
            subject: { type: 'string' },
            topic: { type: 'string' },
            understanding: { type: 'integer' },
            note: { type: 'string' }
          },
          required: [
            'reported',
            'evidence',
            'task_no',
            'catalog_id',
            'subject',
            'topic',
            'understanding',
            'note'
          ]
        }
      }
    },
    required: [
      'is_study_report',
      'summary',
      'task_updates',
      'extra_progress',
      'weaknesses'
    ]
  };

  const taskText = tasks.map((t, i) => {
    const o = t.obj;
    return [
      'task_no=' + (i + 1),
      'plan_id=' + String(o['plan_id'] || ''),
      'catalog_id=' + String(o['catalog_id'] || ''),
      '科目=' + String(o['科目'] || ''),
      '教材=' + String(o['教材'] || ''),
      '範囲=' + String(o['範囲'] || ''),
      '予定分=' + number_(o['予定分']),
      '現在完了率=' + number_(o['完了率%']),
      '現在理解度=' + number_(o['理解度1-5'])
    ].join(', ');
  }).join('\n');

  const systemText = [
    'あなたは汎用の学習管理アシスタントです。',
    'ユーザーの自由文の学習報告を、提示された当日の予定と照合して構造化してください。',
    '最重要ルール：今回のメッセージで明示的に言及されたタスクだけ task_updates に含めてください。',
    '未言及タスクを0%・未着手などとして補完することは禁止です。',
    'task_updates の reported は、今回のメッセージで明示されたものだけ true にしてください。',
    'evidence には、今回のユーザーメッセージから、そのタスクを報告したことが分かる短い文字列を「原文そのまま」で入れてください。要約・言い換えは禁止です。',
    '例：「2番も終わった」と書かれていれば、task_no=2 の evidence は「2番も終わった」とできます。',
    '「1番不安」のように番号が付いた感想・不安・理解度は、その番号のタスクに必ず紐づけてください。',
    '番号が付いていない「理解度2」等は、文脈上もっとも近い明示済みタスクにだけ紐づけ、曖昧ならunderstanding=0にしてください。',
    '同じ日に複数回報告するため、前回までの進捗を0に戻す推測は禁止です。',
    'completion_rate は0〜100、understandingは不明なら0、actual_minutesは不明なら-1にしてください。',
    '「半分」は50程度としてよいですが、明示された割合を優先してください。',
    '計画外に進めた範囲は extra_progress に入れ、reported=true、evidenceは必ず原文そのままにしてください。',
    'catalogに一致しそうならcatalog_idを入れ、不明なら空文字にしてください。',
    'weaknessesも今回のメッセージに根拠がある場合だけ出してください。',
    'weaknesses の evidence もユーザー原文そのまま、task_no は対応する予定番号。不明なら0にしてください。',
    '理解度2以下、または「難しい・分からない・苦手・不安」等があればweakness候補です。',
    '学習報告ではないメッセージなら is_study_report=false とし、配列は空にしてください。',
    '書かれていない内容を断定しないでください。'
  ].join('\n');

  const userText = [
    '対象日: ' + planDate,
    '',
    '当日の予定:',
    taskText || 'なし',
    '',
    'catalog抜粋:',
    catalogText || 'なし',
    '',
    '今回の報告:',
    messageText
  ].join('\n');

  return callOpenAIStructured_(systemText, userText, schema, 'study_report_v21');
}


function callOpenAIStructured_(systemText, userText, schema, formatName) {
  const apiKey = getSecret_(APP.SECRET_KEYS.OPENAI_KEY);
  const model = getConfig_('OPENAI_MODEL') || 'gpt-4o-mini';

  const payload = {
    model,
    store: false,
    input: [
      {
        role: 'system',
        content: [{ type: 'input_text', text: systemText }]
      },
      {
        role: 'user',
        content: [{ type: 'input_text', text: userText }]
      }
    ],
    text: {
      format: {
        type: 'json_schema',
        name: formatName,
        strict: true,
        schema
      }
    }
  };

  let lastError = null;

  for (let attempt = 0; attempt < 2; attempt++) {
    const res = UrlFetchApp.fetch('https://api.openai.com/v1/responses', {
      method: 'post',
      contentType: 'application/json',
      headers: {
        Authorization: 'Bearer ' + apiKey
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const code = res.getResponseCode();
    const text = res.getContentText();

    if (code >= 200 && code < 300) {
      const body = JSON.parse(text);
      const outputText = extractResponseText_(body);
      if (!outputText) throw new Error('OpenAI response text is empty');
      return JSON.parse(outputText);
    }

    lastError = new Error('OpenAI API error: ' + code + ' ' + text);

    if (![429, 500, 502, 503, 504].includes(code) || attempt === 1) break;
    Utilities.sleep(1500 * (attempt + 1));
  }

  throw lastError || new Error('OpenAI API error');
}

function extractResponseText_(body) {
  let out = '';

  (body.output || []).forEach(item => {
    if (item.type !== 'message') return;

    (item.content || []).forEach(part => {
      if (part.type === 'output_text') out += String(part.text || '');
      if (part.type === 'refusal') throw new Error('OpenAI refusal: ' + String(part.refusal || ''));
    });
  });

  return out;
}

// =====================================================
// 4. 進捗の累積適用・logs upsert
// =====================================================

function applyStudyReport_(planDate, tasks, parsed, originalText) {
  let updated = 0;
  let extra = 0;
  let ignored = 0;

  const byPlanId = {};
  const byTaskNo = {};

  tasks.forEach((t, i) => {
    byPlanId[String(t.obj['plan_id'] || '')] = t;
    byTaskNo[i + 1] = t;
  });

  (parsed.task_updates || []).forEach(update => {
    if (update.reported !== true || !hasGroundedEvidence_(originalText, update.evidence)) {
      ignored++;
      return;
    }

    let target = null;

    const planId = String(update.plan_id || '').trim();
    if (planId && byPlanId[planId]) target = byPlanId[planId];

    if (!target) {
      const no = clampInt_(update.task_no, 1, 9999);
      if (byTaskNo[no]) target = byTaskNo[no];
    }

    if (!target) {
      ignored++;
      return;
    }

    applyOnePlanUpdate_(target, update, originalText);
    updated++;
  });

  (parsed.extra_progress || []).forEach(ep => {
    if (ep.reported !== true || !hasGroundedEvidence_(originalText, ep.evidence)) {
      ignored++;
      return;
    }

    const row = createOrFindExtraProgressPlan_(planDate, ep);
    applyOnePlanUpdate_(row, ep, originalText);
    extra++;
  });

  (parsed.weaknesses || []).forEach(w => {
    if (w.reported !== true || !hasGroundedEvidence_(originalText, w.evidence)) {
      ignored++;
      return;
    }

    const taskNo = clampInt_(w.task_no, 0, 9999);
    if (taskNo > 0 && byTaskNo[taskNo]) {
      if (!String(w.catalog_id || '').trim()) {
        w.catalog_id = String(byTaskNo[taskNo].obj['catalog_id'] || '').trim();
      }
      if (!String(w.subject || '').trim()) {
        w.subject = String(byTaskNo[taskNo].obj['科目'] || '').trim();
      }
      if (!String(w.topic || '').trim()) {
        w.topic = String(byTaskNo[taskNo].obj['範囲'] || '').trim();
      }
    }

    upsertWeakness_(w);
  });

  return { updated, extra, ignored };
}

function hasGroundedEvidence_(originalText, evidence) {
  const source = normalizeEvidenceText_(originalText);
  const quote = normalizeEvidenceText_(evidence);

  if (!quote || quote.length < 2) return false;
  return source.indexOf(quote) !== -1;
}

function normalizeEvidenceText_(text) {
  return String(text || '')
    .normalize('NFKC')
    .replace(/\s+/g, '')
    .replace(/[。．.、,！!？?]/g, '')
    .trim();
}


function applyOnePlanUpdate_(planRow, update, originalText) {
  const table = readTable_(APP.SHEETS.PLANS);
  const sh = getSheet_(APP.SHEETS.PLANS);

  const current = planRow.obj;
  const oldRate = clampInt_(current['完了率%'], 0, 100);
  const incomingRate = clampInt_(update.completion_rate, 0, 100);
  const newRate = Math.max(oldRate, incomingRate);

  const oldActual = number_(current['実績分']);
  const incomingActual = number_(update.actual_minutes);
  const newActual = incomingActual >= 0 ? Math.max(oldActual, incomingActual) : oldActual;

  const oldUnderstanding = clampInt_(current['理解度1-5'], 0, 5);
  const incomingUnderstanding = clampInt_(update.understanding, 0, 5);
  const newUnderstanding = incomingUnderstanding > 0 ? incomingUnderstanding : oldUnderstanding;

  const oldMemo = String(current['メモ'] || '').trim();
  const note = String(update.note || '').trim();
  const evidence = String(update.evidence || '').trim();
  const memoText = note || evidence;
  const newMemo = memoText ? appendMemo_(oldMemo, memoText) : oldMemo;

  const newValues = planRow.values.slice();
  newValues[table.col('完了')] = newRate >= 100;
  newValues[table.col('完了率%')] = newRate;
  newValues[table.col('実績分')] = newActual;
  newValues[table.col('理解度1-5')] = newUnderstanding;
  newValues[table.col('メモ')] = newMemo;
  newValues[table.col('翌日繰越')] = newRate < 100 ? 'YES' : 'NO';
  newValues[table.col('report_received')] = true;

  sh.getRange(planRow.rowNumber, 1, 1, table.headers.length).setValues([newValues]);

  const refreshed = rowToObject_(table.headers, newValues);

  upsertStudyLog_({
    timestamp: new Date(),
    date: normalizeDate_(refreshed['日付']),
    plan_id: refreshed['plan_id'],
    catalog_id: refreshed['catalog_id'],
    科目: refreshed['科目'],
    教材: refreshed['教材'],
    範囲: refreshed['範囲'],
    予定分: number_(refreshed['予定分']),
    実績分: newActual,
    '完了率%': newRate,
    '理解度1-5': newUnderstanding,
    つまずき: newUnderstanding > 0 && newUnderstanding <= 2 ? note : '',
    コメント: newMemo,
    調整: newRate < 100 ? '翌日へ繰越' : '次へ進む'
  });

  updateCatalogProgress_(refreshed['catalog_id'], newRate);

  if (
    String(refreshed['作成区分'] || '') === '弱点復習' &&
    newRate >= 100
  ) {
    advanceWeaknessReview_(refreshed['catalog_id'], newUnderstanding);
  }
}

function createOrFindExtraProgressPlan_(planDate, ep) {
  const catalogId = String(ep.catalog_id || '').trim();
  const catalogRow = catalogId ? findCatalogById_(catalogId) : findCatalogByText_(ep);

  const effectiveCatalogId = catalogRow ? String(catalogRow.obj['id'] || '') : catalogId;
  const existing = getPlanRowsByDate_(planDate).find(r => {
    if (effectiveCatalogId && String(r.obj['catalog_id'] || '') === effectiveCatalogId) return true;
    return (
      String(r.obj['教材'] || '') === String(ep.material || '') &&
      String(r.obj['範囲'] || '') === String(ep.range || '')
    );
  });

  if (existing) return existing;

  const subject = catalogRow ? catalogRow.obj['科目'] : ep.subject;
  const material = catalogRow ? catalogRow.obj['教材'] : ep.material;
  const range = catalogRow ? catalogRow.obj['範囲'] : ep.range;
  const planned = catalogRow ? number_(catalogRow.obj['目安分']) : Math.max(number_(ep.actual_minutes), 0);
  const purpose = catalogRow ? catalogRow.obj['目的'] : '計画外で進めた学習を記録';

  appendPlanRow_(planDate, {
    科目: subject,
    教材: material,
    範囲: range,
    予定分: planned,
    目的: purpose,
    catalog_id: effectiveCatalogId,
    作成区分: '計画外進捗',
    fixed: true
  });

  const rows = getPlanRowsByDate_(planDate);
  return rows[rows.length - 1];
}

// =====================================================
// 5. 翌日計画
// =====================================================

function generateInitialPlanForDate_(dateStr) {
  removeRegeneratableRowsForDate_(dateStr);
  buildPlanForDate_(dateStr, null);
}

function generateAdaptivePlanForDate_(dateStr, sourceDate) {
  removeRegeneratableRowsForDate_(dateStr);
  buildPlanForDate_(dateStr, sourceDate);
}

function buildPlanForDate_(dateStr, sourceDate) {
  const budget = getBudgetForDate_(dateStr);
  let used = getFixedMinutesForDate_(dateStr);

  if (sourceDate) {
    used += addIncompleteCarryovers_(sourceDate, dateStr, budget - used);
  }

  if (getConfig_('WEAKNESS_REVIEW') === 'ON') {
    used += addDueWeaknessReviews_(dateStr, budget - used);
  }

  const rules = parsePlanRules_(
    isWeekend_(dateStr)
      ? getConfig_('PLAN_RULES_WEEKEND')
      : getConfig_('PLAN_RULES_WEEKDAY')
  );

  for (const rule of rules) {
    for (let i = 0; i < rule.count; i++) {
      const added = addNextCatalogItem_(dateStr, rule.subject, budget - used, '自動計画');
      if (!added) break;
      used += added;
    }
  }

  if (isWeekend_(dateStr)) {
    const fillSubjects = String(getConfig_('FILL_SUBJECTS_WEEKEND') || '')
      .split(',')
      .map(v => v.trim())
      .filter(Boolean);

    let guard = 0;
    while (used < budget && guard < 50) {
      guard++;
      let addedAny = false;

      for (const subject of fillSubjects) {
        const added = addNextCatalogItem_(dateStr, subject, budget - used, '自動計画');
        if (added) {
          used += added;
          addedAny = true;
        }
        if (used >= budget) break;
      }

      if (!addedAny) break;
    }
  }
}

function addIncompleteCarryovers_(sourceDate, targetDate, remainingBudget) {
  let used = 0;

  getPlanRowsByDate_(sourceDate).forEach(row => {
    const o = row.obj;
    const rate = clampInt_(o['完了率%'], 0, 100);

    if (rate >= 100) return;
    if (remainingBudget - used <= 0) return;

    const planned = number_(o['予定分']);
    const remaining = Math.max(15, Math.round(planned * (100 - rate) / 100));
    const minutes = Math.min(remaining, remainingBudget - used);

    if (minutes <= 0) return;

    appendPlanRow_(targetDate, {
      科目: o['科目'],
      教材: o['教材'],
      範囲: cleanCarryoverRange_(o['範囲']),
      予定分: minutes,
      目的: rate > 0 ? '前回の残りを完了して定着させる' : '前回未着手のため同じ範囲を優先する',
      catalog_id: o['catalog_id'],
      作成区分: '未完了繰越',
      fixed: false
    });

    used += minutes;
  });

  return used;
}

function copyUnreportedPlan_(sourceDate, targetDate) {
  removeRegeneratableRowsForDate_(targetDate);

  getPlanRowsByDate_(sourceDate).forEach(row => {
    const o = row.obj;

    appendPlanRow_(targetDate, {
      科目: o['科目'],
      教材: o['教材'],
      範囲: cleanCarryoverRange_(o['範囲']),
      予定分: number_(o['予定分']),
      目的: '前回の報告がないため同じメニューを継続',
      catalog_id: o['catalog_id'],
      作成区分: '未報告再送',
      fixed: false
    });
  });
}

function addNextCatalogItem_(dateStr, subject, remainingBudget, source) {
  if (remainingBudget <= 0) return 0;

  const candidates = getCatalogCandidates_(subject);
  const usedCatalogIds = getUsedCatalogIds_(dateStr);

  for (const c of candidates) {
    const id = String(c.obj['id'] || '').trim();
    const minutes = number_(c.obj['目安分']);

    if (!id || usedCatalogIds[id]) continue;
    if (minutes <= 0 || minutes > remainingBudget) continue;

    appendPlanRow_(dateStr, {
      科目: c.obj['科目'],
      教材: c.obj['教材'],
      範囲: c.obj['範囲'],
      予定分: minutes,
      目的: c.obj['目的'],
      catalog_id: id,
      作成区分: source,
      fixed: false
    });

    return minutes;
  }

  return 0;
}

function addDueWeaknessReviews_(dateStr, remainingBudget) {
  if (remainingBudget <= 0) return 0;

  const table = readTable_(APP.SHEETS.WEAKNESS);
  let used = 0;

  const userId = getCurrentUserId_();

  table.rows.forEach(r => {
    if (remainingBudget - used <= 0) return;

    const o = r.obj;
    if (String(o['user_id'] || '').trim() !== userId) return;
    if (String(o['状態'] || '') === '克服') return;

    const due = normalizeDate_(o['次回復習日']);
    if (!due || due > dateStr) return;

    const catalog = findCatalogById_(o['catalog_id']);
    const minutes = Math.min(20, remainingBudget - used);

    if (minutes <= 0) return;

    appendPlanRow_(dateStr, {
      科目: o['科目'] || (catalog ? catalog.obj['科目'] : '復習'),
      教材: catalog ? catalog.obj['教材'] : '弱点復習',
      範囲: o['テーマ'],
      予定分: minutes,
      目的: '弱点を間隔反復で再確認する',
      catalog_id: o['catalog_id'],
      作成区分: '弱点復習',
      fixed: false
    });

    used += minutes;
  });

  return used;
}

function removeRegeneratableRowsForDate_(dateStr) {
  const table = readTable_(APP.SHEETS.PLANS);
  const sh = getSheet_(APP.SHEETS.PLANS);
  const rowsToDelete = [];

  table.rows.forEach(r => {
    const o = r.obj;
    if (normalizeDate_(o['日付']) !== dateStr) return;
    if (String(o['user_id'] || '').trim() !== getCurrentUserId_()) return;

    const fixed = bool_(o['fixed']);
    if (!fixed) rowsToDelete.push(r.rowNumber);
  });

  rowsToDelete.sort((a, b) => b - a).forEach(rowNum => sh.deleteRow(rowNum));
}

function appendPlanRow_(dateStr, data) {
  const table = readTable_(APP.SHEETS.PLANS);
  const values = new Array(table.headers.length).fill('');

  setByHeader_(values, table, '日付', dateStr);
  setByHeader_(values, table, 'plan_id', makePlanId_(dateStr, new Date().getTime() + Math.random()));
  setByHeader_(values, table, '科目', data.科目 || '');
  setByHeader_(values, table, '教材', data.教材 || '');
  setByHeader_(values, table, '範囲', data.範囲 || '');
  setByHeader_(values, table, '予定分', number_(data.予定分));
  setByHeader_(values, table, '目的', data.目的 || '');
  setByHeader_(values, table, '完了', false);
  setByHeader_(values, table, '完了率%', 0);
  setByHeader_(values, table, '実績分', 0);
  setByHeader_(values, table, '理解度1-5', 0);
  setByHeader_(values, table, 'メモ', '');
  setByHeader_(values, table, '翌日繰越', '');
  setByHeader_(values, table, 'catalog_id', data.catalog_id || '');
  setByHeader_(values, table, '作成区分', data.作成区分 || '自動計画');
  setByHeader_(values, table, 'fixed', !!data.fixed);
  setByHeader_(values, table, 'report_received', false);
  setByHeader_(values, table, 'user_id', getCurrentUserId_());

  getSheet_(APP.SHEETS.PLANS)
    .getRange(getSheet_(APP.SHEETS.PLANS).getLastRow() + 1, 1, 1, values.length)
    .setValues([values]);
}

// =====================================================
// 6. weakness管理
// =====================================================

function upsertWeakness_(w) {
  const table = readTable_(APP.SHEETS.WEAKNESS);
  const sh = getSheet_(APP.SHEETS.WEAKNESS);

  const resolved = resolveWeaknessIdentity_(w);
  const catalogId = resolved.catalogId;
  const subject = resolved.subject;
  const topic = resolved.topic;

  if (!topic) return;

  const userId = getCurrentUserId_();

  const existing = table.rows.find(r => {
    if (String(r.obj['user_id'] || '').trim() !== userId) return false;
    if (String(r.obj['状態'] || '') === '克服') return false;

    const rowCatalogId = String(r.obj['catalog_id'] || '').trim();
    const rowSubject = normalizeEvidenceText_(r.obj['科目']);
    const rowTopic = normalizeEvidenceText_(r.obj['テーマ']);

    if (catalogId && rowCatalogId && rowCatalogId === catalogId) return true;

    return (
      rowTopic === normalizeEvidenceText_(topic) &&
      (!subject || !rowSubject || rowSubject === normalizeEvidenceText_(subject))
    );
  });

  const understanding = clampInt_(w.understanding, 0, 5);
  const nextReview = nextDate_(today_(), 1);

  if (existing) {
    const values = existing.values.slice();

    if (catalogId) setByHeader_(values, table, 'catalog_id', catalogId);
    if (subject) setByHeader_(values, table, '科目', subject);
    setByHeader_(values, table, 'テーマ', topic);
    setByHeader_(values, table, '理解度', understanding);
    setByHeader_(values, table, '原因/メモ', appendMemo_(existing.obj['原因/メモ'], w.note));
    setByHeader_(values, table, '次回復習日', nextReview);
    setByHeader_(values, table, '復習間隔日', 1);
    setByHeader_(values, table, '状態', '要復習');
    setByHeader_(values, table, '最終更新', new Date());

    sh.getRange(existing.rowNumber, 1, 1, table.headers.length).setValues([values]);
    return;
  }

  const values = new Array(table.headers.length).fill('');
  setByHeader_(values, table, 'weakness_id', 'W-' + Utilities.getUuid());
  setByHeader_(values, table, 'catalog_id', catalogId);
  setByHeader_(values, table, '科目', subject);
  setByHeader_(values, table, 'テーマ', topic);
  setByHeader_(values, table, '理解度', understanding);
  setByHeader_(values, table, '原因/メモ', w.note || '');
  setByHeader_(values, table, '登録日', today_());
  setByHeader_(values, table, '次回復習日', nextReview);
  setByHeader_(values, table, '復習間隔日', 1);
  setByHeader_(values, table, '状態', '要復習');
  setByHeader_(values, table, '最終更新', new Date());
  setByHeader_(values, table, 'user_id', getCurrentUserId_());

  sh.getRange(sh.getLastRow() + 1, 1, 1, table.headers.length).setValues([values]);
}

function resolveWeaknessIdentity_(w) {
  let catalogId = String(w.catalog_id || '').trim();
  let subject = String(w.subject || '').trim();
  let topic = String(w.topic || '').trim();

  let catalog = catalogId ? findCatalogById_(catalogId) : null;

  if (!catalog && topic) {
    const table = readTable_(APP.SHEETS.CATALOG);
    const normTopic = normalizeEvidenceText_(topic);
    const normSubject = normalizeEvidenceText_(subject);

    catalog = table.rows.find(r => {
      const rowTopic = normalizeEvidenceText_(r.obj['範囲']);
      const rowSubject = normalizeEvidenceText_(r.obj['科目']);

      return rowTopic === normTopic && (!normSubject || rowSubject === normSubject);
    }) || null;
  }

  if (catalog) {
    catalogId = String(catalog.obj['id'] || catalogId).trim();
    subject = String(catalog.obj['科目'] || subject).trim();
    topic = String(catalog.obj['範囲'] || topic).trim();
  }

  return { catalogId, subject, topic };
}


function advanceWeaknessReview_(catalogId, understanding) {
  if (!catalogId) return;

  const table = readTable_(APP.SHEETS.WEAKNESS);
  const sh = getSheet_(APP.SHEETS.WEAKNESS);

  const userId = getCurrentUserId_();

  table.rows.forEach(r => {
    if (String(r.obj['user_id'] || '').trim() !== userId) return;
    if (String(r.obj['catalog_id'] || '') !== String(catalogId)) return;
    if (String(r.obj['状態'] || '') === '克服') return;

    const values = r.values.slice();
    let interval = number_(r.obj['復習間隔日']) || 1;

    if (understanding > 0 && understanding <= 2) {
      interval = 1;
      setByHeader_(values, table, '状態', '要復習');
    } else {
      interval = nextReviewInterval_(interval);
      if (interval >= 30 && understanding >= 4) {
        setByHeader_(values, table, '状態', '克服');
      } else {
        setByHeader_(values, table, '状態', '要復習');
      }
    }

    setByHeader_(values, table, '復習間隔日', interval);
    setByHeader_(values, table, '次回復習日', nextDate_(today_(), interval));
    setByHeader_(values, table, '理解度', understanding);
    setByHeader_(values, table, '最終更新', new Date());

    sh.getRange(r.rowNumber, 1, 1, table.headers.length).setValues([values]);
  });
}

function nextReviewInterval_(current) {
  if (current < 3) return 3;
  if (current < 7) return 7;
  if (current < 14) return 14;
  return 30;
}

// =====================================================
// 7. catalog / logs / event
// =====================================================

function updateCatalogProgress_(catalogId, rate) {
  // v3: catalogは全ユーザー共通の教材マスタなので、個人の完了状態では書き換えない。
  // 進捗はplans/logsのuser_id単位で管理する。
  return;
}

function upsertStudyLog_(record) {
  const table = readTable_(APP.SHEETS.LOGS);
  const sh = getSheet_(APP.SHEETS.LOGS);

  const userId = getCurrentUserId_();

  const existing = table.rows.find(r =>
    normalizeDate_(r.obj['date']) === normalizeDate_(record.date) &&
    String(r.obj['plan_id'] || '') === String(record.plan_id || '') &&
    String(r.obj['user_id'] || '').trim() === userId
  );

  record.user_id = userId;

  const values = new Array(table.headers.length).fill('');
  table.headers.forEach((h, i) => {
    if (Object.prototype.hasOwnProperty.call(record, h)) values[i] = record[h];
  });

  if (existing) {
    sh.getRange(existing.rowNumber, 1, 1, table.headers.length).setValues([values]);
  } else {
    sh.getRange(sh.getLastRow() + 1, 1, 1, table.headers.length).setValues([values]);
  }
}

function upsertLineEvent_(record) {
  const table = readTable_(APP.SHEETS.EVENTS);
  const sh = getSheet_(APP.SHEETS.EVENTS);

  const eventId = String(record.webhookEventId || '').trim();
  const existing = eventId
    ? table.rows.find(r => String(r.obj['webhook_event_id'] || '') === eventId)
    : null;

  const map = {
    timestamp: new Date(),
    event_type: record.eventType || '',
    webhook_event_id: eventId,
    user_id: record.userId || '',
    message_text: record.messageText || '',
    reply_token: record.replyToken || '',
    status: record.status || '',
    parsed_json: record.parsedJson || '',
    note: record.note || ''
  };

  const values = new Array(table.headers.length).fill('');
  table.headers.forEach((h, i) => values[i] = map[h] !== undefined ? map[h] : '');

  if (existing) {
    sh.getRange(existing.rowNumber, 1, 1, table.headers.length).setValues([values]);
  } else {
    sh.getRange(sh.getLastRow() + 1, 1, 1, table.headers.length).setValues([values]);
  }
}

function isWebhookEventProcessed_(eventId) {
  if (!eventId) return false;
  const table = readTable_(APP.SHEETS.EVENTS);

  const row = table.rows.find(r =>
    String(r.obj['webhook_event_id'] || '') === eventId
  );

  return !!row && String(row.obj['status'] || '') === 'processed';
}

// =====================================================
// 8. LINE送信
// =====================================================

function sendLineResponse_(replyToken, userId, text) {
  let replied = false;

  if (replyToken) {
    try {
      replied = replyLine_(replyToken, text);
    } catch (_) {
      replied = false;
    }
  }

  if (!replied && userId) pushLine_(userId, text);
}

function replyLine_(replyToken, text) {
  const token = getSecret_(APP.SECRET_KEYS.LINE_TOKEN);
  if (!token) throw new Error('LINE_CHANNEL_ACCESS_TOKEN がScript Propertiesにありません。');

  const res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify({
      replyToken,
      messages: [{ type: 'text', text: String(text).substring(0, 5000) }]
    }),
    muteHttpExceptions: true
  });

  return res.getResponseCode() >= 200 && res.getResponseCode() < 300;
}

function pushLine_(userId, text) {
  const token = getSecret_(APP.SECRET_KEYS.LINE_TOKEN);
  if (!token) throw new Error('LINE_CHANNEL_ACCESS_TOKEN がScript Propertiesにありません。');

  const res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify({
      to: userId,
      messages: [{ type: 'text', text: String(text).substring(0, 5000) }]
    }),
    muteHttpExceptions: true
  });

  const code = res.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error('LINE push error: ' + code + ' ' + res.getContentText());
  }
}

function notifyAdminsError_(title, err) {
  try {
    getActiveRegisteredUsers_()
      .filter(u => String(u.obj['role'] || '') === 'admin')
      .forEach(u => {
        const userId = String(u.obj['user_id'] || '').trim();
        if (userId) pushLine_(userId, '⚠️ ' + title + '\n' + String(err).substring(0, 1000));
      });
  } catch (_) {}
}

function logSystemErrorForUser_(userId, message) {
  try {
    upsertLineEvent_({
      eventType: 'system_error',
      webhookEventId: '',
      userId: userId || '',
      messageText: '',
      replyToken: '',
      status: 'error',
      parsedJson: '',
      note: String(message)
    });
  } catch (_) {}
  notifyAdminsError_('ユーザー処理エラー', message);
}

// =====================================================
// 9. トリガー
// =====================================================

function createDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'sendTodayPlan') ScriptApp.deleteTrigger(t);
  });

  const hour = clampInt_(getConfig_('MORNING_HOUR') || 7, 0, 23);
  const minute = clampInt_(getConfig_('MORNING_MINUTE') || 0, 0, 59);

  ScriptApp.newTrigger('sendTodayPlan')
    .timeBased()
    .everyDays(1)
    .atHour(hour)
    .nearMinute(minute)
    .inTimezone(APP.TIMEZONE)
    .create();
}

// =====================================================
// 10. レポート返信
// =====================================================

function buildReportReply_(parsed, result, tomorrowTasks) {
  const lines = [
    '📘 学習報告を更新したよ。',
    '今回更新した予定：' + result.updated + '件',
    '計画外の進捗：' + result.extra + '件',
    result.ignored ? '根拠不足で反映しなかった候補：' + result.ignored + '件' : ''
  ];

  if (parsed.summary) lines.push('内容：' + parsed.summary);

  if (tomorrowTasks.length) {
    lines.push('', '明日のメニューも今の進捗に合わせて作り直したよ。');
  }

  lines.push('追加で進めた内容があれば、同じ日に何回送っても累積して反映するよ。');

  return lines.filter(v => v !== '').join('\n');
}

// =====================================================
// 11. 汎用テーブル操作
// =====================================================

function readTable_(sheetName) {
  const sh = getSheet_(sheetName);
  const range = sh.getDataRange();
  const values = range.getValues();

  if (!values.length) throw new Error(sheetName + ' has no data');

  const headers = values[0].map(v => String(v || '').trim());
  const headerMap = {};

  headers.forEach((h, i) => {
    if (h) headerMap[h] = i;
  });

  const rows = [];

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (row.every(v => v === '' || v === null)) continue;

    rows.push({
      rowNumber: i + 1,
      values: row,
      obj: rowToObject_(headers, row)
    });
  }

  return {
    sheet: sh,
    headers,
    rows,
    col(name) {
      if (headerMap[name] === undefined) {
        throw new Error(sheetName + ' に見出し「' + name + '」がありません。setupV2Schema()を実行してください。');
      }
      return headerMap[name];
    }
  };
}

function rowToObject_(headers, values) {
  const obj = {};
  headers.forEach((h, i) => {
    if (h) obj[h] = values[i];
  });
  return obj;
}

function setByHeader_(values, table, header, value) {
  values[table.col(header)] = value;
}

function getSheet_(sheetName) {
  const sh = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (!sh) throw new Error('sheet not found: ' + sheetName);
  return sh;
}

function getPlanRowsByDate_(dateStr) {
  const userId = getCurrentUserId_();
  return readTable_(APP.SHEETS.PLANS).rows.filter(r =>
    normalizeDate_(r.obj['日付']) === dateStr &&
    String(r.obj['user_id'] || '').trim() === userId
  );
}

function markPlanDateReported_(dateStr) {
  const table = readTable_(APP.SHEETS.PLANS);
  const sh = getSheet_(APP.SHEETS.PLANS);
  const col = table.col('report_received');
  const userId = getCurrentUserId_();

  table.rows.forEach(r => {
    if (
      normalizeDate_(r.obj['日付']) === dateStr &&
      String(r.obj['user_id'] || '').trim() === userId
    ) {
      sh.getRange(r.rowNumber, col + 1).setValue(true);
    }
  });
}

function planDateHasReport_(dateStr) {
  return getPlanRowsByDate_(dateStr).some(r => bool_(r.obj['report_received']));
}

function findLatestPlanDateBefore_(dateStr) {
  const userId = getCurrentUserId_();
  const dates = readTable_(APP.SHEETS.PLANS).rows
    .filter(r => String(r.obj['user_id'] || '').trim() === userId)
    .map(r => normalizeDate_(r.obj['日付']))
    .filter(d => d && d < dateStr)
    .sort();
  return dates.length ? dates[dates.length - 1] : '';
}

function findLatestPlanDateOnOrBefore_(dateStr) {
  const userId = getCurrentUserId_();
  const dates = readTable_(APP.SHEETS.PLANS).rows
    .filter(r => String(r.obj['user_id'] || '').trim() === userId)
    .map(r => normalizeDate_(r.obj['日付']))
    .filter(d => d && d <= dateStr)
    .sort();
  return dates.length ? dates[dates.length - 1] : '';
}

function getFixedMinutesForDate_(dateStr) {
  return getPlanRowsByDate_(dateStr)
    .filter(r => bool_(r.obj['fixed']))
    .reduce((sum, r) => sum + number_(r.obj['予定分']), 0);
}

function getCatalogCandidates_(subject) {
  return readTable_(APP.SHEETS.CATALOG).rows
    .filter(r =>
      String(r.obj['科目'] || '').trim() === String(subject || '').trim()
    )
    .sort((a, b) => {
      const pa = number_(a.obj['優先度']);
      const pb = number_(b.obj['優先度']);
      if (pa !== pb) return pb - pa;
      return a.rowNumber - b.rowNumber;
    });
}

function getUsedCatalogIds_(targetDate) {
  const used = {};

  // 実際に進捗があるもの、手動固定したものは「既に使った教材」とみなす。
  const userId = getCurrentUserId_();

  readTable_(APP.SHEETS.PLANS).rows.forEach(r => {
    if (String(r.obj['user_id'] || '').trim() !== userId) return;
    const id = String(r.obj['catalog_id'] || '').trim();
    if (!id) return;

    const rate = clampInt_(r.obj['完了率%'], 0, 100);
    const fixed = bool_(r.obj['fixed']);
    const rowDate = normalizeDate_(r.obj['日付']);

    if (rate > 0 || fixed || (targetDate && rowDate === targetDate)) {
      used[id] = true;
    }
  });

  return used;
}

function findCatalogById_(id) {
  if (!id) return null;
  return readTable_(APP.SHEETS.CATALOG).rows.find(r =>
    String(r.obj['id'] || '') === String(id)
  ) || null;
}

function findCatalogByText_(ep) {
  const material = String(ep.material || '').trim();
  const range = String(ep.range || '').trim();

  return readTable_(APP.SHEETS.CATALOG).rows.find(r =>
    (!material || String(r.obj['教材'] || '').trim() === material) &&
    (!range || String(r.obj['範囲'] || '').trim() === range)
  ) || null;
}

function readCatalogForPrompt_() {
  const rows = readTable_(APP.SHEETS.CATALOG).rows;

  return rows.slice(0, 300).map(r => [
    'id=' + String(r.obj['id'] || ''),
    '科目=' + String(r.obj['科目'] || ''),
    '教材=' + String(r.obj['教材'] || ''),
    '範囲=' + String(r.obj['範囲'] || ''),
    '状態=' + String(r.obj['状態'] || '')
  ].join(', ')).join('\n');
}

// =====================================================
// 11.5 旧仕様の「未報告なのに勝手に進んだ予定」を整理
// =====================================================

function cleanupLegacyUnreportedAutoPlans() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const table = readTable_(APP.SHEETS.PLANS);
    const sh = getSheet_(APP.SHEETS.PLANS);
    const rowsToDelete = [];

    table.rows.forEach(r => {
      const o = r.obj;
      const source = String(o['作成区分'] || '').trim();
      const rate = clampInt_(o['完了率%'], 0, 100);
      const actual = number_(o['実績分']);
      const reported = bool_(o['report_received']);
      const completed = bool_(o['完了']);

      const legacyAuto = ['catalog自動生成', '未報告再送'].includes(source);

      if (legacyAuto && rate === 0 && actual === 0 && !reported && !completed) {
        rowsToDelete.push(r.rowNumber);
      }
    });

    rowsToDelete.sort((a, b) => b - a).forEach(n => sh.deleteRow(n));
    return 'deleted ' + rowsToDelete.length + ' legacy rows';
  } finally {
    lock.releaseLock();
  }
}


function getAppVersion() {
  return '3.0.0-multiuser';
}


function setCurrentUserId_(userId) {
  CURRENT_USER_ID_ = String(userId || '').trim();
}

function clearCurrentUserId_() {
  CURRENT_USER_ID_ = '';
}

function getCurrentUserId_() {
  const userId = String(CURRENT_USER_ID_ || '').trim();
  if (!userId) throw new Error('current user is not set');
  return userId;
}

function getUsersTable_() {
  return readTable_(APP.SHEETS.USERS);
}

function findUserByLineUserId_(userId) {
  const id = String(userId || '').trim();
  if (!id) return null;
  return getUsersTable_().rows.find(r =>
    String(r.obj['user_id'] || '').trim() === id
  ) || null;
}

function getCurrentUser_() {
  const id = String(CURRENT_USER_ID_ || '').trim();
  return id ? findUserByLineUserId_(id) : null;
}

function getActiveRegisteredUsers_() {
  return getUsersTable_().rows.filter(r =>
    bool_(r.obj['active']) &&
    String(r.obj['user_id'] || '').trim()
  );
}

function registerUserByCode_(lineUserId, messageText) {
  const codeText = String(messageText || '').trim();
  if (!codeText) return null;

  const table = getUsersTable_();
  const sh = getSheet_(APP.SHEETS.USERS);

  // 同じLINE userIdが既に別行へ登録されていたら再登録しない。
  if (table.rows.some(r => String(r.obj['user_id'] || '').trim() === String(lineUserId || '').trim())) {
    return findUserByLineUserId_(lineUserId);
  }

  const target = table.rows.find(r =>
    String(r.obj['registration_code'] || '').trim() === codeText &&
    !String(r.obj['user_id'] || '').trim() &&
    bool_(r.obj['active'])
  );

  if (!target) return null;

  const values = target.values.slice();
  setByHeader_(values, table, 'user_id', String(lineUserId || '').trim());
  setByHeader_(values, table, 'registered_at', new Date());

  sh.getRange(target.rowNumber, 1, 1, table.headers.length).setValues([values]);

  return {
    rowNumber: target.rowNumber,
    values,
    obj: rowToObject_(table.headers, values)
  };
}

function setUserLastSentDate_(rowNumber, dateStr) {
  const table = getUsersTable_();
  getSheet_(APP.SHEETS.USERS)
    .getRange(rowNumber, table.col('last_sent_date') + 1)
    .setValue(dateStr);
}

function resetUserRegistrationByRow(rowNumber) {
  const table = getUsersTable_();
  const sh = getSheet_(APP.SHEETS.USERS);
  const row = clampInt_(rowNumber, 2, sh.getMaxRows());

  sh.getRange(row, table.col('user_id') + 1).clearContent();
  sh.getRange(row, table.col('registered_at') + 1).clearContent();
  sh.getRange(row, table.col('last_sent_date') + 1).clearContent();

  return 'registration reset: row ' + row;
}

// =====================================================
// 12. config / secret / 日付
// =====================================================

function getConfigMap_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('CONFIG_MAP');

  if (cached) return JSON.parse(cached);

  const table = readTable_(APP.SHEETS.CONFIG);
  const map = {};

  table.rows.forEach(r => {
    const key = String(r.obj['key'] || '').trim();
    if (key) map[key] = String(r.obj['value'] || '').trim();
  });

  cache.put('CONFIG_MAP', JSON.stringify(map), 300);
  return map;
}

function getConfig_(key) {
  return getConfigMap_()[key] || '';
}

function clearConfigCache() {
  CacheService.getScriptCache().remove('CONFIG_MAP');
}

function getSecret_(key) {
  return String(PropertiesService.getScriptProperties().getProperty(key) || '').trim();
}


function getBudgetForDate_(dateStr) {
  const user = getCurrentUser_();
  const key = isWeekend_(dateStr) ? 'weekend_target_min' : 'weekday_target_min';
  const fallback = isWeekend_(dateStr)
    ? number_(getConfig_('WEEKEND_TARGET_MIN') || 120)
    : number_(getConfig_('WEEKDAY_TARGET_MIN') || 60);

  if (!user) return fallback;

  const value = number_(user.obj[key]);
  return value > 0 ? value : fallback;
}

function parsePlanRules_(text) {
  return String(text || '')
    .split(',')
    .map(s => s.trim())
    .filter(Boolean)
    .map(part => {
      const idx = part.lastIndexOf(':');
      if (idx < 0) return { subject: part, count: 1 };

      return {
        subject: part.substring(0, idx).trim(),
        count: Math.max(1, clampInt_(part.substring(idx + 1), 1, 20))
      };
    });
}

function isWeekend_(dateStr) {
  const d = parseDate_(dateStr);
  const day = d.getDay();
  return day === 0 || day === 6;
}

function today_() {
  return Utilities.formatDate(new Date(), APP.TIMEZONE, 'yyyy/MM/dd');
}

function nextDate_(dateStr, days) {
  const d = parseDate_(dateStr);
  d.setDate(d.getDate() + Number(days || 0));
  return Utilities.formatDate(d, APP.TIMEZONE, 'yyyy/MM/dd');
}

function parseDate_(dateStr) {
  const s = normalizeDate_(dateStr);
  const parts = s.split('/');
  return new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]), 12, 0, 0);
}

function normalizeDate_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, APP.TIMEZONE, 'yyyy/MM/dd');
  }

  const s = String(value || '').trim().replace(/-/g, '/');
  if (!s) return '';

  const m = s.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})$/);
  if (m) {
    return [
      m[1],
      String(m[2]).padStart(2, '0'),
      String(m[3]).padStart(2, '0')
    ].join('/');
  }

  const d = new Date(s);
  if (!isNaN(d.getTime())) return Utilities.formatDate(d, APP.TIMEZONE, 'yyyy/MM/dd');

  return s;
}

function makePlanId_(dateStr, suffix) {
  const compact = String(dateStr || '').replace(/\//g, '');
  return compact + '-' + String(suffix || Utilities.getUuid()).replace(/[^A-Za-z0-9]/g, '').slice(-12);
}

function cleanCarryoverRange_(range) {
  return String(range || '')
    .replace(/（残り・復習）/g, '')
    .replace(/\(残り・復習\)/g, '')
    .trim();
}

function appendMemo_(oldMemo, newMemo) {
  const oldText = String(oldMemo || '').trim();
  const newText = String(newMemo || '').trim();
  if (!newText) return oldText;
  if (!oldText) return '[' + today_() + '] ' + newText;
  if (oldText.includes(newText)) return oldText;
  return oldText + '\n[' + today_() + '] ' + newText;
}

function number_(v) {
  const n = Number(v);
  return isNaN(n) ? 0 : n;
}

function clampInt_(v, min, max) {
  const n = Math.round(number_(v));
  return Math.max(min, Math.min(max, n));
}

function bool_(v) {
  if (v === true) return true;
  const s = String(v || '').trim().toUpperCase();
  return s === 'TRUE' || s === 'YES' || s === '1' || s === 'ON';
}
