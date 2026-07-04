/**
 * PROJECT BRAIN OS - Google Apps Script Backend
 * Version: 1.2.0
 *
 * CẬP NHẬT:
 * 1. Telegram gửi thông báo DONE chi tiết hơn.
 * 2. Telegram có thể nhận lệnh điều khiển:
 *    /start, /help, /status, /stop, /resume, /run, /reset_failed, /test, /webhook
 *
 * LUỒNG:
 * HTML Dashboard
 * -> Google Apps Script Web App
 * -> Google Sheet
 * -> Agent Queue
 * -> Gemini / Claude API
 * -> Google Docs / Drive
 * -> Telegram Notification / Telegram Command
 *
 * LƯU Ý:
 * - Không chạy setupProjectBrainOS() nếu muốn giữ dữ liệu sheet cũ.
 * - Chỉ chạy setupProjectBrainOS() khi tạo database mới từ đầu.
 * - Không dán API key vào chat hoặc ảnh chụp màn hình.
 */

const PBOS = {
  VERSION: '1.2.0',
  SHEETS: {
    PROJECTS: 'Project_Intake',
    TASKS: 'Agent_Tasks',
    OUTPUTS: 'Agent_Outputs',
    ACTIONS: 'Action_Plan',
    RISKS: 'Risk_Register',
    FINANCE: 'Finance_Model',
    FINAL: 'Final_Output',
    LOGS: 'Run_Logs',
    SETTINGS: 'Settings'
  },
  DEFAULT_MODELS: {
    GEMINI: 'gemini-2.5-flash',
    CLAUDE: 'claude-sonnet-4-5'
  }
};

/************************************************************
 * WEB APP
 ************************************************************/

function doGet(e) {
  return HtmlService
    .createTemplateFromFile('index')
    .evaluate()
    .setTitle('Project Brain OS')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents || '{}');

    /**
     * Nếu request đến từ Telegram webhook
     */
    if (payload.message || payload.edited_message || payload.callback_query) {
      return json_(handleTelegramUpdate_(payload));
    }

    /**
     * Nếu request đến từ HTML Dashboard
     */
    const action = payload.action || 'submitProject';

    if (action === 'submitProject') return json_(submitProject(payload));
    if (action === 'status') return json_(checkProjectBrainOSStatus());
    if (action === 'stop') return json_(stopProjectBrainOS());
    if (action === 'resume') return json_(resumeProjectBrainOS());
    if (action === 'run') return json_(runAgentQueue());
    if (action === 'reset_failed') return json_(resetFailedTasks());
    if (action === 'testTelegram') return json_(testTelegram());

    return json_({
      ok: false,
      error: 'Unknown action: ' + action
    });

  } catch (err) {
    return json_({
      ok: false,
      error: String(err),
      stack: err.stack || ''
    });
  }
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj, null, 2))
    .setMimeType(ContentService.MimeType.JSON);
}

/************************************************************
 * SETUP DATABASE
 ************************************************************/

function setupProjectBrainOS() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const headers = {
    [PBOS.SHEETS.PROJECTS]: [
      'Project_ID',
      'Created_At',
      'Project_Name',
      'Stage',
      'Idea',
      'Goal_30',
      'Target_User',
      'Budget',
      'Deadline',
      'Resources',
      'Constraints',
      'Desired_Output',
      'Health_Score',
      'Folder_URL',
      'Master_Doc_URL',
      'Status'
    ],

    [PBOS.SHEETS.TASKS]: [
      'Task_ID',
      'Project_ID',
      'Agent',
      'Complexity',
      'Provider',
      'Model',
      'Mission',
      'Input_Summary',
      'Expected_Output',
      'Priority',
      'Deadline',
      'Status',
      'Attempt',
      'Last_Error',
      'Output_Doc_URL',
      'Updated_At'
    ],

    [PBOS.SHEETS.OUTPUTS]: [
      'Output_ID',
      'Project_ID',
      'Task_ID',
      'Agent',
      'Provider',
      'Model',
      'Output',
      'Quality_Score',
      'Created_At'
    ],

    [PBOS.SHEETS.ACTIONS]: [
      'Action_ID',
      'Project_ID',
      'Day',
      'Action',
      'Owner',
      'Output',
      'Priority',
      'Status',
      'Notes'
    ],

    [PBOS.SHEETS.RISKS]: [
      'Risk_ID',
      'Project_ID',
      'Risk',
      'Impact',
      'Probability',
      'Prevention',
      'Response',
      'Owner',
      'Status'
    ],

    [PBOS.SHEETS.FINANCE]: [
      'Project_ID',
      'Initial_Cost',
      'Monthly_Fixed_Cost',
      'Variable_Cost',
      'Expected_Revenue',
      'Expected_Profit',
      'Break_Even',
      'Worst_Case',
      'Base_Case',
      'Best_Case',
      'Notes'
    ],

    [PBOS.SHEETS.FINAL]: [
      'Project_ID',
      'Brief_Doc_URL',
      'SOP_Doc_URL',
      'Checklist_Doc_URL',
      'Agent_Output_Folder_URL',
      'Status',
      'Updated_At'
    ],

    [PBOS.SHEETS.LOGS]: [
      'Created_At',
      'Level',
      'Project_ID',
      'Task_ID',
      'Message',
      'Data'
    ],

    [PBOS.SHEETS.SETTINGS]: [
      'Key',
      'Value',
      'Note'
    ]
  };

  Object.keys(headers).forEach(name => {
    let sh = ss.getSheetByName(name) || ss.insertSheet(name);
    sh.clear();
    sh.getRange(1, 1, 1, headers[name].length).setValues([headers[name]]);
    sh.setFrozenRows(1);
    sh.autoResizeColumns(1, headers[name].length);
  });

  const settings = ss.getSheetByName(PBOS.SHEETS.SETTINGS);

  settings.getRange(2, 1, 14, 3).setValues([
    ['GEMINI_MODEL', PBOS.DEFAULT_MODELS.GEMINI, 'Model dùng cho tác vụ nhẹ/nhanh'],
    ['CLAUDE_MODEL', PBOS.DEFAULT_MODELS.CLAUDE, 'Model dùng cho tác vụ phức tạp'],
    ['ROOT_FOLDER_NAME', 'Project Brain OS Outputs', 'Tên thư mục Drive chứa output'],
    ['QUEUE_BATCH_SIZE', '2', 'Số Agent chạy mỗi lần để tránh timeout'],
    ['ENABLE_TELEGRAM', 'FALSE', 'TRUE nếu muốn gửi thông báo Telegram'],
    ['TELEGRAM_CHAT_ID', '', 'Chat ID nếu dùng Telegram'],
    ['TELEGRAM_WEBHOOK_URL', '', 'Web App URL dùng cho Telegram webhook'],
    ['ALLOW_TELEGRAM_ANY_CHAT', 'FALSE', 'TRUE nếu cho phép mọi chat điều khiển bot, không khuyến nghị'],
    ['ENABLE_GMAIL', 'FALSE', 'TRUE nếu muốn gửi email thông báo'],
    ['NOTIFY_EMAIL', '', 'Email nhận thông báo nếu dùng Gmail'],
    ['MAX_OUTPUT_TOKENS_GEMINI', '4096', 'Giới hạn output Gemini'],
    ['MAX_OUTPUT_TOKENS_CLAUDE', '4096', 'Giới hạn output Claude'],
    ['AUTO_STOP_WHEN_DONE', 'FALSE', 'TRUE nếu muốn tự xóa trigger khi không còn task PENDING/RETRY'],
    ['PROJECT_BRAIN_OS_PAUSED', 'FALSE', 'Trạng thái dừng tạm thời']
  ]);

  PropertiesService.getScriptProperties().setProperty('PROJECT_BRAIN_OS_PAUSED', 'FALSE');

  ensureQueueTrigger_();

  return {
    ok: true,
    message: 'Project Brain OS sheets and trigger created.'
  };
}

/************************************************************
 * API KEYS + TELEGRAM CONFIG
 ************************************************************/

function setApiKeys(geminiApiKey, claudeApiKey, telegramBotToken) {
  const props = PropertiesService.getScriptProperties();

  if (geminiApiKey) {
    props.setProperty('GEMINI_API_KEY', geminiApiKey);
  }

  if (claudeApiKey) {
    props.setProperty('CLAUDE_API_KEY', claudeApiKey);
  }

  if (telegramBotToken) {
    props.setProperty('TELEGRAM_BOT_TOKEN', telegramBotToken);
  }

  return {
    ok: true,
    message: 'API keys saved in Script Properties.'
  };
}

function saveMyApiKeys() {
  /**
   * THAY 4 DÒNG BÊN DƯỚI BẰNG KEY MỚI CỦA BẠN.
   * Không gửi key qua chat hoặc ảnh chụp màn hình.
   */


  PropertiesService.getScriptProperties().setProperty(
    'ENABLE_TELEGRAM',
    'TRUE'
  );

  return {
    ok: true,
    message: 'Saved API keys and Telegram config to Script Properties.'
  };
}

function saveTelegramWebhookUrl() {
  /**
   * Dán Web App URL của Apps Script vào biến webAppUrl bên dưới.
   * Ví dụ:
   * https://script.google.com/macros/s/AKfycbxxxxxxxxxxxxxxxx/exec
   */
  const webAppUrl = 'https://script.google.com/macros/s/AKfycbwt5Y1KZO37uocp-IK3V0Csww3SzM8NB6SyekWSAWiS7Qoky76zL4_sowfmqdyajIt2Sw/exec';

  PropertiesService.getScriptProperties().setProperty(
    'TELEGRAM_WEBHOOK_URL',
    webAppUrl
  );

  setSetting_(
    'TELEGRAM_WEBHOOK_URL',
    webAppUrl,
    'Web App URL dùng cho Telegram webhook'
  );

  return {
    ok: true,
    webAppUrl
  };
}

/************************************************************
 * SUBMIT PROJECT
 ************************************************************/

function submitProject(project) {
  ensureSheets_();

  const projectId =
    'P-' +
    Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      'yyyyMMdd-HHmmss'
    ) +
    '-' +
    Math.floor(Math.random() * 9000 + 1000);

  const createdAt = new Date();
  const score = scoreProject_(project);
  const p = normalizeProject_(project);

  const folder = createProjectFolder_(
    projectId,
    p.name || 'Untitled Project'
  );

  const masterDoc = DocumentApp.create(
    projectId + ' - Master Project Doc'
  );

  DriveApp.getFileById(masterDoc.getId()).moveTo(folder);

  append_(PBOS.SHEETS.PROJECTS, [
    projectId,
    createdAt,
    p.name,
    p.stage,
    p.idea,
    p.goal30,
    p.targetUser,
    p.budget,
    p.deadline,
    p.resources,
    p.constraints,
    p.desiredOutput,
    score,
    folder.getUrl(),
    masterDoc.getUrl(),
    'QUEUED'
  ]);

  const tasks = buildAgentTasks_(projectId, p);

  const taskRows = tasks.map(t => [
    t.taskId,
    projectId,
    t.agent,
    t.complexity,
    t.provider,
    t.model,
    t.mission,
    short_(p.idea, 240),
    t.expectedOutput,
    t.priority,
    t.deadline,
    'PENDING',
    0,
    '',
    '',
    createdAt
  ]);

  const taskSheet = getSheet_(PBOS.SHEETS.TASKS);

  taskSheet
    .getRange(
      taskSheet.getLastRow() + 1,
      1,
      taskRows.length,
      taskRows[0].length
    )
    .setValues(taskRows);

  seedActionPlan_(projectId);

  log_('INFO', projectId, '', 'Project submitted and agent tasks queued.', {
    score: score,
    folderUrl: folder.getUrl()
  });

  notify_(
    projectId,
    '🚀 Dự án mới đã được gửi vào Project Brain OS.\n\n' +
      '📌 Project: ' + p.name + '\n' +
      '🆔 ID: ' + projectId + '\n' +
      '📊 Health Score: ' + score + '\n' +
      '📁 Drive Folder:\n' + folder.getUrl() + '\n\n' +
      '📄 Master Doc:\n' + masterDoc.getUrl()
  );

  ensureQueueTrigger_();

  return {
    ok: true,
    projectId: projectId,
    healthScore: score,
    status: 'QUEUED',
    message: 'Project received. Agent queue is ready.',
    folderUrl: folder.getUrl(),
    masterDocUrl: masterDoc.getUrl(),
    nextStep: 'Run runAgentQueue() manually once, or wait for the time trigger.'
  };
}

/************************************************************
 * AGENT QUEUE
 ************************************************************/

function runAgentQueue() {
  if (isPaused_()) {
    Logger.log('Project Brain OS is paused. Queue stopped.');
    return {
      ok: true,
      paused: true,
      processed: 0
    };
  }

  ensureSheets_();

  const settings = getSettings_();
  const batchSize = Number(settings.QUEUE_BATCH_SIZE || 2);

  const sh = getSheet_(PBOS.SHEETS.TASKS);
  const data = sh.getDataRange().getValues();

  if (data.length <= 1) {
    return {
      ok: true,
      processed: 0,
      message: 'No tasks.'
    };
  }

  const headers = data[0];
  const idx = indexMap_(headers);

  let processed = 0;

  for (let r = 1; r < data.length && processed < batchSize; r++) {
    const row = data[r];

    if (row[idx.Status] !== 'PENDING' && row[idx.Status] !== 'RETRY') {
      continue;
    }

    const task = rowToObject_(headers, row);

    try {
      sh.getRange(r + 1, idx.Status + 1).setValue('RUNNING');
      sh.getRange(r + 1, idx.Updated_At + 1).setValue(new Date());

      const project = findProject_(task.Project_ID);
      const prompt = buildAgentPrompt_(project, task);

      const output = callModel_(
        task.Provider,
        task.Model,
        prompt,
        Number(settings['MAX_OUTPUT_TOKENS_' + task.Provider] || 4096)
      );

      const outputDocUrl = saveAgentOutput_(project, task, output);

      append_(PBOS.SHEETS.OUTPUTS, [
        'O-' + task.Task_ID,
        task.Project_ID,
        task.Task_ID,
        task.Agent,
        task.Provider,
        task.Model,
        output,
        estimateQuality_(output),
        new Date()
      ]);

      sh.getRange(r + 1, idx.Status + 1).setValue('DONE');
      sh.getRange(r + 1, idx.Last_Error + 1).setValue('');
      sh.getRange(r + 1, idx.Output_Doc_URL + 1).setValue(outputDocUrl);
      sh.getRange(r + 1, idx.Updated_At + 1).setValue(new Date());

      log_('INFO', task.Project_ID, task.Task_ID, task.Agent + ' completed.', {
        outputDocUrl: outputDocUrl
      });

      processed++;

    } catch (err) {
      const attempt = Number(row[idx.Attempt] || 0) + 1;

      sh.getRange(r + 1, idx.Attempt + 1).setValue(attempt);
      sh.getRange(r + 1, idx.Last_Error + 1).setValue(String(err));
      sh.getRange(r + 1, idx.Status + 1).setValue(
        attempt >= 3 ? 'FAILED' : 'RETRY'
      );
      sh.getRange(r + 1, idx.Updated_At + 1).setValue(new Date());

      log_('ERROR', task.Project_ID, task.Task_ID, 'Agent failed: ' + String(err), {
        stack: err.stack || ''
      });

      notify_(
        task.Project_ID,
        '⚠️ Agent lỗi: ' + task.Agent + '\n' + String(err).slice(0, 900)
      );

      processed++;
    }
  }

  finalizeCompletedProjects_();
  autoStopIfIdle_();

  return {
    ok: true,
    processed: processed
  };
}

/************************************************************
 * FINALIZE PROJECT
 ************************************************************/

function finalizeCompletedProjects_() {
  const projectsSh = getSheet_(PBOS.SHEETS.PROJECTS);
  const projects = projectsSh.getDataRange().getValues();

  if (projects.length <= 1) {
    return;
  }

  const ph = projects[0];
  const pi = indexMap_(ph);

  const tasksSh = getSheet_(PBOS.SHEETS.TASKS);
  const tasks = tasksSh.getDataRange().getValues();

  if (tasks.length <= 1) {
    return;
  }

  const th = tasks[0];
  const ti = indexMap_(th);

  for (let r = 1; r < projects.length; r++) {
    const projectId = projects[r][pi.Project_ID];
    const status = projects[r][pi.Status];

    if (status === 'DONE') {
      continue;
    }

    const related = tasks
      .slice(1)
      .filter(t => t[ti.Project_ID] === projectId);

    if (
      related.length &&
      related.every(t => ['DONE', 'FAILED'].includes(t[ti.Status]))
    ) {
      const project = rowToObject_(ph, projects[r]);
      const docUrl = synthesizeFinalDoc_(projectId, project);

      projectsSh.getRange(r + 1, pi.Status + 1).setValue('DONE');

      append_(PBOS.SHEETS.FINAL, [
        projectId,
        docUrl,
        '',
        '',
        project.Folder_URL,
        'DONE',
        new Date()
      ]);

      log_('INFO', projectId, '', 'Project finalized.', {
        finalDocUrl: docUrl
      });

      const message = buildProjectDoneTelegramMessage_(
        project,
        related,
        ti,
        docUrl
      );

      notify_(projectId, message);
    }
  }
}

function buildProjectDoneTelegramMessage_(project, relatedTasks, taskIdx, finalDocUrl) {
  const totalCount = relatedTasks.length;
  const doneCount = relatedTasks.filter(t => t[taskIdx.Status] === 'DONE').length;
  const failedCount = relatedTasks.filter(t => t[taskIdx.Status] === 'FAILED').length;

  const failedAgents = relatedTasks
    .filter(t => t[taskIdx.Status] === 'FAILED')
    .map(t => '- ' + t[taskIdx.Agent])
    .join('\n');

  let message =
    '✅ PROJECT BRAIN OS HOÀN TẤT\n\n' +
    '📌 Project: ' + (project.Project_Name || '') + '\n' +
    '🆔 ID: ' + (project.Project_ID || '') + '\n' +
    '📊 Health Score: ' + (project.Health_Score || '') + '\n' +
    '🤖 Agent hoàn tất: ' + doneCount + '/' + totalCount + '\n' +
    '⚠️ Agent lỗi: ' + failedCount + '\n\n' +
    '📄 Final Doc:\n' + finalDocUrl + '\n\n' +
    '📁 Drive Folder:\n' + (project.Folder_URL || '') + '\n\n' +
    '📋 Việc tiếp theo:\n' +
    '1. Mở Final Doc kiểm tra nội dung.\n' +
    '2. Đọc QA Review và Action_Plan.\n' +
    '3. Đánh dấu Action_Plan DONE nếu đã nghiệm thu.\n' +
    '4. Chạy thử dự án tiếp theo để kiểm tra khả năng tái sử dụng.';

  if (failedCount > 0) {
    message += '\n\n⚠️ Agent FAILED:\n' + failedAgents;
  }

  return message;
}

function autoStopIfIdle_() {
  const settings = getSettings_();

  if (String(settings.AUTO_STOP_WHEN_DONE || '').toUpperCase() !== 'TRUE') {
    return;
  }

  const sh = getSheet_(PBOS.SHEETS.TASKS);
  const data = sh.getDataRange().getValues();

  if (data.length <= 1) {
    return;
  }

  const h = data[0];
  const i = indexMap_(h);

  const hasPending = data
    .slice(1)
    .some(r => ['PENDING', 'RETRY', 'RUNNING'].includes(r[i.Status]));

  if (!hasPending) {
    stopProjectBrainOS();
  }
}

/************************************************************
 * SYNTHESIZE FINAL DOC
 ************************************************************/

function synthesizeFinalDoc_(projectId, project) {
  const outputsSh = getSheet_(PBOS.SHEETS.OUTPUTS);
  const values = outputsSh.getDataRange().getValues();

  const h = values[0];
  const i = indexMap_(h);

  const outputs = values
    .slice(1)
    .filter(r => r[i.Project_ID] === projectId);

  const doc = DocumentApp.create(
    projectId + ' - FINAL PROJECT BRAIN OS OUTPUT'
  );

  const body = doc.getBody();

  body
    .appendParagraph('PROJECT BRAIN OS - FINAL OUTPUT')
    .setHeading(DocumentApp.ParagraphHeading.TITLE);

  body.appendParagraph('Project ID: ' + projectId);
  body.appendParagraph('Project: ' + project.Project_Name);
  body.appendParagraph('Created: ' + new Date());
  body.appendHorizontalRule();

  outputs.forEach(r => {
    body
      .appendParagraph(String(r[i.Agent] || ''))
      .setHeading(DocumentApp.ParagraphHeading.HEADING1);

    body.appendParagraph(String(r[i.Output] || ''));
  });

  doc.saveAndClose();

  const folderId = extractDriveId_(project.Folder_URL);

  if (folderId) {
    DriveApp.getFileById(doc.getId()).moveTo(
      DriveApp.getFolderById(folderId)
    );
  }

  return doc.getUrl();
}

/************************************************************
 * TELEGRAM COMMAND HANDLER
 ************************************************************/

function handleTelegramUpdate_(update) {
  const updateId = String(update.update_id || '');

  // Chống Telegram retry cùng một update nhiều lần
  if (updateId) {
    const cache = CacheService.getScriptCache();
    const cacheKey = 'TG_UPDATE_' + updateId;

    if (cache.get(cacheKey)) {
      Logger.log('Duplicate Telegram update ignored: ' + updateId);
      return {
        ok: true,
        duplicate: true,
        updateId: updateId
      };
    }

    cache.put(cacheKey, '1', 600);
  }

  const lock = LockService.getScriptLock();

  try {
    lock.tryLock(5000);

    const props = PropertiesService.getScriptProperties();

    const message = update.message || update.edited_message || {};
    const from = message.from || {};
    const chat = message.chat || {};
    const chatId = String(chat.id || '');
    const text = String(message.text || '').trim();

    // Không xử lý message từ bot khác
    if (from.is_bot) {
      return {
        ok: true,
        skipped: true,
        reason: 'FROM_BOT'
      };
    }

    const settings = getSettings_();

    const allowedChatId = String(
      settings.TELEGRAM_CHAT_ID ||
      props.getProperty('TELEGRAM_CHAT_ID') ||
      ''
    );

    const allowAnyChat =
      String(settings.ALLOW_TELEGRAM_ANY_CHAT || '').toUpperCase() === 'TRUE';

    if (!chatId) {
      return {
        ok: false,
        error: 'NO_CHAT_ID'
      };
    }

    if (!allowAnyChat && allowedChatId && chatId !== allowedChatId) {
      sendTelegramMessage_(
        chatId,
        '⛔ Bạn không có quyền điều khiển Project Brain OS.'
      );

      log_('ERROR', 'TELEGRAM', '', 'Unauthorized Telegram chat.', {
        chatId: chatId,
        allowedChatId: allowedChatId
      });

      return {
        ok: false,
        error: 'UNAUTHORIZED_CHAT_ID',
        chatId: chatId
      };
    }

    if (!text) {
      sendTelegramMessage_(
        chatId,
        'Mình chỉ nhận lệnh dạng text. Gõ /help để xem lệnh.'
      );

      return {
        ok: true,
        message: 'No text command.'
      };
    }

    let reply = '';

    if (text === '/start' || text === '/help') {
      reply =
        '🤖 Project Brain OS Bot\n\n' +
        'Các lệnh hiện có:\n' +
        '/status - Xem trạng thái hệ thống\n' +
        '/stop - Dừng tạm thời hệ thống\n' +
        '/resume - Bật lại hệ thống\n' +
        '/run - Chạy Agent Queue một lần\n' +
        '/reset_failed - Reset task FAILED về PENDING\n' +
        '/test - Test Telegram\n' +
        '/webhook - Xem thông tin webhook\n' +
        '/help - Xem hướng dẫn';
    }

    else if (text === '/status') {
      const status = checkProjectBrainOSStatus();

      reply =
        '📊 Trạng thái Project Brain OS\n\n' +
        'Paused: ' + status.paused + '\n' +
        'Version: ' + status.version + '\n\n' +
        'Task Stats:\n' +
        JSON.stringify(status.taskStats || {}, null, 2) + '\n\n' +
        'Triggers: ' + ((status.triggers || []).length);
    }

    else if (text === '/stop') {
      const result = stopProjectBrainOS();

      reply =
        '⏸ Đã dừng tạm thời Project Brain OS.\n' +
        'Deleted triggers: ' + result.deletedTriggers;
    }

    else if (text === '/resume') {
      resumeProjectBrainOS();

      reply =
        '▶️ Đã bật lại Project Brain OS.\n' +
        'Queue trigger đã sẵn sàng.';
    }

    else if (text === '/run') {
      const result = runAgentQueue();

      reply =
        '⚙️ Đã chạy Agent Queue một lần.\n' +
        'Processed: ' + (result.processed || 0) + '\n' +
        'Paused: ' + (result.paused || false);
    }

    else if (text === '/reset_failed') {
      const result = resetFailedTasks();

      reply =
        '♻️ Đã reset task lỗi.\n' +
        'Reset count: ' + result.resetCount;
    }

    else if (text === '/test') {
      reply = '✅ Telegram command bot đang hoạt động.';
    }

    else if (text === '/webhook') {
      const info = getTelegramWebhookInfo();

      reply =
        '🔗 Webhook info:\n' +
        JSON.stringify(info, null, 2).slice(0, 3500);
    }

    else {
      reply =
        'Không hiểu lệnh: ' + text + '\n\n' +
        'Gõ /help để xem danh sách lệnh.';
    }

    sendTelegramMessage_(chatId, reply);

    return {
      ok: true,
      command: text,
      updateId: updateId,
      reply: reply
    };

  } catch (err) {
    try {
      const message = update.message || {};
      const chat = message.chat || {};
      const chatId = String(chat.id || '');

      if (chatId) {
        sendTelegramMessage_(
          chatId,
          '❌ Lỗi khi xử lý lệnh Telegram:\n' + String(err).slice(0, 1000)
        );
      }
    } catch (innerErr) {
      Logger.log('Failed to send Telegram error message: ' + innerErr);
    }

    log_('ERROR', 'TELEGRAM', '', 'Telegram command error: ' + String(err), {
      stack: err.stack || ''
    });

    return {
      ok: false,
      error: String(err),
      stack: err.stack || ''
    };

  } finally {
    try {
      lock.releaseLock();
    } catch (e) {}
  }
}

function sendTelegramMessage_(chatId, text) {
  const token = PropertiesService
    .getScriptProperties()
    .getProperty('TELEGRAM_BOT_TOKEN');

  if (!token) {
    throw new Error('Missing TELEGRAM_BOT_TOKEN in Script Properties.');
  }

  const url =
    'https://api.telegram.org/bot' +
    token +
    '/sendMessage';

  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    muteHttpExceptions: true,
    payload: JSON.stringify({
      chat_id: chatId,
      text: text,
      disable_web_page_preview: false
    })
  });

  const code = res.getResponseCode();
  const body = res.getContentText();

  if (code >= 300) {
    throw new Error('Telegram sendMessage error ' + code + ': ' + body);
  }

  return body;
}

function setTelegramWebhook() {
  const props = PropertiesService.getScriptProperties();

  const token = props.getProperty('TELEGRAM_BOT_TOKEN');

  const webAppUrl =
    props.getProperty('TELEGRAM_WEBHOOK_URL') ||
    setting_('TELEGRAM_WEBHOOK_URL');

  if (!token) {
    throw new Error('Missing TELEGRAM_BOT_TOKEN in Script Properties.');
  }

  if (!webAppUrl || webAppUrl === 'DAN_GOOGLE_WEB_APP_URL_VAO_DAY') {
    throw new Error(
      'Missing TELEGRAM_WEBHOOK_URL. Run saveTelegramWebhookUrl() after adding your Web App URL.'
    );
  }

  const url =
  'https://api.telegram.org/bot' +
  token +
  '/setWebhook?url=' +
  encodeURIComponent(webAppUrl) +
  '&drop_pending_updates=true';

  const res = UrlFetchApp.fetch(url, {
    method: 'get',
    muteHttpExceptions: true
  });

  const code = res.getResponseCode();
  const text = res.getContentText();

  log_(
    code >= 200 && code < 300 ? 'INFO' : 'ERROR',
    'TELEGRAM',
    '',
    'setWebhook response: ' + code,
    {
      response: text
    }
  );

  if (code >= 300) {
    throw new Error('Telegram setWebhook error ' + code + ': ' + text);
  }

  return JSON.parse(text);
}

function deleteTelegramWebhook() {
  const token = PropertiesService
    .getScriptProperties()
    .getProperty('TELEGRAM_BOT_TOKEN');

  if (!token) {
    throw new Error('Missing TELEGRAM_BOT_TOKEN in Script Properties.');
  }

  const url =
    'https://api.telegram.org/bot' +
    token +
    '/deleteWebhook';

  const res = UrlFetchApp.fetch(url, {
    method: 'get',
    muteHttpExceptions: true
  });

  const text = res.getContentText();

  log_('INFO', 'TELEGRAM', '', 'deleteWebhook response', {
    response: text
  });

  return JSON.parse(text);
}

function getTelegramWebhookInfo() {
  const token = PropertiesService
    .getScriptProperties()
    .getProperty('TELEGRAM_BOT_TOKEN');

  if (!token) {
    throw new Error('Missing TELEGRAM_BOT_TOKEN in Script Properties.');
  }

  const url =
    'https://api.telegram.org/bot' +
    token +
    '/getWebhookInfo';

  const res = UrlFetchApp.fetch(url, {
    method: 'get',
    muteHttpExceptions: true
  });

  return JSON.parse(res.getContentText());
}

/************************************************************
 * TELEGRAM NOTIFY
 ************************************************************/

function notify_(projectId, message) {
  const settings = getSettings_();
  const props = PropertiesService.getScriptProperties();

  const enabledFromSheet =
    String(settings.ENABLE_TELEGRAM || '').toUpperCase() === 'TRUE';

  const enabledFromProps =
    String(props.getProperty('ENABLE_TELEGRAM') || '').toUpperCase() === 'TRUE';

  if (!enabledFromSheet && !enabledFromProps) {
    log_('INFO', projectId, '', 'Telegram skipped: ENABLE_TELEGRAM is not TRUE.', {});
    return {
      ok: false,
      skipped: true,
      reason: 'ENABLE_TELEGRAM_NOT_TRUE'
    };
  }

  const token = props.getProperty('TELEGRAM_BOT_TOKEN');

  const chatId =
    settings.TELEGRAM_CHAT_ID ||
    props.getProperty('TELEGRAM_CHAT_ID');

  if (!token) {
    log_('ERROR', projectId, '', 'Telegram skipped: missing TELEGRAM_BOT_TOKEN.', {});
    return {
      ok: false,
      error: 'MISSING_TELEGRAM_BOT_TOKEN'
    };
  }

  if (!chatId) {
    log_('ERROR', projectId, '', 'Telegram skipped: missing TELEGRAM_CHAT_ID.', {});
    return {
      ok: false,
      error: 'MISSING_TELEGRAM_CHAT_ID'
    };
  }

  const result = sendTelegramMessage_(
    chatId,
    '[' + projectId + '] ' + message
  );

  log_('INFO', projectId, '', 'Telegram sent.', {
    response: result
  });

  return {
    ok: true,
    response: result
  };
}

function enableTelegram() {
  const props = PropertiesService.getScriptProperties();

  props.setProperty('ENABLE_TELEGRAM', 'TRUE');

  const chatId =
    props.getProperty('TELEGRAM_CHAT_ID') ||
    setting_('TELEGRAM_CHAT_ID') ||
    '';

  setSetting_(
    'ENABLE_TELEGRAM',
    'TRUE',
    'TRUE nếu muốn gửi thông báo Telegram'
  );

  setSetting_(
    'TELEGRAM_CHAT_ID',
    chatId,
    'Chat ID nếu dùng Telegram'
  );

  Logger.log('Telegram enabled.');

  return {
    ok: true,
    message: 'Telegram enabled.'
  };
}

function disableTelegram() {
  PropertiesService
    .getScriptProperties()
    .setProperty('ENABLE_TELEGRAM', 'FALSE');

  setSetting_(
    'ENABLE_TELEGRAM',
    'FALSE',
    'TRUE nếu muốn gửi thông báo Telegram'
  );

  return {
    ok: true,
    message: 'Telegram disabled.'
  };
}

function testTelegram() {
  return notify_(
    'TEST',
    '✅ Project Brain OS Telegram đã kết nối thành công.'
  );
}

function resendDoneProjectTelegram() {
  const finalSh = getSheet_(PBOS.SHEETS.FINAL);
  const values = finalSh.getDataRange().getValues();

  if (values.length <= 1) {
    throw new Error('Chưa có dữ liệu trong Final_Output.');
  }

  const headers = values[0];
  const idx = indexMap_(headers);

  let sent = 0;

  for (let r = 1; r < values.length; r++) {
    const projectId = values[r][idx.Project_ID];
    const briefDocUrl = values[r][idx.Brief_Doc_URL];
    const status = values[r][idx.Status];

    if (projectId && status === 'DONE') {
      const project = findProject_(projectId);
      const taskSummary = getProjectTaskSummary_(projectId);

      const summary =
        '✅ PROJECT ĐÃ HOÀN TẤT\n\n' +
        '📌 Project: ' + project.Project_Name + '\n' +
        '🆔 ID: ' + projectId + '\n' +
        '📊 Health Score: ' + project.Health_Score + '\n' +
        '🤖 Agent hoàn tất: ' + taskSummary.done + '/' + taskSummary.total + '\n' +
        '⚠️ Agent lỗi: ' + taskSummary.failed + '\n\n' +
        '📄 Final Doc:\n' + briefDocUrl + '\n\n' +
        '📁 Drive Folder:\n' + project.Folder_URL + '\n\n' +
        '📋 Việc tiếp theo:\n' +
        '1. Mở Final Doc kiểm tra nội dung.\n' +
        '2. Đọc QA Review và Action_Plan.\n' +
        '3. Đánh dấu Action_Plan DONE nếu đã nghiệm thu.\n' +
        '4. Chạy thử dự án tiếp theo.';

      notify_(projectId, summary);
      sent++;
    }
  }

  return {
    ok: true,
    sent: sent
  };
}

/************************************************************
 * AGENTS
 ************************************************************/

function buildAgentTasks_(projectId, p) {
  const gemini = setting_('GEMINI_MODEL') || PBOS.DEFAULT_MODELS.GEMINI;
  const claude = setting_('CLAUDE_MODEL') || PBOS.DEFAULT_MODELS.CLAUDE;

  const list = [
    [
      'Master Orchestrator Agent',
      'complex',
      'CLAUDE',
      claude,
      'Chia dự án thành module, giao việc, xác định thứ tự ưu tiên, tạo Project Brief.',
      'Project Brief + Agent Assignment',
      'P1',
      'Ngày 1'
    ],
    [
      'Strategy Agent',
      'complex',
      'CLAUDE',
      claude,
      'Phân tích vấn đề, khách hàng, định vị, mô hình kinh doanh, lợi thế cạnh tranh.',
      'Strategy Plan',
      'P1',
      'Ngày 1-2'
    ],
    [
      'Research Agent',
      'simple',
      'GEMINI',
      gemini,
      'Xác định giả định cần kiểm chứng, dữ liệu cần tìm, đối thủ/công cụ cần so sánh.',
      'Research Checklist + Insight Map',
      'P2',
      'Ngày 2'
    ],
    [
      'Product Agent',
      'complex',
      'CLAUDE',
      claude,
      'Thiết kế MVP, tính năng cốt lõi, trải nghiệm người dùng, roadmap.',
      'MVP Spec + Roadmap',
      'P1',
      'Ngày 2'
    ],
    [
      'Operations Agent',
      'simple',
      'GEMINI',
      gemini,
      'Tạo SOP vận hành, checklist ngày/tuần, quy trình xử lý lỗi.',
      'SOP + Operating Checklist',
      'P2',
      'Ngày 3'
    ],
    [
      'Automation Agent',
      'complex',
      'CLAUDE',
      claude,
      'Thiết kế workflow Google Forms/Sheets/Apps Script/Docs/Drive/API/Telegram.',
      'Automation Blueprint',
      'P1',
      'Ngày 3'
    ],
    [
      'Finance Agent',
      'simple',
      'GEMINI',
      gemini,
      'Tạo mô hình chi phí, doanh thu, điểm hòa vốn, kịch bản xấu/trung bình/tốt.',
      'Finance Model Draft',
      'P2',
      'Ngày 4'
    ],
    [
      'Marketing Agent',
      'simple',
      'GEMINI',
      gemini,
      'Tạo USP, kênh marketing, content 7 ngày, cách tìm khách đầu tiên.',
      'Marketing Plan',
      'P2',
      'Ngày 4-5'
    ],
    [
      'Sales Agent',
      'simple',
      'GEMINI',
      gemini,
      'Tạo script tư vấn, báo giá, xử lý từ chối, follow-up.',
      'Sales Scripts',
      'P2',
      'Ngày 5'
    ],
    [
      'Legal & Risk Agent',
      'complex',
      'CLAUDE',
      claude,
      'Phát hiện rủi ro pháp lý, tài chính, vận hành, dữ liệu, đối tác.',
      'Risk Register',
      'P1',
      'Ngày 5'
    ],
    [
      'QA Agent',
      'complex',
      'CLAUDE',
      claude,
      'Kiểm tra toàn bộ đầu ra, phát hiện thiếu sót, mâu thuẫn, điểm phi thực tế.',
      'QA Review',
      'P1',
      'Ngày 6'
    ],
    [
      'Documentation Agent',
      'simple',
      'GEMINI',
      gemini,
      'Đóng gói thành tài liệu, template, checklist, prompt dùng lại cho dự án sau.',
      'Reusable Docs',
      'P2',
      'Ngày 7'
    ]
  ];

  return list.map((x, n) => ({
    taskId: projectId + '-T' + String(n + 1).padStart(2, '0'),
    agent: x[0],
    complexity: x[1],
    provider: x[2],
    model: x[3],
    mission: x[4],
    expectedOutput: x[5],
    priority: x[6],
    deadline: x[7]
  }));
}

function buildAgentPrompt_(project, task) {
  return `Bạn là ${task.Agent} trong hệ thống PROJECT BRAIN OS.

DỮ LIỆU DỰ ÁN:
- Project ID: ${project.Project_ID}
- Tên dự án: ${project.Project_Name}
- Giai đoạn: ${project.Stage}
- Ý tưởng: ${project.Idea}
- Mục tiêu 30 ngày: ${project.Goal_30}
- Người dùng/khách hàng: ${project.Target_User}
- Ngân sách: ${project.Budget}
- Deadline: ${project.Deadline}
- Nguồn lực: ${project.Resources}
- Ràng buộc/rủi ro: ${project.Constraints}
- Đầu ra mong muốn: ${project.Desired_Output}

NHIỆM VỤ CỦA BẠN:
${task.Mission}

ĐẦU RA BẮT BUỘC:
${task.Expected_Output}

YÊU CẦU CHẤT LƯỢNG:
- Viết bằng tiếng Việt.
- Không trả lời lý thuyết chung chung.
- Tạo bảng hành động cụ thể.
- Có checklist triển khai.
- Có rủi ro/cách xử lý nếu liên quan.
- Kết thúc bằng 3 bước hành động tiếp theo.
- Định dạng Markdown rõ ràng.`;
}

/************************************************************
 * MODEL CALLERS
 ************************************************************/

function callModel_(provider, model, prompt, maxTokens) {
  if (provider === 'GEMINI') {
    return callGemini_(model, prompt, maxTokens);
  }

  if (provider === 'CLAUDE') {
    return callClaude_(model, prompt, maxTokens);
  }

  throw new Error('Unknown provider: ' + provider);
}

function callGemini_(model, prompt, maxTokens) {
  const key = PropertiesService
    .getScriptProperties()
    .getProperty('GEMINI_API_KEY');

  if (!key) {
    throw new Error('Missing GEMINI_API_KEY in Script Properties.');
  }

  const url =
    'https://generativelanguage.googleapis.com/v1beta/models/' +
    encodeURIComponent(model) +
    ':generateContent?key=' +
    encodeURIComponent(key);

  const payload = {
    contents: [
      {
        role: 'user',
        parts: [
          {
            text: prompt
          }
        ]
      }
    ],
    generationConfig: {
      temperature: 0.4,
      maxOutputTokens: maxTokens || 4096
    }
  };

  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const code = res.getResponseCode();
  const text = res.getContentText();

  if (code >= 300) {
    throw new Error('Gemini API error ' + code + ': ' + text);
  }

  const json = JSON.parse(text);

  return (
    json.candidates?.[0]?.content?.parts
      ?.map(p => p.text || '')
      .join('\n') ||
    text
  );
}

function callClaude_(model, prompt, maxTokens) {
  const key = PropertiesService
    .getScriptProperties()
    .getProperty('CLAUDE_API_KEY');

  if (!key) {
    throw new Error('Missing CLAUDE_API_KEY in Script Properties.');
  }

  const payload = {
    model: model,
    max_tokens: maxTokens || 4096,
    temperature: 0.35,
    messages: [
      {
        role: 'user',
        content: prompt
      }
    ]
  };

  const res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
    headers: {
      'x-api-key': key,
      'anthropic-version': '2023-06-01'
    }
  });

  const code = res.getResponseCode();
  const text = res.getContentText();

  if (code >= 300) {
    throw new Error('Claude API error ' + code + ': ' + text);
  }

  const json = JSON.parse(text);

  return (
    (json.content || [])
      .map(c => c.text || '')
      .join('\n') ||
    text
  );
}

/************************************************************
 * SAVE OUTPUT
 ************************************************************/

function saveAgentOutput_(project, task, output) {
  const doc = DocumentApp.create(
    task.Task_ID + ' - ' + task.Agent
  );

  const body = doc.getBody();

  body
    .appendParagraph(task.Agent)
    .setHeading(DocumentApp.ParagraphHeading.TITLE);

  body.appendParagraph('Project: ' + project.Project_Name);
  body.appendParagraph('Provider: ' + task.Provider + ' | Model: ' + task.Model);
  body.appendHorizontalRule();
  body.appendParagraph(output);

  doc.saveAndClose();

  const folderId = extractDriveId_(project.Folder_URL);

  if (folderId) {
    DriveApp.getFileById(doc.getId()).moveTo(
      DriveApp.getFolderById(folderId)
    );
  }

  return doc.getUrl();
}

function createProjectFolder_(projectId, name) {
  const rootName =
    setting_('ROOT_FOLDER_NAME') ||
    'Project Brain OS Outputs';

  const root = getOrCreateFolder_(rootName);

  return root.createFolder(
    projectId + ' - ' + cleanName_(name)
  );
}

function getOrCreateFolder_(name) {
  const it = DriveApp.getFoldersByName(name);

  return it.hasNext()
    ? it.next()
    : DriveApp.createFolder(name);
}

function seedActionPlan_(projectId) {
  const rows = [
    [
      projectId + '-A01',
      projectId,
      'Ngày 1',
      'Kiểm tra Project Brief và mục tiêu thật sự',
      'Master Orchestrator',
      'Project Brief bản 1',
      'P1',
      'TODO',
      ''
    ],
    [
      projectId + '-A02',
      projectId,
      'Ngày 2',
      'Chốt MVP và khách hàng mục tiêu',
      'Strategy/Product',
      'MVP Spec',
      'P1',
      'TODO',
      ''
    ],
    [
      projectId + '-A03',
      projectId,
      'Ngày 3',
      'Tạo SOP và automation blueprint',
      'Ops/Automation',
      'SOP + Workflow',
      'P1',
      'TODO',
      ''
    ],
    [
      projectId + '-A04',
      projectId,
      'Ngày 4',
      'Tạo bảng tài chính và risk register',
      'Finance/Risk',
      'Finance + Risk',
      'P2',
      'TODO',
      ''
    ],
    [
      projectId + '-A05',
      projectId,
      'Ngày 5',
      'Tạo marketing + sales scripts',
      'Marketing/Sales',
      'Content + Script',
      'P2',
      'TODO',
      ''
    ],
    [
      projectId + '-A06',
      projectId,
      'Ngày 6',
      'QA toàn bộ hệ thống',
      'QA',
      'QA Review',
      'P1',
      'TODO',
      ''
    ],
    [
      projectId + '-A07',
      projectId,
      'Ngày 7',
      'Đóng gói template dùng lại',
      'Documentation',
      'Reusable Kit',
      'P1',
      'TODO',
      ''
    ]
  ];

  const sh = getSheet_(PBOS.SHEETS.ACTIONS);

  sh.getRange(
    sh.getLastRow() + 1,
    1,
    rows.length,
    rows[0].length
  ).setValues(rows);
}

/************************************************************
 * CONTROL COMMANDS
 ************************************************************/

function stopProjectBrainOS() {
  const props = PropertiesService.getScriptProperties();

  props.setProperty('PROJECT_BRAIN_OS_PAUSED', 'TRUE');

  setSetting_(
    'PROJECT_BRAIN_OS_PAUSED',
    'TRUE',
    'Trạng thái dừng tạm thời'
  );

  const triggers = ScriptApp.getProjectTriggers();

  let deleted = 0;

  triggers.forEach(trigger => {
    const fn = trigger.getHandlerFunction();

    if (
      fn === 'runAgentQueue' ||
      fn === 'ensureQueueTrigger' ||
      fn === 'processAgentQueue'
    ) {
      ScriptApp.deleteTrigger(trigger);
      deleted++;
    }
  });

  Logger.log('Project Brain OS paused. Deleted triggers: ' + deleted);

  return {
    ok: true,
    paused: true,
    deletedTriggers: deleted,
    message: 'Project Brain OS đã được dừng tạm thời.'
  };
}

function resumeProjectBrainOS() {
  const props = PropertiesService.getScriptProperties();

  props.setProperty('PROJECT_BRAIN_OS_PAUSED', 'FALSE');

  setSetting_(
    'PROJECT_BRAIN_OS_PAUSED',
    'FALSE',
    'Trạng thái dừng tạm thời'
  );

  ensureQueueTrigger_();

  Logger.log('Project Brain OS resumed.');

  return {
    ok: true,
    paused: false,
    message: 'Project Brain OS đã được bật chạy lại.'
  };
}

function checkProjectBrainOSStatus() {
  const paused = isPaused_();

  const triggers = ScriptApp
    .getProjectTriggers()
    .map(t => ({
      functionName: t.getHandlerFunction(),
      eventType: String(t.getEventType())
    }));

  const taskStats = getTaskStats_();

  const status = {
    ok: true,
    version: PBOS.VERSION,
    paused: paused,
    triggers: triggers,
    taskStats: taskStats
  };

  Logger.log(JSON.stringify(status, null, 2));

  return status;
}

function resetFailedTasks() {
  const sheet = getSheet_(PBOS.SHEETS.TASKS);

  if (!sheet) {
    throw new Error('Không tìm thấy sheet Agent_Tasks');
  }

  const data = sheet.getDataRange().getValues();

  if (data.length <= 1) {
    return {
      ok: true,
      resetCount: 0
    };
  }

  const headers = data[0];

  const statusCol = headers.indexOf('Status') + 1;
  const attemptCol = headers.indexOf('Attempt') + 1;
  const lastErrorCol = headers.indexOf('Last_Error') + 1;
  const updatedAtCol = headers.indexOf('Updated_At') + 1;

  if (!statusCol || !attemptCol || !lastErrorCol) {
    throw new Error('Thiếu cột Status, Attempt hoặc Last_Error trong Agent_Tasks');
  }

  let resetCount = 0;

  for (let r = 2; r <= data.length; r++) {
    const status = sheet.getRange(r, statusCol).getValue();

    if (status === 'FAILED') {
      sheet.getRange(r, statusCol).setValue('PENDING');
      sheet.getRange(r, attemptCol).setValue(0);
      sheet.getRange(r, lastErrorCol).setValue('');

      if (updatedAtCol) {
        sheet.getRange(r, updatedAtCol).setValue(new Date());
      }

      resetCount++;
    }
  }

  Logger.log('Đã reset ' + resetCount + ' task FAILED về PENDING');

  return {
    ok: true,
    resetCount: resetCount
  };
}

function resetRetryAndFailedTasks() {
  const sheet = getSheet_(PBOS.SHEETS.TASKS);

  if (!sheet) {
    throw new Error('Không tìm thấy sheet Agent_Tasks');
  }

  const data = sheet.getDataRange().getValues();

  if (data.length <= 1) {
    return {
      ok: true,
      resetCount: 0
    };
  }

  const headers = data[0];
  const idx = indexMap_(headers);

  let resetCount = 0;

  for (let r = 1; r < data.length; r++) {
    const status = data[r][idx.Status];

    if (
      status === 'FAILED' ||
      status === 'RETRY' ||
      status === 'RUNNING'
    ) {
      sheet.getRange(r + 1, idx.Status + 1).setValue('PENDING');
      sheet.getRange(r + 1, idx.Attempt + 1).setValue(0);
      sheet.getRange(r + 1, idx.Last_Error + 1).setValue('');
      sheet.getRange(r + 1, idx.Updated_At + 1).setValue(new Date());

      resetCount++;
    }
  }

  return {
    ok: true,
    resetCount: resetCount
  };
}

/************************************************************
 * HELPERS
 ************************************************************/

function ensureQueueTrigger_() {
  if (isPaused_()) {
    return;
  }

  const exists = ScriptApp
    .getProjectTriggers()
    .some(t => t.getHandlerFunction() === 'runAgentQueue');

  if (!exists) {
    ScriptApp
      .newTrigger('runAgentQueue')
      .timeBased()
      .everyMinutes(1)
      .create();
  }
}

function ensureSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss.getSheetByName(PBOS.SHEETS.PROJECTS)) {
    setupProjectBrainOS();
  }
}

function isPaused_() {
  const props = PropertiesService.getScriptProperties();

  const propPaused =
    String(props.getProperty('PROJECT_BRAIN_OS_PAUSED') || '')
      .toUpperCase() === 'TRUE';

  const settingPaused =
    String(setting_('PROJECT_BRAIN_OS_PAUSED') || '')
      .toUpperCase() === 'TRUE';

  return propPaused || settingPaused;
}

function getTaskStats_() {
  const sh = getSheet_(PBOS.SHEETS.TASKS);

  if (!sh) {
    return {};
  }

  const data = sh.getDataRange().getValues();

  if (data.length <= 1) {
    return {};
  }

  const h = data[0];
  const i = indexMap_(h);

  const stats = {};

  data.slice(1).forEach(r => {
    const s = r[i.Status] || 'EMPTY';
    stats[s] = (stats[s] || 0) + 1;
  });

  return stats;
}

function getProjectTaskSummary_(projectId) {
  const sh = getSheet_(PBOS.SHEETS.TASKS);
  const data = sh.getDataRange().getValues();

  if (data.length <= 1) {
    return {
      total: 0,
      done: 0,
      failed: 0
    };
  }

  const h = data[0];
  const i = indexMap_(h);

  const rows = data
    .slice(1)
    .filter(r => r[i.Project_ID] === projectId);

  return {
    total: rows.length,
    done: rows.filter(r => r[i.Status] === 'DONE').length,
    failed: rows.filter(r => r[i.Status] === 'FAILED').length
  };
}

function getSheet_(name) {
  return SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(name);
}

function append_(sheetName, row) {
  getSheet_(sheetName).appendRow(row);
}

function indexMap_(headers) {
  return headers.reduce((m, h, i) => {
    m[String(h).replace(/\s+/g, '_')] = i;
    return m;
  }, {});
}

function rowToObject_(headers, row) {
  const o = {};

  headers.forEach((h, i) => {
    o[String(h).replace(/\s+/g, '_')] = row[i];
  });

  return o;
}

function findProject_(projectId) {
  const data = getSheet_(PBOS.SHEETS.PROJECTS)
    .getDataRange()
    .getValues();

  const h = data[0];
  const i = indexMap_(h);

  const row = data
    .slice(1)
    .find(r => r[i.Project_ID] === projectId);

  if (!row) {
    throw new Error('Project not found: ' + projectId);
  }

  return rowToObject_(h, row);
}

function getSettings_() {
  const sh = getSheet_(PBOS.SHEETS.SETTINGS);

  if (!sh) {
    return {};
  }

  const rows = sh
    .getDataRange()
    .getValues()
    .slice(1);

  const out = {};

  rows.forEach(r => {
    if (r[0]) {
      out[r[0]] = r[1];
    }
  });

  return out;
}

function setting_(key) {
  return getSettings_()[key];
}

function setSetting_(key, value, note) {
  const sh = getSheet_(PBOS.SHEETS.SETTINGS);

  if (!sh) {
    return;
  }

  const data = sh.getDataRange().getValues();

  for (let r = 1; r < data.length; r++) {
    if (data[r][0] === key) {
      sh.getRange(r + 1, 2).setValue(value);

      if (note !== undefined) {
        sh.getRange(r + 1, 3).setValue(note);
      }

      return;
    }
  }

  sh.appendRow([
    key,
    value,
    note || ''
  ]);
}

function normalizeProject_(p) {
  return {
    name: p.projectName || p.name || 'Untitled Project',
    stage: p.projectStage || p.stage || 'Planning',
    idea: p.projectIdea || p.idea || '',
    goal30: p.goal30 || p.goal || '',
    targetUser: p.targetUser || p.user || '',
    budget: p.budget || '',
    deadline: p.deadline || '',
    resources: p.resources || '',
    constraints: p.constraints || '',
    desiredOutput: p.desiredOutput || p.output || ''
  };
}

function scoreProject_(p) {
  const d = normalizeProject_(p);

  let score = 45;

  if (d.name.length > 3) score += 5;
  if (d.idea.length > 80) score += 8;
  if (d.goal30.length > 30) score += 8;
  if (d.targetUser.length > 30) score += 8;
  if (d.budget.length > 3) score += 5;
  if (d.deadline.length > 2) score += 4;
  if (d.resources.length > 30) score += 6;
  if (d.constraints.length > 20) score += 5;
  if (d.desiredOutput.length > 30) score += 6;

  return Math.min(100, score);
}

function estimateQuality_(output) {
  const len = String(output || '').length;

  if (len > 6000) return 90;
  if (len > 3000) return 80;
  if (len > 1200) return 70;

  return 55;
}

function short_(s, n) {
  return String(s || '').slice(0, n || 200);
}

function cleanName_(s) {
  return String(s || '')
    .replace(/[\\/:*?"<>|#{}%~&]/g, '-')
    .slice(0, 80);
}

function extractDriveId_(url) {
  const m = String(url || '').match(/[-\w]{25,}/);
  return m ? m[0] : '';
}

function log_(level, projectId, taskId, message, data) {
  try {
    append_(PBOS.SHEETS.LOGS, [
      new Date(),
      level,
      projectId,
      taskId,
      message,
      JSON.stringify(data || {})
    ]);
  } catch (e) {
    Logger.log('Log failed: ' + e);
  }
}
