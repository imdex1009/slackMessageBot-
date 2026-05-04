// =============================================
// ⚙️ 설정값 - 본인 환경에 맞게 수정하세요
// =============================================
const CONFIG = {
  SLACK_BOT_TOKEN: 'xoxb-your-slack-bot-token',   // Slack Bot User OAuth Token

  // 발송 목록 시트 탭 이름 목록 (순서 무관, 개수 자유롭게 추가/삭제 가능)
  SHEET_NAMES: [
    '발송목록',
  ],

  LOG_SHEET_NAME: '발송로그',  // 통합 로그 시트 탭 이름 (자동 생성)

  // 컬럼 위치 (모든 시트 동일 구조)
  COL_CHANNEL:  1,   // A열: 발송 채널 ID (ex. C0XXXXXXXXX)
  COL_DATETIME: 2,   // B열: 발송 예약 일시
  COL_MESSAGE:  3,   // C열: 메시지 내용
  COL_STATUS:   4,   // D열: 발송완료여부 (자동 기재)
  COL_MENTION:  5,   // E열: 멘션 대상 (User ID / Group ID / @here 등)

  DONE_LABEL:       '✅ 발송완료',
  FAIL_LABEL:       '❌ 발송실패',
  EXPIRE_LABEL:     '⏰ 발송기한 만료',
  NO_CHANNEL_LABEL: '⚠️ 채널 ID 누락',

  RETRY_COUNT:    3,    // 실패 시 최대 재시도 횟수
  RETRY_DELAY:    1000, // 재시도 간격 (밀리초)
  EXPIRE_MINUTES: 3,    // 발송 기한 만료 기준 (분) - 1분 트리거 기준 2~3분 권장
  TIMEZONE:       'Asia/Seoul',
};

// =============================================
// 🚀 메인 함수 - 트리거로 주기적 실행됨
// =============================================
function sendScheduledMessages() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const now = new Date();

  CONFIG.SHEET_NAMES.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      Logger.log(`⚠️ 시트를 찾을 수 없습니다: ${sheetName} (스킵)`);
      return;
    }

    Logger.log(`\n📋 [${sheetName}] 처리 시작`);
    processSheet(ss, sheet, sheetName, now);
    Logger.log(`📋 [${sheetName}] 처리 완료`);
  });
}

// =============================================
// 📄 시트별 처리 함수
// =============================================
function processSheet(ss, sheet, sheetName, now) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('  ℹ️ 발송할 데이터가 없습니다.');
    return;
  }

  const data = sheet.getRange(2, 1, lastRow - 1, 5).getValues(); // A~E열 읽기

  data.forEach((row, i) => {
    const [channelId, datetime, message, status, mention] = row;
    const rowNum = i + 2;

    // ① 이미 처리된 행 스킵
    if (String(status).includes('발송완료')    ||
        String(status).includes('발송실패')    ||
        String(status).includes('발송기한 만료') ||
        String(status).includes('채널 ID 누락')) return;

    // ② 필수값 누락 스킵
    if (!datetime || !message) return;

    // ③ 채널 ID 누락 처리
    if (!channelId) {
      sheet.getRange(rowNum, CONFIG.COL_STATUS).setValue(CONFIG.NO_CHANNEL_LABEL);
      writeLog(ss, {
        sheetName,
        requestedAt: formatDate(new Date(datetime)),
        sentAt:      '-',
        message,
        mention:     mention || '-',
        channel:     '-',
        status:      '채널누락',
        error:       'A열 채널 ID가 비어 있음',
        retries:     0,
      });
      Logger.log(`  [${rowNum}행] ⚠️ 채널 ID 누락`);
      return;
    }

    const sendAt   = new Date(datetime);
    const expireAt = new Date(sendAt.getTime() + CONFIG.EXPIRE_MINUTES * 60 * 1000);

    // ④ 발송 시각 1분 초과로 남았으면 스킵, 1분 이내면 즉시 발송 진행
    if (sendAt.getTime() - now.getTime() > 60 * 1000) return;

    // ⑤ 발송 기한 만료 처리
    if (now > expireAt) {
      sheet.getRange(rowNum, CONFIG.COL_STATUS).setValue(CONFIG.EXPIRE_LABEL);
      writeLog(ss, {
        sheetName,
        requestedAt: formatDate(sendAt),
        sentAt:      '-',
        message,
        mention:     mention || '-',
        channel:     channelId,
        status:      '만료',
        error:       `발송 예정 시간 ${CONFIG.EXPIRE_MINUTES}분 초과`,
        retries:     0,
      });
      Logger.log(`  [${rowNum}행] ⏰ 발송기한 만료 - ${message.substring(0, 30)}`);
      return;
    }

    // ⑥ 멘션 + 메시지 조합 (볼드 처리)
    const mentionText = buildMentionText(mention);
    const fullMessage = mentionText
      ? `${mentionText}\n*${message}*`
      : `*${message}*`;

    // ⑦ 즉시 발송 (재시도 포함)
    const result = sendWithRetry(fullMessage, String(channelId).trim());

    // ⑧ 시트 상태 업데이트
    const statusValue = result.success
      ? `${CONFIG.DONE_LABEL} (${formatDate(now)})`
      : `${CONFIG.FAIL_LABEL} (${result.error})`;
    sheet.getRange(rowNum, CONFIG.COL_STATUS).setValue(statusValue);

    // ⑨ 통합 로그 기록
    writeLog(ss, {
      sheetName,
      requestedAt: formatDate(sendAt),
      sentAt:      result.success ? formatDate(now) : '-',
      message,
      mention:     mention || '-',
      channel:     channelId,
      status:      result.success ? '성공' : '실패',
      error:       result.error   || '-',
      retries:     result.retries,
    });

    Logger.log(`  [${rowNum}행] ${result.success ? '✅ 발송 성공' : '❌ 발송 실패'} → 채널: ${channelId} / ${message.substring(0, 30)}`);
  });
}

// =============================================
// 🔁 재시도 포함 즉시 발송 함수
// =============================================
function sendWithRetry(message, channelId) {
  let lastError = '';

  for (let attempt = 1; attempt <= CONFIG.RETRY_COUNT; attempt++) {
    const result = sendSlackMessage(message, channelId);

    if (result.success) {
      return { success: true, retries: attempt - 1 };
    }

    lastError = result.error;
    Logger.log(`  ⚠️ 발송 시도 ${attempt}/${CONFIG.RETRY_COUNT} 실패: ${lastError}`);

    if (attempt < CONFIG.RETRY_COUNT) {
      Utilities.sleep(CONFIG.RETRY_DELAY);
    }
  }

  return { success: false, error: lastError, retries: CONFIG.RETRY_COUNT };
}

// =============================================
// 📤 Slack 메시지 즉시 발송
// =============================================
function sendSlackMessage(message, channelId) {
  const url = 'https://slack.com/api/chat.postMessage';

  const options = {
    method:             'post',
    contentType:        'application/json',
    headers:            { Authorization: `Bearer ${CONFIG.SLACK_BOT_TOKEN}` },
    payload:            JSON.stringify({
                          channel: channelId,
                          text:    message,
                        }),
    muteHttpExceptions: true,
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const result   = JSON.parse(response.getContentText());

    if (result.ok) return { success: true };
    return { success: false, error: result.error };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// =============================================
// 💬 멘션 텍스트 생성
// -----------------------------------------------
// E열 입력 형식 - 쉼표(,)로 여러 명 지정 가능
//
//   단일:   U0XXXXXXXXX
//   복수:   U0XXXXXXXXX, U0YYYYYYYYY, @here
//
//   U0XXXXXXXXX  → 개인 유저 멘션       <@U0XXXXXXXXX>
//   S0XXXXXXXXX  → User Group 멘션     <!subteam^S0XXXXXXXXX>
//   @here        → 온라인 채널 멤버     <!here>
//   @channel     → 채널 전체 멤버      <!channel>
//   @everyone    → 워크스페이스 전체   <!everyone>
// =============================================
function buildMentionText(mention) {
  if (!mention) return '';

  const tokens = String(mention).split(',').map(t => t.trim()).filter(Boolean);

  const mentions = tokens.map(token => {
    if (token === '@here')     return '<!here>';
    if (token === '@channel')  return '<!channel>';
    if (token === '@everyone') return '<!everyone>';
    if (token.startsWith('S')) return `<!subteam^${token}>`;  // User Group
    if (token.startsWith('U')) return `<@${token}>`;          // 개인 User
    return token;
  });

  return mentions.join(' ');
}

// =============================================
// 📋 통합 로그 시트 기록 (없으면 자동 생성)
// =============================================
function writeLog(ss, log) {
  let logSheet = ss.getSheetByName(CONFIG.LOG_SHEET_NAME);

  if (!logSheet) {
    logSheet = ss.insertSheet(CONFIG.LOG_SHEET_NAME);
    const headers = [
      '시트명', '예약발송일시', '실제발송일시', '메시지내용',
      '멘션대상', '발송채널ID', '발송결과', '오류내용', '재시도횟수',
    ];
    logSheet.appendRow(headers);

    const headerRange = logSheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#1a73e8')
               .setFontColor('#ffffff')
               .setFontWeight('bold');
    logSheet.setFrozenRows(1);
    logSheet.setColumnWidth(1, 120); // 시트명
    logSheet.setColumnWidth(4, 300); // 메시지 내용
    logSheet.setColumnWidth(6, 150); // 채널 ID
  }

  logSheet.appendRow([
    log.sheetName,
    log.requestedAt,
    log.sentAt,
    log.message,
    log.mention,
    log.channel,
    log.status,
    log.error,
    log.retries,
  ]);
}

// =============================================
// ⏰ 트리거 설정 (최초 1회만 실행)
// =============================================
function createTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('sendScheduledMessages')
    .timeBased()
    .everyMinutes(1)
    .create();

  Logger.log('✅ 트리거 생성 완료 - sendScheduledMessages 1분마다 실행');
}

// =============================================
// 🛠️ 유틸: 날짜 포맷
// =============================================
function formatDate(date) {
  return Utilities.formatDate(date, CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
}

// =============================================
// 🔍 채널 목록 조회 (채널 ID 확인용)
// Apps Script 편집기에서 직접 실행 후 로그 확인
// =============================================
function listChannels() {
  const url      = 'https://slack.com/api/conversations.list';
  const response = UrlFetchApp.fetch(url, {
    headers: { Authorization: `Bearer ${CONFIG.SLACK_BOT_TOKEN}` },
  });
  const data = JSON.parse(response.getContentText());

  if (!data.ok) {
    Logger.log(`❌ 조회 실패: ${data.error}`);
    return;
  }

  data.channels.forEach(c => {
    Logger.log(`채널명: #${c.name} | ID: ${c.id}`);
  });
}

// =============================================
// 🔍 User Group 목록 조회 (ID 확인용)
// Apps Script 편집기에서 직접 실행 후 로그 확인
// =============================================
function listUserGroups() {
  const url      = 'https://slack.com/api/usergroups.list';
  const response = UrlFetchApp.fetch(url, {
    headers: { Authorization: `Bearer ${CONFIG.SLACK_BOT_TOKEN}` },
  });
  const data = JSON.parse(response.getContentText());

  if (!data.ok) {
    Logger.log(`❌ 조회 실패: ${data.error}`);
    return;
  }

  data.usergroups.forEach(g => {
    Logger.log(`이름: ${g.name} | 핸들: @${g.handle} | ID: ${g.id}`);
  });
}

// =============================================
// 🔍 전체 멤버 목록 조회 (User ID 확인용)
// =============================================
function listAllMembers() {
  const url      = 'https://slack.com/api/users.list';
  const response = UrlFetchApp.fetch(url, {
    headers: { Authorization: `Bearer ${CONFIG.SLACK_BOT_TOKEN}` },
  });
  const data = JSON.parse(response.getContentText());

  if (!data.ok) {
    Logger.log(`❌ 조회 실패: ${data.error}`);
    return;
  }

  data.members
    .filter(m => !m.is_bot && !m.deleted)
    .forEach(m => {
      Logger.log(`이름: ${m.real_name} | ID: ${m.id}`);
    });
}
