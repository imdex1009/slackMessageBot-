// =============================================
// ⚙️ 설정값 - 본인 환경에 맞게 수정하세요
// =============================================
const CONFIG = {
  SLACK_BOT_TOKEN: 'xoxb-your-slack-bot-token',   // Slack Bot User OAuth Token

  // 발송 목록 시트 탭 이름 목록 (순서 무관, 개수 자유롭게 추가/삭제 가능)
  SHEET_NAMES: [
    '발송목록',
  ],

  LOG_SHEET_NAME:      '발송로그',      // 통합 로그 시트 탭 이름 (자동 생성)
  LOSS_LOG_SHEET_NAME: '예약소실로그',  // 예약 소실 전용 로그 시트 탭 이름 (자동 생성)

  // 컬럼 위치 (모든 시트 동일 구조)
  COL_CHANNEL:      1,   // A열: 발송 채널 ID (ex. C0XXXXXXXXX)
  COL_DATETIME:     2,   // B열: 발송 예약 일시
  COL_MESSAGE:      3,   // C열: 메시지 내용
  COL_STATUS:       4,   // D열: 발송완료여부 (자동 기재)
  COL_SCHEDULE_ID:  5,   // E열: Slack 예약 메시지 ID (자동 기재)
  COL_MENTION:      6,   // F열: 멘션 대상 (User ID / Group ID / @here 등)

  DONE_LABEL:       '✅ 발송완료',
  FAIL_LABEL:       '❌ 발송실패',
  SCHEDULED_LABEL:  '📅 발송예약됨',
  EXPIRE_LABEL:     '⏰ 발송기한 만료',
  NO_CHANNEL_LABEL: '⚠️ 채널 ID 누락',

  RETRY_COUNT:           3,    // 실패 시 최대 재시도 횟수
  RETRY_DELAY:           1000, // 재시도 간격 (밀리초)
  EXPIRE_MINUTES:        10,   // 발송 기한 만료 기준 (분) - 트리거 간격의 2배 권장
  SCHEDULE_WINDOW_MINUTES: 30, // 발송 시각 기준 이 시간(분) 이내 메시지만 예약 등록
  TIMEZONE:              'Asia/Seoul',
};

// =============================================
// 🚀 메인 함수 - 트리거로 주기적 실행됨
// =============================================
function sendScheduledMessages() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const now = new Date();

  // Slack 예약 목록 1회 조회 (폐기 예약 취소 + 발송 완료 확인에 공유)
  const slackScheduled    = fetchSlackScheduledMessages();
  const slackScheduledIds = slackScheduled ? new Set(slackScheduled.map(m => m.id)) : null;

  // 모든 시트를 대상으로 폐기 예약 자동 취소
  cancelDiscardedScheduledMessages(ss, slackScheduled);

  // 설정된 모든 시트를 순회하며 처리
  CONFIG.SHEET_NAMES.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      Logger.log(`⚠️ 시트를 찾을 수 없습니다: ${sheetName} (스킵)`);
      return;
    }

    Logger.log(`\n📋 [${sheetName}] 처리 시작`);
    processSheet(ss, sheet, sheetName, now, slackScheduledIds);
    Logger.log(`📋 [${sheetName}] 처리 완료`);
  });
}

// =============================================
// 📄 시트별 처리 함수
// =============================================
function processSheet(ss, sheet, sheetName, now, slackScheduledIds) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('  ℹ️ 발송할 데이터가 없습니다.');
    return;
  }

  const data = sheet.getRange(2, 1, lastRow - 1, 6).getValues(); // A~F열 읽기

  data.forEach((row, i) => {
    const [channelId, datetime, message, status, scheduleId, mention] = row;
    const rowNum = i + 2; // 실제 시트 행 번호

    // ① 완전히 처리된 행 스킵
    if (String(status).includes('발송완료')    ||
        String(status).includes('발송실패')    ||
        String(status).includes('발송기한 만료') ||
        String(status).includes('채널 ID 누락')) return;

    // ① - 예약됨 상태 처리
    if (String(status).includes('발송예약됨')) {
      if (!datetime) return;

      const sendAt = new Date(datetime);

      // Slack API 조회 실패 → 확인 불가, 다음 트리거에서 재확인
      if (slackScheduledIds === null) return;

      // Slack 목록에 예약 존재 → 정상 대기 중
      if (slackScheduledIds.has(String(scheduleId))) return;

      // Slack 목록에 예약이 없는 경우
      if (now < sendAt) {
        // 발송 시각 전인데 예약이 사라짐 → 상태·ID 초기화 후 즉시 재등록 시도
        sheet.getRange(rowNum, CONFIG.COL_STATUS).setValue('');
        sheet.getRange(rowNum, CONFIG.COL_SCHEDULE_ID).setValue('');

        const timeUntil = sendAt.getTime() - now.getTime();
        const windowMs  = CONFIG.SCHEDULE_WINDOW_MINUTES * 60 * 1000;

        if (timeUntil > 60 * 1000 && timeUntil <= windowMs) {
          // 30분 이내 → 즉시 재등록
          const mentionText = buildMentionText(mention);
          const fullMessage = mentionText ? `${mentionText}\n*${message}*` : `*${message}*`;
          const result = scheduleSlackMessage(fullMessage, String(channelId).trim(), sendAt);

          if (result.success) {
            sheet.getRange(rowNum, CONFIG.COL_STATUS).setValue(CONFIG.SCHEDULED_LABEL);
            sheet.getRange(rowNum, CONFIG.COL_SCHEDULE_ID).setValue(result.scheduleId);
            writeLog(ss, {
              sheetName,
              requestedAt: formatDate(sendAt),
              sentAt:      '-',
              message,
              mention:     mention || '-',
              channel:     channelId,
              status:      '예약재등록',
              error:       '-',
              retries:     0,
            });
            writeLossLog(ss, {
              detectedAt:    formatDate(now),
              sheetName,
              lostId:        scheduleId,
              channel:       channelId,
              scheduledAt:   formatDate(sendAt),
              message,
              result:        '재등록 성공',
              newScheduleId: result.scheduleId,
            });
            Logger.log(`  [${rowNum}행] 🔄 예약 소실 → 즉시 재등록 완료 → ${formatDate(sendAt)}`);
          } else {
            writeLossLog(ss, {
              detectedAt:    formatDate(now),
              sheetName,
              lostId:        scheduleId,
              channel:       channelId,
              scheduledAt:   formatDate(sendAt),
              message,
              result:        `재등록 실패: ${result.error}`,
              newScheduleId: '-',
            });
            Logger.log(`  [${rowNum}행] 🔄 예약 소실 → 재등록 실패 (다음 트리거 재시도): ${result.error}`);
          }
        } else if (timeUntil > windowMs) {
          // 30분 초과 → 아직 예약 시점 아님
          writeLossLog(ss, {
            detectedAt:    formatDate(now),
            sheetName,
            lostId:        scheduleId,
            channel:       channelId,
            scheduledAt:   formatDate(sendAt),
            message,
            result:        `대기 (${CONFIG.SCHEDULE_WINDOW_MINUTES}분 이내 도달 시 재등록)`,
            newScheduleId: '-',
          });
          Logger.log(`  [${rowNum}행] 🔄 예약 소실 감지 → 상태 초기화 (${CONFIG.SCHEDULE_WINDOW_MINUTES}분 이내 도달 시 재등록)`);
        } else {
          // 1분 미만 → 즉시 발송 대기 (이미 상태 초기화됨, 다음 루프에서 즉시 발송)
          writeLossLog(ss, {
            detectedAt:    formatDate(now),
            sheetName,
            lostId:        scheduleId,
            channel:       channelId,
            scheduledAt:   formatDate(sendAt),
            message,
            result:        '발송 임박 - 즉시 발송 대기',
            newScheduleId: '-',
          });
          Logger.log(`  [${rowNum}행] 🔄 예약 소실 → 발송 임박, 즉시 발송 대기`);
        }
      } else {
        // 발송 시각이 지남 → 발송 완료 처리
        sheet.getRange(rowNum, CONFIG.COL_STATUS).setValue(`${CONFIG.DONE_LABEL} (${formatDate(sendAt)})`);
        writeLog(ss, {
          sheetName,
          requestedAt: formatDate(sendAt),
          sentAt:      formatDate(sendAt),
          message,
          mention:     mention || '-',
          channel:     channelId,
          status:      '성공(예약발송)',
          error:       `예약ID: ${scheduleId}`,
          retries:     0,
        });
        Logger.log(`  [${rowNum}행] ✅ 예약 발송 완료 확인 → ID: ${scheduleId} / ${message.substring(0, 30)}`);
      }
      return;
    }

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

    // ④ 멘션 + 메시지 조합 (볼드 처리)
    const mentionText = buildMentionText(mention);
    const fullMessage = mentionText
      ? `${mentionText}\n*${message}*`
      : `*${message}*`;

    const timeUntil = sendAt.getTime() - now.getTime();
    const windowMs  = CONFIG.SCHEDULE_WINDOW_MINUTES * 60 * 1000;

    // ⑤ 발송 시각 30분 초과 → 아직 예약 시점 아님, 스킵
    if (timeUntil > windowMs) return;

    // ⑥ 발송 시각 1분~30분 이내 → Slack에 예약 등록
    if (timeUntil > 60 * 1000) {
      const result = scheduleSlackMessage(fullMessage, String(channelId).trim(), sendAt);

      if (result.success) {
        sheet.getRange(rowNum, CONFIG.COL_STATUS).setValue(CONFIG.SCHEDULED_LABEL);
        sheet.getRange(rowNum, CONFIG.COL_SCHEDULE_ID).setValue(result.scheduleId);
        writeLog(ss, {
          sheetName,
          requestedAt: formatDate(sendAt),
          sentAt:      '-',
          message,
          mention:     mention || '-',
          channel:     channelId,
          status:      '예약등록',
          error:       '-',
          retries:     0,
        });
        Logger.log(`  [${rowNum}행] 📅 Slack 예약 등록 완료 → ${formatDate(sendAt)} / ${message.substring(0, 30)}`);
      } else {
        Logger.log(`  [${rowNum}행] ⚠️ 예약 등록 실패 (다음 트리거에서 재시도): ${result.error}`);
      }
      return;
    }

    // ⑥ 발송 기한 만료 처리
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

    // ⑦ 발송 시각 1분 미만 or 이미 도달 → 즉시 발송 (fallback)
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
// 📅 Slack 메시지 예약 등록
// =============================================
function scheduleSlackMessage(message, channelId, sendAt) {
  const url    = 'https://slack.com/api/chat.scheduleMessage';
  const postAt = Math.floor(sendAt.getTime() / 1000); // Unix timestamp (초)

  const options = {
    method:             'post',
    contentType:        'application/json',
    headers:            { Authorization: `Bearer ${CONFIG.SLACK_BOT_TOKEN}` },
    payload:            JSON.stringify({
                          channel: channelId,
                          text:    message,
                          post_at: postAt,
                        }),
    muteHttpExceptions: true,
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const result   = JSON.parse(response.getContentText());

    if (result.ok) return { success: true, scheduleId: result.scheduled_message_id };
    return { success: false, error: result.error };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// =============================================
// 🗑️ 폐기 예약 메시지 취소
// - 시트에서 삭제된 행에 걸려있던 Slack 예약을 자동 취소
// =============================================
function cancelDiscardedScheduledMessages(ss, slackScheduled) {
  if (!slackScheduled) return;

  // 모든 시트에서 "발송예약됨" 상태인 scheduleId 수집
  const activeIds = new Set();
  CONFIG.SHEET_NAMES.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 2) return;

    sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues()
      .filter(row => String(row[CONFIG.COL_STATUS - 1]).includes('발송예약됨') && row[CONFIG.COL_SCHEDULE_ID - 1])
      .forEach(row => activeIds.add(String(row[CONFIG.COL_SCHEDULE_ID - 1])));
  });

  // 시트에 없는 ID → 삭제된 행이므로 Slack에서도 취소
  slackScheduled.forEach(msg => {
    if (activeIds.has(msg.id)) return;

    const result = deleteSlackScheduledMessage(msg.channel_id, msg.id);

    if (result.success) {
      writeLog(ss, {
        sheetName:   '-',
        requestedAt: '-',
        sentAt:      '-',
        message:     msg.text || '-',
        mention:     '-',
        channel:     msg.channel_id,
        status:      '폐기취소',
        error:       `시트에서 행 삭제됨 / 예약ID: ${msg.id}`,
        retries:     0,
      });
      Logger.log(`🗑️ 폐기 예약 취소 완료 → ID: ${msg.id} / 채널: ${msg.channel_id}`);
    } else {
      Logger.log(`⚠️ 폐기 예약 취소 실패 → ID: ${msg.id} / 오류: ${result.error}`);
    }
  });
}

// =============================================
// 📋 Slack 예약 메시지 목록 전체 조회 (페이지네이션 처리)
// =============================================
function fetchSlackScheduledMessages() {
  const messages = [];
  let cursor = '';

  // Slack API는 기본값으로 가까운 시간대 메시지만 반환하므로
  // 최대 예약 가능 기간(120일)까지 명시적으로 지정
  const latest = Math.floor(Date.now() / 1000) + (120 * 24 * 60 * 60);

  do {
    const base   = `https://slack.com/api/chat.scheduledMessages.list?latest=${latest}`;
    const url    = base + (cursor ? `&cursor=${cursor}` : '');
    const response = UrlFetchApp.fetch(url, {
      headers:            { Authorization: `Bearer ${CONFIG.SLACK_BOT_TOKEN}` },
      muteHttpExceptions: true,
    });
    const data = JSON.parse(response.getContentText());

    if (!data.ok) {
      Logger.log(`⚠️ Slack 예약 목록 조회 실패: ${data.error}`);
      return null;
    }

    messages.push(...data.scheduled_messages);
    cursor = (data.response_metadata && data.response_metadata.next_cursor) || '';
  } while (cursor);

  return messages;
}

// =============================================
// 🗑️ Slack 예약 메시지 취소
// =============================================
function deleteSlackScheduledMessage(channelId, scheduleId) {
  const url = 'https://slack.com/api/chat.deleteScheduledMessage';

  const options = {
    method:             'post',
    contentType:        'application/json',
    headers:            { Authorization: `Bearer ${CONFIG.SLACK_BOT_TOKEN}` },
    payload:            JSON.stringify({
                          channel:              channelId,
                          scheduled_message_id: scheduleId,
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
// F열 입력 형식 - 쉼표(,)로 여러 명 지정 가능
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
// 🚨 예약 소실 전용 로그 시트 기록 (없으면 자동 생성)
// =============================================
function writeLossLog(ss, log) {
  let sheet = ss.getSheetByName(CONFIG.LOSS_LOG_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.LOSS_LOG_SHEET_NAME);
    const headers = [
      '감지 시각', '시트명', '소실 예약ID', '채널ID',
      '발송 예정 일시', '메시지 내용', '처리 결과', '새 예약ID',
    ];
    sheet.appendRow(headers);

    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#e8430a')
               .setFontColor('#ffffff')
               .setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 160); // 감지 시각
    sheet.setColumnWidth(3, 150); // 소실 예약ID
    sheet.setColumnWidth(5, 160); // 발송 예정 일시
    sheet.setColumnWidth(6, 280); // 메시지 내용
    sheet.setColumnWidth(7, 200); // 처리 결과
    sheet.setColumnWidth(8, 150); // 새 예약ID
  }

  sheet.appendRow([
    log.detectedAt,
    log.sheetName,
    log.lostId,
    log.channel,
    log.scheduledAt,
    log.message,
    log.result,
    log.newScheduleId,
  ]);
}

// =============================================
// ⏰ 트리거 설정 (최초 1회만 실행)
// - 다음 5분 단위 정각(예: 0:05, 0:10)부터 실행되도록 정렬
// =============================================
function createTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));

  // 다음 5분 단위 정각 계산 (현재 시각이 0:02 → 0:05 / 0:07 → 0:10)
  const now         = new Date();
  const interval    = 5 * 60 * 1000;
  const nextAligned = new Date(Math.ceil((now.getTime() + 1) / interval) * interval);

  // ① 다음 정각에 일회성 트리거로 initTrigger 호출
  ScriptApp.newTrigger('initTrigger')
    .timeBased()
    .at(nextAligned)
    .create();

  Logger.log(`✅ 첫 실행 예약: ${formatDate(nextAligned)} → 이후 5분마다 자동 실행`);
}

// =============================================
// ⏰ 5분 단위 정렬 초기화 - createTrigger에서 자동 호출됨
// =============================================
function initTrigger() {
  // 일회성 트리거 제거
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'initTrigger')
    .forEach(t => ScriptApp.deleteTrigger(t));

  // ② 이 시점(5분 정각)부터 5분 주기 트리거 시작
  ScriptApp.newTrigger('sendScheduledMessages')
    .timeBased()
    .everyMinutes(5)
    .create();

  // 즉시 첫 실행
  sendScheduledMessages();

  Logger.log('✅ 5분 단위 정렬 완료 - sendScheduledMessages 주기 실행 시작');
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

// =============================================
// 🔍 봇이 등록한 예약 메시지 목록 조회
// Apps Script 편집기에서 직접 실행 후 로그 확인
// =============================================
function listScheduledMessages() {
  const messages = fetchSlackScheduledMessages();

  if (!messages) {
    Logger.log('❌ 예약 메시지 목록 조회 실패');
    return;
  }

  if (messages.length === 0) {
    Logger.log('ℹ️ 현재 등록된 예약 메시지가 없습니다.');
    return;
  }

  Logger.log(`📋 예약 메시지 총 ${messages.length}건\n`);

  messages
    .sort((a, b) => a.post_at - b.post_at) // 발송 시각 오름차순 정렬
    .forEach((msg, idx) => {
      const postAt     = new Date(msg.post_at * 1000);
      const postAtStr  = formatDate(postAt);
      const text       = (msg.text || '').substring(0, 50);
      Logger.log(
        `[${idx + 1}] 예약ID: ${msg.id}\n` +
        `     채널ID: ${msg.channel_id}\n` +
        `     발송시각: ${postAtStr}\n` +
        `     메시지: ${text}${msg.text && msg.text.length > 50 ? '...' : ''}`
      );
    });
}
