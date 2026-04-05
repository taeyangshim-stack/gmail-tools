/**
 * Gmail 중복 메일 정리 스크립트 (v3 - 개선판)
 *
 * Gmail API 일일 한도를 고려하여 배치 처리 적용
 * GmailApp.moveThreadsToTrash() 로 한 번에 여러 스레드 처리
 *
 * v3 변경사항:
 *   - 다중 메시지 스레드 내 개별 중복 메시지 삭제 지원
 *   - 각 단계별 실행 결과 요약 리포트
 *   - API 호출 에러 처리 (재시도 + 안전 중단)
 *   - DRY_RUN 시 전체 대상 수 미리보기
 *
 * 사용법:
 *   1) https://script.google.com 접속
 *   2) 새 프로젝트 생성
 *   3) 이 코드를 붙여넣기
 *   4) 단계별로 실행: step1_exosphere() -> step2_outlook() -> step3_duplicates()
 *   5) 한도 초과 시 24시간 후 다시 실행
 */

// ============================================================
// 설정
// ============================================================
var DRY_RUN = false;  // true: 미리보기만, false: 실제 삭제
var MAX_RETRIES = 2;  // API 호출 실패 시 재시도 횟수

// ============================================================
// 유틸리티: 안전한 배치 삭제 (재시도 포함)
// ============================================================
function safeTrashThreads_(threads) {
  for (var attempt = 0; attempt <= MAX_RETRIES; attempt++) {
    try {
      GmailApp.moveThreadsToTrash(threads);
      return true;
    } catch (e) {
      Logger.log("  ⚠ 삭제 실패 (시도 " + (attempt + 1) + "/" + (MAX_RETRIES + 1) + "): " + e.message);
      if (attempt < MAX_RETRIES) {
        Utilities.sleep(2000 * (attempt + 1));
      }
    }
  }
  Logger.log("  ✖ 배치 삭제 최종 실패 — 건너뜀 (" + threads.length + "스레드)");
  return false;
}

function safeTrashMessages_(messages) {
  for (var attempt = 0; attempt <= MAX_RETRIES; attempt++) {
    try {
      for (var i = 0; i < messages.length; i++) {
        messages[i].moveToTrash();
      }
      return true;
    } catch (e) {
      Logger.log("  ⚠ 메시지 삭제 실패 (시도 " + (attempt + 1) + "/" + (MAX_RETRIES + 1) + "): " + e.message);
      if (attempt < MAX_RETRIES) {
        Utilities.sleep(2000 * (attempt + 1));
      }
    }
  }
  Logger.log("  ✖ 메시지 삭제 최종 실패 — 건너뜀 (" + messages.length + "건)");
  return false;
}

// ============================================================
// STEP 1: 엑소스피어 보안알림 전체 삭제
// 별도 실행하여 한도 분산
// ============================================================
function step1_exosphere() {
  Logger.log("=== [STEP 1] 엑소스피어 보안알림 삭제 ===");
  Logger.log("DRY_RUN = " + DRY_RUN);

  var query = "from:no-reply@exosp.com";
  var totalThreads = 0;
  var totalMessages = 0;
  var failedThreads = 0;

  while (true) {
    var threads = GmailApp.search(query, 0, 100);
    if (threads.length === 0) break;

    var msgCount = 0;
    for (var i = 0; i < threads.length; i++) {
      msgCount += threads[i].getMessageCount();
    }

    if (!DRY_RUN) {
      if (!safeTrashThreads_(threads)) {
        failedThreads += threads.length;
        break;
      }
    }

    totalThreads += threads.length;
    totalMessages += msgCount;
    Logger.log("  배치 처리: " + threads.length + "스레드 (" + msgCount + "메시지), 누적 " + totalMessages + "건");

    if (DRY_RUN) {
      // DRY_RUN: 전체 대상 수 확인을 위해 계속 스캔
      var moreThreads = GmailApp.search(query, 100, 100);
      while (moreThreads.length > 0) {
        for (var j = 0; j < moreThreads.length; j++) {
          totalMessages += moreThreads[j].getMessageCount();
        }
        totalThreads += moreThreads.length;
        moreThreads = GmailApp.search(query, totalThreads, 100);
      }
      break;
    }
    Utilities.sleep(1000);
  }

  Logger.log("──────────────────────────────────");
  Logger.log("📊 STEP 1 결과 요약");
  Logger.log("  대상 스레드: " + totalThreads + "개");
  Logger.log("  대상 메시지: " + totalMessages + "건");
  Logger.log("  실패: " + failedThreads + "개");
  Logger.log("  모드: " + (DRY_RUN ? "미리보기" : "실제 삭제"));
  Logger.log("──────────────────────────────────");
}

// ============================================================
// STEP 2: Outlook 테스트 메시지 삭제
// ============================================================
function step2_outlook() {
  Logger.log("=== [STEP 2] Outlook 테스트 메시지 삭제 ===");
  Logger.log("DRY_RUN = " + DRY_RUN);

  var query = 'subject:"Microsoft Outlook 테스트 메시지"';
  var threads = GmailApp.search(query, 0, 100);
  var failed = false;

  if (threads.length === 0) {
    Logger.log("  대상 없음");
    Logger.log("──────────────────────────────────");
    Logger.log("📊 STEP 2 결과 요약: 대상 없음");
    Logger.log("──────────────────────────────────");
    return;
  }

  var msgCount = 0;
  for (var i = 0; i < threads.length; i++) {
    Logger.log("  대상: " + threads[i].getFirstMessageSubject() + " (" + threads[i].getMessageCount() + "건)");
    msgCount += threads[i].getMessageCount();
  }

  if (!DRY_RUN) {
    if (!safeTrashThreads_(threads)) {
      failed = true;
    }
  }

  Logger.log("──────────────────────────────────");
  Logger.log("📊 STEP 2 결과 요약");
  Logger.log("  대상 스레드: " + threads.length + "개");
  Logger.log("  대상 메시지: " + msgCount + "건");
  Logger.log("  실패: " + (failed ? "있음" : "없음"));
  Logger.log("  모드: " + (DRY_RUN ? "미리보기" : "실제 삭제"));
  Logger.log("──────────────────────────────────");
}

// ============================================================
// STEP 3: IMAP 중복 메일 탐지 및 삭제
// 동일 제목 + 발신자 + 시간(분 단위) 중 첫 번째만 보존
// 다중 메시지 스레드 내 개별 중복 메시지도 처리
// ============================================================
function step3_duplicates() {
  Logger.log("=== [STEP 3] IMAP 중복 메일 탐지 ===");
  Logger.log("DRY_RUN = " + DRY_RUN);

  var query = "-from:no-reply@exosp.com -subject:'Microsoft Outlook 테스트 메시지' -in:trash";
  var startIndex = 0;
  var batchSize = 50;
  var scannedCount = 0;
  var duplicateCount = 0;
  var deletedThreads = 0;
  var deletedMessages = 0;
  var failedCount = 0;
  var seen = {};
  var trashTargets = [];       // 단일 메시지 스레드 → 배치 삭제
  var messageTrashTargets = []; // 다중 메시지 스레드 내 개별 중복 메시지

  while (true) {
    var threads;
    try {
      threads = GmailApp.search(query, startIndex, batchSize);
    } catch (e) {
      Logger.log("  ⚠ 검색 실패: " + e.message);
      break;
    }
    if (threads.length === 0) break;

    for (var i = 0; i < threads.length; i++) {
      var messages = threads[i].getMessages();
      var duplicatesInThread = [];

      for (var j = 0; j < messages.length; j++) {
        var msg = messages[j];
        var subject = (msg.getSubject() || "").trim();
        var from = (msg.getFrom() || "").trim().toLowerCase();
        var dateKey = Utilities.formatDate(msg.getDate(), "Asia/Seoul", "yyyy-MM-dd HH:mm");

        var key = subject + "|" + from + "|" + dateKey;

        if (seen[key]) {
          duplicatesInThread.push(msg);
          duplicateCount++;
          Logger.log("  중복: [" + dateKey + "] " + subject.substring(0, 60));
        } else {
          seen[key] = true;
        }
        scannedCount++;
      }

      if (duplicatesInThread.length === 0) continue;

      if (messages.length === 1) {
        // 단일 메시지 스레드 → 스레드 단위 배치 삭제
        trashTargets.push(threads[i]);
      } else if (duplicatesInThread.length > 0) {
        // 다중 메시지 스레드 → 중복 메시지만 개별 삭제
        for (var k = 0; k < duplicatesInThread.length; k++) {
          messageTrashTargets.push(duplicatesInThread[k]);
        }
      }
    }

    // 스레드 배치 삭제 (100개 단위)
    if (trashTargets.length >= 100) {
      if (!DRY_RUN) {
        if (safeTrashThreads_(trashTargets)) {
          deletedThreads += trashTargets.length;
        } else {
          failedCount += trashTargets.length;
        }
        Logger.log("  >> " + trashTargets.length + "스레드 배치 삭제");
      }
      trashTargets = [];
    }

    // 개별 메시지 삭제 (50개 단위)
    if (messageTrashTargets.length >= 50) {
      if (!DRY_RUN) {
        if (safeTrashMessages_(messageTrashTargets)) {
          deletedMessages += messageTrashTargets.length;
        } else {
          failedCount += messageTrashTargets.length;
        }
        Logger.log("  >> " + messageTrashTargets.length + "건 개별 메시지 삭제");
      }
      messageTrashTargets = [];
    }

    startIndex += batchSize;
    Logger.log("  진행: " + scannedCount + "건 스캔, 중복 " + duplicateCount + "건");

    if (scannedCount > 2000) {
      Logger.log("  >> 2000건 도달. 나머지는 다시 실행하세요.");
      break;
    }

    Utilities.sleep(300);
  }

  // 남은 스레드 삭제
  if (trashTargets.length > 0 && !DRY_RUN) {
    if (safeTrashThreads_(trashTargets)) {
      deletedThreads += trashTargets.length;
    } else {
      failedCount += trashTargets.length;
    }
    Logger.log("  >> 마지막 " + trashTargets.length + "스레드 삭제");
  }

  // 남은 개별 메시지 삭제
  if (messageTrashTargets.length > 0 && !DRY_RUN) {
    if (safeTrashMessages_(messageTrashTargets)) {
      deletedMessages += messageTrashTargets.length;
    } else {
      failedCount += messageTrashTargets.length;
    }
    Logger.log("  >> 마지막 " + messageTrashTargets.length + "건 개별 메시지 삭제");
  }

  Logger.log("──────────────────────────────────");
  Logger.log("📊 STEP 3 결과 요약");
  Logger.log("  스캔: " + scannedCount + "건");
  Logger.log("  중복 발견: " + duplicateCount + "건");
  Logger.log("  삭제 스레드: " + deletedThreads + "개 (단일 메시지 스레드)");
  Logger.log("  삭제 메시지: " + deletedMessages + "건 (다중 메시지 스레드 내)");
  Logger.log("  실패: " + failedCount + "건");
  Logger.log("  모드: " + (DRY_RUN ? "미리보기" : "실제 삭제"));
  Logger.log("──────────────────────────────────");
}

// ============================================================
// 전체 실행 (한도 여유가 있을 때 사용)
// ============================================================
function runAll() {
  step1_exosphere();
  step2_outlook();
  step3_duplicates();
}
