/**
 * Gmail 중복 메일 정리 스크립트 (v2 - API 호출 최소화)
 *
 * Gmail API 일일 한도를 고려하여 배치 처리 적용
 * GmailApp.moveThreadsToTrash() 로 한 번에 여러 스레드 처리
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

// ============================================================
// STEP 1: 엑소스피어 보안알림 전체 삭제
// 별도 실행하여 한도 분산
// ============================================================
function step1_exosphere() {
  Logger.log("--- [STEP 1] 엑소스피어 보안알림 삭제 ---");
  Logger.log("DRY_RUN = " + DRY_RUN);

  var query = "from:no-reply@exosp.com";
  var totalCount = 0;

  // 한 번에 최대 100 스레드씩, 배치로 휴지통 이동
  while (true) {
    var threads = GmailApp.search(query, 0, 100);
    if (threads.length === 0) break;

    var msgCount = 0;
    for (var i = 0; i < threads.length; i++) {
      msgCount += threads[i].getMessageCount();
    }

    if (!DRY_RUN) {
      // 핵심: 배치로 한 번에 처리 (API 호출 1회)
      GmailApp.moveThreadsToTrash(threads);
    }

    totalCount += msgCount;
    Logger.log("  배치 처리: " + threads.length + "스레드 (" + msgCount + "메시지), 누적 " + totalCount + "건");

    if (DRY_RUN) break;
    Utilities.sleep(1000);
  }

  Logger.log("완료: 엑소스피어 " + totalCount + "건 처리");
}

// ============================================================
// STEP 2: Outlook 테스트 메시지 삭제
// ============================================================
function step2_outlook() {
  Logger.log("--- [STEP 2] Outlook 테스트 메시지 삭제 ---");
  Logger.log("DRY_RUN = " + DRY_RUN);

  var query = 'subject:"Microsoft Outlook 테스트 메시지"';
  var threads = GmailApp.search(query, 0, 100);

  if (threads.length === 0) {
    Logger.log("  대상 없음");
    return;
  }

  var msgCount = 0;
  for (var i = 0; i < threads.length; i++) {
    Logger.log("  대상: " + threads[i].getFirstMessageSubject() + " (" + threads[i].getMessageCount() + "건)");
    msgCount += threads[i].getMessageCount();
  }

  if (!DRY_RUN) {
    GmailApp.moveThreadsToTrash(threads);
  }

  Logger.log("완료: Outlook 테스트 " + msgCount + "건 처리");
}

// ============================================================
// STEP 3: IMAP 중복 메일 탐지 및 삭제
// 동일 제목 + 발신자 + 시간(분 단위) 중 첫 번째만 보존
// ============================================================
function step3_duplicates() {
  Logger.log("--- [STEP 3] IMAP 중복 메일 탐지 ---");
  Logger.log("DRY_RUN = " + DRY_RUN);

  // 엑소스피어/Outlook은 이미 처리했으므로 제외
  var query = "-from:no-reply@exosp.com -subject:'Microsoft Outlook 테스트 메시지' -in:trash";
  var startIndex = 0;
  var batchSize = 50;  // 한 번에 50 스레드씩 (API 절약)
  var scannedCount = 0;
  var duplicateCount = 0;
  var seen = {};
  var trashTargets = [];  // 삭제 대상 스레드 모아두기

  while (true) {
    var threads = GmailApp.search(query, startIndex, batchSize);
    if (threads.length === 0) break;

    for (var i = 0; i < threads.length; i++) {
      var messages = threads[i].getMessages();
      var threadHasDuplicate = false;

      for (var j = 0; j < messages.length; j++) {
        var msg = messages[j];
        var subject = (msg.getSubject() || "").trim();
        var from = (msg.getFrom() || "").trim().toLowerCase();
        var dateKey = Utilities.formatDate(msg.getDate(), "Asia/Seoul", "yyyy-MM-dd HH:mm");

        var key = subject + "|" + from + "|" + dateKey;

        if (seen[key]) {
          threadHasDuplicate = true;
          duplicateCount++;
          Logger.log("  중복: [" + dateKey + "] " + subject.substring(0, 60));
        } else {
          seen[key] = true;
        }
        scannedCount++;
      }

      // 스레드 내 모든 메시지가 중복인 경우만 스레드를 휴지통으로
      // (스레드에 원본+중복이 섞여있으면 개별 메시지 단위 처리)
      if (threadHasDuplicate && messages.length === 1) {
        trashTargets.push(threads[i]);
      }
    }

    // 삭제 대상이 100개 쌓이면 배치 처리
    if (trashTargets.length >= 100) {
      if (!DRY_RUN) {
        GmailApp.moveThreadsToTrash(trashTargets);
        Logger.log("  >> " + trashTargets.length + "스레드 배치 삭제");
      }
      trashTargets = [];
    }

    startIndex += batchSize;
    Logger.log("  진행: " + scannedCount + "건 스캔, 중복 " + duplicateCount + "건");

    // 실행 시간 제한 방지 (5분 안에 끝내기)
    if (scannedCount > 2000) {
      Logger.log("  >> 2000건 도달. 나머지는 다시 실행하세요.");
      break;
    }

    Utilities.sleep(300);
  }

  // 남은 삭제 대상 처리
  if (trashTargets.length > 0 && !DRY_RUN) {
    GmailApp.moveThreadsToTrash(trashTargets);
    Logger.log("  >> 마지막 " + trashTargets.length + "스레드 삭제");
  }

  Logger.log("완료: " + scannedCount + "건 스캔, 중복 " + duplicateCount + "건 처리");
}

// ============================================================
// 전체 실행 (한도 여유가 있을 때 사용)
// ============================================================
function runAll() {
  step1_exosphere();
  step2_outlook();
  step3_duplicates();
}
