/**
 * 전주솔내고등학교 좌석 중심 컴퓨터·모니터 관리 웹앱 (Google Apps Script)
 * 
 * - 시트 구성: [좌석], [기기], [교무실]
 * - 사진 파일: 구글 드라이브 전용 폴더에 자동 업로드하여 링크로 저장 (5만자 한도 해결)
 * - Gemini Vision API: PC/모니터 라벨 자동 판독
 * - 초기데이터생성: 전주솔내고 2026 좌석배치도 98석 자동 생성
 */

const SHEET_SEATS = '좌석';
const SHEET_DEVICES = '기기';
const SHEET_ROOMS = '교무실';
const PHOTO_FOLDER_NAME = '[전주솔내고] 기기관리_사진';

/** 웹앱 진입점 (일반교사용: ?page=teacher, 관리자용: ?page=admin 또는 기본) */
function doGet(e) {
  // TEMP DEBUG: ?debug=1 로 접속하면 진단용 순수 텍스트 결과를 바로 반환 (원인 확인 후 제거 예정)
  if (e && e.parameter && e.parameter.debug === '1') {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const seatSheet = ss.getSheetByName(SHEET_SEATS);
      const seats = seatSheet ? getSheetObjects_(seatSheet) : [];
      const g5 = seats.find(s => s.seat_id === 'GYOMU_05');
      const lines = [
        'ss.getName()=' + ss.getName(),
        'seats.length=' + seats.length,
        'GYOMU_05=' + JSON.stringify(g5)
      ];
      return ContentService.createTextOutput(lines.join('\n'));
    } catch (err) {
      return ContentService.createTextOutput('진단 오류: ' + err.message);
    }
  }

  const page = (e && e.parameter && e.parameter.page) ? e.parameter.page : 'admin';
  const file = (page === 'teacher') ? 'teacher' : 'index';
  const title = (page === 'teacher')
    ? '전주솔내고 교직원 컴퓨터·모니터 현황 조사 (교사용)'
    : '전주솔내고 좌석별 PC·모니터 관리 시스템 (관리자)';

  return HtmlService.createHtmlOutputFromFile(file)
    .setTitle(title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** GitHub Pages 등 외부 정적 페이지에서 오는 fetch 요청 처리 (index.html/teacher.html의 gasRunner()가 호출) */
const ALLOWED_FUNCTIONS_ = {
  getInitialData, saveDevice, deleteDevice, updateSeatUser, saveSeatDevices,
  batchUpdateSeatUsers, analyzeLabelImage, addSeat, deleteSeat, updateSeatInfo,
  updateSeatPosition, batchUpdateSeatPositions
};
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const fn = ALLOWED_FUNCTIONS_[body.fn];
    if (!fn) throw new Error('허용되지 않은 함수: ' + body.fn);
    const result = fn.apply(null, body.args || []);
    return ContentService.createTextOutput(JSON.stringify(result === undefined ? { success: true } : result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/** HTML 파일 include 헬퍼 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/** 전체 초기 데이터(좌석, 기기, 교무실) 일괄 반환 */
function getInitialData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ensureSheetsInitialized_(ss);

    let seatSheet = ss.getSheetByName(SHEET_SEATS);
    let seats = seatSheet ? getSheetObjects_(seatSheet) : [];

    if (!seats || seats.length === 0) {
      초기데이터생성();
      seatSheet = ss.getSheetByName(SHEET_SEATS);
      seats = seatSheet ? getSheetObjects_(seatSheet) : [];
    }

    // 최후의 안전장치: 시트가 여전히 비어있다면 메모리 프리셋 98석 즉시 반환
    if (!seats || seats.length === 0) {
      const presetRooms = getPresetRooms_();
      seats = [];
      presetRooms.forEach(room => {
        room.seats.forEach(s => {
          seats.push({
            seat_id: s.id,
            room_id: room.id,
            room_name: room.name,
            seat_label: s.label,
            current_user: s.user || '',
            extension: s.ext || '',
            updated_at: ''
          });
        });
      });
    }

    const devSheet = ss.getSheetByName(SHEET_DEVICES);
    const roomSheet = ss.getSheetByName(SHEET_ROOMS);
    const devices = getSheetObjects_(devSheet);
    const rooms = getSheetObjects_(roomSheet);

    let webAppUrl = '';
    try {
      webAppUrl = ScriptApp.getService().getUrl();
    } catch (e) {}

    // TEMP DEBUG: 동기화 진단용 로그 (원인 확인 후 제거 예정)
    const g5 = seats.find(s => s.seat_id === 'GYOMU_05');
    Logger.log('getInitialData 진단: seats.length=' + seats.length + ', GYOMU_05=' + JSON.stringify(g5));

    return {
      success: true,
      seats: seats,
      devices: devices,
      rooms: rooms,
      webAppUrl: webAppUrl,
      serverTime: new Date().toISOString()
    };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

/** 기기 정보 저장 (신규 등록 및 기존 수정) */
function saveDevice(payload) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const devSheet = ss.getSheetByName(SHEET_DEVICES);
    const headers = devSheet.getRange(1, 1, 1, devSheet.getLastColumn()).getValues()[0];
    
    // device_id가 없으면 신규 생성
    let deviceId = payload.device_id;
    if (!deviceId) {
      deviceId = 'DEV_' + Utilities.getUuid().substring(0, 8).toUpperCase();
      payload.device_id = deviceId;
      payload.registered_at = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
    }
    payload.updated_at = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');

    // 만약 base64 이미지 데이터가 포함되어 있다면 구글 드라이브에 저장
    if (payload.photo_base64 && payload.photo_base64.startsWith('data:image')) {
      const photoResult = uploadPhotoToDrive_(payload.seat_id, payload.device_type, payload.photo_base64);
      if (photoResult.success) {
        payload.photo_url = photoResult.viewUrl;
        payload.photo_id = photoResult.fileId;
      }
      delete payload.photo_base64; // 시트에는 base64 본문을 저장하지 않음
    }

    const data = devSheet.getDataRange().getValues();
    let rowIndex = -1;

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(deviceId)) {
        rowIndex = i + 1;
        break;
      }
    }

    const rowValues = headers.map(h => payload[h] !== undefined ? payload[h] : '');

    if (rowIndex > 0) {
      devSheet.getRange(rowIndex, 1, 1, headers.length).setValues([rowValues]);
    } else {
      devSheet.appendRow(rowValues);
    }

    // 좌석 시트 최종 수정일시 갱신
    touchSeat_(ss, payload.seat_id);

    return {
      success: true,
      deviceId: deviceId,
      photoUrl: payload.photo_url || '',
      device: payload
    };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

/** 기기 삭제 */
function deleteDevice(deviceId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const devSheet = ss.getSheetByName(SHEET_DEVICES);
    const data = devSheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(deviceId)) {
        const seatId = data[i][1];
        devSheet.deleteRow(i + 1);
        if (seatId) touchSeat_(ss, seatId);
        return { success: true };
      }
    }
    return { success: false, error: '해당 기기를 찾을 수 없습니다.' };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

/** 좌석 사용자(선생님 성명, 내선번호) 갱신 */
function updateSeatUser(seatId, teacherName, extension) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const seatSheet = ss.getSheetByName(SHEET_SEATS);
    const data = seatSheet.getDataRange().getValues();
    const headers = data[0];

    const seatIdCol = headers.indexOf('seat_id');
    const userCol = headers.indexOf('current_user');
    const extCol = headers.indexOf('extension');
    const updateCol = headers.indexOf('updated_at');

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][seatIdCol]) === String(seatId)) {
        if (userCol !== -1) seatSheet.getRange(i + 1, userCol + 1).setValue(teacherName || '');
        if (extCol !== -1 && extension !== undefined) seatSheet.getRange(i + 1, extCol + 1).setValue(extension || '');
        if (updateCol !== -1) {
          const nowStr = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
          seatSheet.getRange(i + 1, updateCol + 1).setValue(nowStr);
        }
        return { success: true, seatId: seatId, user: teacherName, ext: extension };
      }
    }
    return { success: false, error: '해당 좌석을 찾을 수 없습니다.' };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

/** 일반교사용 원터치 제출: 좌석 정보 및 귀속 기기 목록 일괄 저장 */
function saveSeatDevices(seatId, seatInfo, deviceList) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. 좌석 정보(선생님 성명, 내선번호) 갱신
    if (seatInfo) {
      updateSeatUser(seatId, seatInfo.current_user, seatInfo.extension);
    }
    
    // 2. 기기 시트에서 해당 좌석의 기존 기기 삭제
    const devSheet = ss.getSheetByName(SHEET_DEVICES);
    const headers = devSheet.getRange(1, 1, 1, devSheet.getLastColumn()).getValues()[0];
    const devData = devSheet.getDataRange().getValues();
    const seatIdCol = headers.indexOf('seat_id');

    for (let i = devData.length - 1; i >= 1; i--) {
      if (String(devData[i][seatIdCol]) === String(seatId)) {
        devSheet.deleteRow(i + 1);
      }
    }

    // 3. 새로 입력된 기기들 추가
    const nowStr = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
    if (deviceList && Array.isArray(deviceList)) {
      deviceList.forEach(dev => {
        if (!dev.device_id) {
          dev.device_id = 'DEV_' + seatId + '_' + (dev.device_type || 'DEV') + '_' + Utilities.getUuid().substring(0, 4);
        }
        dev.registered_at = nowStr;
        dev.updated_at = nowStr;
        const rowValues = headers.map(h => dev[h] !== undefined ? dev[h] : '');
        devSheet.appendRow(rowValues);
      });
    }

    touchSeat_(ss, seatId);
    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

/** 인사이동 시 여러 좌석 일괄 사용자 갱신 */
function batchUpdateSeatUsers(seatList) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const seatSheet = ss.getSheetByName(SHEET_SEATS);
    const data = seatSheet.getDataRange().getValues();
    const headers = data[0];

    const seatIdCol = headers.indexOf('seat_id');
    const userCol = headers.indexOf('current_user');
    const extCol = headers.indexOf('extension');
    const updateCol = headers.indexOf('updated_at');

    const seatMap = {};
    seatList.forEach(item => { seatMap[item.seat_id] = item; });

    const nowStr = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
    let updatedCount = 0;

    for (let i = 1; i < data.length; i++) {
      const sId = String(data[i][seatIdCol]);
      if (seatMap[sId]) {
        const item = seatMap[sId];
        if (userCol !== -1 && item.current_user !== undefined) data[i][userCol] = item.current_user;
        if (extCol !== -1 && item.extension !== undefined) data[i][extCol] = item.extension;
        if (updateCol !== -1) data[i][updateCol] = nowStr;
        updatedCount++;
      }
    }

    seatSheet.getDataRange().setValues(data);
    return { success: true, count: updatedCount };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

/** Gemini Vision API로 라벨 사진 판독 */
function analyzeLabelImage(base64Data, seatId, deviceType) {
  try {
    const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!apiKey) {
      return {
        success: false,
        error: '스크립트 속성에 GEMINI_API_KEY가 설정되어 있지 않습니다.'
      };
    }

    // data:image/...;base64, 부분 추출
    const mimeMatch = base64Data.match(/^data:([^;]+);base64,(.+)$/);
    const mimeType = mimeMatch ? mimeMatch[1] : 'image/jpeg';
    const rawBase64 = mimeMatch ? mimeMatch[2] : base64Data;

    // 최신 고속 비전 모델 Gemini 3.7 Flash 우선 호출 (2.0 Flash는 구글이 서비스 종료함)
    let modelName = 'gemini-3.7-flash';
    let url = `https://generativelanguage.googleapis.com/v1beta/models/${modelName}:generateContent?key=${apiKey}`;

    const promptText = `당신은 대한민국 공립학교의 전산장비(컴퓨터 본체, 모니터, 노트북) 라벨 및 명판 전문 판독관입니다.
첨부된 사진은 다음 두 가지 형태 중 하나입니다:

1. [학교 K-에듀파인 / RFID 물품관리 라벨] (흰색 바탕 격자무늬 스티커)
   - "규격명" 항목을 세심히 분석하십시오. (예: "데스크톱컴퓨터, 엠컴, T273-3210, Intel Core i7 12700" -> 제조사: "엠컴", 모델명: "T273-3210")
   - "취득일자" 항목에서 연월일만 추출하십시오. 괄호 안 내용(예: 괄호 안 연수)은 제거하십시오. (예: "2023-07-14(5년)" -> "2023-07-14")
   - "비고" 란을 정밀하게 분석하여 다음 두 가지 번호를 각각 분리 추출하십시오:
     * "KKR-GAP-..." 번호 또는 조달청 바코드 식별번호 -> 물품관리번호 (RFID) (asset_number)
     * "M000018563" 등 "M"으로 시작하는 8~10자리 학교 자체 물품대장번호 -> 자체 물품관리대장번호 (school_asset_number)
   - CPU 사양(Intel Core i7...), 취득단가 등은 특이사항(notes)에 기재하십시오.

2. [제조사 KC인증 / 제품 규격 명판 스티커] (검은색 또는 은색 명판)
   - 상단 브랜드 로고(MCOM, SAMSUNG, LG 등) 또는 "인증받은자의 상호/제조자"(예: (주)엠텍정보, 삼성전자 등)를 제조사(manufacturer)로 추출하십시오.
   - "기기명칭(모델명)" 항목에서 모델명(예: "T273-3210")만 깔끔하게 모델명(model_name)으로 추출하십시오. "퍼스널컴퓨터" 같은 일반 명칭은 제외하십시오.
   - "제조년월" 항목(예: "2023.05")을 취득일자/제조연월(acquired_date)로 추출하십시오.
   - "Serial No."(시리얼번호)나 "인증번호"는 특이사항(notes)에 기재하십시오.

반드시 아래 JSON 양식으로만 결과를 반환하세요. 앞뒤에 markdown backtick이나 부연 설명 없이 오직 순수한 JSON 객체만 출력하십시오:
{
  "device_type": "본체",
  "manufacturer": "제조사명 (예: 엠텍정보, 엠컴, 삼성전자, LG전자, 삼보컴퓨터 등)",
  "model_name": "모델명 (예: T273-3210, DM500SDA 등)",
  "acquired_date": "취득일자 또는 제조연월 (예: 2023-07-14 또는 2023.05)",
  "asset_number": "물품관리번호 (RFID / KKR-GAP-... 번호, 없으면 빈 문자열)",
  "school_asset_number": "자체 물품관리대장번호 (M으로 시작하는 학교 자체 번호, 없으면 빈 문자열)",
  "notes": "특이사항 (CPU 사양, 취득단가, 시리얼번호 등)"
}
식별할 수 없는 필드는 빈 문자열("")로 두십시오.`;
    const requestBody = {
      contents: [{
        parts: [
          { text: promptText },
          { inlineData: { mimeType: mimeType, data: rawBase64 } }
        ]
      }],
      generationConfig: {
        temperature: 0.1,
        responseMimeType: "application/json"
      }
    };

    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(requestBody),
      muteHttpExceptions: true
    };

    let res = UrlFetchApp.fetch(url, options);
    let code = res.getResponseCode();
    // 만약 3.7 모델 지원 불가 시 2.5 Flash로 자동 폴백
    if (code !== 200 && code === 404) {
      modelName = 'gemini-2.5-flash';
      url = `https://generativelanguage.googleapis.com/v1beta/models/${modelName}:generateContent?key=${apiKey}`;
      res = UrlFetchApp.fetch(url, options);
      code = res.getResponseCode();
    }
    if (code !== 200) {
      return { success: false, error: `Gemini API 오류 (${code}): ` + res.getContentText() };
    }

    const resJson = JSON.parse(res.getContentText());
    const textOut = resJson.candidates[0].content.parts[0].text;
    const parsedData = JSON.parse(textOut.replace(/```json/g, '').replace(/```/g, '').trim());

    // 촬영 즉시 관리자 드라이브에 사진 저장 (이후 "저장" 버튼을 누르지 않아도 사진은 보존됨)
    let photoUrl = '';
    let photoId = '';
    if (seatId && base64Data && base64Data.startsWith('data:image')) {
      const photoResult = uploadPhotoToDrive_(seatId, deviceType, base64Data);
      if (photoResult.success) {
        photoUrl = photoResult.viewUrl;
        photoId = photoResult.fileId;
      }
    }

    return { success: true, data: parsedData, photo_url: photoUrl, photo_id: photoId };
  } catch (err) {
    return { success: false, error: '라벨 분석 실패: ' + err.message };
  }
}

/** 구글 드라이브에 기기 사진 저장 */
function uploadPhotoToDrive_(seatId, deviceType, base64Data) {
  try {
    const folder = getOrCreatePhotoFolder_();
    const mimeMatch = base64Data.match(/^data:([^;]+);base64,(.+)$/);
    const mimeType = mimeMatch ? mimeMatch[1] : 'image/jpeg';
    const rawData = Utilities.base64Decode(mimeMatch ? mimeMatch[2] : base64Data);

    const ext = mimeType.includes('png') ? 'png' : 'jpg';
    const timestamp = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyyMMdd_HHmmss');
    const filename = `${seatId}_${deviceType || 'DEVICE'}_${timestamp}.${ext}`;

    const blob = Utilities.newBlob(rawData, mimeType, filename);
    const file = folder.createFile(blob);
    
    // 학교 내부 또는 링크 소유자 열람 가능 설정
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (e) {}

    return {
      success: true,
      fileId: file.getId(),
      viewUrl: file.getUrl(),
      downloadUrl: `https://drive.google.com/uc?export=view&id=${file.getId()}`
    };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

/** 사진 저장 전용 구글 드라이브 폴더 조회/생성 */
function getOrCreatePhotoFolder_() {
  const folders = DriveApp.getFoldersByName(PHOTO_FOLDER_NAME);
  if (folders.hasNext()) {
    return folders.next();
  }
  const newFolder = DriveApp.createFolder(PHOTO_FOLDER_NAME);
  try {
    newFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e) {}
  return newFolder;
}

/** 좌석 수정일시 갱신 및 교무실별 탭 실시간 동기화 헬퍼 */
function touchSeat_(ss, seatId) {
  try {
    const seatSheet = ss.getSheetByName(SHEET_SEATS);
    const data = seatSheet.getDataRange().getValues();
    const headers = data[0];
    const idCol = headers.indexOf('seat_id');
    const updateCol = headers.indexOf('updated_at');
    if (idCol !== -1 && updateCol !== -1) {
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][idCol]) === String(seatId)) {
          const nowStr = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
          seatSheet.getRange(i + 1, updateCol + 1).setValue(nowStr);
          break;
        }
      }
    }

    // 실시간으로 '전체교무실배치도' 및 '해당 교무실 개별 탭'에 즉시 반영
    syncSeatToSpreadsheetViews_(ss, seatId);
  } catch (e) {
    Logger.log('touchSeat_ error: ' + e.message);
  }
}

// =========================================================================
// 교무실별 개별 탭 & 전체배치도 탭 자동 생성 및 실시간 동기화 시스템
// =========================================================================

const SHEET_ALL_VIEW = '전체교무실배치도';

const ROOM_TAB_DEFS = [
  { id: 'gyomu_center', tabName: '교무센터' },
  { id: 'grade3', tabName: '3학년교무실' },
  { id: 'grade2', tabName: '2학년교무실' },
  { id: 'grade1', tabName: '1학년교무실' },
  { id: 'safety', tabName: '학생안전부' },
  { id: 'edu_info', tabName: '교육정보부' },
  { id: 'support', tabName: '통합지원반' },
  { id: 'library', tabName: '도서실' },
  { id: 'counsel_health_meal', tabName: '상담·보건·급식실' },
  { id: 'admin', tabName: '행정실' },
  { id: 'special', tabName: '특별실' },
  { id: 'class_common', tabName: '학급및공용실' }
];

const VIEW_HEADERS = [
  '교무실', '좌석 ID', '좌석 명칭', '선생님 성명', '내선번호', '조사상태',
  '본체 제조사', '본체 모델명', '본체 취득일자', '본체 관리번호(RFID)', '본체 자체대장번호',
  '모니터1 제조사', '모니터1 모델명', '모니터1 취득일자', '모니터1 관리번호(RFID)', '모니터1 자체대장번호',
  '모니터2 제조사', '모니터2 모델명', '모니터2 취득일자', '모니터2 관리번호(RFID)', '모니터2 자체대장번호',
  '특이사항', '라벨 사진', '최근 수정일시'
];

/** 교무실 ID를 탭 명칭으로 변환 */
function getRoomTabName_(roomId, seatId) {
  if (seatId === 'PRIN_01' || roomId === 'gyomu_principal' || roomId === 'gyomu_center') {
    return '교무센터';
  }
  const found = ROOM_TAB_DEFS.find(r => r.id === roomId);
  return found ? found.tabName : '기타교무실';
}

/** 좌석 및 기기 정보로부터 스프레드시트용 통합 1행 데이터 생성 */
function buildSeatViewRow_(seat, devices) {
  const pc = devices.find(d => d.device_type === 'PC');
  const mon1 = devices.find(d => d.device_type === 'MONITOR_1' || d.device_type === '모니터');
  const mon2 = devices.find(d => d.device_type === 'MONITOR_2');

  let status = '⚪ 미조사';
  if (pc && mon1) status = '✅ 조사완료';
  else if (pc || mon1 || mon2) status = '🟡 부분등록';

  const notesArr = [];
  if (pc && pc.notes) notesArr.push('본체: ' + pc.notes);
  if (mon1 && mon1.notes) notesArr.push('모니터1: ' + mon1.notes);
  if (mon2 && mon2.notes) notesArr.push('모니터2: ' + mon2.notes);
  const notesText = notesArr.join(' / ');

  const photoUrls = devices.map(d => d.photo_url).filter(Boolean);
  const photoText = photoUrls.length > 0 ? photoUrls.join('\n') : '';

  const tabName = getRoomTabName_(seat.room_id, seat.seat_id);

  return [
    tabName,
    seat.seat_id,
    seat.seat_label,
    seat.current_user || '',
    seat.extension || '',
    status,
    pc ? (pc.manufacturer || '') : '',
    pc ? (pc.model_name || '') : '',
    pc ? (pc.acquired_date || '') : '',
    pc ? (pc.asset_number || '') : '',
    pc ? (pc.school_asset_number || '') : '',
    mon1 ? (mon1.manufacturer || '') : '',
    mon1 ? (mon1.model_name || '') : '',
    mon1 ? (mon1.acquired_date || '') : '',
    mon1 ? (mon1.asset_number || '') : '',
    mon1 ? (mon1.school_asset_number || '') : '',
    mon2 ? (mon2.manufacturer || '') : '',
    mon2 ? (mon2.model_name || '') : '',
    mon2 ? (mon2.acquired_date || '') : '',
    mon2 ? (mon2.asset_number || '') : '',
    mon2 ? (mon2.school_asset_number || '') : '',
    notesText,
    photoText,
    seat.updated_at || Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss')
  ];
}

/** 단일 좌석 변경 시 '전체교무실배치도' 및 '해당 교무실 개별 탭' 실시간 동기화 */
function syncSeatToSpreadsheetViews_(ss, seatId) {
  try {
    const seatSheet = ss.getSheetByName(SHEET_SEATS);
    const devSheet = ss.getSheetByName(SHEET_DEVICES);
    if (!seatSheet || !devSheet) return;

    const seats = getSheetObjects_(seatSheet);
    const seat = seats.find(s => String(s.seat_id) === String(seatId));
    if (!seat) return;

    const allDevices = getSheetObjects_(devSheet);
    const seatDevices = allDevices.filter(d => String(d.seat_id) === String(seatId));

    const viewRow = buildSeatViewRow_(seat, seatDevices);
    const targetRoomTab = getRoomTabName_(seat.room_id, seat.seat_id);

    // 1. '전체교무실배치도' 시트에 갱신
    updateOrAppendViewRow_(ss, SHEET_ALL_VIEW, seatId, viewRow, '#1b4332');

    // 2. 해당 교무실 개별 탭에 갱신 (예: 3학년교무실 탭)
    updateOrAppendViewRow_(ss, targetRoomTab, seatId, viewRow, '#1e3a8a');
  } catch (err) {
    Logger.log('syncSeatToSpreadsheetViews_ error: ' + err.message);
  }
}

/** 뷰 시트에서 seat_id 찾아서 행 갱신 또는 추가 */
function updateOrAppendViewRow_(ss, sheetName, seatId, viewRow, headerColor) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    setupSheet_(sheet, VIEW_HEADERS, [], headerColor);
  }

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    sheet.appendRow(viewRow);
    formatViewSheet_(sheet);
    return;
  }

  const idCol = 1; // 0-indexed index 1: '좌석 ID'
  let targetRow = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idCol]) === String(seatId)) {
      targetRow = i + 1;
      break;
    }
  }

  if (targetRow > 0) {
    sheet.getRange(targetRow, 1, 1, viewRow.length).setValues([viewRow]);
  } else {
    sheet.appendRow(viewRow);
  }
  formatViewSheet_(sheet);
}

/** 뷰 시트 서식 맞춤 (열 정렬, 줄바꿈) */
function formatViewSheet_(sheet) {
  try {
    sheet.setFrozenRows(1);
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow <= 1) return;

    sheet.getRange(2, 1, lastRow - 1, Math.min(6, lastCol)).setHorizontalAlignment('center');
    if (lastCol >= 24) {
      sheet.getRange(2, 24, lastRow - 1, 1).setHorizontalAlignment('center');
    }
    if (lastCol >= 22) {
      sheet.getRange(2, 22, lastRow - 1, 1).setWrap(true);
    }
  } catch (e) {
    Logger.log('formatViewSheet_ error: ' + e.message);
  }
}

/**
 * [관리자 메뉴] 전체 교무실별 탭 및 전체배치도 탭 일괄 생성 및 동기화
 * '전체교무실배치도' 탭 1개 + 각 교무실별 11개 탭을 생성하고
 * 최신 좌석 및 PC/모니터 정보를 예쁜 표로 정돈하여 채웁니다.
 */
function syncAllRoomSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const seatSheet = ss.getSheetByName(SHEET_SEATS);
  if (!seatSheet || seatSheet.getLastRow() <= 1) return;
  const devSheet = ss.getSheetByName(SHEET_DEVICES);
  const seats = getSheetObjects_(seatSheet);
  const devices = getSheetObjects_(devSheet);

  const devMap = {};
  devices.forEach(d => {
    if (!devMap[d.seat_id]) devMap[d.seat_id] = [];
    devMap[d.seat_id].push(d);
  });

  // 1. 전체교무실배치도 탭
  const allRows = seats.map(s => buildSeatViewRow_(s, devMap[s.seat_id] || []));
  let allViewSheet = ss.getSheetByName(SHEET_ALL_VIEW);
  if (!allViewSheet) {
    allViewSheet = ss.insertSheet(SHEET_ALL_VIEW);
  }
  setupSheet_(allViewSheet, VIEW_HEADERS, allRows, '#1b4332');
  formatViewSheet_(allViewSheet);

  // 2. 각 교무실별 개별 탭
  ROOM_TAB_DEFS.forEach((rDef, idx) => {
    const roomSeats = seats.filter(s => getRoomTabName_(s.room_id, s.seat_id) === rDef.tabName);
    const roomRows = roomSeats.map(s => buildSeatViewRow_(s, devMap[s.seat_id] || []));

    let roomTab = ss.getSheetByName(rDef.tabName);
    if (!roomTab) {
      roomTab = ss.insertSheet(rDef.tabName);
    }
    setupSheet_(roomTab, VIEW_HEADERS, roomRows, '#1e3a8a');
    formatViewSheet_(roomTab);
  });

  // 3. 탭 순서 정렬 (전체교무실배치도 -> 11개 교무실 탭 -> 좌석 -> 기기 -> 교무실)
  orderSheetsOrder_(ss);

  SpreadsheetApp.flush();
  Logger.log('교무실별 개별 탭 및 전체배치도 탭 동기화 완료!');
  return { success: true };
}

/** 시트 탭 순서 정렬 */
function orderSheetsOrder_(ss) {
  try {
    const desiredOrder = [
      SHEET_ALL_VIEW,
      ...ROOM_TAB_DEFS.map(r => r.tabName),
      SHEET_SEATS,
      SHEET_DEVICES,
      SHEET_ROOMS
    ];

    desiredOrder.forEach((name, idx) => {
      const sh = ss.getSheetByName(name);
      if (sh) {
        ss.setActiveSheet(sh);
        ss.moveActiveSheet(idx + 1);
      }
    });

    const firstSheet = ss.getSheetByName(SHEET_ALL_VIEW);
    if (firstSheet) ss.setActiveSheet(firstSheet);
  } catch (e) {
    Logger.log('orderSheetsOrder_ error: ' + e.message);
  }
}

/** 구글 스프레드시트 열릴 때 커스텀 관리자 메뉴 생성 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🏫 솔내고 교직원 기기관리')
    .addItem('🔄 전체 교무실별 탭 동기화/새로고침', 'syncAllRoomSheets')
    .addItem('➕ [신규] 학급 및 공용실 등 교무실 추가', '신규교무실_학급및공용실_추가')
    .addItem('⚙️ 초기 데이터 및 탭 전체 재설정 (⚠️ 기존 데이터 초기화됨)', '초기데이터생성')
    .addToUi();
}

/**
 * [관리자 전용] 기존 좌석·기기 데이터는 전혀 건드리지 않고,
 * '교무실' 시트에 "학급 및 공용실 등" 행 1개만 안전하게 추가합니다.
 * (초기데이터생성과 달리 시트를 clear()하지 않으므로 이미 등록된 98석 데이터가 보존됩니다)
 */
function 신규교무실_학급및공용실_추가() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const roomSheet = ss.getSheetByName(SHEET_ROOMS);
  if (!roomSheet) {
    ui.alert('"교무실" 시트를 찾을 수 없습니다. 먼저 초기 데이터를 생성해 주세요.');
    return;
  }

  const data = roomSheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf('room_id');
  const orderCol = headers.indexOf('order_index');

  const already = data.some((row, i) => i > 0 && String(row[idCol]) === 'class_common');
  if (already) {
    ui.alert('"학급 및 공용실 등" 교무실은 이미 등록되어 있습니다.');
    return;
  }

  let maxOrder = 0;
  for (let i = 1; i < data.length; i++) {
    const v = Number(data[i][orderCol]) || 0;
    if (v > maxOrder) maxOrder = v;
  }

  const newRow = headers.map(h => {
    if (h === 'room_id') return 'class_common';
    if (h === 'room_name') return '학급 및 공용실 등';
    if (h === 'floor_info') return '교내 각 층';
    if (h === 'phone_fax') return '';
    if (h === 'seat_count') return 0;
    if (h === 'order_index') return maxOrder + 1;
    return '';
  });
  roomSheet.appendRow(newRow);

  ui.alert('"학급 및 공용실 등" 교무실이 추가되었습니다!\n기존 좌석/기기 데이터는 전혀 변경되지 않았습니다.\n\n이제 웹앱에서 교무실 선택 드롭다운의 "학급 및 공용실 등"을 고른 뒤 "좌석 추가" 버튼으로 교실/공용실을 하나씩 등록하면 됩니다.');
}


/** 시트 데이터를 객체 배열로 변환 */
function getSheetObjects_(sheet) {
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  const results = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const obj = {};
    let hasContent = false;
    for (let j = 0; j < headers.length; j++) {
      let val = row[j];
      // 시트가 날짜로 자동 인식한 셀은 google.script.run 전송 오류를 막기 위해 문자열로 변환
      if (val instanceof Date) {
        val = Utilities.formatDate(val, 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
      }
      obj[headers[j]] = val;
      if (val !== '') hasContent = true;
    }
    if (hasContent) results.push(obj);
  }
  return results;
}

/** 필요한 시트가 없거나 좌석 데이터가 비어있으면 자동 생성 */
function ensureSheetsInitialized_(ss) {
  let seatSheet = ss.getSheetByName(SHEET_SEATS);
  let devSheet = ss.getSheetByName(SHEET_DEVICES);
  let roomSheet = ss.getSheetByName(SHEET_ROOMS);

  // 시트가 없거나, 좌석 시트 행 개수가 1개 이하(헤더만 있거나 빈 시트)이면 즉시 자동 생성
  if (!seatSheet || !devSheet || !roomSheet || seatSheet.getLastRow() <= 1) {
    초기데이터생성();
  }
}

/**
 * [관리자 전용] 전주솔내고등학교 2026 좌석배치도 초기 데이터 생성 함수
 * 시트가 초기화되어 있지 않을 때 스프레드시트에 표준 시트 3종 및 98개 좌석을 자동 등록합니다.
 */
function 초기데이터생성() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. 좌석 시트 생성 및 초기화
  let seatSheet = ss.getSheetByName(SHEET_SEATS);
  if (!seatSheet) {
    seatSheet = ss.insertSheet(SHEET_SEATS);
  }
  
  const seatHeaders = ['seat_id', 'room_id', 'room_name', 'seat_label', 'current_user', 'extension', 'updated_at'];
  
  // 2. 기기 시트 생성 및 초기화
  let devSheet = ss.getSheetByName(SHEET_DEVICES);
  if (!devSheet) {
    devSheet = ss.insertSheet(SHEET_DEVICES);
  }
  const devHeaders = [
    'device_id', 'seat_id', 'device_type', 'manufacturer', 'model_name',
    'acquired_date', 'asset_number', 'school_asset_number', 'item_code', 'serial_number',
    'photo_url', 'photo_id', 'notes', 'registered_at', 'updated_at'
  ];

  // 3. 교무실 시트 생성 및 초기화
  let roomSheet = ss.getSheetByName(SHEET_ROOMS);
  if (!roomSheet) {
    roomSheet = ss.insertSheet(SHEET_ROOMS);
  }
  const roomHeaders = ['room_id', 'room_name', 'floor_info', 'phone_fax', 'seat_count', 'order_index'];

  // 프리셋 데이터 불러오기
  const presetRooms = getPresetRooms_();
  const seatRows = [];
  const roomRows = [];

  let order = 1;
  presetRooms.forEach(room => {
    roomRows.push([
      room.id,
      room.name,
      room.floor,
      room.phone || '',
      room.seats.length,
      order++
    ]);

    room.seats.forEach(s => {
      seatRows.push([
        s.id,
        room.id,
        room.name + (room.floor ? ` (${room.floor})` : ''),
        s.label,
        s.user || '',
        s.ext || '',
        Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss')
      ]);
    });
  });

  // 서식 적용 및 데이터 쓰기
  setupSheet_(seatSheet, seatHeaders, seatRows, '#4a7c59');
  setupSheet_(devSheet, devHeaders, [], '#2b5876');
  setupSheet_(roomSheet, roomHeaders, roomRows, '#5d4037');

  Logger.log(`초기 데이터 설정 완료: 교무실 ${roomRows.length}개, 총 좌석 ${seatRows.length}개`);
  
  // 전체배치도 및 각 교무실별 탭 일괄 생성 및 동기화
  syncAllRoomSheets();
}

/** 시트 헤더 서식 및 데이터 일괄 주입 */
function setupSheet_(sheet, headers, rows, headerColor) {
  sheet.clear();
  sheet.appendRow(headers);

  if (rows && rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }

  // 헤더 스타일
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground(headerColor)
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  sheet.setFrozenRows(1);
}

/** 2026학년도 전주솔내고 교직원 좌석배치도 기준 프리셋 */
function getPresetRooms_() {
  return [
    {
      id: "gyomu_principal",
      name: "교장실",
      floor: "",
      phone: "270-8200 / 8623-2914",
      seats: [{ id: "PRIN_01", label: "교장", user: "정진복", ext: "200" }]
    },
    {
      id: "gyomu_center",
      name: "교무센터",
      floor: "",
      phone: "",
      seats: [
        { id: "GYOMU_01", label: "교감", user: "", ext: "" },
        { id: "GYOMU_02", label: "상단 1열-1", user: "", ext: "" },
        { id: "GYOMU_03", label: "상단 2열-1", user: "", ext: "" },
        { id: "GYOMU_04", label: "상단 1열-2", user: "", ext: "" },
        { id: "GYOMU_05", label: "상단 2열-2", user: "", ext: "" },
        { id: "GYOMU_06", label: "하단 1열-1", user: "", ext: "" },
        { id: "GYOMU_07", label: "하단 2열-1", user: "", ext: "" },
        { id: "GYOMU_08", label: "하단 3열-1", user: "", ext: "" },
        { id: "GYOMU_09", label: "하단 1열-2", user: "", ext: "" },
        { id: "GYOMU_10", label: "하단 2열-2", user: "", ext: "" },
        { id: "GYOMU_11", label: "하단 3열-2", user: "", ext: "" },
        { id: "GYOMU_12", label: "중앙 1열-1", user: "", ext: "" },
        { id: "GYOMU_13", label: "중앙 2열-1", user: "", ext: "" },
        { id: "GYOMU_14", label: "중앙 1열-2", user: "", ext: "" },
        { id: "GYOMU_15", label: "중앙 2열-2", user: "", ext: "" },
        { id: "GYOMU_16", label: "중앙 1열-3", user: "", ext: "" },
        { id: "GYOMU_17", label: "중앙 2열-3", user: "", ext: "" },
        { id: "GYOMU_18", label: "우측 1열-1", user: "", ext: "" },
        { id: "GYOMU_19", label: "우측 2열-1", user: "", ext: "" },
        { id: "GYOMU_20", label: "우측 1열-2", user: "", ext: "" },
        { id: "GYOMU_21", label: "우측 2열-2", user: "", ext: "" }
      ]
    },
    {
      id: "grade3",
      name: "3학년 교무실",
      floor: "",
      phone: "",
      seats: [
        { id: "G3_01", label: "3학년부장", user: "", ext: "" },
        { id: "G3_02", label: "좌측 중단-1", user: "", ext: "" },
        { id: "G3_03", label: "좌측 중단-2", user: "", ext: "" },
        { id: "G3_04", label: "좌측 하단-1", user: "", ext: "" },
        { id: "G3_05", label: "좌측 하단-2", user: "", ext: "" },
        { id: "G3_06", label: "우측 상단-1", user: "", ext: "" },
        { id: "G3_07", label: "우측 상단-2", user: "", ext: "" },
        { id: "G3_08", label: "우측 중단-1", user: "", ext: "" },
        { id: "G3_09", label: "우측 중단-2", user: "", ext: "" },
        { id: "G3_10", label: "우측 하단-1", user: "", ext: "" },
        { id: "G3_11", label: "우측 하단-2", user: "", ext: "" }
      ]
    },
    {
      id: "grade2",
      name: "2학년 교무실",
      floor: "",
      phone: "",
      seats: [
        { id: "G2_01", label: "2학년부장", user: "", ext: "" },
        { id: "G2_02", label: "좌측 중단-1", user: "", ext: "" },
        { id: "G2_03", label: "좌측 중단-2", user: "", ext: "" },
        { id: "G2_04", label: "좌측 하단-1", user: "", ext: "" },
        { id: "G2_05", label: "좌측 하단-2", user: "", ext: "" },
        { id: "G2_06", label: "우측 상단-1", user: "", ext: "" },
        { id: "G2_07", label: "우측 상단-2", user: "", ext: "" },
        { id: "G2_08", label: "우측 중단-1", user: "", ext: "" },
        { id: "G2_09", label: "우측 중단-2", user: "", ext: "" },
        { id: "G2_10", label: "우측 하단-1", user: "", ext: "" },
        { id: "G2_11", label: "우측 하단-2", user: "", ext: "" }
      ]
    },
    {
      id: "grade1",
      name: "1학년 교무실",
      floor: "",
      phone: "",
      seats: [
        { id: "G1_01", label: "1학년부장", user: "", ext: "" },
        { id: "G1_02", label: "좌측 중단-1", user: "", ext: "" },
        { id: "G1_03", label: "좌측 중단-2", user: "", ext: "" },
        { id: "G1_04", label: "좌측 하단-1", user: "", ext: "" },
        { id: "G1_05", label: "좌측 하단-2", user: "", ext: "" },
        { id: "G1_06", label: "우측 상단-1", user: "", ext: "" },
        { id: "G1_07", label: "우측 상단-2", user: "", ext: "" },
        { id: "G1_08", label: "우측 중단-1", user: "", ext: "" },
        { id: "G1_09", label: "우측 중단-2", user: "", ext: "" },
        { id: "G1_10", label: "우측 하단-1", user: "", ext: "" },
        { id: "G1_11", label: "우측 하단-2", user: "", ext: "" }
      ]
    },
    {
      id: "safety",
      name: "학생안전부",
      floor: "",
      phone: "",
      seats: [
        { id: "SAFE_01", label: "학생안전 1열-1", user: "", ext: "" },
        { id: "SAFE_02", label: "학생안전 2열-1", user: "", ext: "" },
        { id: "SAFE_03", label: "학생안전 1열-2", user: "", ext: "" },
        { id: "SAFE_04", label: "학생안전 2열-2", user: "", ext: "" },
        { id: "SAFE_05", label: "학생안전 1열-3", user: "", ext: "" },
        { id: "SAFE_06", label: "학생안전 2열-3", user: "", ext: "" },
        { id: "SAFE_07", label: "상담실", user: "", ext: "" }
      ]
    },
    {
      id: "support",
      name: "통합지원반",
      floor: "",
      phone: "",
      seats: [
        { id: "SUPP_01", label: "통합지원 1", user: "", ext: "" },
        { id: "SUPP_02", label: "통합지원 2", user: "", ext: "" },
        { id: "SUPP_03", label: "통합지원 3", user: "", ext: "" }
      ]
    },
    {
      id: "edu_info",
      name: "교육정보부",
      floor: "",
      phone: "",
      seats: [
        { id: "INFO_01", label: "정보부 1열-1", user: "", ext: "" },
        { id: "INFO_02", label: "정보부 2열-1", user: "", ext: "" },
        { id: "INFO_03", label: "정보부 1열-2", user: "", ext: "" },
        { id: "INFO_04", label: "정보부 2열-2", user: "", ext: "" }
      ]
    },
    {
      id: "library",
      name: "도서실",
      floor: "",
      phone: "",
      seats: [
        { id: "LIB_01", label: "사서", user: "", ext: "" }
      ]
    },
    {
      id: "counsel_health_meal",
      name: "상담·보건·급식실",
      floor: "각 층",
      phone: "",
      seats: [
        { id: "WEE_01", label: "위클래스(본관 전관 3층)", user: "", ext: "" },
        { id: "CAREER_01", label: "진로상담실(본관 후관 2층)", user: "", ext: "" },
        { id: "HEALTH_01", label: "보건실-1(본관 전관 1층)", user: "", ext: "" },
        { id: "HEALTH_02", label: "보건실-2(본관 전관 1층)", user: "", ext: "" },
        { id: "MEAL_01", label: "급식실-1", user: "", ext: "" },
        { id: "MEAL_02", label: "급식실-2", user: "", ext: "" }
      ]
    },
    {
      id: "admin",
      name: "행정실",
      floor: "",
      phone: "",
      seats: [
        { id: "ADM_01", label: "행정실장", user: "", ext: "" },
        { id: "ADM_02", label: "상단 중앙", user: "", ext: "" },
        { id: "ADM_03", label: "상단 우측", user: "", ext: "" },
        { id: "ADM_04", label: "하단좌 1열-1", user: "", ext: "" },
        { id: "ADM_05", label: "하단좌 2열-1", user: "", ext: "" },
        { id: "ADM_06", label: "하단좌 1열-2", user: "", ext: "" },
        { id: "ADM_07", label: "하단좌 2열-2", user: "", ext: "" },
        { id: "ADM_08", label: "하단우 1열-1", user: "", ext: "" },
        { id: "ADM_09", label: "하단우 2열-1", user: "", ext: "" },
        { id: "ADM_10", label: "하단우 1열-2", user: "", ext: "" },
        { id: "ADM_11", label: "하단우 2열-2", user: "", ext: "" }
      ]
    },
    {
      id: "special",
      name: "특별실",
      floor: "교내 각 층",
      phone: "",
      seats: [
        { id: "SPEC_ART", label: "미술실(별관)", user: "", ext: "" },
        { id: "SPEC_MUSIC", label: "음악실(별관)", user: "", ext: "" },
        { id: "SPEC_PE_1", label: "체육과실-1(본관 전관1층)", user: "", ext: "" },
        { id: "SPEC_PE_2", label: "체육과실-2(본관 전관1층)", user: "", ext: "" },
        { id: "SPEC_PE_3", label: "체육과실-3(본관 전관1층)", user: "", ext: "" },
        { id: "SPEC_CONF_S", label: "소회의실(교무센터옆)", user: "", ext: "" },
        { id: "SPEC_CONF_L", label: "대회의실(본관3층)", user: "", ext: "" },
        { id: "SPEC_DUTY", label: "당직실(본관1층)", user: "", ext: "" },
        { id: "SPEC_DINE", label: "식생활관", user: "", ext: "" },
        { id: "SPEC_PRINT", label: "인쇄실(본관전관1층)", user: "", ext: "" },
        { id: "SPEC_DORM", label: "기숙사", user: "", ext: "" },
        { id: "SPEC_SCIENCE", label: "과학실", user: "", ext: "" }
      ]
    },
    {
      id: "class_common",
      name: "학급 및 공용실 등",
      floor: "교내 각 층",
      phone: "",
      seats: []
    }
  ];
}

/** 신규 좌석 추가 */
function addSeat(seatPayload) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const seatSheet = ss.getSheetByName(SHEET_SEATS);
    let headers = seatSheet.getRange(1, 1, 1, seatSheet.getLastColumn()).getValues()[0];

    // 페이로드에 아직 없는 열(예: special_subroom_id)이 있으면 자동으로 열 추가
    Object.keys(seatPayload).forEach(k => {
      if (headers.indexOf(k) === -1) {
        seatSheet.getRange(1, headers.length + 1).setValue(k);
        headers.push(k);
      }
    });

    let seatId = seatPayload.seat_id;
    if (!seatId) {
      const prefix = (seatPayload.room_id || 'SEAT').toUpperCase().replace(/[^A-Z0-9]/g, '').substring(0, 6);
      seatId = prefix + '_' + Utilities.getUuid().substring(0, 6).toUpperCase();
      seatPayload.seat_id = seatId;
    }
    seatPayload.updated_at = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');

    const rowValues = headers.map(h => seatPayload[h] !== undefined ? seatPayload[h] : '');
    seatSheet.appendRow(rowValues);

    // 교무실 좌석 수 갱신
    updateRoomSeatCount_(ss, seatPayload.room_id);

    return {
      success: true,
      seat: seatPayload
    };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

/** 좌석 삭제 (해당 좌석 및 귀속 기기 일괄 삭제) */
function deleteSeat(seatId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const seatSheet = ss.getSheetByName(SHEET_SEATS);
    const data = seatSheet.getDataRange().getValues();
    const headers = data[0];
    const idCol = headers.indexOf('seat_id');
    const roomCol = headers.indexOf('room_id');

    let roomId = '';
    let found = false;

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idCol]) === String(seatId)) {
        roomId = data[i][roomCol];
        seatSheet.deleteRow(i + 1);
        found = true;
        break;
      }
    }

    if (!found) {
      return { success: false, error: '해당 좌석을 찾을 수 없습니다.' };
    }

    // 귀속된 기기도 함께 삭제
    const devSheet = ss.getSheetByName(SHEET_DEVICES);
    if (devSheet) {
      const devData = devSheet.getDataRange().getValues();
      const devSeatCol = devData[0].indexOf('seat_id');
      if (devSeatCol !== -1) {
        for (let j = devData.length - 1; j >= 1; j--) {
          if (String(devData[j][devSeatCol]) === String(seatId)) {
            devSheet.deleteRow(j + 1);
          }
        }
      }
    }

    if (roomId) updateRoomSeatCount_(ss, roomId);

    return { success: true, seatId: seatId };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

/** 교무실 시트 좌석 수 자동 계산 및 갱신 헬퍼 */
function updateRoomSeatCount_(ss, roomId) {
  try {
    const roomSheet = ss.getSheetByName(SHEET_ROOMS);
    const seatSheet = ss.getSheetByName(SHEET_SEATS);
    if (!roomSheet || !seatSheet) return;

    const seatData = seatSheet.getDataRange().getValues();
    const roomCol = seatData[0].indexOf('room_id');
    let count = 0;
    for (let i = 1; i < seatData.length; i++) {
      if (String(seatData[i][roomCol]) === String(roomId)) count++;
    }

    const roomData = roomSheet.getDataRange().getValues();
    const rIdCol = roomData[0].indexOf('room_id');
    const rCountCol = roomData[0].indexOf('seat_count');
    if (rIdCol === -1 || rCountCol === -1) return;

    for (let j = 1; j < roomData.length; j++) {
      if (String(roomData[j][rIdCol]) === String(roomId)) {
        roomSheet.getRange(j + 1, rCountCol + 1).setValue(count);
        break;
      }
    }
  } catch (e) {}
}


/** 단일 좌석 위치(pos_x, pos_y) 저장 */
/** 좌석 정보(이름, 내선, 위치, 고정여부 등) 범용 갱신 */
function updateSeatInfo(seatId, seatData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const seatSheet = ss.getSheetByName(SHEET_SEATS);
    const data = seatSheet.getDataRange().getValues();
    let headers = data[0];
    const idCol = headers.indexOf('seat_id');

    Object.keys(seatData).forEach(k => {
      if (headers.indexOf(k) === -1) {
        seatSheet.getRange(1, headers.length + 1).setValue(k);
        headers.push(k);
      }
    });

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idCol]) === String(seatId)) {
        Object.keys(seatData).forEach(k => {
          const colIdx = headers.indexOf(k);
          if (colIdx !== -1) {
            seatSheet.getRange(i + 1, colIdx + 1).setValue(seatData[k]);
          }
        });
        return { success: true };
      }
    }
    return { success: false, error: '좌석을 찾을 수 없습니다.' };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function updateSeatPosition(seatId, posX, posY) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const seatSheet = ss.getSheetByName(SHEET_SEATS);
    const data = seatSheet.getDataRange().getValues();
    const headers = data[0];
    const idCol = headers.indexOf('seat_id');
    let xCol = headers.indexOf('pos_x');
    let yCol = headers.indexOf('pos_y');

    // pos_x, pos_y 열이 없으면 추가
    if (xCol === -1 || yCol === -1) {
      if (xCol === -1) {
        seatSheet.getRange(1, headers.length + 1).setValue('pos_x');
        xCol = headers.length;
        headers.push('pos_x');
      }
      if (yCol === -1) {
        seatSheet.getRange(1, headers.length + 1).setValue('pos_y');
        yCol = headers.length;
        headers.push('pos_y');
      }
    }

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idCol]) === String(seatId)) {
        seatSheet.getRange(i + 1, xCol + 1).setValue(posX);
        seatSheet.getRange(i + 1, yCol + 1).setValue(posY);
        break;
      }
    }

    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

/** 다중 좌석 위치·크기 일괄 저장 */
function batchUpdateSeatPositions(positions) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const seatSheet = ss.getSheetByName(SHEET_SEATS);
    const data = seatSheet.getDataRange().getValues();
    const headers = data[0];
    const idCol = headers.indexOf('seat_id');
    let xCol = headers.indexOf('pos_x');
    let yCol = headers.indexOf('pos_y');
    let wCol = headers.indexOf('width');
    let hCol = headers.indexOf('height');

    if (xCol === -1) {
      seatSheet.getRange(1, headers.length + 1).setValue('pos_x');
      xCol = headers.length;
      headers.push('pos_x');
    }
    if (yCol === -1) {
      seatSheet.getRange(1, headers.length + 1).setValue('pos_y');
      yCol = headers.length;
      headers.push('pos_y');
    }
    if (wCol === -1) {
      seatSheet.getRange(1, headers.length + 1).setValue('width');
      wCol = headers.length;
      headers.push('width');
    }
    if (hCol === -1) {
      seatSheet.getRange(1, headers.length + 1).setValue('height');
      hCol = headers.length;
      headers.push('height');
    }

    const posMap = {};
    positions.forEach(p => { posMap[p.seat_id] = p; });

    for (let i = 1; i < data.length; i++) {
      const sId = String(data[i][idCol]);
      const p = posMap[sId];
      if (p) {
        if (p.pos_x !== undefined) seatSheet.getRange(i + 1, xCol + 1).setValue(p.pos_x);
        if (p.pos_y !== undefined) seatSheet.getRange(i + 1, yCol + 1).setValue(p.pos_y);
        if (p.width !== undefined) seatSheet.getRange(i + 1, wCol + 1).setValue(p.width);
        if (p.height !== undefined) seatSheet.getRange(i + 1, hCol + 1).setValue(p.height);
      }
    }

    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  }
}
