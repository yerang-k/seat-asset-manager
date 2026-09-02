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

/** 웹앱 진입점 */
function doGet(e) {
  const template = HtmlService.createTemplateFromFile('index');
  return template.evaluate()
    .setTitle('전주솔내고 좌석별 PC·모니터 관리 시스템')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
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

    const seatSheet = ss.getSheetByName(SHEET_SEATS);
    const devSheet = ss.getSheetByName(SHEET_DEVICES);
    const roomSheet = ss.getSheetByName(SHEET_ROOMS);

    const seats = getSheetObjects_(seatSheet);
    const devices = getSheetObjects_(devSheet);
    const rooms = getSheetObjects_(roomSheet);

    return {
      success: true,
      seats: seats,
      devices: devices,
      rooms: rooms,
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
function analyzeLabelImage(base64Data) {
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

    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${apiKey}`;

    const promptText = `당신은 대한민국 공립학교의 전산장비(컴퓨터 본체, 모니터, 노트북) 라벨 및 명판 전문 판독관입니다.
첨부된 사진은 학교 컴퓨터/모니터에 부착된 K-에듀파인 물품관리 라벨, 조달청 라벨, 또는 제조사(삼성, LG, 삼보, 주연테크, HP 등) 스펙 명판 스티커입니다.
사진을 정밀하게 분석하여 아래 JSON 양식으로만 결과를 반환하세요. 앞뒤에 markdown backtick이나 부연 설명 없이 오직 순수한 JSON 객체만 출력하십시오.

{
  "device_type": "본체" 또는 "모니터" 또는 "노트북" 또는 "기타",
  "manufacturer": "제조사명 (예: 삼성전자, LG전자, 삼보컴퓨터, 주연테크, 한성컴퓨터, HP, DELL 등)",
  "model_name": "모델명 (예: DM500SDA, 24MB35V 등)",
  "acquired_date": "취득일자 또는 제조연월 (YYYY-MM-DD 또는 YYYY.MM)",
  "asset_number": "물품관리번호 (K-에듀파인 물품관리번호 등)",
  "item_code": "물품목록번호 (조달청 물품목록/식별번호 등)",
  "serial_number": "제조번호/S/N",
  "notes": "특이사항 (기재된 전압, 소비전력, 기타 참고사항)"
}
읽을 수 없는 필드는 빈 문자열("")로 채우십시오.`;

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

    const res = UrlFetchApp.fetch(url, options);
    const code = res.getResponseCode();
    if (code !== 200) {
      return { success: false, error: `Gemini API 오류 (${code}): ` + res.getContentText() };
    }

    const resJson = JSON.parse(res.getContentText());
    const textOut = resJson.candidates[0].content.parts[0].text;
    const parsedData = JSON.parse(textOut.replace(/```json/g, '').replace(/```/g, '').trim());

    return { success: true, data: parsedData };
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

/** 좌석 수정일시 갱신 헬퍼 */
function touchSeat_(ss, seatId) {
  try {
    const seatSheet = ss.getSheetByName(SHEET_SEATS);
    const data = seatSheet.getDataRange().getValues();
    const headers = data[0];
    const idCol = headers.indexOf('seat_id');
    const updateCol = headers.indexOf('updated_at');
    if (idCol === -1 || updateCol === -1) return;

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idCol]) === String(seatId)) {
        const nowStr = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
        seatSheet.getRange(i + 1, updateCol + 1).setValue(nowStr);
        break;
      }
    }
  } catch (e) {}
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
      obj[headers[j]] = row[j];
      if (row[j] !== '') hasContent = true;
    }
    if (hasContent) results.push(obj);
  }
  return results;
}

/** 필요한 시트가 없으면 생성 */
function ensureSheetsInitialized_(ss) {
  let seatSheet = ss.getSheetByName(SHEET_SEATS);
  let devSheet = ss.getSheetByName(SHEET_DEVICES);
  let roomSheet = ss.getSheetByName(SHEET_ROOMS);

  if (!seatSheet || !devSheet || !roomSheet) {
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
    'acquired_date', 'asset_number', 'item_code', 'serial_number',
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
  for (let c = 1; c <= headers.length; c++) {
    sheet.autoResizeColumn(c);
  }
}

/** 2026학년도 전주솔내고 교직원 좌석배치도 기준 프리셋 */
function getPresetRooms_() {
  return [
    {
      id: "gyomu_principal",
      name: "교장실",
      floor: "본관 전관 1층",
      phone: "270-8200 / 8623-2914",
      seats: [{ id: "PRIN_01", label: "교장", user: "정진복", ext: "200" }]
    },
    {
      id: "gyomu_center",
      name: "교무센터",
      floor: "본관 후관 1층",
      phone: "270-8208, FAX 276-3015",
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
      floor: "본관 후관 4층",
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
      floor: "본관 후관 2층",
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
      floor: "본관 후관 3층",
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
      floor: "본관 후관 1층",
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
      floor: "본관 후관 1층",
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
      floor: "본관 전관 4층",
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
      floor: "본관 후관 3층",
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
      floor: "본관 1층",
      phone: "270-8209 FAX 276-3014",
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
        { id: "SPEC_DORM", label: "기숙사", user: "", ext: "" }
      ]
    }
  ];
}

/** 신규 좌석 추가 */
function addSeat(seatPayload) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const seatSheet = ss.getSheetByName(SHEET_SEATS);
    const headers = seatSheet.getRange(1, 1, 1, seatSheet.getLastColumn()).getValues()[0];

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

/** 다중 좌석 위치 일괄 저장 */
function batchUpdateSeatPositions(positions) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const seatSheet = ss.getSheetByName(SHEET_SEATS);
    const data = seatSheet.getDataRange().getValues();
    const headers = data[0];
    const idCol = headers.indexOf('seat_id');
    let xCol = headers.indexOf('pos_x');
    let yCol = headers.indexOf('pos_y');

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

    const posMap = {};
    positions.forEach(p => { posMap[p.seat_id] = p; });

    for (let i = 1; i < data.length; i++) {
      const sId = String(data[i][idCol]);
      if (posMap[sId]) {
        seatSheet.getRange(i + 1, xCol + 1).setValue(posMap[sId].pos_x);
        seatSheet.getRange(i + 1, yCol + 1).setValue(posMap[sId].pos_y);
      }
    }

    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  }
}
