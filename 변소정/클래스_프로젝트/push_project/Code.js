/**
 * 쇼핑몰 배차노트 자동 파싱 시스템 (최종본)
 */

/**
 * 1. 단순 트리거 (onEdit)
 */
function onEdit(e) {
  if (!e || !e.range) return;
  
  const range = e.range;
  const col = range.getColumn();
  
  // A열(1번)이 아니면 종료
  if (col !== 1) return;

  parseAColumn(e);
}

/**
 * 2. 실제 파싱 및 입력 실행 함수
 */
function parseAColumn(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const startRow = range.getRow();
  const values = range.getDisplayValues(); // 다중 행 대응

  for (let i = 0; i < values.length; i++) {
    const text = values[i][0].trim();
    const currentRow = startRow + i;

    // 시작 조건 확인: '결제완료 안내]' 문구 포함 여부
    if (!text || !text.includes('결제완료 안내]')) continue;

    // --- 데이터 추출 (정규식) ---

    // 1. 주문번호 (L열)
    let order = '';
    const orderMatch = text.match(/주문번호\s+(\d+)/);
    if (orderMatch) order = orderMatch[1];

    // 2. 품목명 (M열) - 첫 번째 줄(헤더) 이후의 첫 실질적인 내용 줄에서 추출
    let item = '';
    const allLines = text.split('\n').map(l => l.trim()).filter(l => l.length > 0);
    if (allLines.length > 1) {
      const contentLine = allLines[1]; // [결제완료 안내] 다음 줄
      // 첫 번째 '-' 이후의 내용만 추출 (이름/이모지 등 제거)
      const hyphenIndex = contentLine.indexOf('-');
      if (hyphenIndex !== -1) {
        item = contentLine.substring(hyphenIndex + 1).trim();
      } else {
        item = contentLine; // '-'가 없으면 줄 전체 사용
      }
    }

    // 3. 금액 (J열)
    let price = '';
    const priceMatch = text.match(/결제금액:\s*([\d,]+)/);
    if (priceMatch) price = priceMatch[1].replace(/,/g, '');

    // 4. 날짜 (E열)
    let date = '';
    const dateMatch = text.match(/(\d{4}-\d{2}-\d{2})/);
    if (dateMatch) date = dateMatch[1];

    // 5. 요일 (F열)
    let day = '';
    const dayMatch = text.match(/\(([월화수목금토일])\)/);
    if (dayMatch) day = dayMatch[1];

    // --- 데이터 입력 ---
    try {
      if (date)  sheet.getRange(currentRow, 5).setValue(date);   // E열
      if (day)   sheet.getRange(currentRow, 6).setValue(day);    // F열
      if (price) sheet.getRange(currentRow, 10).setValue(price); // J열
      
      if (order) {
        // 주문번호 텍스트 형식 유지 (앞자리 0)
        sheet.getRange(currentRow, 12).setNumberFormat('@').setValue(order); // L열
      }
      
      if (item)  sheet.getRange(currentRow, 13).setValue(item);  // M열
    } catch (err) {
      console.error('Row ' + currentRow + ' 입력 실패: ' + err.message);
    }
  }
}

/**
 * 3. 수동 테스트용 함수 (필요 시 실행)
 */
function debugManualTest() {
  const range = SpreadsheetApp.getActiveRange();
  parseAColumn({ range: range });
}
