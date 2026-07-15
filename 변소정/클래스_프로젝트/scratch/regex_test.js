const sampleText = `:large_blue_circle:차봉호-위니드인테리어 / 도어-숲속길마을 월드메르디앙 센트럴파크-707동 504호
시공/납품 진행 부탁드립니다.
- 총 결제금액: 796,400원
- 납기완료 예정 날짜
1. 2026-04-20 (월)

상세내역 확인하기 (주문번호 00280115)`;

function testRegex(text) {
  console.log("--- Testing Regex ---");
  
  // M column: Item
  const itemMatch = text.match(/-(.+?)\r?\n시공\/납품/s);
  console.log("Item Match:", itemMatch ? itemMatch[1].trim() : "FAILED");

  // J column: Price
  const priceMatch = text.match(/결제금액:\s*([\d,]+)/);
  console.log("Price Match:", priceMatch ? priceMatch[1].replace(/,/g, '') : "FAILED");

  // E column: Date
  const dateMatch = text.match(/(\d{4}[-.]\d{2}[-.]\d{2})/);
  console.log("Date Match:", dateMatch ? dateMatch[1].replace(/\./g, '-') : "FAILED");

  // F column: Day
  const dayMatch = text.match(/\(([월화수목금토일])\)/);
  console.log("Day Match:", dayMatch ? dayMatch[1] : "FAILED");

  // L column: Order No
  const orderMatch = text.match(/주문번호\s*(\d+)/);
  console.log("Order Match:", orderMatch ? orderMatch[1] : "FAILED");
}

testRegex(sampleText);
