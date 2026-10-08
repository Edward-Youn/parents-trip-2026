/**
 * 부모님 수원·서울 나들이 — 공유 체크리스트/메모 저장소
 * 구글 시트 > 확장 프로그램 > Apps Script 에 이 코드를 붙여넣고
 * 배포 > 새 배포 > 웹 앱 (실행: 나, 액세스 권한: 모든 사용자) 로 배포하세요.
 */
const SHEET_NAME = 'state';
const MAX_TEXT = 1000;

const DEFAULT_TODOS = [
  {id:'b1', cat:'book', order:1, urgent:true, done:false, text:'렌터카 예약 (8일 저녁 또는 9일 아침 픽업, 10일 아침 반납)', memo:'4명 + 짐이 들어가는 차급으로'},
  {id:'b2', cat:'book', order:2, urgent:true, done:false, text:'수원화성문화제 9일 야간 공연 좌석 확인 (야조·선유몽·행궁야화)', link:'https://www.shfestival.com'},
  {id:'b3', cat:'book', order:3, urgent:true, done:false, text:'9일 점심 광교산 바베큐 식당 정해서 전화 예약', memo:'공휴일이라 붐빔 · 12:30 전후'},
  {id:'b4', cat:'book', order:4, done:false, text:'9일 저녁 수원 왕갈비 예약 / 원격 줄서기'},
  {id:'b5', cat:'book', order:5, done:false, text:'에어비앤비 호스트에게 10일 오전 짐 보관 · 체크인 시간 문의'},
  {id:'b6', cat:'book', order:6, done:false, text:'10일 창덕궁 달빛기행 / 경복궁 별빛야행 취소표 확인', memo:'없으면 청계천·광화문광장 야경 산책으로'},
  {id:'b7', cat:'book', order:7, done:false, text:'10일 한복 대여 예약 (부모님이 원하시면)'},
  {id:'b8', cat:'book', order:8, done:false, text:'11일 10:00 서울역 출발 열차 승차권 재확인'},
  {id:'s1', cat:'shop', order:10, done:false, text:'육회·육사시미용 소고기 (우둔/홍두깨, 당일 구매)'},
  {id:'s2', cat:'shop', order:11, done:false, text:'육회 양념: 배, 참기름, 마늘, 잣, 계란 노른자'},
  {id:'s3', cat:'shop', order:12, done:false, text:'계란국 · 죽순나물 재료'},
  {id:'s4', cat:'shop', order:13, done:false, text:'부모님 주무실 이부자리·수건 준비'}
];

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
}
function read_() {
  const v = sheet_().getRange('A1').getValue();
  if (!v) return {todos: DEFAULT_TODOS, memos: []};
  try { return JSON.parse(v); } catch (e) { return {todos: DEFAULT_TODOS, memos: []}; }
}
function write_(s) { sheet_().getRange('A1').setValue(JSON.stringify(s)); }
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function str_(v, n) { return String(v == null ? '' : v).slice(0, n || MAX_TEXT); }

function doGet() { return out_(read_()); }

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const b = JSON.parse(e.postData.contents);
    const s = read_();
    const id = str_(b.id, 40);
    if (b.op === 'addTodo' && b.item) {
      const it = b.item;
      s.todos.push({id: str_(it.id, 40), text: str_(it.text, 200), cat: ['book','shop','etc'].indexOf(it.cat) >= 0 ? it.cat : 'etc',
                    done: !!it.done, order: Number(it.order) || 0});
    } else if (b.op === 'toggleTodo') {
      s.todos.forEach(t => { if (t.id === id) t.done = !!b.done; });
    } else if (b.op === 'delTodo') {
      s.todos = s.todos.filter(t => t.id !== id);
    } else if (b.op === 'addMemo' && b.item) {
      const m = b.item;
      s.memos.unshift({id: str_(m.id, 40), text: str_(m.text), who: str_(m.who, 20), at: Number(m.at) || Date.now()});
      s.memos = s.memos.slice(0, 200);
    } else if (b.op === 'delMemo') {
      s.memos = s.memos.filter(m => m.id !== id);
    }
    write_(s);
    return out_(s);
  } finally {
    lock.releaseLock();
  }
}
