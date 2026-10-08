/**
 * 부모님 수원·서울 나들이 — 공유 체크리스트/메모 저장소
 * 구글 시트 > 확장 프로그램 > Apps Script 에 이 코드를 붙여넣고
 * 배포 > 새 배포 > 웹 앱 (실행: 나, 액세스 권한: 모든 사용자) 로 배포하세요.
 */
const SHEET_NAME = 'state';
const MAX_TEXT = 1000;

const DEFAULT_TODOS = [
  {id:'b1', cat:'book', order:1, urgent:true, done:false, text:'수원화성문화제 9일 야간 공연 좌석 확인 (선유몽·야조·행궁야화)', link:'https://www.shfestival.com'},
  {id:'b2', cat:'book', order:2, urgent:true, done:false, text:'9일 점심 폭포농원식당 전화로 자리 확인', memo:'031-256-9774 · 12:15 전후'},
  {id:'b3', cat:'book', order:3, urgent:true, done:false, text:'10일 수원역 → 서울역 기차 예매 (ITX-새마을·무궁화)', memo:'코레일톡 · 09:10 전후'},
  {id:'b4', cat:'book', order:4, done:false, text:'9일 저녁 연포갈비 웨이팅 방법 확인'},
  {id:'b5', cat:'book', order:5, done:false, text:'에어비앤비 호스트에게 10일 오전 짐 보관 · 체크인 시간 문의'},
  {id:'b6', cat:'book', order:6, done:false, text:'10일 창덕궁 달빛기행 / 경복궁 별빛야행 취소표 확인'},
  {id:'b7', cat:'book', order:7, done:false, text:'10일 한복 대여 예약 (부모님이 원하시면)'},
  {id:'b8', cat:'book', order:8, done:false, text:'11일 10:00 서울역 출발 열차 승차권 재확인'},
  {id:'s1', cat:'shop', order:10, done:false, text:'꾸리살 300g (육회) · 꾸리살·설깃살 600g (뭉티기)'},
  {id:'s2', cat:'shop', order:11, done:false, text:'배 · 계란 · 구운 김 · 무순 · 청양고추 · 통마늘'},
  {id:'s3', cat:'shop', order:12, done:false, text:'순두부 1팩 · 대파 · 묵은지 · 냉동 죽순 해동'},
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
