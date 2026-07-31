# 고유명·혐오 표현 이진 체크 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 신어 판별(1차·2차)에 고유명 여부/혐오 표현 여부 Y·N 체크를 추가하고, 두 값이 없으면 서버가 저장을 거부하게 한다.

**Architecture:** Google Apps Script 백엔드(`Code.gs`, 시트 바인딩) + 단일 파일 프론트(`index.html`, GitHub Pages). 프론트는 `google.script.run` 셰임 → `fetch`/`doPost` JSON API 호출. 모든 시트 접근은 `headerIndex_` 헤더명 기준이라 컬럼 위치와 무관.

**Tech Stack:** Apps Script(ES5 스타일 JS), 바닐라 JS/CSS, clasp 배포, git/GitHub Pages.

**Spec:** `docs/superpowers/specs/2026-07-31-verdict-binary-checks-design.md`

## Global Constraints

- 컬럼명(정확히): `1차 고유명 여부`, `1차 혐오 표현 여부`, `2차 고유명 여부`, `2차 혐오 표현 여부`
- 저장값: `Y` / `N` (빈 문자열 = 미판단)
- 필터 필드 이름: **주석 결과** (id `f-chk`)
- 서버 오류 문구: `고유명·혐오 표현 여부를 모두 선택해야 저장됩니다.`
- 코드 스타일: 기존 파일과 동일하게 ES5(`var`, 함수 선언), 한국어 주석·문구, 세미콜론 사용
- 이 프로젝트에는 테스트 프레임워크가 없음. 각 태스크는 **문법 검증(node) + 코드 리뷰**로 검증하고, 동작 검증은 Task 4 배포 후 스모크 테스트 체크리스트로 수행
- 배포는 Task 4에서만. **백엔드·프론트를 같은 시점에 배포**(구 프론트+신 백엔드 조합은 저장 거부, 신 프론트+구 백엔드는 체크값 유실). 배포 전 사용자 확인 필수
- dev(`KNO-workbench-dev`) 미러링은 범위 외

---

### Task 1: 백엔드 — 스키마·검증·마이그레이션 (Code.gs)

**Files:**
- Modify: `Code.gs` (16-22 HEADERS, 349-350 createProject CLEAR, 399 getTemplate, 467 LIST_FIELDS, 528-547 genAgree, 553-556 genReal, 588-630 saveFirst/saveSecond)

**Interfaces:**
- Produces: `saveFirst(token, payload)` / `saveSecond(token, payload)`의 payload가 `{ row_id, verdict, memo, proper, hate, op_id }`로 확장됨. `proper`·`hate`는 `'Y'|'N'` 필수, 아니면 throw. `getItems`/`getItem` 응답에 4개 체크 컬럼 포함(값 `'Y'|'N'|''`).
- Consumes: 없음 (기존 코드만 수정)

- [ ] **Step 1: HEADERS에 체크 컬럼 삽입 + 상수 추가**

`Code.gs:16-22`의 `HEADERS`를 다음으로 교체(1차 메모 뒤, 2차 메모 뒤에 삽입):

```js
var HEADERS = [
  'ID', '신어 후보', '작업자', '검수자', '배정 주차',
  '1차 판별', '1차 일시', '1차 메모', '1차 고유명 여부', '1차 혐오 표현 여부',
  '2차 판별', '2차 일시', '2차 메모', '2차 고유명 여부', '2차 혐오 표현 여부',
  '상태', '작업 구분', '출처', '추출 시기',
  'LLM 판단 결과', 'LLM 판단 기준', 'LLM 판단 근거',
  '용례', '용례 일자', '용례 URL', '검색 URL'
];
```

바로 아래 `var VERDICTS = ...` 줄 다음에 추가:

```js
var CHECK_COLS = ['1차 고유명 여부', '1차 혐오 표현 여부', '2차 고유명 여부', '2차 혐오 표현 여부'];
function chkOk_(v) { return v === 'Y' || v === 'N'; }
```

- [ ] **Step 2: 기존 시트 자동 마이그레이션 헬퍼 추가**

`headerIndex_` 함수(`Code.gs:83-87`) 바로 아래에 추가:

```js
// 체크 컬럼(고유명·혐오)이 없는 기존 프로젝트 시트에 자동 추가 후 최신 헤더 인덱스 반환.
function ensureCheckCols_(sh) {
  var idx = headerIndex_(sh), missing = [];
  for (var i = 0; i < CHECK_COLS.length; i++) if (!(CHECK_COLS[i] in idx)) missing.push(CHECK_COLS[i]);
  if (missing.length) {
    sh.getRange(1, sh.getLastColumn() + 1, 1, missing.length).setValues([missing]);
    styleHeader_(sh, sh.getLastColumn());
    idx = headerIndex_(sh);
  }
  return idx;
}
```

- [ ] **Step 3: saveFirst 검증·기록 확장**

`Code.gs:588-609`의 `saveFirst`를 다음으로 교체:

```js
function saveFirst(token, payload) {
  var me = me_(token);
  if (VERDICTS.indexOf(payload.verdict) === -1) throw new Error('판별 값 오류');
  if (!chkOk_(payload.proper) || !chkOk_(payload.hate)) throw new Error('고유명·혐오 표현 여부를 모두 선택해야 저장됩니다.');
  var lock = LockService.getDocumentLock(); lock.waitLock(20000);
  try {
    if (opSeen_(payload.op_id)) return { ok: true, dup: true };
    var sh = projSheetOfRow_(payload.row_id); if (!sh) throw new Error('프로젝트 없음');
    var idx = ensureCheckCols_(sh), rownum = findRow_(sh, idx, payload.row_id);
    if (rownum < 0) throw new Error('행 없음: ' + payload.row_id);
    assertCanEditStage_(me, sh.getRange(rownum, idx['작업자'] + 1).getValue(), sh.getRange(rownum, idx['검수자'] + 1).getValue(), 1);
    var cand = String(sh.getRange(rownum, idx['신어 후보'] + 1).getValue());
    var prev = String(sh.getRange(rownum, idx['상태'] + 1).getValue()).trim();
    var prevVerdict = String(sh.getRange(rownum, idx['1차 판별'] + 1).getValue()).trim();
    var prevTs = String(sh.getRange(rownum, idx['1차 일시'] + 1).getValue()).trim();
    setCell_(sh, rownum, idx, '1차 판별', payload.verdict);
    setCell_(sh, rownum, idx, '1차 메모', payload.memo || '');
    setCell_(sh, rownum, idx, '1차 고유명 여부', payload.proper);
    setCell_(sh, rownum, idx, '1차 혐오 표현 여부', payload.hate);
    if (!(prevVerdict === payload.verdict && prevTs)) setCell_(sh, rownum, idx, '1차 일시', now_());   // 같은 판별 재저장(체크 추가)이면 원래 일시 보존
    var ns = String(sh.getRange(rownum, idx['2차 판별'] + 1).getValue()).trim() ? STATUS.SECOND : STATUS.FIRST;
    setCell_(sh, rownum, idx, '상태', ns);
    appendLog_(me, payload.row_id, cand, '1차', payload.verdict + ' (고유명' + payload.proper + '/혐오' + payload.hate + ')', payload.memo || '', prev, ns);
    SpreadsheetApp.flush(); opMark_(payload.op_id);
    return { ok: true };
  } finally { lock.releaseLock(); }
}
```

- [ ] **Step 4: saveSecond 동일 확장**

`Code.gs:610-630`의 `saveSecond`를 다음으로 교체:

```js
function saveSecond(token, payload) {
  var me = me_(token);
  if (VERDICTS.indexOf(payload.verdict) === -1) throw new Error('판별 값 오류');
  if (!chkOk_(payload.proper) || !chkOk_(payload.hate)) throw new Error('고유명·혐오 표현 여부를 모두 선택해야 저장됩니다.');
  var lock = LockService.getDocumentLock(); lock.waitLock(20000);
  try {
    if (opSeen_(payload.op_id)) return { ok: true, dup: true };
    var sh = projSheetOfRow_(payload.row_id); if (!sh) throw new Error('프로젝트 없음');
    var idx = ensureCheckCols_(sh), rownum = findRow_(sh, idx, payload.row_id);
    if (rownum < 0) throw new Error('행 없음: ' + payload.row_id);
    assertCanEditStage_(me, sh.getRange(rownum, idx['작업자'] + 1).getValue(), sh.getRange(rownum, idx['검수자'] + 1).getValue(), 2);
    var cand = String(sh.getRange(rownum, idx['신어 후보'] + 1).getValue());
    var prev = String(sh.getRange(rownum, idx['상태'] + 1).getValue()).trim();
    var prevVerdict = String(sh.getRange(rownum, idx['2차 판별'] + 1).getValue()).trim();
    var prevTs = String(sh.getRange(rownum, idx['2차 일시'] + 1).getValue()).trim();
    setCell_(sh, rownum, idx, '2차 판별', payload.verdict);
    setCell_(sh, rownum, idx, '2차 메모', payload.memo || '');
    setCell_(sh, rownum, idx, '2차 고유명 여부', payload.proper);
    setCell_(sh, rownum, idx, '2차 혐오 표현 여부', payload.hate);
    if (!(prevVerdict === payload.verdict && prevTs)) setCell_(sh, rownum, idx, '2차 일시', now_());   // 같은 판별 재저장(체크 추가)이면 원래 일시 보존
    setCell_(sh, rownum, idx, '상태', STATUS.SECOND);
    appendLog_(me, payload.row_id, cand, '2차', payload.verdict + ' (고유명' + payload.proper + '/혐오' + payload.hate + ')', payload.memo || '', prev, STATUS.SECOND);
    SpreadsheetApp.flush(); opMark_(payload.op_id);
    return { ok: true };
  } finally { lock.releaseLock(); }
}
```

- [ ] **Step 5: 배분·목록·템플릿 반영**

5-1. `createProject`(`Code.gs:350`)의 `CLEAR` 배열에 4컬럼 추가:

```js
var CLEAR = ['작업자', '검수자', '배정 주차', '1차 판별', '1차 메모', '1차 일시', '1차 고유명 여부', '1차 혐오 표현 여부', '2차 판별', '2차 메모', '2차 일시', '2차 고유명 여부', '2차 혐오 표현 여부'];   // 배정·판별 추적만 초기화(집필 내용·1차/2차 뜻풀이·메모는 보존)
```

5-2. `genAgree`(`Code.gs:524-548`): `var sh = projItemSheet_(projectId); if (!sh) throw new Error('시트 없음');` 다음 줄의 `var idx = headerIndex_(sh), ...`를 `var idx = ensureCheckCols_(sh), ...`로 변경. 그리고 리셋 줄(`Code.gs:539`)에 4컬럼 추가:

```js
      o[C['1차 판별']] = ''; o[C['1차 메모']] = ''; o[C['1차 일시']] = ''; o[C['2차 판별']] = ''; o[C['2차 메모']] = ''; o[C['2차 일시']] = '';
      o[C['1차 고유명 여부']] = ''; o[C['1차 혐오 표현 여부']] = ''; o[C['2차 고유명 여부']] = ''; o[C['2차 혐오 표현 여부']] = '';
```

5-3. `genReal`(`Code.gs:549-578`): `var idx = headerIndex_(sh), ...`를 `var idx = ensureCheckCols_(sh), ...`로 변경만 한다. **체크값을 지우지 않음** — genReal은 판별값을 보존하므로 체크값도 함께 보존(판별·체크 원자성).

5-4. `LIST_FIELDS`(`Code.gs:467`)에 4컬럼 추가:

```js
var LIST_FIELDS = ['ID', '신어 후보', '출처', '추출 시기', '작업 구분', '작업자', '검수자', '배정 주차', '상태', '1차 판별', '2차 판별', '1차 고유명 여부', '1차 혐오 표현 여부', '2차 고유명 여부', '2차 혐오 표현 여부'];
```

5-5. `getTemplate`(`Code.gs:399`)의 반환 문자열에 4컬럼 추가:

```js
  return '작업 구분,작업자,검수자,배정 주차,상태,1차 판별,1차 메모,1차 일시,1차 고유명 여부,1차 혐오 표현 여부,2차 판별,2차 메모,2차 일시,2차 고유명 여부,2차 혐오 표현 여부,ID,신어 후보,출처,추출 시기,LLM 판단 결과,LLM 판단 기준,LLM 판단 근거,용례,용례 일자,용례 URL,검색 URL';
```

- [ ] **Step 6: 문법 검증**

Run: `node --check Code.gs`
Expected: 출력 없음(문법 오류 없음)

추가 확인(4컬럼이 정확한 이름으로 등장하는지):
Run: `grep -c "고유명 여부" Code.gs` → 8 이상, `grep -c "혐오 표현 여부" Code.gs` → 8 이상

- [ ] **Step 7: Commit**

```bash
git add Code.gs
git commit -m "feat: 판별 저장에 고유명·혐오 표현 Y/N 체크 필수화(서버 검증·시트 자동 마이그레이션)"
```

---

### Task 2: 프론트 — 판별 패널 체크 카드 + 저장 게이트 (index.html)

**Files:**
- Modify: `index.html` (CSS ~74-79 부근에 스타일 추가, 883-891 selectItem, 901-921 stagePanel/SEL, 923-949 renderDetail, 1060 pick 부근, 1065-1066 markDirty/autosave, 1129-1145 doSave)

**Interfaces:**
- Consumes: Task 1의 `saveFirst`/`saveSecond` payload `{ row_id, verdict, memo, proper, hate, op_id }`, 항목 필드명 `1차 고유명 여부` 등 4종.
- Produces: 전역 `CHK` 상태 `{'1차':{proper:'',hate:''},'2차':{proper:'',hate:''}}`, 함수 `checkCard(stage, allowed)`, `pickChk(stage, field, val)`, `stageComplete(stage)`. Task 3이 `CHK`는 쓰지 않고 항목 필드값만 사용.

- [ ] **Step 1: CSS 추가**

`<style>` 내 `.verdicts`/`.vcard` 스타일 근처(`.filters` 규칙 뒤 아무 곳)에 추가:

```css
    .chk-card { border:1px solid var(--line); border-radius:10px; padding:8px 10px; margin:8px 0 4px; display:flex; flex-direction:column; gap:6px; }
    .chk-row { display:flex; align-items:center; gap:6px; }
    .chk-lbl { font-size:13px; flex:1; }
    .chk-btn { width:44px; height:28px; border:1px solid var(--line); border-radius:8px; background:transparent; cursor:pointer; font-weight:bold; }
    .chk-btn.on { background:var(--sineo, #2f6fed); color:#fff; border-color:transparent; }
    .chk-btn:disabled { cursor:default; opacity:.55; }
```

- [ ] **Step 2: CHK 상태·카드 빌더·선택 함수 추가**

`index.html:901`의 `var SEL={'1차':'','2차':''};` 바로 아래에 추가:

```js
    var CHK = { '1차': { proper: '', hate: '' }, '2차': { proper: '', hate: '' } };
    var CHK_FIELDS = [['고유명 여부', 'proper'], ['혐오 표현 여부', 'hate']];
    function stageComplete(stage){ return !!(SEL[stage] && CHK[stage].proper && CHK[stage].hate); }
    function checkCard(stage, allowed){
      return '<div class="chk-card" data-stage="'+stage+'">'+CHK_FIELDS.map(function(cf){
        return '<div class="chk-row"><span class="chk-lbl">'+cf[0]+'</span>'+['Y','N'].map(function(v){
          var on = CHK[stage][cf[1]]===v;
          return '<button type="button" class="chk-btn'+(on?' on':'')+'" data-f="'+cf[1]+'" data-v="'+v+'"'
            +(allowed?' onclick="pickChk(\''+stage+'\',\''+cf[1]+'\',\''+v+'\')"':' disabled')+'>'+v+'</button>';
        }).join('')+'</div>';
      }).join('')+'</div>';
    }
    function pickChk(stage, f, v){
      CHK[stage][f]=v;
      document.querySelectorAll('.chk-card[data-stage="'+stage+'"] .chk-btn[data-f="'+f+'"]').forEach(function(b){ b.classList.toggle('on', b.getAttribute('data-v')===v); });
      markDirty();
    }
```

- [ ] **Step 3: stagePanel에 체크 카드 삽입 + renderDetail에서 CHK 초기화**

3-1. `stagePanel`(`index.html:917-921`)의 return 문에서 `verdictCards(stage, SEL[stage], allowed)` 뒤에 `+checkCard(stage, allowed)` 삽입:

```js
      return '<div class="stage-panel '+cls+'"><div class="sp-head">'+title+'</div>'+verdictCards(stage, SEL[stage], allowed)+checkCard(stage, allowed)+memoBox(stage,memoLabel,memoVal,!allowed)+actionRow(stage,allowed,isReviewer,when)+'</div>';
```

3-2. `renderDetail`(`index.html:926`)의 `SEL={...};` 바로 아래에 추가:

```js
      CHK = { '1차': { proper: String(it['1차 고유명 여부']||'').trim(), hate: String(it['1차 혐오 표현 여부']||'').trim() },
              '2차': { proper: String(it['2차 고유명 여부']||'').trim(), hate: String(it['2차 혐오 표현 여부']||'').trim() } };
```

- [ ] **Step 4: 저장 게이트 — doSave·autosave·selectItem**

4-1. `doSave`(`index.html:1129-1143`)를 다음으로 교체:

```js
    function doSave(stage, advance){
      var it=ITEMS[CUR]; if(!it) return;
      if(!canEdit(it, stage==='1차'?1:2)) return;
      var v=SEL[stage], c=CHK[stage], miss=[];
      if(!v) miss.push('판별');
      if(!c.proper) miss.push('고유명 여부');
      if(!c.hate) miss.push('혐오 표현 여부');
      if(miss.length){
        if(advance) toast(miss.join(' · ')+'를 선택해야 저장됩니다 (판별 Ctrl+1/2/3)');
        else setSaveBadge('고유명·혐오 여부 선택 필요', true);
        return;
      }
      var memo=(($('memo-'+stage)||{}).value)||'';
      clearTimeout(AUTO_TIMER);
      var fn = stage==='1차'?'saveFirst':'saveSecond';
      var payload = { row_id:it['ID'], verdict:v, memo:memo, proper:c.proper, hate:c.hate, op_id:newOpId() };
      // 낙관적 반영 + 즉시 다음(저장은 백그라운드)
      if(stage==='1차'){ it['1차 판별']=v; it['1차 메모']=memo; it['1차 고유명 여부']=c.proper; it['1차 혐오 표현 여부']=c.hate; it['상태']=it['2차 판별']?'2차완료':'1차완료'; }
      else { it['2차 판별']=v; it['2차 메모']=memo; it['2차 고유명 여부']=c.proper; it['2차 혐오 표현 여부']=c.hate; it['상태']='2차완료'; }
      DIRTY=false; applyResults();
      enqueueSave(fn, payload);   // 무손실 큐(실패 시 자동 재시도, 탭 닫아도 보존)
      if(advance) nextItem();
    }
```

4-2. `autosave`(`index.html:1066`)를 교체:

```js
    function autosave(){ if(!DIRTY) return; var stage=window.__activeStage||'1차'; if(stageComplete(stage)) doSave(stage, false); else setSaveBadge('고유명·혐오 여부 선택 필요', true); }
```

4-3. `selectItem`(`index.html:884`)의 첫 줄 `if(DIRTY){ doSave(window.__activeStage||'1차', false); }`를 교체:

```js
      if(DIRTY){ var st=window.__activeStage||'1차';
        if(stageComplete(st)) doSave(st, false);
        else { toast('저장되지 않음: 판별·고유명·혐오 여부를 모두 선택해야 저장됩니다.'); DIRTY=false; clearTimeout(AUTO_TIMER); setSaveBadge('', false); } }
```

- [ ] **Step 5: 문법 검증**

Run(인라인 스크립트 블록 컴파일 검사):

```bash
node -e "var s=require('fs').readFileSync('index.html','utf8');var re=/<script>([\s\S]*?)<\/script>/g,m,n=0;while((m=re.exec(s))){new Function(m[1]);n++;}console.log('scripts OK:',n)"
```

Expected: `scripts OK: 1` (문법 오류 시 SyntaxError로 실패)

- [ ] **Step 6: Commit**

```bash
git add index.html
git commit -m "feat: 판별 패널에 고유명·혐오 Y/N 체크 카드 추가, 미선택 시 저장 차단"
```

---

### Task 3: 프론트 — 필터 2열 개편 + 주석 결과 필터 + 체크 필요 배지 (index.html)

**Files:**
- Modify: `index.html` (74-79 CSS, 351-375 필터 마크업, 545 showView 리셋 줄, 853-858 itemPasses, 863-882 renderList)

**Interfaces:**
- Consumes: Task 1이 `getItems` 목록 응답에 포함시킨 `1차 고유명 여부` 등 4개 필드(값 `'Y'|'N'|''`).
- Produces: 신규 `#f-chk` 셀렉트(옵션 value: `1p:Y`,`1p:N`,`1h:Y`,`1h:N`,`1:miss`,`2p:Y`,`2p:N`,`2h:Y`,`2h:N`,`2:miss`), `chkOptions()` 함수.

- [ ] **Step 1: 필터 CSS를 2열 그리드로 변경**

`index.html:74-79`를 다음으로 교체:

```css
    .filters { display:grid; grid-template-columns:1fr 1fr; gap:6px; margin-bottom:10px; }
    .filters input, .filters select { height:34px; padding:6px 8px; border:1px solid var(--line); border-radius:8px; font-size:13px; line-height:1.2; box-sizing:border-box; vertical-align:middle; width:100%; min-width:0; }
    #f-q, #f-assignee, #f-chk { grid-column:1 / -1; }
```

(기존 `#f-status, #f-v1, #f-v2 { width:124px; }`와 `.filters input { flex:1; min-width:110px; }` 규칙은 삭제. 74-79 사이의 다른 규칙이 있으면 보존.)

- [ ] **Step 2: 필터 마크업 재배치 + f-chk 추가**

`index.html:351-375`의 `.filters` 내부를 다음 순서로 교체(검색 → 팀 → 주차·상태 1행 → 1차·2차 결과 1행 → 주석 결과):

```html
          <div class="filters">
            <input id="f-q" placeholder="신어 후보 항목 검색" onkeydown="if(event.key==='Enter')loadItems()">
            <select id="f-assignee" onchange="loadItems()" style="display:none"><option value="">팀 전체</option></select>
            <select id="f-week" onchange="loadItems()" style="display:none">
              <option value="">전체 주차</option>
            </select>
            <select id="f-status" onchange="loadItems()">
              <option value="">전체 상태</option>
              <option value="미작업">미작업</option>
              <option value="1차완료">1차 작업 완료</option>
              <option value="2차완료">2차 검수 완료</option>
            </select>
            <select id="f-v1" onchange="applyResults()" style="display:none" title="1차 작업 결과">
              <option value="">1차 작업 결과</option>
              <option value="신어">신어</option>
              <option value="비신어">비신어</option>
              <option value="판단 보류">판단 보류</option>
            </select>
            <select id="f-v2" onchange="applyResults()" style="display:none" title="2차 검수 결과">
              <option value="">2차 검수 결과</option>
              <option value="신어">신어</option>
              <option value="비신어">비신어</option>
              <option value="판단 보류">판단 보류</option>
            </select>
            <select id="f-chk" onchange="applyResults()" style="display:none" title="주석 결과"></select>
          </div>
```

- [ ] **Step 3: chkOptions() 추가 + showView 연동**

3-1. `itemPasses` 함수 위(`index.html:853` 부근)에 추가:

```js
    function chkOptions(){
      function grp(st, label){
        return '<optgroup label="'+label+'">'
          +'<option value="'+st+'p:Y">고유명 Y</option><option value="'+st+'p:N">고유명 N</option>'
          +'<option value="'+st+'h:Y">혐오 표현 Y</option><option value="'+st+'h:N">혐오 표현 N</option>'
          +'<option value="'+st+':miss">체크 미완료</option></optgroup>';
      }
      return '<option value="">주석 결과</option>'+grp('1','1차 작업')+(MODE==='real'?grp('2','2차 검수'):'');
    }
```

3-2. `showView`(`index.html:545`)의 `else if(v==='mine' || v==='all'){ ... }` 줄에서 `$('f-v2').value='';` 뒤에 다음을 삽입(같은 줄 유지):

```js
$('f-chk').innerHTML=chkOptions(); $('f-chk').value=''; $('f-chk').style.display=(MODE!=='write')?'':'none';
```

- [ ] **Step 4: itemPasses에 주석 결과 필터 로직 추가**

`index.html:853-858`의 `itemPasses`를 다음으로 교체:

```js
    function itemPasses(it){
      var v1=$('f-v1').value, v2=$('f-v2').value;
      if(v1 && verdictOf(it,1)!==v1) return false;
      if(MODE==='real' && v2 && verdictOf(it,2)!==v2) return false;
      var ck=$('f-chk').value;
      if(ck){
        var st = ck.charAt(0)==='2' ? 2 : 1, pre = st===2 ? '2차 ' : '1차 ';
        var p=String(it[pre+'고유명 여부']||'').trim(), h=String(it[pre+'혐오 표현 여부']||'').trim();
        if(ck.indexOf(':miss')>0){ if(!(verdictOf(it,st) && (!p || !h))) return false; }   // 판별은 했는데 체크가 빈 항목
        else { var val = ck.charAt(1)==='p' ? p : h; if(val!==ck.slice(3)) return false; }
      }
      return true;
    }
```

- [ ] **Step 5: 목록에 ⚠ 체크 필요 배지**

5-1. CSS(`Task 3 Step 1`에서 추가한 `.filters` 규칙 뒤)에 추가:

```css
    .pill.warnchk { background:#fdecea; color:#c0392b; }
```

5-2. `renderList`(`index.html:863-882`)에서 `var meta = ...` 줄 앞에 추가:

```js
        var needChk = MODE!=='write' && (
          (v1 && (!String(it['1차 고유명 여부']||'').trim() || !String(it['1차 혐오 표현 여부']||'').trim())) ||
          (MODE==='real' && v2 && (!String(it['2차 고유명 여부']||'').trim() || !String(it['2차 혐오 표현 여부']||'').trim())) );
```

그리고 rows.push 줄의 `</span></div>` 앞(상태 pill 뒤)에 배지 삽입:

```js
        rows.push('<div class="item'+(i===CUR?' sel':'')+'" data-i="'+i+'" onclick="selectItem('+i+')"><div class="cand"><span class="nm">'+esc(it['신어 후보'])+'</span><span class="pill '+pillClass(st)+'">'+esc(label)+'</span>'+(needChk?'<span class="pill warnchk">⚠ 체크 필요</span>':'')+'</div><div class="meta">'+meta+'</div></div>');
```

- [ ] **Step 6: 문법 검증**

Run:

```bash
node -e "var s=require('fs').readFileSync('index.html','utf8');var re=/<script>([\s\S]*?)<\/script>/g,m,n=0;while((m=re.exec(s))){new Function(m[1]);n++;}console.log('scripts OK:',n)"
```

Expected: `scripts OK: 1`

- [ ] **Step 7: Commit**

```bash
git add index.html
git commit -m "feat: 필터 2열 개편 + 주석 결과(고유명·혐오) 필터 + 체크 필요 배지"
```

---

### Task 4: 동시 배포 + 스모크 테스트

**Files:**
- 없음(배포만). 작업 디렉터리: `KNO-workbench/`

**Interfaces:**
- Consumes: Task 1-3의 커밋 완료 상태.

⚠ **이 태스크는 실행 전 사용자에게 배포 여부를 명시적으로 확인받는다** (운영 데이터·연구원 사용 중 시스템).
⚠ **push·clasp deploy 전에 반드시 변경 내용(diff 요약)을 사용자에게 먼저 보여주고 승인받는다. 승인 없이 절대 push하지 않는다.**

- [ ] **Step 1: 배포 전 점검**

- `appsscript.json`에 `"webapp": { "executeAs": "USER_DEPLOYING", "access": "ANYONE_ANONYMOUS" }` 블록 존재 확인 (CLAUDE.md 함정 항목)
- `git log --oneline -5`로 Task 1-3 커밋 확인
- 사용자에게 배포 진행 확인 요청 (가능하면 연구원 작업이 없는 시간대)

- [ ] **Step 2: 백엔드 배포**

```bash
clasp push --force
clasp deploy -i AKfycbxGQ25QDvzAdXOCdYWXihv3Lkdj6zVXyq5M0KiGjccGJTRbiY1XRMvRjCHKrmlFdWLZ -d "고유명·혐오 체크 필수화"
```

반드시 기존 배포 ID(`-i`) 사용 — 새 배포를 만들면 `/exec` URL이 바뀌어 프론트가 깨짐.

- [ ] **Step 3: 프론트 배포(즉시 연달아)**

```bash
git push
```

GitHub Pages 반영 1~2분. 이 사이 구 프론트 저장은 서버가 거부함(무손실 큐가 재시도하므로 새로고침 후 체크 선택하면 복구됨).

- [ ] **Step 4: 스모크 테스트 (사용자 또는 관리자 계정으로 브라우저에서)**

1. 판별 프로젝트 항목 열기 → 판별 카드 아래 `고유명 여부 Y/N · 혐오 표현 여부 Y/N` 카드 표시 확인
2. 판별만 선택, 체크 미선택 → 배지 `고유명·혐오 여부 선택 필요`, Ctrl+S 시 토스트, 저장 안 됨 확인
3. 체크까지 선택 → 저장 → 시트에서 `1차 고유명 여부`/`1차 혐오 표현 여부` = Y/N 기록 확인 (기존 시트라면 컬럼이 끝에 자동 생성됐는지 확인)
4. 기존 판별 완료 항목 열기 → 기존 판별 선택된 채 표시 → 체크만 추가 저장 → `1차 일시`가 원래 시각 그대로인지 확인
5. 목록에서 판별 완료·체크 미기록 항목에 `⚠ 체크 필요` 배지 확인
6. `주석 결과` 필터: `1차 체크 미완료` 선택 → 재작업 대상만 목록에 남는지 확인, `1차 고유명 Y` 등 값 필터 확인
7. 일치도 프로젝트: 단일 패널에 체크 카드 표시 + `주석 결과` 필터에 2차 옵션군 없음 확인
8. 집필 프로젝트: 체크 카드·주석 결과 필터가 보이지 않는지 확인(회귀)
9. 변경로그 시트: 판별 컬럼에 `신어 (고유명N/혐오N)` 형태 기록 확인

- [ ] **Step 5: 실패 시 롤백**

- 백엔드: `clasp deploy -i AKfycbxGQ25Q... -d "롤백"` 전에 `git revert`로 Code.gs 되돌린 후 `clasp push --force`
- 프론트: `git revert <커밋>` 후 `git push`
