// ============================================================
// LV2 N카탈로그 수집 프로그램 - 버전 2
//
//  버전 2의 핵심: "키워드 대기줄"
//   - 사이드 패널에서 키워드를 언제든 추가하면 대기줄에 쌓이고,
//     작업 중에도 뒤에 계속 추가할 수 있다.
//   - 대기줄이 빌 때까지 하나씩 처리하고, 빌 때 CSV를 저장한다.
//   - 진행 상황은 페이지 가운데 모달 + 사이드 패널 양쪽에 보여준다.
//   - 시작 전에 회원 로그인(구글 시트 연동)을 반드시 확인한다.
// ============================================================

importScripts('config.js');
importScripts('xlsx.full.min.js');   // 엑셀(.xlsx) 파일 생성 라이브러리 (로컬 포함)

const searchUrlBase = 'https://search.shopping.naver.com/search/all?query=';
const NAVER_LOGIN_URL = 'https://nid.naver.com/nidlogin.login';

// 로그인 창을 "한 번만" 연다. 이미 열어 둔 창이 살아 있으면 다시 열지 않고(앞으로 끌어오지도 않는다) 그대로 둔다.
//  ★실측(2026-11-02): 로그인이 풀린 채로 3초마다 다시 돌면서 창을 계속 새로 열어 탭이 수십 개가 됐다.
let _loginTabId = null;
async function openNaverLoginOnce() {
    try {
        if (_loginTabId) {
            const t = await chrome.tabs.get(_loginTabId);
            if (t && /nid\.naver\.com/.test(String(t.url || ''))) return false;   // 이미 열려 있다
        }
    } catch (_) { _loginTabId = null; }
    try {
        const t = await chrome.tabs.create({ url: NAVER_LOGIN_URL, active: true });
        _loginTabId = t && t.id;
        return true;
    } catch (_) { return false; }
}

// 로그인·회원 확인에서 멈췄을 때는 5분에 한 번만 다시 해 본다 (그 사이 사용자가 로그인하면 다음 시도에 바로 이어간다)
const HALT_RETRY_MS = 5 * 60 * 1000;
let _haltAt = 0;
const BOM = String.fromCharCode(0xFEFF);

const CONFIG = {
    // --- 수집량 ---
    MAX_CATALOGS_PER_KEYWORD: 60, // 키워드당 상한 (1페이지 전체가 잘리지 않게 여유 있게)
    MAX_QTY_COLUMN: 10,           // CSV에 만들 수량 칸 (1개~10개)

    // --- 사람처럼 천천히 (숫자를 키우면 더 안전, 더 느림) ---
    READ_PAUSE_MIN: 2200,         // 페이지 열고 "읽는" 시간
    READ_PAUSE_MAX: 5200,
    CATALOG_DELAY_MIN: 5000,      // 카탈로그 사이 간격
    CATALOG_DELAY_MAX: 11000,
    KEYWORD_DELAY_MIN: 15000,     // 키워드 사이 간격
    KEYWORD_DELAY_MAX: 30000,
    LONG_BREAK_EVERY: 25,         // 카탈로그 N개마다 긴 휴식
    LONG_BREAK_MIN: 60000,
    LONG_BREAK_MAX: 120000,

    OPTION_READ_TRIES: 3,         // 옵션을 못 읽으면 펼치고 기다렸다 몇 번까지 다시 볼지
    CHECK_BUY_LINK: true,         // 등록 후보의 매입처 페이지를 열어 판매중단 여부를 확인할지
    SAVE_FAIL_REPORT: true,       // 못 읽은 카탈로그의 진단 파일을 같이 저장할지

    // --- 리뷰 필터 (실제로 잘 팔리는 카탈로그만) ---
    //  등록은 AI 토큰이 들지만 소싱은 안 든다. 그래서 "팔리는 것"만 등록 쪽으로 넘기도록
    //  리뷰 수로 거른다. 기준은 2주(14일) 리뷰 수. 패널의 구간(10·30·50…)에서 고른다.
    //  한 달로 보면 이미 트렌드에서 밀린 상품도 통과하므로 2주로 본다 (루크 지시 2026-11-01).
    REVIEW_FILTER: true,          // 켜면: 2주 리뷰가 기준 미만인 카탈로그는 건너뜀
    REVIEW_DAYS: 14,              // 최근 며칠 = 2주
    REVIEW_MIN_COUNT: 10,         // 2주 리뷰가 이 개수 미만이면 건너뜀 (패널에서 바꾼다)
    //  패널에서 고를 수 있는 구간. 100 미만은 10·30·50·70, 500 까지 50 단위, 그 위는 100 단위.
    REVIEW_TIERS: [10, 30, 50, 70, 100, 150, 200, 250, 300, 350, 400, 450, 500,
                   600, 700, 800, 900, 1000],
    // 리뷰를 '몇 페이지까지' 세는지. 기준(10개)에 닿아도 멈추지 않고 끝까지 센다
    //  — 판매량을 가늠하려면 "10개 넘음" 이 아니라 실제 개수가 필요하기 때문이다.
    //  한 페이지 20개. 상한에 닿으면 그 숫자에 '+' 를 붙여 "더 있다"를 알린다.
    //  네이버에 요청을 너무 많이 보내면 IP 가 막히므로 상한을 둔다.
    REVIEW_COUNT_PAGES: 6,       // 6페이지 = 최대 120개까지 정확히 셈

    // --- 낱개 신규 (묶음만 팔고 "1개" 옵션이 아직 없는 카탈로그) ---
    //  제목 끝에 ", 1개" 를 붙여 등록하면 네이버가 1개 옵션을 새로 만들어 준다.
    //  참고할 1개 최저가가 없으니, 목표순이익이 그대로 남는 값을 우리가 정해서 올린다.
    SOLO_NEW: true,               // 끄면 예전처럼 '1개 옵션 없음'으로만 두고 넘어간다
    SOLO_NEW_MAX_RATIO: 2.0,      // 낱개 판매가 ÷ 개당원가. 이보다 크면 '확인필요'로 표시
    SOLO_NEW_HARD_RATIO: 3.0,     // 이보다 크면 담지 않는다 (낱개로 팔기엔 너무 비싸다)

    // --- 빈 수량 신규 (옵션 사이에 비어 있는 수량 자리) ---
    //  택배비는 몇 개를 팔든 한 건만 나간다. 그래서 개당 차익이 얇은 상품은
    //  1개로는 영원히 목표순이익을 못 넘지만, 몇 개를 묶으면 넘긴다.
    //      최소수량 n >= (택배비 + 목표순이익) / (개당시세 x (1-수수료율) - 개당원가)
    //  비어 있는 자리면 깎을 상대가 없어 판매가를 우리가 정한다 (낱개신규와 같은 생각).
    GAP_QTY: true,                // 끄면 예전처럼 1개·메인수량만 본다
    GAP_QTY_MAX: 12,              // 이보다 많이 묶는 제안은 하지 않는다 (사 가는 사람이 없다)

    // --- 수량 칸 전부 (그룹상품 등록용, v4.97) ---
    //  네이버 그룹상품은 1개·2개·3개… 칸을 한 상품으로 묶어 올린다(등록 한도 1개로 산정).
    //  그래서 "제일 작은 묶음 하나"가 아니라 1~12개 칸 전부를 계산해 작업지시서에 담는다.
    //  칸마다 목표순이익 = 개당원가 x 수량 x 비율, 단 최소 1,000원 · 최대 10,000원 (루크 규칙).
    //    남이 있는 칸  → 그 값보다 10원 싸게. 그래도 최소 순이익이 안 남으면 그 칸은 뺀다.
    //    비어 있는 칸  → 목표순이익이 그대로 남는 값을 우리가 정한다.
    //  칸마다 개당 값이 내려갈 필요는 없다(쿠팡도 그렇다). 다만 더 큰 묶음보다 비싸면 뺀다.
    SLOT_MAX: 12,
    SLOT_PROFIT_MIN: 1000,
    SLOT_PROFIT_MAX: 10000,
    SLOT_PROFIT_RATE: 15,         // % — 패널에서 바꾼다
    WORKORDER_SPEC: '1.1',        // 1.1 = 매입.경로 '낱개'(순방향) 항목이 들어간다

    // --- 낱개 사입(순방향) (v4.98) ---
    //  네이버 1개 최저가 판매처(스마트스토어)에서 "최대할인가 − 적립(구매·멤버십·현대카드)" 로 1개를 사서
    //  2~12개 칸에 판다. 표시 최저가와 실제 사는 값의 차이가 마진이다 (루크 설계 2026-11-01).
    //  켜면 낱개 전용 카탈로그도 페이지를 연다(리뷰 기준은 똑같이 적용). 패널 상자로 끈다.
    FORWARD_BUY: true,
    //  공식몰 규칙(루크): 카탈로그 판매처에 공식몰이 있어야 하고, 그 값이 최저가이거나 최저가와 이 금액 미만 차이여야 한다.
    //  (공식몰이 아니면 멤버십·브랜드 혜택이 안 붙고, 최저가보다 많이 비싸면 혜택을 빼도 남지 않는다)
    FORWARD_OFFICIAL_GAP: 1000,        // 작업지시서 규격 번호. 칸 모양이 바뀌면 올린다 (자동등록이 본다)

    REQUIRE_NAVER_LOGIN: true,    // 네이버 로그인 안 되어 있으면 시작하지 않음
    CLOSE_WORKER_TAB: true
};

// 확장 아이콘을 누르면 사이드 패널이 열리게
try { chrome.sidePanel.setPanelBehavior({ openPanelOnActionClick: true }); } catch (e) {}

// ------------------------------------------------------------
// 버전 2 상태 (대기줄)
// ------------------------------------------------------------
const S = {
    running: false,
    stopRequested: false,
    queue: [],        // 다음에 추적할 키워드들
    current: null,    // { keyword, phase, catalogIndex, catalogTotal, collected }
    done: [],         // { keyword, total, hits }  ->  "콜라겐 (4/18)"
    results: [],      // 지금까지 모은 모든 카탈로그 행
    failures: [],     // 옵션을 못 읽은 건들 (진단 파일용)
    pending: [],      // 찾아놓고 아직 방문 못 한 카탈로그들 (중지돼도 저장됨)
    workerTabId: null,
    lastError: '',
    note: '',         // 준비 단계 안내 ("회원 확인 중..." 등) - 패널에 표시
    autoAdded: 0,     // [꼬리물기] 이번 실행에서 자동추가한 키워드 수 (연료통)
    deepAdded: 0,     // [꼬리물기] 이번 실행에 태운 심화(브랜드/제조사) 키워드 수 - 실행당 2개 제한
    nextTargets: [],  // [꼬리물기] 상한을 넘겨 이번엔 못 도는 키워드 ("다음사냥감" 파일)
    runStamp: '',     // 이번 실행의 표시 - 같은 실행의 파일들을 한 묶음으로 묶는다
    runRound: 1,      // 꼬리물기 차수: 「이어서 찾기」를 누를 때마다 1씩 커진다 (1차 -> 2차 -> ...)
    runHits: 0,       // [통계] 이번 실행에서 찾은 등록가능 상품 수
    runActiveMs: 0,   // [통계] 이번 실행의 실제 작업 시간 누적 (중지한 동안은 안 센다)
    skipN: newSkipN(),// [통계] 이번 실행에서 건너뛴 이유별 개수 (화면 요약용)
    catStat: {},      // [카테고리 성적] 경로별 {seen,bundle,ok,solo} — 판이 바뀌어도 쌓인다
    update: null,     // [업데이트] { latest, ready, note } — 새 버전이 있을 때만 채워진다
    updCheckedAt: 0,  // 마지막으로 새 버전을 확인한 시각 (저장 안 함)
    lastStateAt: 0,   // 패널이 마지막으로 상태를 물어본 시각 = 패널이 열려 있나 (저장 안 함)
    sessionStartMs: 0 // 지금 켜져 있는 작업 구간의 시작 시각 (저장 안 함)
};

// 건너뛴 이유를 세는 칸. 규칙을 새로 넣었을 때 "너무 많이 거르고 있지 않은지"를
// 바로 보려면, 몇 개를 왜 버렸는지가 숫자로 남아 있어야 한다. (v4.91)
function newSkipN() {
    return {
        review: 0,       // 최근 리뷰가 기준 미만 (잘 안 팔리는 카탈로그)
        solo: 0,         // 낱개(1개)만 파는 카탈로그 - 묶음 사입 자체가 불가능
        soloNew: 0,      // ★찾은 것: 묶음만 있어서 "1개" 를 새로 만들어 올릴 수 있는 카탈로그
        soloPricey: 0,   // 낱개로 올리면 너무 비싸져서 담지 않은 것
        noOne: 0,        // 1개 옵션이 없는데 낱개 신규로도 못 쓴 것
        bundleless: 0,   // 옵션을 읽었는데 2개 이상 묶음이 없음
        mismatch: 0,     // 매입처 실제 상품이 카탈로그와 다름
        notOfficial: 0   // 화장품인데 매입처가 공식 채널이 아님
    };
}

// ------------------------------------------------------------
// [카테고리 성적] 어느 카테고리에서 실제로 건지는지 쌓는다. (v4.93)
//
//  "묶음으로 파는 상품이 많은 카테고리"를 짐작으로 적어두는 대신,
//  방문한 카탈로그를 카테고리별로 세어 우리 데이터가 알려주게 한다.
//  방문마다 경로의 1~3단계를 각각 한 칸씩 올린다.
//    "식품 > 건강식품 > 홍삼" → "식품" / "식품 > 건강식품" / 전체, 세 줄
//
//  ⚠️ 낱개 전용으로 '방문 전에' 걸러낸 것은 카테고리를 모른다(페이지를 안 열었다).
//     그래서 묶음비율은 "방문한 것 중" 비율이다. 전체 낱개전용 수는 skipN.solo 로 본다.
// ------------------------------------------------------------
const CATSTAT_DEPTH = 3;      // 몇 단계까지 묶어서 셀지
const CATSTAT_MIN_SEEN = 3;   // 이만큼은 봐야 성적으로 친다 (1~2건은 운)
const CATSTAT_MAX_KEYS = 400; // 너무 불어나지 않게 상한

function bumpCatStat(cat) {
    const path = String(cat.catPath || cat.catSmall || '').trim();
    if (!path) return;
    const parts = path.split('>').map(s => s.trim()).filter(Boolean);
    if (!parts.length) return;

    const hasBundle = Number(cat.bestQty) > 0;
    const v = String(cat.verdict || '');
    const isOk = v.indexOf('등록가능') === 0;
    // 새 수량 자리 = 낱개신규(1개) + 빈수량신규(옵션 사이의 빈 칸).
    //  gapQty 는 목표순이익을 넘긴 경우에만 채워지므로 숫자만 보면 된다.
    const isSolo = v === '낱개신규' || Number(cat.gapQty) > 0;

    for (let d = 1; d <= Math.min(CATSTAT_DEPTH, parts.length); d++) {
        const key = parts.slice(0, d).join(' > ');
        const e = S.catStat[key] || (S.catStat[key] = { seen: 0, bundle: 0, ok: 0, solo: 0 });
        e.seen++;
        if (hasBundle) e.bundle++;
        if (isOk) e.ok++;
        if (isSolo) e.solo++;
    }
}

// 건진 수가 많은 순으로 정렬해서 돌려준다
function catStatList(minSeen) {
    const need = minSeen === undefined ? CATSTAT_MIN_SEEN : minSeen;
    const out = [];
    for (const key of Object.keys(S.catStat || {})) {
        const e = S.catStat[key];
        if (!e || e.seen < need) continue;
        const hit = e.ok + e.solo;
        out.push({
            path: key,
            seen: e.seen, bundle: e.bundle, ok: e.ok, solo: e.solo, hit: hit,
            bundleRate: Math.round((e.bundle / e.seen) * 1000) / 10,
            hitRate: Math.round((hit / e.seen) * 1000) / 10
        });
    }
    // 윗단계가 아랫단계와 숫자가 똑같으면(= 그 가지 하나가 전부면) 윗단계는 뺀다.
    //   안 그러면 "식품 / 식품 > 건강식품 / 식품 > 건강식품 > 홍삼" 이 같은 숫자로
    //   세 줄 나와서 목록이 같은 말만 반복하게 된다. 더 구체적인 쪽만 남긴다.
    const kept = out.filter(a => !out.some(b =>
        b !== a && b.path.indexOf(a.path + ' > ') === 0 && b.seen === a.seen));

    kept.sort((a, b) => (b.hit - a.hit) || (b.hitRate - a.hitRate) || (b.seen - a.seen));
    return kept;
}

// 엑셀 「카테고리 성적」 시트
function catStatAoa() {
    const list = catStatList();
    const aoa = [['카테고리', '방문', '묶음있음', '묶음비율(%)', '등록가능', '새수량자리', '건진수', '건진비율(%)']];
    for (const r of list) aoa.push([r.path, r.seen, r.bundle, r.bundleRate, r.ok, r.solo, r.hit, r.hitRate]);
    if (!list.length) {
        aoa.push(['아직 모인 게 없어요 — 같은 카테고리를 ' + CATSTAT_MIN_SEEN + '개 이상 봐야 나옵니다']);
    }
    return aoa;
}

// 패널에 보여줄 "지금까지 건진 것" 미리보기 — 엑셀을 안 열어도 감이 오게 (v4.94)
function topFinds(n) {
    const out = [];
    for (const r of (S.results || [])) {
        const v = String(r.verdict || '');
        const solo = v === '낱개신규';
        const ok = v.indexOf('등록가능') === 0;
        // 빈 수량 자리가 더 많이 남으면 그쪽을 보여준다 (안 보여주면 있는 줄을 모른다)
        const gapN = Number(r.gapQty) || 0;
        const gapWins = gapN > 1
            && Number(r.gapProfit) > Math.max(Number(solo ? r.soloProfit : r.profit) || -Infinity, -Infinity);
        if (!ok && !solo && !gapWins) continue;
        const useGap = gapWins;
        out.push({
            title: String(useGap ? qtyTitleOf(r.title, gapN)
                : (solo ? soloTitleOf(r.title) : (r.title || ''))).slice(0, 40),
            price: useGap ? r.gapPrice : (solo ? r.soloPrice : r.basePrice),
            profit: useGap ? r.gapProfit : (solo ? r.soloProfit : r.profit),
            solo: solo || useGap,
            rv: r.reviewRecent === '' || r.reviewRecent === undefined ? ''
                : (String(r.reviewRecent) + (r.reviewMore ? '+' : (r.reviewGuess ? '?' : '')))
        });
    }
    out.sort((a, b) => (Number(b.profit) || 0) - (Number(a.profit) || 0));
    return out.slice(0, n || 10);
}

// 너무 불어나면 많이 본 것부터 남기고 자른다
function trimCatStat() {
    const keys = Object.keys(S.catStat || {});
    if (keys.length <= CATSTAT_MAX_KEYS) return;
    keys.sort((a, b) => (S.catStat[b].seen || 0) - (S.catStat[a].seen || 0));
    const keep = {};
    for (const k of keys.slice(0, CATSTAT_MAX_KEYS)) keep[k] = S.catStat[k];
    S.catStat = keep;
}

function publicState() {
    return {
        running: S.running,
        stopRequested: S.stopRequested,
        queue: S.queue.slice(),
        current: S.current,
        done: S.done.slice(),
        resultCount: S.results.length,
        pendingCount: S.pending.length,
        lastError: S.lastError,
        note: S.note,
        autoAdded: S.autoAdded,
        nextCount: S.nextTargets.length,
        runHits: S.runHits,
        skipN: Object.assign({}, S.skipN),
        catTop: catStatList().slice(0, 10),
        topFinds: topFinds(10),
        update: S.update || null,
        appVer: appVer(),
        runActiveMs: S.runActiveMs + (S.running && S.sessionStartMs ? Date.now() - S.sessionStartMs : 0)
    };
}

// ------------------------------------------------------------
// [업데이트] 새 버전이 나왔는지 보고, 폴더가 이미 갈렸으면 스스로 새로고침한다.
//
//  역할 나눔은 config.js 의 설명을 보세요. 여기서는 두 가지만 합니다.
//   ① 깃허브의 최신 버전 번호를 가끔 확인해 패널에 알린다
//   ② 업데이트기가 남긴 `디스크버전.json` 이 지금 돌고 있는 버전과 다르면
//      → 폴더는 새 버전인데 크롬이 옛 코드를 돌리고 있다는 뜻 → 새로고침
//
//  ★새로고침하면 모아둔 결과가 날아갑니다. 그래서 **쉬는 중 + 모은 게 없을 때만** 합니다.
//    돌리는 중이거나 결과가 있으면 패널에 "준비됨" 만 띄우고 기다립니다.
// ------------------------------------------------------------
// 지금 저절로 갈아타도 되는 상태인가.
//  ① 돌리는 중이거나 모아둔 결과가 있으면 안 된다 (새로고침하면 날아간다)
//  ② 패널이 열려 있으면 안 된다 — 보고 있는 화면이 갑자기 닫혀 놀라게 된다.
//     이때는 ⬆️ 칸의 [지금 적용] 을 누르게 한다.
//  ※ 어차피 크롬을 다시 켜면 폴더에 있는 새 버전으로 뜬다. 그게 가장 흔한 경로다.
function updateIdle() {
    const panelOpen = S.lastStateAt && Date.now() - S.lastStateAt < 15000;
    return !S.running && !panelOpen
        && (S.results || []).length === 0 && (S.queue || []).length === 0;
}

// 새로고침해도 버전이 안 바뀌는 상황(업데이트기가 쪽지는 남겼는데 파일 교체가 덜 된 등)에서
// 끝없이 새로고침하는 사고를 막는다. 같은 버전으로는 두 번까지만 시도한다.
async function tryReloadFor(ver) {
    try {
        const st = await chrome.storage.local.get('updReload');
        const r = st.updReload || {};
        const n = (r.ver === ver ? (Number(r.n) || 0) : 0) + 1;
        await chrome.storage.local.set({ updReload: { ver: ver, n: n, at: Date.now() } });
        if (n > 2) {
            S.lastError = '새 버전(' + ver + ')으로 바꾸지 못했습니다. 설치파일을 다시 받아 설치해 주세요.';
            return false;
        }
    } catch (e) { /* 저장을 못 하면 세는 것만 못 하는 것 — 새로고침은 해본다 */ }
    chrome.runtime.reload();
    return true;
}

// 켜질 때: 목표 버전이 됐으면 시도 횟수를 지운다 (다음 업데이트를 위해)
async function clearReloadCountIfArrived() {
    try {
        const st = await chrome.storage.local.get('updReload');
        const r = st.updReload;
        if (r && r.ver && !verGreater(r.ver, appVer())) await chrome.storage.local.remove('updReload');
    } catch (e) {}
}

async function maybeCheckUpdate() {
    const now = Date.now();
    if (S.updCheckedAt && now - S.updCheckedAt < UPDATE_CONFIG.CHECK_EVERY_MIN * 60000) return;
    S.updCheckedAt = now;
    try {
        const running = appVer();
        const disk = await fetchDiskVersion();
        // 폴더가 이미 새 버전 → 돌고 있는 코드만 옛것. 쉬는 중이면 바로 갈아탄다.
        if (disk && disk.version !== running && verGreater(disk.version, running)) {
            S.update = { latest: disk.version, ready: true, note: disk.note || '' };
            if (updateIdle() && await tryReloadFor(disk.version)) return;
            broadcast();
            return;
        }
        const j = await fetchLatestVersion();
        if (!j) return;
        if (verGreater(j.latest, running)) {
            // 아직 폴더는 안 갈렸다 = 업데이트기가 받는 중이거나 아직 돌 차례가 아니다
            //  url 은 자동 업데이트가 등록되지 않은 PC 를 위한 대비책(직접 받는 곳)
            S.update = { latest: j.latest, ready: false, note: j.note || '', url: j.url || '',
                         prog: await fetchUpdateProgress() };   // 몇/몇 받았는지 (v5.2)
        } else {
            S.update = null;
        }
        broadcast();
    } catch (e) { /* 업데이트 확인 실패가 프로그램을 막으면 안 된다 */ }
}

// 받는 중일 때는 진행 쪽지를 자주(5초) 다시 읽는다 — 패널이 "몇/몇 (몇 %)" 를 보이게. (v5.2)
//  쪽지가 '완료' 가 되면 폴더가 갈린 것이니 바로 새 버전 확인으로 넘어간다.
let _progAt = 0;
async function refreshUpdateProgress() {
    if (!S.update || S.update.ready) return;
    const now = Date.now();
    if (now - _progAt < 5000) return;
    _progAt = now;
    try {
        const p = await fetchUpdateProgress();
        const before = JSON.stringify(S.update.prog || null);
        S.update.prog = p;
        if (p && p.stage === '완료' && p.latest && verGreater(p.latest, appVer())) { S.updCheckedAt = 0; maybeCheckUpdate(); return; }
        if (JSON.stringify(p || null) !== before) broadcast();
    } catch (e) {}
}

// 패널의 [지금 적용]. 폴더가 갈렸으면 새로고침, 아직이면 왜 아직인지 알려준다.
async function applyUpdateNow() {
    const running = appVer();
    const disk = await fetchDiskVersion();
    if (!disk || !verGreater(disk.version, running)) {
        return { ok: false, code: 'NOT_READY',
                 message: '새 파일이 아직 안 받아졌습니다. 잠시 뒤 다시 눌러주세요.' };
    }
    if (!updateIdle()) {
        return { ok: false, code: 'BUSY',
                 message: '모아둔 결과가 날아갑니다. 엑셀로 내보낸 뒤 「🧹 기록 비우기」를 하고 눌러주세요.' };
    }
    chrome.runtime.reload();
    return { ok: true };
}

function broadcast() {
    chrome.runtime.sendMessage({ action: 'v2state', state: publicState() }).catch(() => {});
}

async function persist() {
    try {
        // failDebug의 HTML 조각은 크므로 결과 저장본에서는 뺀다
        const slim = S.results.map(r => {
            const c = Object.assign({}, r);
            delete c.failDebug;
            return c;
        });
        trimCatStat();   // 카테고리 성적이 너무 불어나지 않게 정리하고 저장한다
        await chrome.storage.local.set({
            v2queue: { queue: S.queue, done: S.done, pending: S.pending, autoAdded: S.autoAdded, deepAdded: S.deepAdded, nextTargets: S.nextTargets, runStamp: S.runStamp, runRound: S.runRound, runHits: S.runHits, runActiveMs: S.runActiveMs, skipN: S.skipN, catStat: S.catStat,
                       // (v5.4) 다시 켜질 때 "하던 작업부터" 이어가기 위한 것들
                       current: S.current ? String(S.current.keyword || '') : '',   // 하고 있던 키워드 (죽으면 대기줄 맨 앞으로 되돌린다)
                       workerTabId: S.workerTabId || null,                           // 작업 창 — 같은 창을 다시 쓴다 (2개 뜨는 것 방지)
                       userStopped: !!S.userStopped,                                 // 사람이 [중지]를 눌러 멈춘 판은 저절로 되살리지 않는다
                       haltedFor: S.haltedFor || '', haltAt: _haltAt || 0 },        // 로그인·회원 멈춤은 5분 간격 규칙을 다시 켜져도 지킨다
            v2results: { results: slim, failures: S.failures }
        });
    } catch (e) { console.error('[저장] 실패:', e); }
}

// 서비스워커가 재시작되면 상태 복구
(async function restore() {
    try {
        const st = await chrome.storage.local.get(['v2queue', 'v2results']);
        if (st.v2queue) {
            S.queue = st.v2queue.queue || [];
            S.done = st.v2queue.done || [];
            S.pending = st.v2queue.pending || [];
            S.autoAdded = st.v2queue.autoAdded || 0;
            S.deepAdded = st.v2queue.deepAdded || 0;
            S.nextTargets = st.v2queue.nextTargets || [];
            S.runStamp = st.v2queue.runStamp || '';
            S.runRound = st.v2queue.runRound || 1;
            S.runHits = st.v2queue.runHits || 0;
            S.runActiveMs = st.v2queue.runActiveMs || 0;
            S.skipN = Object.assign(newSkipN(), st.v2queue.skipN || {});
            S.catStat = st.v2queue.catStat || {};
            // (v5.4) 하던 키워드는 대기줄 맨 앞으로 — 검색하다 죽은 키워드가 사라지지 않게
            const cur = String(st.v2queue.current || '');
            if (cur && !S.queue.some(k => bareKw(k) === bareKw(cur))) S.queue.unshift(cur);
            S.workerTabId = st.v2queue.workerTabId || null;   // 살아 있으면 그 창을 다시 쓴다 (ensureWorkerTab 이 확인)
            S.userStopped = !!st.v2queue.userStopped;
            S.haltedFor = st.v2queue.haltedFor || '';
            _haltAt = Number(st.v2queue.haltAt) || 0;
        }
        if (st.v2results) {
            S.results = st.v2results.results || [];
            S.failures = st.v2results.failures || [];
        }
    } catch (e) {}

    // ★(v5.4, 루크) 다시 켜지면 하던 작업부터 이어간다.
    //   크롬을 다시 켰거나(업데이트 뒤 포함), 크롬이 확장의 뒷일꾼(서비스워커)을 중간에 꺼 버렸을 때
    //   예전엔 대기줄만 남고 가만히 있었다 — 사람이 [시작]을 다시 눌러야 했고, 그때 작업 창이 하나 더 생겼다.
    //   사람이 [중지]로 멈춘 판은 그대로 둔다. 로그인·회원 멈춤은 5분 규칙(ensureWorker)이 지킨다.
    try {
        if (!S.userStopped && (S.queue.length > 0 || S.pending.length > 0)) {
            S.note = '다시 켜져서 하던 작업을 이어갑니다...';
            setTimeout(() => { try { ensureWorker(); } catch (_) {} }, 4000);
        }
    } catch (e) {}
})();

// ------------------------------------------------------------
// 메시지 핸들러
// ------------------------------------------------------------
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
    if (request.action === 'v2:addKeywords') {
        (async () => {
            const incoming = (request.keywords || []).map(k => String(k).trim()).filter(Boolean);

            // 최근 1개월 안에 검색한 키워드는 다시 받지 않는다
            // (패널의 "이미 검색한 키워드" 목록에서 지우면 바로 다시 검색 가능)
            const st = await chrome.storage.local.get('kwHistory');
            const hist = st.kwHistory || {};
            const RESEARCH_MS = RESEARCH_MS_G;

            const known = new Set(S.queue.map(k => bareKw(k)));
            if (S.current) known.add(bareKw(S.current.keyword));

            let added = 0;
            const rejected = [];
            for (const k of incoming) {
                const bk = bareKw(k);
                if (known.has(bk)) { rejected.push({ keyword: k, reason: 'dup' }); continue; }
                const t = hist[bk];
                if (t && (Date.now() - t) < RESEARCH_MS) {
                    const until = new Date(t + RESEARCH_MS);
                    rejected.push({
                        keyword: k, reason: 'recent',
                        until: (until.getMonth() + 1) + '월 ' + until.getDate() + '일'
                    });
                    continue;
                }
                known.add(bk);
                S.queue.push(k);
                added++;
            }
            persist(); broadcast();
            if (added > 0) ensureWorker(true);
            sendResponse({ ok: true, added: added, rejected: rejected });
        })();
        return true;
    }
    if (request.action === 'v2:kwHistoryList') {
        // 최근 1개월 안에 검색한 키워드 목록 (패널의 펼침 목록용)
        (async () => {
            const st = await chrome.storage.local.get('kwHistory');
            const hist = st.kwHistory || {};
            const now = Date.now();
            const items = Object.keys(hist)
                .filter(k => (now - hist[k]) < RESEARCH_MS_G)
                .sort((a, b) => hist[b] - hist[a])
                .map(k => {
                    const d = new Date(hist[k]);
                    const h = d.getHours();
                    const ampm = h < 12 ? '오전' : '오후';
                    const h12 = h % 12 === 0 ? 12 : h % 12;
                    const mm = String(d.getMinutes()).padStart(2, '0');
                    return { keyword: k, when: (d.getMonth() + 1) + '/' + d.getDate() + ' ' + ampm + ' ' + h12 + ':' + mm };
                });
            sendResponse({ ok: true, items: items });
        })();
        return true;
    }
    if (request.action === 'v2:huntNext') {
        // "다음 사냥감"(연료 초과로 못 돈 키워드)을 대기줄에 태워 이어서 찾는다
        (async () => {
            if (!S.nextTargets.length) { sendResponse({ ok: false, error: '남은 사냥감이 없어요.' }); return; }
            const moved = await boardNextTargets();
            if (moved > 0) ensureWorker(true);
            sendResponse({ ok: true, moved: moved });
        })();
        return true;
    }
    if (request.action === 'v2:clearDone') {
        // 완료된 키워드 전체 삭제 (수집 기록 + 1개월 제한까지 함께 해제)
        (async () => {
            if (S.running) { sendResponse({ ok: false, error: '작업 중에는 삭제할 수 없어요.' }); return; }
            const st = await chrome.storage.local.get('kwHistory');
            const hist = st.kwHistory || {};
            for (const d of S.done) delete hist[bareKw(d.keyword)];
            await chrome.storage.local.set({ kwHistory: hist });
            const kws = new Set(S.done.map(d => d.keyword));
            S.results = S.results.filter(r => !kws.has(r.searchKeyword));
            S.pending = S.pending.filter(p => !kws.has(p.searchKeyword));
            S.done = [];
            await persist(); broadcast();
            sendResponse({ ok: true });
        })();
        return true;
    }
    if (request.action === 'v2:clearHistory') {
        // 검색 기록 전체 삭제 = 모든 키워드 재검색 제한 해제
        (async () => {
            await chrome.storage.local.set({ kwHistory: {} });
            sendResponse({ ok: true });
        })();
        return true;
    }
    if (request.action === 'v2:removeHistory') {
        // 목록에서 삭제 = 재검색 제한 즉시 해제
        (async () => {
            const st = await chrome.storage.local.get('kwHistory');
            const hist = st.kwHistory || {};
            delete hist[bareKw(request.keyword)];
            await chrome.storage.local.set({ kwHistory: hist });
            sendResponse({ ok: true });
        })();
        return true;
    }
    if (request.action === 'v2:removeDone') {
        (async () => {
            if (S.running) { sendResponse({ ok: false, error: '작업 중에는 삭제할 수 없어요.' }); return; }
            const k = request.keyword;
            S.done = S.done.filter(d => d.keyword !== k);
            S.results = S.results.filter(r => r.searchKeyword !== k);
            S.pending = S.pending.filter(p => p.searchKeyword !== k);
            // 3개월 제한도 함께 풀어서 다시 검색할 수 있게 한다
            const st = await chrome.storage.local.get('kwHistory');
            const hist = st.kwHistory || {};
            delete hist[bareKw(k)];
            await chrome.storage.local.set({ kwHistory: hist });
            await persist(); broadcast();
            sendResponse({ ok: true });
        })();
        return true;
    }
    if (request.action === 'v2:start') {
        S.haltedFor = '';      // 사람이 [시작]을 눌렀다 — 멈춤 상태를 풀고 바로 본다
        ensureWorker(true);
        sendResponse({ ok: true });
        return true;
    }
    if (request.action === 'v2:stop') {
        S.stopRequested = true;
        broadcast();
        // 모달에도 바로 "중지하는 중"을 보여준다 (눌렀는데 반응 없다는 느낌 방지)
        if (S.workerTabId) {
            updateModal(S.workerTabId, {
                phase: '중지하는 중...',
                sub: '읽고 있던 카탈로그만 마무리하고 엑셀을 저장해요'
            });
        }
        sendResponse({ ok: true });
        return true;
    }
    if (request.action === 'v2:clearNext') {
        // 다음 사냥감 전체 비우기 - 매 판 엑셀로 백업돼 있고,
        // 지운 키워드도 새 꼬리물기에서 재발굴되므로 잃는 것 없음
        (async () => {
            const n = S.nextTargets.length;
            S.nextTargets = [];
            await persist(); broadcast();
            sendResponse({ ok: true, cleared: n });
        })();
        return true;
    }
    if (request.action === 'v2:clearQueue') {
        // 대기줄 전체 비우기 - 자동추가 키워드는 버리지 않고 "다음 사냥감"으로 돌려보낸다
        (async () => {
            const parkedKeys = new Set(S.nextTargets.map(k => bareKw(k)));
            let parked = 0;
            for (const k of S.queue) {
                if (!/\(자동추가/.test(k)) continue;          // 직접 입력한 키워드는 그냥 삭제
                if (parkedKeys.has(bareKw(k))) continue;
                parkedKeys.add(bareKw(k));
                S.nextTargets.push(k);
                parked++;
            }
            S.queue = [];
            await persist(); broadcast();
            sendResponse({ ok: true, parked: parked });
        })();
        return true;
    }
    if (request.action === 'v2:removeKeyword') {
        S.queue = S.queue.filter(k => k !== request.keyword);
        persist(); broadcast();
        sendResponse({ ok: true });
        return true;
    }
    if (request.action === 'v2:getState') {
        S.lastStateAt = Date.now();            // 패널이 열려 있다는 표시 (업데이트 판단에 쓴다)
        maybeCheckUpdate();                    // 가끔 새 버전이 나왔는지 본다 (기다리지 않는다)
        refreshUpdateProgress();               // 받는 중이면 몇/몇 받았는지 다시 읽는다 (v5.2)
        maybeRunFileOrder();                   // 실행지시 파일이 있으면 따른다 (기다리지 않는다)
        sendResponse({ ok: true, state: publicState() });
        return true;
    }
    if (request.action === 'v2:applyUpdate') {
        // 패널의 [지금 적용] — 업데이트기가 폴더를 이미 갈아놓은 경우에만 의미가 있다
        applyUpdateNow().then(r => sendResponse(r))
                        .catch(e => sendResponse({ ok: false, error: String(e) }));
        return true;
    }
    if (request.action === 'v2:exportCsv') {
        exportAll().then(() => sendResponse({ ok: true }))
                   .catch(e => sendResponse({ ok: false, error: String(e) }));
        return true;
    }
    if (request.action === 'v2:clear') {
        if (S.running) { sendResponse({ ok: false, error: '작업 중에는 초기화할 수 없습니다.' }); return true; }
        S.queue = []; S.done = []; S.results = []; S.failures = []; S.pending = []; S.lastError = '';
        S.autoAdded = 0; S.deepAdded = 0; S.nextTargets = []; S.runStamp = ''; S.runRound = 1;
        S.runHits = 0; S.runActiveMs = 0; S.skipN = newSkipN(); S.catStat = {};
        try { eraseSavedFile('다음사냥감'); } catch (_) {}   // 다음사냥감 파일도 함께 정리 (v4.90)
        persist(); broadcast();
        sendResponse({ ok: true });
        return true;
    }
    if (request.action === 'checkLogin') {
        isNaverLoggedIn().then(ok => sendResponse({ loggedIn: ok }));
        return true;
    }
    if (request.action === 'openNaverLogin') {
        openNaverLoginOnce();
        sendResponse({ ok: true });
        return true;
    }
});

// ------------------------------------------------------------
// 공용 유틸
// ------------------------------------------------------------

// 서비스워커가 잠들지 않게 쪼개서 대기
async function sleep(ms) {
    while (ms > 0) {
        const chunk = Math.min(ms, 5000);
        await new Promise(r => setTimeout(r, chunk));
        ms -= chunk;
        try { await chrome.runtime.getPlatformInfo(); } catch (e) {}
    }
}

// 중지 버튼에 즉시 반응하는 대기.
// 카탈로그 사이 간격(5~11초)이나 긴 휴식(1~2분) 중에 중지를 눌러도
// 1초 안에 알아차리고 빠져나온다.
async function restSleep(ms) {
    while (ms > 0 && !S.stopRequested) {
        const chunk = Math.min(ms, 1000);
        await new Promise(r => setTimeout(r, chunk));
        ms -= chunk;
        try { await chrome.runtime.getPlatformInfo(); } catch (e) {}
    }
}

const rand = (min, max) => min + Math.random() * (max - min);

// 순이익 값 (계산 안 된 행은 맨 뒤로 가게 -무한대)
const profitOf = (r) => (r.profit === "" || r.profit === undefined) ? -Infinity : r.profit;

// 1개 기준·메인수량 기준·낱개신규 중 가장 좋은 순이익 (완료 집계용 - 상품 단위)
const bestProfitOf = (r) => {
    const a = (r.profit === "" || r.profit === undefined) ? -Infinity : r.profit;
    const b = (r.mainProfit === "" || r.mainProfit === undefined) ? -Infinity : r.mainProfit;
    const c = (r.soloProfit === "" || r.soloProfit === undefined) ? -Infinity : r.soloProfit;
    const d = (r.gapProfit === "" || r.gapProfit === undefined) ? -Infinity : r.gapProfit;
    return Math.max(a, b, c, d);
};

// 확장된 판매행의 순이익 (엑셀 정렬용 - 행 단위)
const sellProfitOf = (r) => (r.sellProfit === "" || r.sellProfit === undefined) ? -Infinity : r.sellProfit;

// 판매 단위별로 행을 만든다:
//   모든 상품 = 1개 기준 1행 (등록가능/조금 남음/PASS)
//   메인수량 기준이 목표 순이익을 넘으면 = "등록가능(메인N개)" 행을 하나 더 추가
function expandSellRows(data, targetProfit) {
    const out = [];
    for (const r of data) {
        // 1개 기준 행: 제목/주소도 "1개 옵션" 기준으로 맞춘다.
        //  - 1개 옵션을 실제로 눌러서 얻은 제목/주소가 있으면 그것을 쓰고
        //  - 없으면 카드 제목 끝의 ", N개"를 ", 1개"로 바꿔서 쓴다
        //  - 낱개 신규(1개 옵션이 없던 카탈로그)는 제목에 ", 1개" 를 새로 붙이고
        //    우리가 정한 판매가·순이익을 쓴다 (v4.92)
        const isSolo = !!r.soloNew;
        const oneTitle = isSolo ? soloTitleOf(r.title)
            : ((r.linkQty === 1 && r.optionTitle) ? r.optionTitle
                : String(r.title || '').replace(/,\s*\d{1,3}\s*개\s*$/, ', 1개'));
        const oneUrl = qtyUrlOf(r, 1);

        out.push(Object.assign({}, r, {
            sellQty: 1,
            sellPrice: isSolo ? r.soloPrice : r.basePrice,
            sellProfit: isSolo ? r.soloProfit : r.profit,
            sellCostRate: isSolo
                ? (r.unitCost > 0 ? Math.round((r.soloProfit / r.unitCost) * 1000) / 10 : '')
                : r.costRate,
            sellVerdict: r.verdict,
            sellTitle: oneTitle,
            sellCatalogUrl: oneUrl
        }));

        // [수량 칸 전부] 2개부터 12개까지, 남는 칸은 전부 한 줄씩. (v4.97)
        //  남이 있는 칸(경쟁) = '등록가능', 비어 있는 칸 = '빈수량신규'.
        //  예전의 "메인수량 행"과 "빈 수량 한 칸"은 이 안에 들어 있다.
        //  칸이 이미 최소 순이익(1,000원)을 넘긴 것만 담겨 있어 여기서 다시 거르지 않는다.
        //  주소는 칸마다 다르다 — 남이 있는 칸(경쟁)은 그 수량을 눌렀을 때의 주소, 빈 칸은 대표 주소. (v5.0)
        for (const s of (r.slots || [])) {
            if (!(s.qty > 1)) continue;
            const isMain = (s.basis === '경쟁' && Number(r.mainQty) === s.qty);
            out.push(Object.assign({}, r, {
                sellQty: s.qty,
                sellPrice: s.price,
                sellProfit: s.profit,
                sellCostRate: r.unitCost > 0 ? Math.round((s.profit / (r.unitCost * s.qty)) * 1000) / 10 : '',
                sellVerdict: s.basis === '경쟁' ? '등록가능' : '빈수량신규',
                // 검색 카드가 곧 그 수량의 제목이면 그대로, 아니면 끝의 수량만 바꾼다
                sellTitle: isMain ? (r.title || '') : qtyTitleOf(r.title, s.qty),
                sellCatalogUrl: s.basis === '경쟁' ? qtyUrlOf(r, s.qty) : (r.catalogUrl || '')
            }));
        }

        // [낱개 사입(순방향)] 네이버 1개 최저가를 실제 매입가로 사서 파는 칸. (v4.98)
        //  매입 칸(수량·총액·개당원가·주소·판매처)을 그 길로 바꿔 적는다 — 엑셀에서 "어디서 얼마에 사는지" 가 바로 보이게.
        for (const s of (r.fwdSlots || [])) {
            out.push(Object.assign({}, r, {
                sellQty: s.qty,
                sellPrice: s.price,
                sellProfit: s.profit,
                sellCostRate: r.fwdUnitCost > 0 ? Math.round((s.profit / (r.fwdUnitCost * s.qty)) * 1000) / 10 : '',
                sellVerdict: '낱개사입·' + (s.basis === '경쟁' ? '등록가능' : '빈수량신규'),
                sellTitle: s.qty === 1
                    ? ((r.linkQty === 1 && r.optionTitle) ? r.optionTitle : qtyTitleOf(r.title, 1))
                    : qtyTitleOf(r.title, s.qty),
                sellCatalogUrl: s.basis === '경쟁' ? qtyUrlOf(r, s.qty) : (r.catalogUrl || ''),
                bestQty: 1, bestTotal: r.fwdUnitCost, unitCost: r.fwdUnitCost,
                buyUrl: r.fwdLink || r.soloBuyLink || '',
                buySeller: (r.fwdSeller || '') + ' (네이버 낱개)',
                buyStatus: r.fwdNote || ''
            }));
        }
    }
    return out;
}

// 그 수량 칸의 주소. (v5.0)
//   ① 칸마다 눌러 받아 적은 것(qtyUrls) → ② 가장 적은 수량(optionUrl) → ③ 매입 묶음(buyUrl) → ④ 대표(카탈로그) 주소.
//   옛 결과(qtyUrls 가 없는 것)도 ②③④ 로 예전과 같이 나온다.
function qtyUrlOf(r, n) {
    const q = Number(n);
    const m = (r && r.qtyUrls) || {};
    if (q > 0 && m[q]) return m[q];
    if (q > 0 && q === Number(r.linkQty) && r.optionUrl) return r.optionUrl;
    if (q > 0 && q === Number(r.bestQty) && r.buyUrl) return r.buyUrl;
    return (r && r.catalogUrl) || '';
}

// 파일명에 못 쓰는 글자 제거
function csvSafeName(s) {
    return String(s).replace(/[\\\/:*?"<>|]/g, '_').slice(0, 40);
}

// [꼬리물기] 키워드 뒤의 "(자동추가)" / "(자동추가·심화)" 꼬리표를 뗀 원형.
// 검색창 입력, 1개월 기록, 중복 확인은 전부 이 원형 기준으로 한다.
function bareKw(k) {
    return String(k).replace(/\(자동추가(·심화)?\)\s*$/, '').trim();
}

// 브랜드/제조사에서 온 "심화" 키워드인지 (카테고리 나무를 끝까지 파고드는 대상)
function isDeepKw(k) {
    return /\(자동추가·심화\)\s*$/.test(String(k));
}

// 이 키워드를 언제 검색했는지 기록 (3개월 재검색 제한용)
async function recordKeywordSearched(keyword) {
    try {
        const st = await chrome.storage.local.get('kwHistory');
        const hist = st.kwHistory || {};
        hist[bareKw(keyword)] = Date.now();
        await chrome.storage.local.set({ kwHistory: hist });
    } catch (e) {}
}

async function loadMoney() {
    const st = await chrome.storage.local.get('money');
    const m = st.money || {};
    const al = Number(m.autoLimit);   // 0이면 꼬리물기 끔. 미설정이면 10.
    return {
        shippingFee: Number(m.shippingFee) || 3000,
        buyShipFee: Number(m.buyShipFee) || 0,   // 묶음 1건 매입 배송비 (기본 0 = 예전 그대로)
        feeRate: (Number(m.feeRate) || 6.7) / 100,
        targetProfit: Number(m.targetProfit) || 1000,
        autoLimit: S.noTail ? 0 : ((isNaN(al) || al < 0) ? 10 : al),   // 실행지시 파일이 꼬리물기를 끈 판이면 0
        autoHunt: !!m.autoHunt,           // 완전 자동: 판이 끝나면 다음 사냥감을 알아서 이어감
        soloNewToCsv: !!m.soloNewToCsv,   // 낱개 신규도 등록대기(자동등록)에 담을지 — 기본 꺼짐
        forwardBuy: m.forwardBuy === undefined ? CONFIG.FORWARD_BUY : !!m.forwardBuy,   // 낱개 사입(순방향) 계산할지
        slotRate: Number(m.slotRate) || CONFIG.SLOT_PROFIT_RATE,      // 칸별 목표순이익 비율(%)
        reviewMin: Number(m.reviewMin) || CONFIG.REVIEW_MIN_COUNT     // 2주 리뷰 기준 (구간에서 고른 값)
    };
}

// ------------------------------------------------------------
// 회원 확인 (구글 시트 연동)
//  - 시작할 때 반드시 성공해야 하고,
//  - 키워드마다 다시 확인해서 사용기간이 끝나면 자동으로 멈춘다.
//  - 일시적인 네트워크 오류로는 진행 중인 작업을 죽이지 않는다.
// ------------------------------------------------------------
const VERIFY_CACHE_MS = 10 * 60 * 1000;  // 10분 안에 확인했으면 다시 묻지 않는다

async function verifyMember() {
    if (!isScriptUrlReady()) return { ok: false, code: 'NO_CONFIG' };
    try {
        const st = await chrome.storage.local.get('member');
        const m = st.member;
        if (!m || !m.phone || !m.token) return { ok: false, code: 'NO_LOGIN' };

        // 방금(10분 내) 패널 로그인이나 이전 확인이 성공했으면 서버에 다시 묻지 않는다.
        // 구글 스크립트가 잠들어 있으면 응답에 5~15초가 걸려서, 매번 물으면 시작이 느려진다.
        // (장시간 작업 중에는 10분마다 재확인되므로 사용기간 만료 자동정지도 유지된다)
        if (m.verifiedAt && (Date.now() - m.verifiedAt) < VERIFY_CACHE_MS) {
            return { ok: true, cached: true, name: m.name };
        }

        const r = await apiCall('verify', {
            id: m.phone, token: m.token, ip: await getPublicIp(), ver: appVer()
        });
        if (r && r.ok) {
            m.name = r.name || m.name;
            m.nick = r.nick || m.nick;
            m.verifiedAt = Date.now();
            await chrome.storage.local.set({ member: m });
        }
        return r || { ok: false, code: 'NO_RESPONSE' };
    } catch (e) {
        return { ok: false, code: 'NETWORK', detail: String(e) };
    }
}

// ── [학습기록] 소싱 결과를 강사님 구글시트로 전송 (키워드·제품·마진) ──
//  같은 결과가 매판 누적되므로, '아직 안 보낸 것'만 골라 보낸다(중복 방지).
//  실패해도 소싱은 계속 — 조용히 넘어간다.
//  ※ 회원확인 서버가 sourcingLog 를 아직 모르면(UNKNOWN) 이번 세션엔 다시 보내지 않는다.
//    같은 폴더의 소싱기록_앱스스크립트_추가분.gs 를 붙이고 새 버전으로 배포하면 켜진다.
let _sourcingLogOff = false;

async function reportSourcingToSheet(rows) {
    try {
        if (_sourcingLogOff) return;
        if (!isScriptUrlReady() || !Array.isArray(rows) || !rows.length) return;
        const stg = await chrome.storage.local.get(['member', 'sentSourcingKeys']);
        const m = stg.member;
        if (!m || !m.phone) return;
        const sent = new Set(stg.sentSourcingKeys || []);
        const last4 = String(m.phone).replace(/[^0-9]/g, '').slice(-4);
        const pick = [];
        for (const r of rows) {
            const url = (r.sellCatalogUrl !== undefined ? r.sellCatalogUrl : r.catalogUrl) || '';
            const verdict = (r.sellVerdict !== undefined ? r.sellVerdict : r.verdict) || '';
            const key = (r.searchKeyword || '') + '|' + url + '|' + verdict + '|' + (r.buyStatus || '');
            if (!url || sent.has(key)) continue;
            pick.push({ r: r, key: key });
            sent.add(key);                   // 같은 판 안의 중복도 즉시 제외
            if (pick.length >= 400) break;   // 한 번에 최대 400줄 (나머지는 다음 판에)
        }
        if (!pick.length) return;
        const compact = pick.map(({ r }) => ({
            k: r.searchKeyword || '',
            t: (r.sellTitle !== undefined ? r.sellTitle : r.title) || '',
            p: (r.sellPrice !== undefined ? r.sellPrice : r.basePrice) || '',
            c: r.unitCost || '',
            pf: (r.sellProfit !== undefined ? r.sellProfit : r.profit) || '',
            mr: r.marginRate || '',
            s: r.buySeller || '',
            v: (r.sellVerdict !== undefined ? r.sellVerdict : r.verdict) || '',
            bs: r.buyStatus || '',
            u: (r.sellCatalogUrl !== undefined ? r.sellCatalogUrl : r.catalogUrl) || '',
            at: r.scrapedAt || ''
        }));
        const res = await apiCall('sourcingLog', {
            id: m.phone, name: m.name || '', last4: last4, rows: JSON.stringify(compact)
        });
        if (res && res.code === 'UNKNOWN') { _sourcingLogOff = true; return; }
        if (res && res.ok) {
            for (const { key } of pick) sent.add(key);
            const arr = [...sent].slice(-5000);   // 최근 5000개만 유지(용량 관리)
            await chrome.storage.local.set({ sentSourcingKeys: arr });
        }
    } catch (e) {
        console.debug('[소싱기록 전송] 건너뜀:', e && e.message);
    }
}

function memberFailText(r) {
    if (r && r.message) return String(r.message);
    const map = {
        'NO_CONFIG': '회원 확인 서버 주소(config.js)가 비어 있습니다.',
        'NO_LOGIN': '로그인이 필요합니다. 사이드 패널에서 로그인해주세요.',
        'NO_USER': '회원 명단에 없는 번호입니다. (루크 수강생: 상담 예약 번호 / 뷰셀·미니쌤 수강생: 뷰셀 앱 번호) 강사에게 문의해 주세요.',
        'NO_PW': '아직 비밀번호가 없습니다. 루크 수강생은 상담 예약 페이지에서, 뷰셀·미니쌤 수강생은 뷰셀 앱에서 먼저 정해주세요.',
        'BAD_PW': '비밀번호가 맞지 않습니다.',
        'BAD_TOKEN': '다시 로그인해 주세요.',
        'BLOCKED': '사용이 멈춰 있습니다. 강사에게 문의해 주세요.',
        'IP_FULL': '다른 컴퓨터에서 로그인해서 이 컴퓨터가 빠졌습니다. 패널에서 다시 로그인하면 이 컴퓨터가 들어오고, 제일 오래 안 쓴 컴퓨터 하나가 빠집니다.',
        'NETWORK': '서버 연결에 실패했습니다. 인터넷 연결을 확인해주세요.'
    };
    return map[(r && r.code) || ''] || '회원 확인에 실패했습니다.';
}

// ------------------------------------------------------------
// 네이버 로그인 확인
//   - 로그인은 사용자가 직접 합니다. 이 프로그램은 아이디/비밀번호를
//     입력하지 않고, "로그인 되어 있는지"만 쿠키로 확인합니다.
// ------------------------------------------------------------
async function isNaverLoggedIn() {
    try {
        const ses = await chrome.cookies.get({ url: 'https://www.naver.com', name: 'NID_SES' });
        if (ses && ses.value) return true;
        const aut = await chrome.cookies.get({ url: 'https://www.naver.com', name: 'NID_AUT' });
        return !!(aut && aut.value);
    } catch (e) {
        console.error('[네이버 로그인 확인] 실패:', e);
        return false;
    }
}

// ------------------------------------------------------------
// 페이지 가운데 모달 (뒤는 어둡게 - 만지면 안 될 것 같은 분위기)
// ------------------------------------------------------------
function injectModal() {
    if (document.getElementById('ext-v2-backdrop')) return;

    const style = document.createElement('style');
    style.id = 'ext-v2-style';
    style.textContent = [
        '@keyframes extHopX { 0% { left: -14%; } 100% { left: 104%; } }',
        '@keyframes extHopY { 0%, 100% { transform: translateY(0) scaleX(-1); } 50% { transform: translateY(-9px) scaleX(-1); } }',
        '@keyframes extPulse { 0%, 100% { opacity: 1; } 50% { opacity: 0.55; } }',
        // 도장 찍히는 느낌: 크게 나타나서 쿵 하고 자리잡음
        '@keyframes extStamp { 0% { transform: scale(2.6); opacity: 0; } 55% { transform: scale(0.92); opacity: 1; } 100% { transform: scale(1); opacity: 1; } }',
        '#ext-v2-rabbit { position: absolute; bottom: 3px; font-size: 22px;',
        '  animation: extHopX 5.5s linear infinite, extHopY 0.5s ease-in-out infinite; }',
        '#ext-v2-phase { animation: extPulse 2s ease-in-out infinite; }'
    ].join('\n');
    document.head.appendChild(style);

    const backdrop = document.createElement('div');
    backdrop.id = 'ext-v2-backdrop';
    backdrop.style.cssText = [
        'position: fixed; inset: 0; z-index: 2147483000;',
        'background: rgba(44, 48, 66, 0.45); backdrop-filter: blur(3px);',
        'display: flex; align-items: center; justify-content: center;'
    ].join(' ');


    const card = document.createElement('div');
    card.id = 'ext-v2-card';
    card.style.cssText = [
        'width: 340px; background: rgba(255,255,255,0.97); border-radius: 24px;',
        'padding: 24px 26px 20px; text-align: center;',
        'box-shadow: 0 18px 50px rgba(31,38,80,0.3);',
        'font-family: -apple-system, "Apple SD Gothic Neo", "Malgun Gothic", sans-serif;',
        'color: #3a3f52;'
    ].join(' ');

    const track = document.createElement('div');
    track.style.cssText = [
        'position: relative; height: 36px; overflow: hidden; margin-bottom: 14px;',
        'background: linear-gradient(90deg, #E8F4EC, #EAF1FB, #F6ECF3); border-radius: 999px;'
    ].join(' ');
    const rabbit = document.createElement('div');
    rabbit.id = 'ext-v2-rabbit';
    rabbit.innerText = '🐇';
    track.appendChild(rabbit);

    const title = document.createElement('div');
    title.innerText = '카탈로그를 부지런히 모으는 중';
    title.style.cssText = 'font-size: 16px; font-weight: 800; margin-bottom: 12px;';

    const kw = document.createElement('div');
    kw.id = 'ext-v2-kw';
    kw.innerText = '-';
    kw.style.cssText = [
        'display: inline-block; padding: 5px 14px; margin-bottom: 12px;',
        'background: #EAF1FB; color: #4A6FB5; border-radius: 999px;',
        'font-size: 14px; font-weight: 700;'
    ].join(' ');

    const big = document.createElement('div');
    big.id = 'ext-v2-big';
    big.innerText = '';
    big.style.cssText = 'font-size: 30px; font-weight: 800; letter-spacing: 1px; margin-bottom: 4px; color: #2f3446;';

    const phase = document.createElement('div');
    phase.id = 'ext-v2-phase';
    phase.innerText = '준비 중...';
    phase.style.cssText = 'font-size: 13px; color: #8a90a5; margin-bottom: 6px;';

    const sub = document.createElement('div');
    sub.id = 'ext-v2-sub';
    sub.innerText = '';
    sub.style.cssText = 'font-size: 12px; color: #a3a9bc; margin-bottom: 14px; white-space: pre-line;';

    const stopBtn = document.createElement('button');
    stopBtn.id = 'ext-v2-stop';
    stopBtn.innerText = '중지하기';
    stopBtn.style.cssText = [
        'width: 100%; padding: 11px 0; border: none; border-radius: 12px; cursor: pointer;',
        'background: #F9D6D6; color: #B5484D; font-weight: 800; font-size: 14px;'
    ].join(' ');
    stopBtn.onclick = function () {
        stopBtn.innerText = '중지하는 중...';
        stopBtn.disabled = true;
        chrome.runtime.sendMessage({ action: 'v2:stop' });
    };

    const note = document.createElement('div');
    note.innerText = '자동으로 조사하고 있어요 · 이 창은 건드리지 않아도 돼요';
    note.style.cssText = 'font-size: 11px; color: #c0c5d4; margin-top: 10px;';

    card.appendChild(track);
    card.appendChild(title);
    card.appendChild(kw);
    card.appendChild(big);
    card.appendChild(phase);
    card.appendChild(sub);
    card.appendChild(stopBtn);
    card.appendChild(note);
    backdrop.appendChild(card);
    document.body.appendChild(backdrop);
}

function updateModalFields(f) {
    const set = (id, v) => { const el = document.getElementById(id); if (el && v !== undefined) el.innerText = v; };
    set('ext-v2-kw', f.keyword);
    set('ext-v2-big', f.big);
    set('ext-v2-phase', f.phase);
    set('ext-v2-sub', f.sub);

    // 전부 완료: 토끼가 뿅! 하고 SUCCESS!
    if (f.success && !document.getElementById('ext-v2-success')) {
        const ov = document.createElement('div');
        ov.id = 'ext-v2-success';
        ov.style.cssText = 'position: fixed; inset: 0; z-index: 2147483001; pointer-events: none;'
            + ' display: flex; flex-direction: column; align-items: center; padding-top: 9vh;';

        const rb = document.createElement('div');
        rb.innerText = '🐇';
        rb.style.cssText = 'font-size: 76px; filter: drop-shadow(0 6px 14px rgba(0,0,0,0.4));'
            + ' animation: extStamp 0.55s ease-out both;';
        ov.appendChild(rb);

        const txt = document.createElement('div');
        txt.innerText = 'SUCCESS!';
        txt.style.cssText = 'font-size: 50px; font-weight: 900; color: #fff; letter-spacing: 4px;'
            + ' margin-top: 14px; text-shadow: 0 6px 22px rgba(0,0,0,0.45);'
            + ' font-family: -apple-system, "Apple SD Gothic Neo", sans-serif;'
            + ' animation: extStamp 0.5s ease-out 0.4s both;';
        ov.appendChild(txt);
        document.body.appendChild(ov);
    }

    if (f.final) {
        const rb = document.getElementById('ext-v2-rabbit');
        if (rb) { rb.style.animationPlayState = 'paused'; rb.innerText = '🐇💤'; }
        const btn = document.getElementById('ext-v2-stop');
        if (btn) {
            btn.innerText = '닫기';
            btn.disabled = false;
            btn.style.background = '#D9EFE1';
            btn.style.color = '#2F7D4F';
            btn.onclick = function () {
                const b = document.getElementById('ext-v2-backdrop');
                if (b) b.remove();
                const s = document.getElementById('ext-v2-style');
                if (s) s.remove();
                const sc = document.getElementById('ext-v2-success');
                if (sc) sc.remove();
            };
        }
    }
}

async function updateModal(tabId, fields) {
    try {
        await chrome.scripting.executeScript({ target: { tabId: tabId }, func: injectModal });
        await chrome.scripting.executeScript({
            target: { tabId: tabId }, func: updateModalFields, args: [fields]
        });
    } catch (e) {
        console.debug('[모달] 갱신 실패:', e.message);
    }
}

function checkReadyState() {
    return document.readyState;
}

// [배경에서 실행] 매입처의 실제 상품명이 카탈로그 제품과 "완전히 다른 제품"인지 판단한다.
//  네이버 가격비교에서 판매자가 엉뚱한 상품을 이 카탈로그에 잘못 매칭해 최저가로 뜨는 경우가 있다.
//  → 의미있는 핵심 키워드가 하나도 안 겹치면 "다른 제품"으로 본다(보수적: 확실한 것만 걸러냄).
function _titleMismatch(catTitle, buyTitle) {
    if (!buyTitle) return false;   // 매입처 제목을 못 읽었으면 판단 보류(제외하지 않음)
    const norm = s => String(s || '').toLowerCase().replace(/[^가-힣a-z0-9]/g, ' ');
    // 제품을 식별하지 못하는 일반/단위 단어는 비교에서 뺀다
    const GENERIC = /^(효소|건강|미식|함유|유기농|무농약|프리미엄|골드|골든|플러스|리얼|파워|스타|정품|더|the|for|and|with|plus|정|포|개|개입|박스|세트|팩|봉|스틱|캡슐|정제|병|캔|구|통|낱개|대용량|증정|사은품|무료배송|배송|택배|당|원|mg|g|ml|kg|l|x|엑스)$/;
    const toks = s => norm(s).split(/\s+/).filter(t => t.length >= 2 && !/^\d+$/.test(t) && !GENERIC.test(t));
    const cat = toks(catTitle);
    if (!cat.length) return false;                 // 카탈로그에 비교할 키워드가 없으면 보류
    const buy = norm(buyTitle);
    const shared = cat.filter(t => buy.indexOf(t) >= 0);
    return shared.length === 0;                    // 의미있는 키워드가 하나도 안 겹치면 = 다른 제품
}

// [매입처 페이지에서 실행] 판매중단/품절 문구가 있는지 확인한다.
// 문구가 있으면 "의심"일 뿐 확정이 아니므로, 표시만 하고 제외하지 않는다.
function checkSellableOnPage() {
    // 매입처 페이지의 실제 상품명도 담아온다.
    // (카탈로그가 "12개"라는데 매입처 상품명이 "1박스"면 매칭오류를 사람 눈으로 바로 확정 가능)
    let pageTitle = '';
    try { pageTitle = (document.title || '').replace(/\s+/g, ' ').trim().slice(0, 80); } catch (e) {}

    try {
        let t = ((document.body && document.body.innerText) || '').slice(0, 8000);
        // 오탐 방지: 마케팅/안내 문구를 먼저 제거
        for (const noise of ['품절임박', '품절 임박', '품절 시', '품절시', '품절될', '품절이 예상',
                             '재입고', '품절되면', '품절 상품 제외']) {
            t = t.split(noise).join('');
        }

        // [강한 신호] 어디에 있든 판매중단으로 본다 (다른 맥락에서 거의 안 나오는 문구)
        const strong = ['구매할 수 없', '판매하지 않', '존재하지 않는 상품', '삭제된 상품',
                        '판매가 종료', '판매중단', '판매 중단', '판매중지', '판매 중지',
                        '진행중인 방송이 아닙', '판매 기간이 아닙'];
        for (const w of strong) {
            if (t.indexOf(w) !== -1) return { sellable: false, hit: w, title: pageTitle };
        }

        // [약한 신호] 리뷰/공지 등 다른 맥락에서도 나올 수 있는 단어는
        // 페이지 위쪽(상품 상태가 표시되는 영역)에서 발견될 때만 인정한다
        const top = t.slice(0, 2500);
        const weak = ['품절', '판매종료', '판매 종료', '일시품절'];
        for (const w of weak) {
            if (top.indexOf(w) !== -1) return { sellable: false, hit: w, title: pageTitle };
        }

        return { sellable: true, title: pageTitle };
    } catch (e) {
        return { sellable: true, title: pageTitle };
    }
}

// marker: 주소에 반드시 포함되어야 하는 문자열(카탈로그 ID).
// 이걸 확인하지 않으면 "직전 페이지가 아직 complete" 인 상태를
// 새 페이지로 착각해 같은 카탈로그를 두 번 읽게 된다.
// ------------------------------------------------------------
// 🚫 네이버 차단 감지 (v4.98, 루크 지시 2026-11-02)
//   IP 가 잠시 막히면 네이버는 보안문자(자동입력 방지) 화면이나 "접근이 제한" 안내를 띄운다.
//   그걸 모르고 계속 돌면 빈 결과만 쌓이고 차단이 길어진다. 그래서:
//     ① 페이지를 열 때마다 차단 화면인지 본다
//     ② 차단이면 30분 쉰다 (연달아 또 막히면 60분 → 120분). 쉬는 동안 1분마다 다시 열어 보고 풀렸으면 바로 이어간다
//     ③ 쉬는 동안 패널·작업 창·윈도우 알림으로 "열린 창에서 보안문자를 풀어 주세요" 라고 알린다 (풀면 바로 이어간다)
//     ④ 막혔던 그 자리(그 카탈로그)부터 다시 한다. 같은 자리에서 3번 막히면 그 카탈로그는 '차단으로 못 봄' 으로 적고 넘어간다
// ------------------------------------------------------------
function detectBlockOnPage() {
    try {
        const title = String(document.title || '');
        const body = String(document.body && document.body.innerText || '').slice(0, 4000);
        const url = String(location.href || '');
        const text = title + '\n' + body;
        const hard = /자동\s*입력\s*방지|보안\s*문자|보안\s*확인|비정상적인\s*(접근|요청|이용)|일시적으로\s*(접근|이용|서비스)[이가]?\s*(제한|중단)|접속이\s*제한|접근이\s*(제한|차단)|너무\s*많은\s*요청|요청이\s*너무\s*많|차단되었습니다|captcha|로봇이\s*아닙니다|잠시\s*후\s*다시\s*(이용|시도)/i;
        if (hard.test(text) || /captcha|\/block|nidlogin\.login\?.*captcha/i.test(url)) {
            return { blocked: true, why: (text.match(hard) || [url.slice(0, 60)])[0], title: title.slice(0, 60) };
        }
        // 카탈로그·검색 주소인데 알맹이가 거의 없다 = 빈 껍데기(차단 응답)일 때가 많다
        if (/search\.shopping\.naver\.com\/(catalog|search)/.test(url) && body.replace(/\s+/g, '').length < 120 && !/catalog_summary_info|product_item/.test(document.documentElement.innerHTML.slice(0, 200000))) {
            return { blocked: true, why: '빈 페이지(알맹이 없음)', title: title.slice(0, 60) };
        }
        return { blocked: false };
    } catch (e) { return { blocked: false, err: String(e && e.message) }; }
}

async function isBlockedTab(tabId) {
    try {
        const r = await chrome.scripting.executeScript({ target: { tabId: tabId }, func: detectBlockOnPage });
        return (r && r[0] && r[0].result) || { blocked: false };
    } catch (e) { return { blocked: false }; }
}

const BLOCK_WAITS_MIN = [30, 60, 120];   // 연달아 막힐 때 쉬는 시간(분)
let _blockStreak = 0;                   // 연달아 막힌 횟수 (한 번 정상으로 지나가면 0)
function blockRested() { _blockStreak = 0; }

// 차단을 만났을 때: 알리고, 쉬고, 풀렸는지 1분마다 본다. 돌아오는 값 = 풀렸으면 true, 중지됐으면 false.
async function blockPause(tabId, where, info) {
    const waitMin = BLOCK_WAITS_MIN[Math.min(_blockStreak, BLOCK_WAITS_MIN.length - 1)];
    _blockStreak++;
    const why = (info && info.why) || '';
    const msg = '🚫 네이버가 잠시 막았습니다 (' + where + (why ? ' · ' + why : '') + '). ' + waitMin + '분 쉬었다가 같은 자리부터 이어갑니다. '
              + '열린 창에 보안문자가 보이면 풀어 주세요 — 풀면 바로 이어갑니다.';
    S.note = msg;
    S.blocked = { at: Date.now(), until: Date.now() + waitMin * 60000, where: where, why: why, streak: _blockStreak };
    broadcast();
    console.warn('[차단]', msg);
    try { await fileOrderNote('[차단] ' + where + ' · ' + why + ' · ' + waitMin + '분 쉼'); } catch (_) {}
    try { await chrome.tabs.update(tabId, { active: true }); } catch (_) {}
    try {
        chrome.notifications.create('selsuda-block-' + Date.now(), {
            type: 'basic', iconUrl: 'icon128.png', title: '셀수다 자동 소싱 — 네이버가 잠시 막았습니다',
            message: waitMin + '분 쉬었다가 이어갑니다. 크롬의 소싱 작업 창에 보안문자가 보이면 풀어 주세요. 풀면 바로 이어갑니다.',
            priority: 2
        });
    } catch (_) {}
    try {
        await updateModal(tabId, { keyword: (S.current && S.current.keyword) || '', big: '🚫',
            phase: '네이버가 잠시 막았어요 — ' + waitMin + '분 쉬는 중', sub: '이 창에 보안문자가 보이면 풀어 주세요. 풀면 바로 이어가요.' });
    } catch (_) {}

    const until = Date.now() + waitMin * 60000;
    let tick = 0;
    while (Date.now() < until) {
        if (S.stopRequested) { S.blocked = null; return false; }
        await sleep(60000);
        tick++;
        // 1분마다 다시 열어 본다 — 사용자가 보안문자를 풀었거나 차단이 일찍 풀렸으면 바로 이어간다
        try {
            await chrome.tabs.reload(tabId);
            await sleep(4000);
            const again = await isBlockedTab(tabId);
            if (!again.blocked) { S.blocked = null; S.note = ''; broadcast(); try { await fileOrderNote('[차단 해제] ' + where + ' · ' + tick + '분 만에 풀림'); } catch (_) {} return true; }
        } catch (_) {}
        const left = Math.max(1, Math.round((until - Date.now()) / 60000));
        try { await updateModal(tabId, { keyword: (S.current && S.current.keyword) || '', big: '🚫', phase: '네이버가 잠시 막았어요 — ' + left + '분 더 쉬는 중', sub: '이 창에 보안문자가 보이면 풀어 주세요. 풀면 바로 이어가요.' }); } catch (_) {}
    }
    S.blocked = null; S.note = ''; broadcast();
    return true;   // 시간이 다 됐다 — 다시 해 본다 (또 막히면 더 길게 쉰다)
}

async function waitForLoad(tabId, marker, maxSeconds = 25) {
    for (let i = 0; i < maxSeconds; i++) {
        if (S.stopRequested) return false;   // 중지 눌렀으면 로딩을 기다리지 않는다
        await sleep(1000);
        try {
            if (marker) {
                const tab = await chrome.tabs.get(tabId);
                if (!tab || tab.status !== 'complete') continue;
                if (!tab.url || tab.url.indexOf(marker) === -1) continue;
            }
            const results = await chrome.scripting.executeScript({
                target: { tabId: tabId }, func: checkReadyState
            });
            if (results && results[0] && results[0].result === 'complete') return true;
        } catch (e) { /* 페이지 전환 중이면 무시 */ }
    }
    return false;
}

// ------------------------------------------------------------
// 사람처럼 읽기: 속도가 들쭉날쭉한 스크롤 + 중간에 멈춤 + 되돌아보기
// ------------------------------------------------------------
function humanScroll(minMs, maxMs) {
    return new Promise(resolve => {
        const budget = minMs + Math.random() * (maxMs - minMs);
        const started = Date.now();
        let goingUp = false;

        function step() {
            if (Date.now() - started > budget) {
                window.scrollTo({ top: 0, behavior: 'smooth' });
                setTimeout(resolve, 400 + Math.random() * 500);
                return;
            }
            const atBottom = window.scrollY + window.innerHeight >= document.body.scrollHeight - 80;
            if (atBottom) goingUp = true;
            if (window.scrollY <= 0) goingUp = false;

            const amount = (120 + Math.random() * 480) * (goingUp ? -1 : 1);
            window.scrollBy({ top: amount, behavior: 'smooth' });

            // 가끔 오래 멈춰서 읽는 척
            const pause = Math.random() < 0.18
                ? 900 + Math.random() * 1600
                : 260 + Math.random() * 520;
            setTimeout(step, pause);
        }
        setTimeout(step, 300 + Math.random() * 700);
    });
}

// 검색 결과는 끝까지 내려가야 40개가 다 로드된다.
// (지연 로딩을 확실히 끝내야 하므로 목록 페이지에서는 이 방식을 쓴다)
// 목록 훑기는 사람도 빠르게 내리므로 카탈로그 읽기보다 속도를 높였다.
function scrollPageToLoadAll() {
    return new Promise(resolve => {
        let totalHeight = 0;
        const timer = setInterval(() => {
            const stepPx = 650 + Math.floor(Math.random() * 350);
            const scrollHeight = document.body.scrollHeight;
            window.scrollBy(0, stepPx);
            totalHeight += stepPx;

            if (totalHeight >= scrollHeight || totalHeight > 50000) {
                clearInterval(timer);
                window.scrollTo(0, 0);
                setTimeout(resolve, 1100 + Math.random() * 900);
            }
        }, 180 + Math.random() * 160);
    });
}

// ------------------------------------------------------------
// 1단계: 검색 결과 파싱
// ------------------------------------------------------------
function extractDataFromPage() {
    const safeText = (t) => t ? t.replace(/[\t\n\r]+/g, " ").replace(/\s{2,}/g, " ").trim() : "";

    let items = [...document.querySelectorAll('div.product_item__KQayS')]
        .filter(el => el.querySelector('[data-shp-contents-type="catalog_nv_mid"]'));

    if (items.length === 0) {
        const seen = new Set();
        items = [...document.querySelectorAll('[data-shp-contents-type="catalog_nv_mid"]')]
            .map(el => el.closest('li, div[class*="product_item"]') || el.parentElement)
            .filter(el => {
                if (!el || seen.has(el)) return false;
                seen.add(el);
                return true;
            });
    }

    const rows = items.map(item => {
        // 제목 링크: 지정 클래스를 먼저 쓰고, 없으면 카탈로그로 가는 링크 중
        // 글자가 가장 긴 것을 고른다. (썸네일/배지 링크가 먼저 잡히는 것 방지)
        let linkEl = item.querySelector('a.product_link__aFnaq');
        if (!linkEl) {
            const cands = [...item.querySelectorAll('a')]
                .filter(a => (a.getAttribute('href') || '').indexOf('/catalog/') !== -1);
            cands.sort((a, b) => (b.innerText || '').trim().length - (a.innerText || '').trim().length);
            linkEl = cands[0] || null;
        }

        let title = safeText(linkEl && linkEl.innerText);
        // 제목이 "3개" 처럼 너무 짧으면 링크를 잘못 고른 것 -> 카드에서 가장 긴 링크 텍스트로 교체
        if (title.length < 8) {
            const all = [...item.querySelectorAll('a')]
                .map(a => safeText(a.innerText))
                .filter(t => t.length >= 8);
            all.sort((a, b) => b.length - a.length);
            if (all[0]) title = all[0];
        }

        let catalogUrl = "";
        const catAnchor = item.querySelector('a[href*="/catalog/"]');
        if (catAnchor && catAnchor.href) catalogUrl = catAnchor.href.split('?')[0];

        let nvMid = "";
        if (catalogUrl) {
            const m = catalogUrl.match(/\/catalog\/(\d+)/);
            if (m) nvMid = m[1];
        }
        if (!nvMid && linkEl && linkEl.dataset && linkEl.dataset.i) nvMid = linkEl.dataset.i;
        // [예비] 특이 구조 카드(수량별 링크형 등)는 대표 주소가 없을 수 있다.
        // 네이버가 카드에 심어둔 데이터 속성에서 카탈로그 번호를 직접 읽는다.
        if (!nvMid) {
            const marker = item.matches && item.matches('[data-shp-contents-type="catalog_nv_mid"]')
                ? item
                : item.querySelector('[data-shp-contents-type="catalog_nv_mid"]');
            if (marker) {
                const cid = marker.getAttribute('data-shp-contents-id')
                    || marker.getAttribute('data-shp-contents-dtl') || '';
                const m2 = String(cid).match(/(\d{8,})/);
                if (m2) nvMid = m2[1];
            }
        }
        if (!catalogUrl && nvMid) catalogUrl = 'https://search.shopping.naver.com/catalog/' + nvMid;

        const priceEl = item.querySelector('span.price_num__Y66T7 em');
        const lowestPrice = (priceEl ? priceEl.innerText : "").replace(/[^0-9]/g, "");

        let deliveryFee = "";
        const feeEl = item.querySelector('div.price_delivery_fee__8n1e5');
        if (feeEl) {
            const feeText = feeEl.innerText;
            if (feeText.includes("무료")) deliveryFee = "0";
            else deliveryFee = feeText.replace(/[^0-9]/g, "");
        }

        let lowestSeller = "";
        const lowIco = item.querySelector('ul.product_mall_list__rYuBz li span.product_ico_low__qVBEJ');
        const lowMallItem = lowIco ? lowIco.closest('a.product_mall_item__MZNBr') : null;
        if (lowMallItem) {
            const n = lowMallItem.querySelector('span.product_mall_name__DuUQV');
            lowestSeller = safeText(n && n.innerText);
        } else {
            const n = item.querySelector('ul.product_mall_list__rYuBz li a span.product_mall_name__DuUQV');
            lowestSeller = safeText(n && n.innerText);
        }

        // [수량 링크] 카드 안에는 "1개" "3개" 처럼 수량만 적힌 링크가 붙어 있다.
        //   = 이 카탈로그에 있는 수량 옵션 목록이다.
        //   낱개(1개)만 파는 카탈로그를 페이지를 열기 전에 걸러내는 데 쓴다. (v4.91)
        //   ※ 단위 글자가 붙은 것만 센다. 단위 없는 숫자(쪽번호 등)를 수량으로
        //     잘못 세면 멀쩡한 카탈로그를 안 보고 버리게 된다.
        const qtySet = new Set();
        for (const a of item.querySelectorAll('a')) {
            // "1개 (30포)" 처럼 괄호가 붙는 표기가 있다 — 옵션 패널 파서와 같은 규칙으로 먼저 떼어낸다
            const t = safeText(a.innerText).replace(/\([^)]*\)/g, "").replace(/\s+/g, "");
            if (!t || t.length > 8) continue;
            const qm = t.match(/^(\d{1,3})(개입|개|세트|팩|박스|입|병|캔|포|묶음)$/);
            if (!qm) continue;
            const q = Number(qm[1]);
            if (q >= 1 && q <= 999) qtySet.add(q);
        }
        const cardQtys = [...qtySet].sort((x, y) => x - y);

        return { title, lowestPrice, deliveryFee, lowestSeller, nvMid, catalogUrl, cardQtys };
    });

    // [자체 점검] 화면의 카탈로그 카드 수와 실제 수집 수를 비교할 수 있게,
    // 걸러진 카드는 "왜 걸러졌는지"를 함께 돌려준다.
    const seenId = new Set();
    const kept = [];
    const skipped = [];
    for (const r of rows) {
        if (!r.title || !r.lowestPrice || !r.catalogUrl) {
            skipped.push({
                nvMid: r.nvMid || '?',
                title: (r.title || '(제목 못 읽음)').slice(0, 40),
                reason: !r.title ? '상품명 못 읽음' : (!r.lowestPrice ? '가격 못 읽음' : '주소 못 읽음')
            });
            continue;
        }
        if (seenId.has(r.nvMid)) continue;   // 같은 카탈로그 중복 카드
        seenId.add(r.nvMid);
        kept.push(r);
    }

    // [자체 점검 2] 페이지에 있는 모든 카탈로그 링크 번호와 대조해서
    // 우리가 못 담은 카탈로그가 있는지 찾는다 (누락 의심 목록)
    const keptIds = new Set(kept.map(r => r.nvMid));
    const missing = [];
    const seenMiss = new Set();
    for (const a of document.querySelectorAll('a[href*="/catalog/"]')) {
        const m = (a.getAttribute('href') || '').match(/\/catalog\/(\d+)/);
        if (!m) continue;
        const id = m[1];
        if (keptIds.has(id) || seenMiss.has(id)) continue;
        seenMiss.add(id);
        const text = (a.innerText || '').replace(/\s+/g, ' ').trim().slice(0, 40);
        // "1개", "2개" 같은 수량 변형 링크는 수집된 카드의 부속물이므로 누락이 아니다
        if (/^\d{1,3}\s*(개입|개|세트|팩|박스|입|병|캔|포|묶음)?$/.test(text)) continue;
        if (missing.length < 10) missing.push({ id: id, text: text });
    }

    return { rows: kept, cardCount: items.length, skipped: skipped, missing: missing };
}

// ------------------------------------------------------------
// [리뷰 필터] 카탈로그 페이지에서 "최근 N일 안에 달린 리뷰 개수"를 세어,
//  실제로 잘 팔리는(리뷰가 활발한) 카탈로그인지 판단한다. (페이지 안에서 실행)
//  - 리뷰 영역을 찾아 최신순으로 정렬 후, 리뷰 날짜를 읽어 N일 이내만 카운트
//  - 네이버는 클래스가 해시라 접두어([class*="review..."])+텍스트 패턴으로 폭넓게 잡는다
//  - 못 읽거나 리뷰 영역을 못 찾으면 pass=false (안전하게 건너뜀)
//  - 첫 카탈로그에서 리뷰 영역 HTML을 sample 로 돌려줘 선택자 보정에 쓴다
// ------------------------------------------------------------
// 페이지 안에서 실행: "쿠폰" 글자가 붙은 가격 요소 주변을 글자로 떠 온다 (진단용, 1회).
//  무엇을 읽을지 아직 모른다 — 루크 화면의 진단 파일을 보고 선택자를 정한다.
function sampleCouponCells() {
    try {
        const out = [];
        const all = document.querySelectorAll('li, tr, div');
        for (const el of all) {
            if (out.length >= 6) break;
            const t = (el.innerText || '').replace(/\s+/g, ' ').trim();
            if (t.length < 8 || t.length > 220) continue;
            if (t.indexOf('쿠폰') < 0 || !/\d{1,3}(,\d{3})+원?/.test(t)) continue;
            // 같은 글이 부모·자식으로 겹쳐 나오지 않게 가장 안쪽만
            if (el.querySelector && [...el.querySelectorAll('*')].some(c => (c.innerText || '').indexOf('쿠폰') >= 0 && (c.innerText || '').trim().length >= 8)) continue;
            out.push(t + '  ⟨' + (el.className && String(el.className).slice(0, 60)) + '⟩');
        }
        return out.length ? 'url=' + location.href + '\n' + out.join('\n') : '';
    } catch (e) { return 'err ' + (e && e.message); }
}

// ------------------------------------------------------------
// [카탈로그 페이지에서 실행, 1개 옵션 상태] 판매처 목록에서 "공식" 표시가 붙은 판매처를 찾는다. (낱개 사입, v4.98)
//   돌려주는 것: 이름 · 그 판매처의 값 · 최대할인가(있으면) · 상품 페이지 링크. 여럿이면 제일 싼 공식몰.
//   네이버 클래스는 해시라 글자로 찾는다: "공식" 글자와 가격이 같이 있는 가장 안쪽 칸.
function readOfficialSellerOnPage(brandName) {
    const norm = (el) => (el && el.innerText || '').replace(/\s+/g, ' ').trim();
    const won = (t) => { const m = String(t || '').match(/(\d{1,3}(?:,\d{3})+|\d+)\s*원/); return m ? Number(m[1].replace(/,/g, '')) : 0; };
    const out = { found: false, name: '', price: 0, ship: 0, maxDiscount: 0, link: '', sample: '' };
    try {
        const all = [...document.querySelectorAll('li, div, a')];
        const cands = all.filter(el => {
            const t = norm(el);
            if (t.length < 8 || t.length > 400) return false;
            if (!/공식/.test(t) || !/\d{1,3}(?:,\d{3})+\s*원/.test(t)) return false;
            // 실측: 요약 블록("최저 14,360원 … 공식인증 사러가기")과 광고 카드("찜하기 … 포인트 최대 700원")도 "공식" 글자가 있다
            if (/사러가기|최저\s*[\d,]+\s*원|^찜하기|포인트 최대/.test(t)) return false;
            // 실측 8차: 딴 상품 추천 카드("구시다플러스 200g 6,100원 … 북설악 공식 스토어")는 **값이 "공식" 보다 앞에** 온다.
            //   이 카탈로그의 판매처 카드는 "닥터지 공식 … 28,900원" 처럼 가게 이름(공식)이 먼저, 값이 뒤다. 값이 먼저면 상품 카드다.
            //   (값만으로 거르는 15% 규칙은 forwardBuyStep 에 따로 있다 — 둘 다 쓴다)
            const _pi = t.search(/\d{1,3}(?:,\d{3})+\s*원/), _oi = t.search(/공식/);
            if (_pi >= 0 && _oi > _pi) return false;
            return el.tagName === 'A' || !!el.querySelector('a[href]');
        });
        const inner = cands.filter(el => !cands.some(o => o !== el && el.contains(o)));
        const picks = [];
        for (const el of inner) {
            // 값: "최대할인가 N원" 을 먼저 떼어낸 뒤, 포인트·배송비가 아닌 첫 금액 (실측: 카드 글자가 한 줄로 이어져 온다)
            let t = norm(el);
            let maxDiscount = 0;
            const md = t.match(/최대할인가\s*(\d{1,3}(?:,\d{3})+|\d+)\s*원/);
            if (md) { maxDiscount = Number(md[1].replace(/,/g, '')); t = t.replace(md[0], ' '); }
            let price = 0;
            const re = /(\d{1,3}(?:,\d{3})+|\d{4,})\s*원/g;
            let mm;
            while ((mm = re.exec(t))) {
                const before = t.slice(Math.max(0, mm.index - 12), mm.index);
                if (/포인트|적립|배송비|최대\s*$/.test(before)) continue;
                price = Number(mm[1].replace(/,/g, '')); break;
            }
            // 이름: 통짜 글자에서 "OO 공식" / "OO 공식스토어" 의 OO. (실측: 줄 나누기가 카드마다 달라 줄로는 못 집는다)
            //   "브랜드 공식"·"도움말"·"최대 혜택"·"단위당 최저가" 같은 배지 글자는 이름이 아니다. "[이솝 공식]" 은 이솝.
            let name = '';
            const full = norm(el);
            const nre = /([^\s\[\]]+)\s*공식(?:\s*(?:몰|스토어|인증|브랜드스토어))?/g;
            let nm;
            while ((nm = nre.exec(full))) {
                // 실측: 글자가 띄어쓰기 없이 붙어 "분말3,500원배송비3,000원하남댁" 처럼 온다 → 마지막 "원" 뒤만 이름
                let cand = nm[1].replace(/^[\[\(]+|[\]\)]+$/g, '').replace(/^.*원/, '').replace(/^.*배송/, '').replace(/^(오늘출발|내일도착|당일출발|도착보장)/, '').trim();
                if (!cand || /^(브랜드|도움말|최대|혜택|단위당|최저가|찜하기|가맹점)$/.test(cand) || /\d/.test(cand)) continue;
                name = cand; break;
            }
            // 배송비: "배송비 3,000원" 이면 그만큼, "무료배송/배송비 무료" 면 0. 카탈로그 최저가는 배송비 포함이라 같이 비교해야 한다.
            let ship = 0;
            const sm = full.match(/배송비\s*(\d{1,3}(?:,\d{3})+|\d+)\s*원/);
            if (sm) ship = Number(sm[1].replace(/,/g, ''));
            else if (/무료배송|배송비\s*무료|배송비무료/.test(full)) ship = 0;
            if (!price) continue;
            const a = el.tagName === 'A' ? el : el.querySelector('a[href]');
            picks.push({ name: name || '공식몰', price: price, ship: ship, maxDiscount: maxDiscount, link: (a && a.href) || '', text: norm(el).slice(0, 160) });
        }
        if (!picks.length) return out;
        // 브랜드 이름이 든 공식몰이 있으면 그것(진짜 브랜드 공식몰), 없으면 배송비 포함 제일 싼 것
        picks.sort((a, b) => (a.price + a.ship) - (b.price + b.ship));
        const bn = String(brandName || '').toLowerCase().replace(/[\s\-_·.]/g, '');
        const p = (bn.length >= 2 && picks.find(x => x.name.toLowerCase().replace(/[\s\-_·.]/g, '').includes(bn))) || picks[0];
        out.found = true; out.name = p.name; out.price = p.price; out.ship = p.ship; out.maxDiscount = p.maxDiscount; out.link = p.link;
        out.sample = picks.map(x => x.text).join('\n');
        return out;
    } catch (e) { out.sample = 'err ' + (e && e.message); return out; }
}

// ------------------------------------------------------------
// [스마트스토어 상품 페이지에서 실행] 1개를 실제로 얼마에 사는지 읽는다. (낱개 사입, v4.98)
//   실제 매입가 = 최종 구매 예상가(최대할인가) − 구매적립(기본+추가) − 멤버십적립(멤버십 추가+네이버 현대카드)
//   리뷰적립·결제적립은 빼지 않는다 (루크 규칙). 멤버십·현대카드는 "다 갖춘다"는 전제로 최대 조건을 쓴다.
//   숫자는 "최대 적립 포인트" 글자를 눌러 열리는 "나의 혜택 전체보기" 창에서 읽는다 (루크 화면 실측 2026-11-01).
//   창 글자(modalText)도 돌려줘 첫 상품은 진단 파일에 담는다 — 네이버가 창을 바꾸면 거기서 보인다.
async function readStoreBenefitOnPage() {
    const sleep = (ms) => new Promise(r => setTimeout(r, ms));
    const norm = (el) => (el && el.innerText || '').replace(/\s+/g, ' ').trim();
    const won = (t) => { const m = String(t || '').match(/(\d{1,3}(?:,\d{3})+|\d+)\s*원/); return m ? Number(m[1].replace(/,/g, '')) : 0; };
    const out = { ok: false, reason: '', storeName: '', shipping: 0, salePrice: 0, maxDiscount: 0, finalPrice: 0, priceBasis: '',
                  points: { basic: 0, extra: 0, membership: 0, card: 0 }, modalText: '' };
    try {
        // ① 페이지 글자에서 최대할인가 (창이 안 열려도 이것만은 건진다)
        const bodyLines = (document.body.innerText || '').split('\n').map(x => x.trim()).filter(Boolean);
        for (let i = 0; i < bodyLines.length; i++) {
            const t = bodyLines[i];
            if (/최대할인가/.test(t)) { out.maxDiscount = won(t) || won(bodyLines[i - 1]) || won(bodyLines[i + 1]); if (out.maxDiscount) break; }
        }
        try { const h = document.querySelector('meta[property="og:site_name"]'); if (h) out.storeName = h.getAttribute('content') || ''; } catch (_) {}
        // 배송비: 1개를 살 때 내는 돈이라 매입가에 들어간다. "무료배송" 이면 0, "배송비 3,000원" 이면 그만큼.
        {
            const bt = bodyLines.join(' ');
            const sm = bt.match(/배송비\s*(\d{1,3}(?:,\d{3})+|\d+)\s*원/);
            if (sm) out.shipping = Number(sm[1].replace(/,/g, ''));
            else if (/무료배송|배송비\s*무료/.test(bt)) out.shipping = 0;
            else out.shipping = 0;
        }

        // ② "최대 적립 포인트" 를 누른다
        const cands = [...document.querySelectorAll('a, button, span, div, strong')]
            .filter(el => /최대\s*적립\s*포인트/.test(norm(el)) && norm(el).length <= 40);
        const btn = cands.length ? cands[cands.length - 1] : null;
        if (!btn) { out.reason = '적립 글자 없음'; if (out.maxDiscount) { out.finalPrice = out.maxDiscount; out.priceBasis = '페이지 최대할인가(적립 못 읽음)'; } return out; }
        btn.click();

        // ③ 창이 뜰 때까지 기다린다 (최대 4초)
        let modal = null;
        for (let k = 0; k < 8 && !modal; k++) {
            await sleep(500);
            const boxes = [...document.querySelectorAll('div, section, aside')]
                .filter(el => { const t = norm(el); return /구매적립/.test(t) && /멤버십적립|최종 구매 예상가/.test(t) && t.length > 60 && t.length < 8000; })
                .sort((a, b) => norm(a).length - norm(b).length);
            // 실측: 제일 작은 상자는 "포인트 혜택" 만 있고 최종 구매 예상가는 그 바깥에 있다 → 예상가까지 품은 것을 먼저 고른다
            modal = boxes.find(el => /최종 구매 예상가/.test(norm(el))) || boxes[0] || null;
        }
        if (!modal) { out.reason = '혜택 창이 안 열림'; if (out.maxDiscount) { out.finalPrice = out.maxDiscount; out.priceBasis = '페이지 최대할인가(창 못 염)'; } return out; }
        const text = (modal.innerText || '');
        out.modalText = text.trim().slice(0, 3000);

        // ④ 줄 단위로 읽는다. 머리글(구매적립/리뷰적립/결제적립/멤버십적립)이 나오면 그 아래 금액을 그 묶음으로 센다.
        //    "기본적립 145원" 처럼 한 줄에 있든, "기본적립" 다음 줄에 "145원" 이 오든 다 센다.
        const lines = text.split('\n').map(x => x.trim()).filter(Boolean);
        //  실측: 창 글자는 "145원 / 290원 / 기본적립 / 추가구매적립" 처럼 금액이 이름보다 먼저 올 수 있다.
        //  그래서 묶음(구매적립/멤버십적립) 안에서 이름 줄과 금액 줄을 따로 모아 순서대로 짝짓는다.
        //  짝이 안 맞으면 합계만 맞추고(기본/멤버십에 몰아 넣음) — 매입가 계산에는 합계만 쓰인다.
        let sec = '';
        const bag = { buy: { labels: [], amounts: [] }, member: { labels: [], amounts: [] } };
        for (let i = 0; i < lines.length; i++) {
            const t = lines[i];
            if (/^최종\s*구매\s*예상가/.test(t)) { out.finalPrice = won(t) || won(lines[i + 1]); sec = 'end'; continue; }
            if (/^구매적립/.test(t)) { sec = 'buy'; continue; }
            if (/리뷰적립/.test(t)) { sec = 'review'; continue; }
            if (/^결제적립/.test(t)) { sec = 'pay'; continue; }
            if (/^멤버십적립/.test(t)) { sec = 'member'; continue; }
            if (/^최대\s*포인트\s*적립|할인 총 합계|상품 판매가|^쿠폰|할인 혜택|포인트 혜택/.test(t)) { sec = ''; continue; }
            if (sec !== 'buy' && sec !== 'member') continue;
            const v = won(t);
            const label = t.replace(/(\d{1,3}(?:,\d{3})+|\d+)\s*원/g, '').replace(/[•·\-–>]/g, '').trim();
            // 항목 이름으로 치는 줄 = "…적립" 또는 "…결제 시". 도움말·네이버페이·네이버플러스멤버십 같은 꼬리표 줄은 아니다.
            if (label.length >= 2 && /적립|결제\s*시/.test(label) && !/^(도움말|네이버페이|네이버플러스멤버십)$/.test(label)) bag[sec].labels.push(label);
            if (v) bag[sec].amounts.push(v);
        }
        const assign = (secName, isSpecial, specialKey, normalKey) => {
            const b = bag[secName];
            if (b.labels.length === b.amounts.length && b.labels.length > 0) {
                b.labels.forEach((lb, k) => { out.points[isSpecial.test(lb) ? specialKey : normalKey] += b.amounts[k]; });
            } else {
                out.points[normalKey] += b.amounts.reduce((a, c) => a + c, 0);
            }
        };
        assign('buy', /추가/, 'extra', 'basic');
        assign('member', /현대카드/, 'card', 'membership');
        if (out.finalPrice) out.priceBasis = '창의 최종 구매 예상가';
        else if (out.maxDiscount) { out.finalPrice = out.maxDiscount; out.priceBasis = '페이지 최대할인가'; }
        out.ok = out.finalPrice > 0;
        if (!out.ok) out.reason = '가격을 못 읽음';

        // ⑤ 닫는다 (ESC → 닫기 버튼 순)
        try { document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', keyCode: 27, bubbles: true })); } catch (_) {}
        await sleep(300);
        try {
            const x = [...document.querySelectorAll('button, a')].find(b => /닫기|close/i.test(norm(b) + ' ' + (b.getAttribute('aria-label') || '')) && norm(b).length <= 6);
            if (x) x.click();
        } catch (_) {}
        return out;
    } catch (e) { out.reason = 'err ' + (e && e.message); return out; }
}

async function scrapeReviewVelocity(days, minCount, maxPages) {
    const sleep = (ms) => new Promise(r => setTimeout(r, ms));
    const now = Date.now(), DAY = 86400000, cutoff = now - days * DAY;

    // ── 1순위: 카탈로그 리뷰 API 로 정확히 센다 (실측 엔드포인트) ──
    //   /catalog/api/{id}/reviews?nvMid={id}&reviewType=ALL&isCatalog=true&trackTotalCount=true&sortType=RECENT&page=N
    //   최신순(RECENT)으로 페이지를 넘기며 14일 밖 리뷰가 나오면 멈춘다(그 뒤는 다 오래됨).
    const _idm = (location.href || '').match(/\/catalog\/(\d+)/);
    const _id = _idm ? _idm[1] : '';
    function _toTs(v) {
        if (v == null) return null;
        if (typeof v === 'number') return v > 1e12 ? v : (v > 1e9 ? v * 1000 : null);
        const s = String(v);
        let mm = s.match(/(\d{4})[-.\/]\s?(\d{1,2})[-.\/]\s?(\d{1,2})/);
        if (mm) { const d = new Date(+mm[1], +mm[2] - 1, +mm[3]).getTime(); if (!isNaN(d)) return d; }
        mm = s.match(/^(\d{2})\.(\d{1,2})\.(\d{1,2})/);
        if (mm) { const d = new Date(2000 + +mm[1], +mm[2] - 1, +mm[3]).getTime(); if (!isNaN(d)) return d; }
        const d = Date.parse(s); return isNaN(d) ? null : d;
    }
    function _dateOf(r) {
        if (!r || typeof r !== 'object') return null;
        const keys = ['registerDate', 'createDate', 'writeDate', 'regDate', 'reviewDate', 'mallCreateDate', 'contentUploadDate', 'registeredDate', 'date'];
        for (const k of keys) if (r[k] != null) { const t = _toTs(r[k]); if (t) return t; }
        for (const k in r) { const t = _toTs(r[k]); if (t && t <= now + DAY && t > now - 3650 * DAY) return t; }
        return null;
    }
    if (_id) {
        try {
            let recent = 0, seen = 0, total = -1, pages = 0, newestTs = 0, oldestTs = 0, sample = '';
            const PAGE_CAP = Math.max(1, Number(maxPages) || 6);
            let page = 1, done = false, capped = false;
            while (page <= PAGE_CAP && !done) {
                if (page > 1) await sleep(120 + Math.random() * 180);   // 연달아 때리지 않는다
                const url = 'https://search.shopping.naver.com/catalog/api/' + _id + '/reviews'
                    + '?nvMid=' + _id + '&reviewType=ALL&isCatalog=true&trackTotalCount=true&sortType=RECENT&page=' + page + '&pageSize=20';
                let j = null;
                try {
                    const res = await fetch(url, { credentials: 'include', headers: { accept: 'application/json' } });
                    if (!res.ok) { sample += 'HTTP' + res.status + '@p' + page + ' '; break; }
                    j = await res.json();
                } catch (e) { sample += 'fetchErr@p' + page + ' '; break; }
                const arr = (j && (j.reviews || j.contents || j.reviewList || j.list || j.reviewInfoList
                    || (j.result && (j.result.reviews || j.result.contents)) || (j.data && (j.data.reviews || j.data.contents)))) || [];
                if (page === 1) {
                    total = (j && (j.totalCount || j.total || (j.reviewInfo && j.reviewInfo.totalCount) || (j.paging && j.paging.totalCount))) || -1;
                    sample += 'keys=' + (j ? Object.keys(j).join(',') : 'null') + ' arr=' + arr.length + '\n';
                    if (arr[0]) sample += 'rev0=' + JSON.stringify(arr[0]).slice(0, 700) + '\n';
                }
                if (!arr.length) { done = true; break; }
                for (const r of arr) {
                    const t = _dateOf(r);
                    if (t == null) continue;
                    seen++;
                    if (t > newestTs) newestTs = t;
                    if (!oldestTs || t < oldestTs) oldestTs = t;
                    if (t >= cutoff) recent++; else done = true;   // RECENT 정렬 → 기간 밖이 나오면 이후는 다 오래됨
                }
                pages = page; page++;
            }
            // 상한에 닿아 멈췄다 = 14일 안쪽 리뷰가 아직 더 있다는 뜻
            if (!done && page > PAGE_CAP) capped = true;
            // 안전장치: API 날짜가 '진짜 최근 날짜'로 제대로 읽힌 경우만 신뢰. 이상하면(미래·너무 옛날·못 읽음) DOM 폴백.
            const sane = seen > 0 && newestTs <= now + 2 * DAY && newestTs > now - 3650 * DAY;
            if (sane) {
                // 상한에 닿았으면(기간 안 리뷰가 더 있다) 본 것의 속도로 기간 전체를 환산한다.
                //  예) 120개가 4일 동안 달렸다 → 2주면 420개. 1,000개짜리 카탈로그도 6페이지로 가늠.
                let projected = 0;
                if (capped && recent > 0) {
                    const spanDays = Math.max(0.5, (newestTs - oldestTs) / DAY);
                    projected = Math.max(recent, Math.round(recent / spanDays * days));
                }
                const pass = (projected || recent) >= minCount;
                sample = 'API pass=' + pass + ' recent=' + recent + ' projected=' + projected + ' seen=' + seen + ' total=' + total
                    + ' pages=' + pages + '/' + PAGE_CAP + ' capped=' + capped
                    + ' / days=' + days + ' min=' + minCount + '\n' + sample;
                return { pass, recentCount: recent, projected, seen, total, capped, why: '', sample };
            }
        } catch (e) { /* API 실패 → 아래 DOM 폴백 */ }
    }

    // ── 2순위(폴백): API 실패 시 DOM 에서 최신 리뷰 날짜로 속도 추정 ──
    // 네이버 카탈로그 리뷰 섹션 (실측: catalogMallReviewContainer_ / #section-mall_review)
    function reviewSection() {
        return document.querySelector('#section-mall_review')
            || document.querySelector('[class*="catalogMallReviewContainer"]')
            || document.querySelector('[class*="reviewList_"]')
            || document.querySelector('[class*="reviewItems_"]')
            || document.querySelector('[class*="review" i]');
    }

    // 리뷰 섹션으로 스크롤 + 여러 번 훑어 지연 로딩된 개별 리뷰(날짜)를 띄운다
    try {
        const sec = reviewSection();
        if (sec) { try { sec.scrollIntoView({ block: 'start' }); } catch (e) {} await sleep(1000); }
        for (let s = 0; s < 5; s++) { window.scrollBy(0, 900); await sleep(500); }
    } catch (e) {}
    // 최신순 정렬 — 최근 리뷰가 위로 오게 (후보 여러 개 시도)
    let sortClicked = false;
    try {
        const cands = [...document.querySelectorAll('a,button,span,li,div')]
            .filter(e => (e.textContent || '').trim() === '최신순' && (e.getBoundingClientRect().height || 0) < 60);
        for (const el of cands) { try { (el.closest('a,button,li') || el).click(); sortClicked = true; } catch (e) {} }
        if (sortClicked) await sleep(1600);
    } catch (e) {}
    // 리뷰 더 로드: 리뷰 목록의 '더보기'/페이지 넘김이 있으면 몇 번 눌러 더 띄운다
    try {
        for (let p = 0; p < 4; p++) {
            const more = [...document.querySelectorAll('a,button')].find(e => {
                const t = (e.textContent || '').replace(/\s+/g, '');
                return (t === '더보기' || /리뷰더보기|더보기리뷰/.test(t))
                    && !!e.closest('[class*="review" i],[class*="Review"]')
                    && (e.getBoundingClientRect().height || 0) < 90;
            });
            if (!more) break;
            try { more.click(); } catch (e) { break; }
            await sleep(1200);
        }
    } catch (e) {}
    try { for (let s = 0; s < 3; s++) { window.scrollBy(0, 800); await sleep(500); } } catch (e) {}

    function parseKDate(c) {
        let m = c.match(/^(\d{2})\.(\d{1,2})\.(\d{1,2})\.?$/);        // 24.09.15 / 24.09.15.
        if (m) { const dt = new Date(2000 + +m[1], +m[2] - 1, +m[3]).getTime(); if (!isNaN(dt)) return dt; }
        m = c.match(/^(20\d{2})\.(\d{1,2})\.(\d{1,2})\.?$/);          // 2024.09.15
        if (m) { const dt = new Date(+m[1], +m[2] - 1, +m[3]).getTime(); if (!isNaN(dt)) return dt; }
        if (/^(방금|오늘)$/.test(c) || /^\d+(시간|분)전$/.test(c)) return now;
        if (/^어제$/.test(c)) return now - DAY;
        m = c.match(/^(\d+)일전$/); if (m) return now - (+m[1]) * DAY;
        m = c.match(/^(\d+)주전$/); if (m) return now - (+m[1]) * 7 * DAY;
        return null;   // 개월/년 전 등은 오래됨(무시)
    }
    function isDateStr(c) {
        return /^\d{2}\.\d{1,2}\.\d{1,2}\.?$/.test(c) || /^(오늘|어제|방금|\d+일전|\d+시간전|\d+분전|\d+주전)$/.test(c);
    }
    function pickDateIn(el) {
        for (const k of el.querySelectorAll('span,em,time,div,p')) {
            const c = (k.textContent || '').replace(/\s+/g, '');
            if (c.length <= 12 && isDateStr(c)) return c;
        }
        return null;
    }

    // 실측 구조: 개별 리뷰 = catalogMallReviewItem_mall_review_item, 그 안 날짜 = reviewItemProfile_data (26.06.23. 형식)
    //  중복 카운트 방지 위해 "리뷰 아이템" 컨테이너만 정확히 겨냥한다 (중첩 요소 제외).
    let items = [...document.querySelectorAll('[class*="MallReviewItem_mall_review_item"],[class*="mall_review_item"]')];
    if (!items.length) items = [...document.querySelectorAll('[class*="ReviewItem_review"],[class*="reviewItems_item"],[class*="review_item_"]')];
    const dates = [];
    const dateStrs = [];
    let firstDateEl = null;
    try {
        if (items.length) {
            for (const it of items) {
                const c = pickDateIn(it);
                if (!c) continue;
                const d = parseKDate(c);
                if (d !== null) { dates.push(d); dateStrs.push(c); if (!firstDateEl) firstDateEl = it; }
            }
        } else {
            // 폴백: 페이지 전체 텍스트에서 날짜 훑기
            const w = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
            let node, guard = 0;
            while ((node = w.nextNode()) && guard < 9000) {
                guard++;
                const c = (node.nodeValue || '').replace(/\s+/g, '');
                if (!c || c.length > 12 || !isDateStr(c)) continue;
                const d = parseKDate(c);
                if (d !== null) { dates.push(d); dateStrs.push(c); if (!firstDateEl) firstDateEl = node.parentElement; }
            }
        }
    } catch (e) {}

    // 최근 days일 내 리뷰 개수. 미리보기가 적으면(예: 4개) 보이는 최신 리뷰들의 '작성 속도'로 추정한다.
    //  - 최신 리뷰 N개가 며칠(span)에 걸쳐 있으면 하루 (N-1)/span 개 → days일 예상 개수를 계산
    //  - 가장 최근 리뷰가 기간 안(살아있음)이고 예상 개수 >= 기준이면 통과
    const recent = dates.filter(d => d >= cutoff).length;
    let pass = recent >= minCount, mode = 'count', projected = 0;
    if (!pass && dates.length >= 3) {
        const sorted = dates.slice().sort((a, b) => b - a);   // 최신순
        const newest = sorted[0], oldest = sorted[sorted.length - 1];
        const spanDays = Math.max((newest - oldest) / DAY, 0.5);
        projected = Math.round(((sorted.length - 1) / spanDays) * days);
        if (newest >= cutoff && projected >= minCount) { pass = true; mode = 'rate'; }
    }

    // 진단: 판정·모드·개수·읽은 날짜들을 다 찍는다 (보정용)
    let sample = 'pass=' + pass + ' mode=' + mode + ' / recent=' + recent + ' / proj=' + projected
        + ' / seen=' + dates.length + ' / items=' + items.length
        + ' / sortClicked=' + sortClicked + ' / days=' + days + ' / min=' + minCount + '\n'
        + 'dates=[' + dateStrs.slice(0, 24).join(', ') + ']\n';
    // 리뷰를 불러오는 네트워크 주소(API)를 찾아 찍는다 → 정확한 API 기반 카운트로 바꾸기 위함
    try {
        const rurls = performance.getEntriesByType('resource').map(e => e.name)
            .filter(u => /review|Review|revList|revsmr|comment/.test(u));
        if (rurls.length) sample += '[NET]\n' + rurls.slice(0, 8).join('\n') + '\n';
        else sample += '[NET] (리뷰 관련 네트워크 없음)\n';
    } catch (e) {}
    // 리뷰 '더보기/페이지 넘김' 구조 캡처: 리뷰 섹션의 끝부분(페이지네이션이 있는 곳)
    try {
        const sec = reviewSection();
        if (sec) { const h = sec.outerHTML || ''; sample += '[TAIL]\n' + h.slice(Math.max(0, h.length - 2600)) + '\n'; }
        // 페이지 넘김처럼 보이는 요소들의 텍스트도 뽑아본다
        const pag = [...document.querySelectorAll('a,button')]
            .filter(e => { const t = (e.textContent || '').replace(/\s+/g, ''); return /^(다음|이전|더보기|\d{1,3})$/.test(t) && e.closest('[class*="review" i],[class*="Review"]'); })
            .map(e => (e.textContent || '').trim() + '@' + (e.className || e.tagName)).slice(0, 12);
        if (pag.length) sample += '[PAG] ' + pag.join(' | ') + '\n';
    } catch (e) {}
    try {
        if (firstDateEl) {
            let box = firstDateEl;
            for (let i = 0; i < 5 && box && (box.textContent || '').length < 300; i++) box = box.parentElement;
            sample += ((box && box.outerHTML) || '').slice(0, 1400);
        } else {
            const sec = reviewSection();
            sample += '(리뷰 날짜 못 찾음) section=' + (sec ? (sec.className || sec.id || 'y') : 'none');
        }
    } catch (e) { sample += '(sample err)'; }

    return { pass, recentCount: recent, projected, mode, total: -1, capped: false,
             seen: dates.length, why: firstDateEl ? '' : 'no-dates', sample };
}

// ------------------------------------------------------------
// 2단계: 카탈로그에서 "수량(N개)" 옵션만 골라 읽는다
//
//  성공률을 올리기 위한 장치들:
//   - 버튼이 2~4줄이어도 "수량처럼 생긴 줄"과 "가격처럼 생긴 줄"을 찾아낸다
//   - 라벨/가격이 다른 요소에 흩어져 있으면 부모를 타고 올라가 짝을 맞춘다
//   - 접혀 있는 옵션은 "옵션 전체보기" 류를 눌러 펼친 뒤 다시 읽는다
//   - 못 읽으면 잠깐 기다렸다 다시 읽는다 (최대 maxTries 회)
//   - 그래도 실패하면 "왜 실패했는지"를 같이 돌려준다
// ------------------------------------------------------------
async function scrapeQuantityOptions(maxTries) {
    const sleep = (ms) => new Promise(r => setTimeout(r, ms));
    const QTY_RE = /^(\d{1,3})(개입|개|세트|팩|박스|입|병|캔|묶음)$/;
    const PRICE_RE = /^[\d,]{3,}원?$/;
    const PENALTY = /판매처|리뷰|쿠폰|카드|무이자|배송/;

    // "1개 (30포)" -> 1 / "2박스" -> 2 / 아니면 0
    function qtyOf(text) {
        if (!text) return 0;
        const t = String(text).replace(/\([^)]*\)/g, '').replace(/\s+/g, '');
        const m = t.match(QTY_RE);
        return m ? Number(m[1]) : 0;
    }

    function priceOf(text) {
        if (!text) return 0;
        const t = String(text).replace(/\s+/g, '');
        if (!PRICE_RE.test(t)) return 0;
        const n = Number(t.replace(/[^0-9]/g, ""));
        return n >= 100 ? n : 0;
    }

    const SELECTOR = 'button, a, li, label, div, span, [role="button"], [role="radio"], [role="tab"]';

    // 1차: 한 요소 안에 수량 줄과 가격 줄이 같이 있는 경우
    function collectPaired() {
        const cand = [];
        for (const el of document.querySelectorAll(SELECTOR)) {
            const rough = el.textContent;
            if (!rough) continue;
            const t = rough.trim();
            if (!t || t.length > 60) continue;      // 큰 덩어리는 옵션 버튼이 아니다
            if (!/\d{3}/.test(t)) continue;         // 가격 숫자가 없으면 볼 것도 없다

            const raw = (el.innerText || "").trim();
            if (!raw || raw.length > 55) continue;

            const lines = raw.split('\n').map(s => s.trim()).filter(Boolean);
            if (lines.length < 2 || lines.length > 4) continue;

            let qty = 0, price = 0;
            for (const line of lines) {
                if (!qty) qty = qtyOf(line);
                if (!price) price = priceOf(line);
            }
            if (!qty || !price) continue;

            cand.push({ el: el, qty: qty, price: price });
        }
        return cand;
    }

    // 2차: "1개" 만 들어있는 요소를 찾고, 부모를 타고 올라가며 가격을 찾는다
    //      (라벨과 가격이 형제 요소로 떨어져 있는 레이아웃 대응)
    function collectSplit() {
        const cand = [];
        for (const el of document.querySelectorAll(SELECTOR)) {
            const raw = (el.innerText || "").trim();
            if (!raw || raw.length > 10) continue;
            const qty = qtyOf(raw);
            if (!qty) continue;

            let node = el.parentElement;
            for (let up = 0; up < 3 && node && node !== document.body; up++, node = node.parentElement) {
                const txt = (node.innerText || "").trim();
                if (txt.length > 60) break;                 // 너무 커지면 그만
                let price = 0;
                for (const line of txt.split('\n')) {
                    price = priceOf(line.trim());
                    if (price) break;
                }
                if (price) { cand.push({ el: node, qty: qty, price: price }); break; }
            }
        }
        return cand;
    }

    // 바깥을 감싸는 중복 요소 제거 (가장 안쪽만 남김)
    function innermost(cand) {
        return cand.filter(c => !cand.some(o => o !== c && c.el !== o.el && c.el.contains(o.el)));
    }

    function containerOf(c, all) {
        let node = c.el.parentElement;
        while (node && node !== document.body) {
            let n = 0;
            for (const o of all) if (node.contains(o.el)) n++;
            if (n >= 2) return node;
            node = node.parentElement;
        }
        return c.el.parentElement;
    }

    function nameOf(container) {
        let node = container;
        for (let up = 0; up < 4 && node && node !== document.body; up++, node = node.parentElement) {
            let sib = node.previousElementSibling;
            while (sib) {
                const t = (sib.innerText || "").trim();
                if (t && t.length <= 30 && !/^[\d,]+\s*원?$/.test(t)) {
                    return t.split('\n')[0].replace(/[\s:]{2,}/g, ' ').replace(/\s*:\s*$/, '').trim();
                }
                sib = sib.previousElementSibling;
            }
        }
        return "수량";
    }

    function pickGroup(cand) {
        if (cand.length < 2) return null;

        const map = new Map();
        for (const c of cand) {
            const box = containerOf(c, cand);
            if (!map.has(box)) map.set(box, []);
            map.get(box).push(c);
        }

        let best = null;
        for (const [box, vals] of map.entries()) {
            if (vals.length < 2) continue;
            const name = nameOf(box);
            // 이름으로 통째로 버리지 않는다. 라벨이 이미 "N개" 형태로 걸러져 있으므로
            // 이름은 점수 가감으로만 쓴다. (진짜 수량 그룹을 실수로 버리는 것을 막기 위함)
            let score = vals.length;
            if (/수량/.test(name)) score += 100;
            if (PENALTY.test(name)) score -= 50;
            if (!best || score > best.score) best = { score: score, name: name, vals: vals };
        }
        if (!best || best.score <= 0) return null;

        // 같은 수량이 여러 번이면 가장 싼 값만
        const byQty = new Map();
        for (const v of best.vals) {
            if (!byQty.has(v.qty) || v.price < byQty.get(v.qty).price) byQty.set(v.qty, v);
        }
        const options = [...byQty.values()]
            .map(v => ({ qty: v.qty, price: v.price }))
            .sort((a, b) => a.qty - b.qty);

        return options.length >= 2 ? { options: options, groupName: best.name } : null;
    }

    // 접혀 있는 옵션 펼치기
    function expandAll() {
        const LABEL = /^(옵션\s*전체보기|전체보기|옵션\s*더보기|더보기|펼치기|옵션\s*펼치기)$/;
        let clicked = 0;
        for (const el of document.querySelectorAll('button, span, div, a, [role="button"]')) {
            const t = (el.innerText || "").trim();
            if (!LABEL.test(t)) continue;
            const href = el.getAttribute && el.getAttribute('href');
            if (el.tagName === 'A' && href && href !== '#' && href !== 'javascript:void(0)') continue;
            if (el.querySelector('button, a')) continue;
            try { el.click(); clicked++; } catch (e) {}
            if (clicked >= 5) break;
        }
        return clicked;
    }

    function readOnce() {
        let cand = innermost(collectPaired());
        let picked = pickGroup(cand);
        if (picked) return { picked: picked, from: 'paired', candCount: cand.length };

        cand = innermost(collectSplit());
        picked = pickGroup(cand);
        if (picked) return { picked: picked, from: 'split', candCount: cand.length };

        return { picked: null, from: '', candCount: cand.length };
    }

    // --- 본 처리: 못 읽으면 펼치고, 기다리고, 다시 읽는다 ---
    let expanded = 0;
    let last = null;

    for (let attempt = 1; attempt <= maxTries; attempt++) {
        last = readOnce();
        if (last.picked) {
            return {
                options: last.picked.options,
                groupName: last.picked.groupName,
                reason: 'ok',
                debug: { attempt: attempt, from: last.from, expanded: expanded }
            };
        }
        if (attempt === 1) expanded = expandAll();
        window.scrollBy(0, 300 + Math.random() * 500);
        await sleep(800 + Math.random() * 900);
    }

    // --- 실패했다. 왜 실패했는지 남긴다 ---
    const bodyText = (document.body.innerText || "");

    let qtyLikeCount = 0;
    for (const el of document.querySelectorAll(SELECTOR)) {
        const t = (el.innerText || "").trim();
        if (t.length <= 10 && qtyOf(t)) qtyLikeCount++;
    }

    let near = "";
    const idx = bodyText.indexOf("수량");
    if (idx >= 0) near = bodyText.slice(Math.max(0, idx - 60), idx + 260).replace(/\s*\n\s*/g, " | ");

    let sampleHtml = "";
    for (const el of document.querySelectorAll(SELECTOR)) {
        const t = (el.innerText || "").trim();
        if (t.length <= 10 && qtyOf(t)) {
            const box = el.closest('ul, ol, section, div');
            sampleHtml = ((box && box.outerHTML) || el.outerHTML || "").slice(0, 1500);
            break;
        }
    }

    // 수량처럼 생긴 글자가 아예 없으면 이 상품엔 정말 수량 옵션이 없는 것.
    // 있는데도 못 읽었다면 우리 파서 문제다.
    const reason = qtyLikeCount === 0 ? 'no-qty-text' : 'qty-without-price';

    return {
        options: [],
        groupName: "",
        reason: reason,
        debug: {
            attempt: maxTries,
            expanded: expanded,
            qtyLikeCount: qtyLikeCount,
            hasQtyWord: /수량/.test(bodyText),
            near: near,
            sampleHtml: sampleHtml
        }
    };
}

// ------------------------------------------------------------
// 특정 수량 옵션을 선택하고, 그 상태의 주소 + 쿠팡 빠른배송 여부를 가져온다.
//   - 옵션을 다 읽은 "뒤에" 따로 실행한다. 클릭이 페이지 이동을 일으켜도
//     이미 확보한 가격 데이터는 잃지 않는다.
//   - 옵션이 링크(<a>)면 클릭하지 않고 주소만 읽는다. (이 경우 배송 감지는 현재 화면 기준)
// ------------------------------------------------------------
async function clickQtyAndInspect(wantQty, lowPrice) {
    const sleep = (ms) => new Promise(r => setTimeout(r, ms));
    const QTY_RE = /^(\d{1,3})(개입|개|세트|팩|박스|입|병|캔|묶음)$/;
    const before = location.href;

    function qtyOfLine(line) {
        const t = String(line || '').replace(/\([^)]*\)/g, '').replace(/\s+/g, '');
        const m = t.match(QTY_RE);
        return m ? Number(m[1]) : 0;
    }

    // [구조 샘플 수집] 최저가 블록 주변의 HTML을 담아온다.
    // (쿠팡 빠른배송 감지를 정확하게 만들기 위한 재료 - 진단 파일에 실린다)
    function captureSellerSample() {
        if (!lowPrice) return '';
        try {
            const el = document.querySelector('[class*="catalogLowestPrice_"]');
            if (!el) return '';
            let node = el, best = el;
            for (let up = 0; up < 5 && node; up++, node = node.parentElement) {
                if ((node.textContent || '').length > 500) break;
                best = node;
            }
            return (best.outerHTML || '').slice(0, 3000);
        } catch (e) {}
        return '';
    }

    // [카탈로그 제목] 옵션을 고르면 페이지 제목도 그 수량으로 바뀐다 (실측 확인됨)
    function readCatalogTitle() {
        try {
            const el = document.querySelector('[class*="catalogInfoTitle_catalog_title"]');
            if (el) {
                const t = (el.innerText || '').replace(/\s+/g, ' ').trim();
                if (t.length >= 5) return t;
            }
        } catch (e) {}
        return '';
    }

    // [쿠팡 빠른배송 여부] 매입처 읽기와 같은 검증된 방식:
    // 요약 영역의 "최저 N원" 줄 다음 몇 줄 안에 "쿠팡"과 "빠른배송"이 함께 있으면 참.
    // (1개 옵션 상태에서 읽으므로 = 1개 최저가 판매처가 쿠팡 빠른배송이라는 뜻)
    function readCoupangFast() {
        try {
            const root = document.querySelector('[class*="catalogSummary_"]')
                || document.getElementById('catalog_summary_info')
                || document.body;
            const lines = (root.innerText || '').split('\n').map(s => s.trim()).filter(Boolean);
            for (let i = 0; i < lines.length; i++) {
                if (!/^최저\s*[\d,]+\s*원?$/.test(lines[i])) continue;
                let hasCoupang = false, hasFast = false;
                for (let j = i + 1; j <= i + 8 && j < lines.length; j++) {
                    if (lines[j].indexOf('쿠팡') !== -1) hasCoupang = true;
                    if (lines[j].indexOf('빠른배송') !== -1) hasFast = true;
                }
                return hasCoupang && hasFast;
            }
        } catch (e) {}
        return false;
    }

    // [매입처 바로가기 링크] "최저가 사러가기"가 최저가 판매처의 상품 페이지로 간다
    function findBuyLinkHref() {
        try {
            for (const a of document.querySelectorAll('a')) {
                if ((a.innerText || '').trim() === '최저가 사러가기') return a.href || '';
            }
            for (const a of document.querySelectorAll('a')) {
                if ((a.innerText || '').trim() === '상품 바로가기') return a.href || '';
            }
        } catch (e) {}
        return '';
    }

    // [최저가 판매처 이름] 카탈로그 요약 영역의 줄 순서가
    // "최저 27,250원 / 배송비 / 무료 / G마켓" 임을 실측으로 확인했다.
    // "최저 N원" 줄 다음의 몇 줄 안에서 판매처처럼 생긴 글자를 찾는다.
    // 못 찾으면 빈칸 - 틀린 이름을 적느니 비워둔다.
    function readLowestSellerName() {
        try {
            const root = document.querySelector('[class*="catalogSummary_"]')
                || document.getElementById('catalog_summary_info')
                || document.body;
            const lines = (root.innerText || '').split('\n').map(s => s.trim()).filter(Boolean);
            for (let i = 0; i < lines.length; i++) {
                if (!/^최저\s*[\d,]+\s*원?$/.test(lines[i])) continue;
                for (let j = i + 1; j <= i + 6 && j < lines.length; j++) {
                    const t = lines[j];
                    if (!t || t.length > 20) continue;
                    if (/원$|포함|^배송비|^무료$|^유료$|^도움말$|^최저|^최대|^쿠폰|^적립|^찜|리뷰|^평점|^배송|^오늘|^내일|^[\d,\.\s]+$/.test(t)) continue;
                    return t;
                }
                break;
            }
        } catch (e) {}
        return '';
    }
    // [화장품 공식몰 확인] 가장 싼(최저가) 판매처 근처에 '공식/공식인증/공식몰/공식스토어/브랜드스토어' 표시가 있으면 공식몰.
    //  최저 판매처 줄부터 다음 '최저' 전까지의 몇 줄만 본다(딴 판매처의 공식표시에 안 속게).
    function readLowestSellerOfficial() {
        try {
            const root = document.querySelector('[class*="catalogSummary_"]')
                || document.getElementById('catalog_summary_info')
                || document.body;
            const lines = (root.innerText || '').split('\n').map(s => s.trim()).filter(Boolean);
            for (let i = 0; i < lines.length; i++) {
                if (!/^최저\s*[\d,]+\s*원?$/.test(lines[i])) continue;
                for (let j = i; j <= i + 8 && j < lines.length; j++) {
                    if (j > i && /^최저\s*[\d,]+\s*원?$/.test(lines[j])) break;   // 다음 판매처 시작 → 중단
                    if (/공식몰|공식스토어|브랜드스토어|공식인증|^공식$/.test(lines[j])) return true;
                }
                break;
            }
        } catch (e) {}
        return false;
    }

    const SELECTOR = 'button, a, li, label, div, span, [role="button"], [role="radio"], [role="tab"]';
    const hits = [];
    for (const el of document.querySelectorAll(SELECTOR)) {
        const raw = (el.innerText || "").trim();
        if (!raw || raw.length > 55) continue;
        const lines = raw.split('\n').map(s => s.trim()).filter(Boolean);
        if (lines.length > 4) continue;
        if (!lines.some(l => qtyOfLine(l) === wantQty)) continue;
        hits.push(el);
    }
    if (!hits.length) return { url: before, clicked: false, found: false, sellerSample: captureSellerSample(), sellerName: readLowestSellerName(), sellerOfficial: readLowestSellerOfficial(), buyLink: findBuyLinkHref(), catalogTitle: readCatalogTitle(), coupangFast: readCoupangFast() };

    // 가장 안쪽 요소를 누른다
    const target = hits.filter(h => !hits.some(o => o !== h && h.contains(o)))[0] || hits[0];

    // 링크라면 클릭 없이 주소만 (페이지 이동 위험 없음)
    const anchor = target.closest('a');
    if (anchor && anchor.href && anchor.href.indexOf('/catalog/') !== -1) {
        return { url: anchor.href, clicked: false, found: true, sellerSample: captureSellerSample(), sellerName: readLowestSellerName(), sellerOfficial: readLowestSellerOfficial(), buyLink: findBuyLinkHref(), catalogTitle: readCatalogTitle(), coupangFast: readCoupangFast() };
    }

    try { target.click(); } catch (e) { return { url: before, clicked: false, found: true, sellerSample: captureSellerSample(), sellerName: readLowestSellerName(), sellerOfficial: readLowestSellerOfficial(), buyLink: findBuyLinkHref(), catalogTitle: readCatalogTitle(), coupangFast: readCoupangFast() }; }

    for (let i = 0; i < 15; i++) {
        await sleep(200);
        if (location.href !== before) break;
    }
    await sleep(600);
    return { url: location.href, clicked: true, found: true, sellerSample: captureSellerSample(), sellerName: readLowestSellerName(), sellerOfficial: readLowestSellerOfficial(), buyLink: findBuyLinkHref(), catalogTitle: readCatalogTitle(), coupangFast: readCoupangFast() };
}

// ------------------------------------------------------------
// 수량 칸마다 "그 수량이 눌린 상태의 주소" 를 읽는다. (v5.0, 루크 요청)
//   ★실측(진단 파일): 수량 옵션은 <button role="radio" class="purchaseOptionFilter_btn_option…"> 라서
//     href 가 없고, 누르면 주소가 바뀐다. 12개·11개·10개 묶음이 각각 다른 주소다.
//   → 칸마다 눌러서 location.href 를 받아 적는다. 주소가 안 바뀌면 그 칸은 비워 둔다(대표 주소로 대신).
//   옵션을 다 읽고 마진까지 계산한 "뒤에" 실행한다 — 이미 확보한 가격은 잃지 않는다.
//   돌려주는 것: { 수량: 주소 } (못 받은 수량은 빠져 있다)
// ------------------------------------------------------------
async function clickQtysAndReadUrls(wantQtys) {
    const sleep = (ms) => new Promise(r => setTimeout(r, ms));
    const QTY_RE = /^(\d{1,3})(개입|개|세트|팩|박스|입|병|캔|묶음)$/;
    const SELECTOR = 'button, a, li, label, div, span, [role="button"], [role="radio"], [role="tab"]';

    function qtyOfLine(line) {
        const t = String(line || '').replace(/\([^)]*\)/g, '').replace(/\s+/g, '');
        const m = t.match(QTY_RE);
        return m ? Number(m[1]) : 0;
    }
    // clickQtyAndInspect 와 같은 규칙으로 "그 수량" 요소를 찾는다 (가장 안쪽 것)
    function findTarget(wantQty) {
        const hits = [];
        for (const el of document.querySelectorAll(SELECTOR)) {
            const raw = (el.innerText || '').trim();
            if (!raw || raw.length > 55) continue;
            const lines = raw.split('\n').map(s => s.trim()).filter(Boolean);
            if (lines.length > 4) continue;
            if (!lines.some(l => qtyOfLine(l) === wantQty)) continue;
            hits.push(el);
        }
        if (!hits.length) return null;
        return hits.filter(h => !hits.some(o => o !== h && h.contains(o)))[0] || hits[0];
    }

    const out = {};
    for (const q0 of (wantQtys || [])) {
        const q = Number(q0);
        if (!(q > 0)) continue;
        const target = findTarget(q);
        if (!target) continue;
        // 링크라면 클릭 없이 주소만
        const anchor = target.closest('a');
        if (anchor && anchor.href && anchor.href.indexOf('/catalog/') !== -1) { out[q] = anchor.href; continue; }
        const before = location.href;
        try { target.click(); } catch (e) { continue; }
        let changed = false;
        for (let i = 0; i < 20; i++) {              // 최대 3초, 보통 0.3초 안에 바뀐다
            await sleep(150);
            if (location.href !== before) { changed = true; break; }
        }
        if (changed) out[q] = location.href;
        await sleep(300 + Math.random() * 300);
    }
    return out;
}

// ------------------------------------------------------------
// [심화] 검색 결과의 카테고리 패널 읽기 (페이지 안에서 실행)
//  "식품 30,363" 처럼 이름+상품수가 붙은 catId 링크만 골라낸다.
//  패널이 접혀 있으면 "카테고리" 헤더를 눌러 펼치고 다시 읽는다.
// ------------------------------------------------------------
async function readCategoryPanel() {
    const zzz = (ms) => new Promise(r => setTimeout(r, ms));
    const items = [];
    const seenKey = new Set();
    // 검색 결과 "탭"(전체/가격비교 등)도 이름+숫자 모양이라 카테고리로 오인된다 - 제외 (2026-09-08 실측)
    const NOT_CATEGORY = ['전체', '가격비교', '네이버페이', '백화점/홈쇼핑', '쇼핑윈도', '해외직구', '브랜드패션', '핫딜'];
    const push = (name, countText, url, clickText) => {
        const count = Number(String(countText).replace(/,/g, ''));
        name = String(name || '').trim();
        if (!name || name.length > 20 || isNaN(count)) return;
        if (NOT_CATEGORY.indexOf(name) >= 0) return;
        const key = name + '#' + countText;
        if (seenKey.has(key)) return;
        seenKey.add(key);
        items.push({ name: name, count: count, url: url || '', clickText: clickText || '' });
    };

    // 접힌 목록이 있으면 필터 영역 안의 "더보기"류 버튼을 펼친다
    try {
        const ext = document.querySelectorAll('[class*="finder"] [class*="btn_extend"], [class*="Filter"] [class*="btn_extend"]');
        let n = 0;
        for (const b of ext) { if (n++ >= 3) break; try { b.click(); } catch (e) {} }
        if (ext.length) await zzz(700);
    } catch (e) {}

    // 방식 A (구형 화면): catId 링크의 글자가 "이름 숫자" 꼴
    try {
        for (const a of document.querySelectorAll('a[href*="catId="]')) {
            const text = (a.textContent || '').replace(/\s+/g, ' ').trim();
            const mm = text.match(/^(.{1,20}?)\s*([\d,]+)$/);
            if (!mm) continue;
            let abs = a.getAttribute('href') || '';
            try { abs = new URL(abs, location.href).href; } catch (e) {}
            push(mm[1], mm[2], abs, '');
        }
    } catch (e) {}

    // 방식 B (신형 화면, 2026-09-08 실측): 개수가 별도 칸에 든 목록
    //  예: <li><a>식품<span class="basicTypeFilter_num__">1,910</span></a></li>
    try {
        for (const numEl of document.querySelectorAll('[class*="Filter"] [class*="num"], [class*="finder"] [class*="num"]')) {
            const cntText = (numEl.textContent || '').trim();
            if (!/^[\d,]+$/.test(cntText)) continue;
            const item = numEl.closest('li, a, button') || numEl.parentElement;
            if (!item) continue;
            const full = (item.textContent || '').replace(/\s+/g, ' ').trim();
            const idx = full.lastIndexOf(cntText);
            if (idx <= 0) continue;
            const name = full.slice(0, idx).trim();
            let url = '';
            const link = item.tagName === 'A' ? item : item.querySelector('a[href]');
            if (link && link.getAttribute('href')) {
                try { url = new URL(link.getAttribute('href'), location.href).href; } catch (e) {}
            }
            push(name, cntText, url, (name + cntText).replace(/\s+/g, ''));
        }
    } catch (e) {}

    return items;
}

// [심화] 신형 화면에서 카테고리 항목(링크 주소가 없는 것)을 직접 눌러 들어간다
function clickCategoryItem(clickText) {
    const els = document.querySelectorAll(
        '[class*="Filter"] li, [class*="Filter"] a, [class*="Filter"] button, '
        + '[class*="finder"] li, [class*="finder"] a, [class*="finder"] button');
    for (const el of els) {
        const t = (el.textContent || '').replace(/\s+/g, '');
        if (t !== clickText) continue;
        const target = (el.tagName === 'A' || el.tagName === 'BUTTON') ? el : (el.querySelector('a, button') || el);
        try { target.click(); return true; } catch (e) { return false; }
    }
    return false;
}

// [심화] 패널을 못 읽었을 때 원인 분석용 화면 재료를 담아온다 (진단 파일용)
function captureCategoryPanelSample() {
    const out = { catIdLinks: 0, texts: [], candidates: [], sampleHtml: '' };
    try {
        const links = document.querySelectorAll('a[href*="catId="]');
        out.catIdLinks = links.length;
        out.texts = Array.from(links).slice(0, 10)
            .map(a => (a.textContent || '').replace(/\s+/g, ' ').trim().slice(0, 40));
        // "이름 숫자" 꼴 글자를 가진 요소들(태그 불문) - 패널이 링크가 아닐 때의 구조 파악용
        for (const el of document.querySelectorAll('a, button, li, span, div')) {
            if (out.candidates.length >= 20) break;
            if (el.children.length > 2) continue;
            const t = (el.textContent || '').replace(/\s+/g, ' ').trim();
            const mm = t.match(/^(.{1,15}?)\s([\d,]{3,9})$/);
            if (!mm) continue;
            out.candidates.push(el.tagName + '|' + t.slice(0, 30) + '|'
                + String(el.getAttribute('href') || el.getAttribute('data-shp-contents-id') || '').slice(0, 60));
        }
        // 필터 쪽 "카테고리" 헤더 주변 HTML (GNB의 카테고리 버튼은 제외)
        let anchor = null;
        for (const el of document.querySelectorAll('div, h3, strong, span, a, button')) {
            const t = (el.textContent || '').replace(/\s+/g, '');
            if (t !== '카테고리' && !(t.indexOf('카테고리') === 0 && t.length <= 6)) continue;
            if (el.closest('[class*="gnb"], [class*="Gnb"], [class*="GNB"], header')) continue;
            anchor = el;
            break;
        }
        let box = anchor;
        for (let i = 0; i < 5 && box && box.parentElement; i++) box = box.parentElement;
        out.sampleHtml = (box ? box.outerHTML : '').slice(0, 4000);

        // class 이름에 category/filter/finder가 든 요소 지도 - 패널의 정체 파악용
        out.classHints = [];
        const seenCls = new Set();
        for (const el of document.querySelectorAll('[class*="ategor"], [class*="ilter"], [class*="inder"]')) {
            if (out.classHints.length >= 25) break;
            const cls = String(el.className && el.className.baseVal !== undefined ? el.className.baseVal : el.className).slice(0, 70);
            if (!cls || seenCls.has(cls)) continue;
            seenCls.add(cls);
            out.classHints.push(el.tagName + '.' + cls + ' >> '
                + (el.textContent || '').replace(/\s+/g, ' ').trim().slice(0, 50));
        }
        // "카테고리" 글자를 못 찾았으면 category 계열 컨테이너라도 덤프
        if (!out.sampleHtml) {
            const cont = document.querySelector('[class*="ategor"]');
            if (cont) out.sampleHtml = cont.outerHTML.slice(0, 4000);
        }
    } catch (e) { out.err = String(e && e.message || e); }
    return out;
}

// ------------------------------------------------------------
// 꼬리물기 재료 읽기 (카탈로그 페이지 안에서 실행)
//  - 소카테고리: 요약 영역의 카테고리 경로(식품>건강식품>영양제>...)의 마지막 항목
//  - 제조사/브랜드: 제품정보의 제품속성 표. 표가 화면에 없으면 "제품정보" 탭을
//    한 번 눌러서 열고 다시 읽는다. 그래도 없으면 빈칸 (지어내지 않는다).
// ------------------------------------------------------------
async function readTailInfoOnPage() {
    const zzz = (ms) => new Promise(r => setTimeout(r, ms));
    const out = { category: '', categoryPath: '', maker: '', brand: '' };
    // 버튼·링크 문구가 값으로 잘못 읽히는 것 차단 ("상품 바로가기" 사고 방지)
    const JUNK = /바로가기|더보기|전체보기|리뷰|판매처|배송|쿠폰|혜택|찜|로그인|최저|적립|할인|\d원/;

    // 1) 소카테고리
    try {
        const scope = document.querySelector('#catalog_summary_info') || document;
        let links = Array.from(scope.querySelectorAll('a[href*="catId="]'));
        if (!links.length) links = Array.from(document.querySelectorAll('a[href*="catId="]')).slice(0, 6);
        const texts = links.map(a => (a.textContent || '').trim())
            .filter(t => t && t.length <= 20 && !JUNK.test(t));
        if (texts.length) { out.category = texts[texts.length - 1]; out.categoryPath = texts.join(' > '); }
    } catch (e) {}

    // 2) 제품속성에서 라벨("제조사"/"브랜드") 바로 옆 값 읽기
    const readAttr = (label) => {
        try {
            const nodes = document.querySelectorAll('th, dt, td, div, span');
            for (const n of nodes) {
                if ((n.textContent || '').trim() !== label) continue;
                const sib = n.nextElementSibling;
                if (sib) {
                    const v = (sib.textContent || '').trim();
                    if (v && v.length <= 30 && v !== label && !JUNK.test(v)) return v;
                }
            }
        } catch (e) {}
        return '';
    };

    out.maker = readAttr('제조사');
    out.brand = readAttr('브랜드');

    // 속성 표가 아직 안 열려 있으면 "제품정보" 탭을 눌러 열어본다
    if (!out.maker && !out.brand) {
        try {
            const tabs = document.querySelectorAll('a, button, [role="tab"], li');
            for (const t of tabs) {
                const txt = (t.textContent || '').replace(/\s+/g, '');
                if (txt === '제품정보') { t.click(); break; }
            }
            await zzz(900);
            out.maker = readAttr('제조사');
            out.brand = readAttr('브랜드');
        } catch (e) {}
    }
    // 값이 "-" 처럼 비어있다는 표시면 빈칸으로
    if (out.maker === '-') out.maker = '';
    if (out.brand === '-') out.brand = '';
    return out;
}

// ------------------------------------------------------------
// 3단계: 마진 계산
//
//  사용자의 "마진 계산기"와 똑같은 공식을 쓴다.
//    개당원가 = 묶음구매가 / 묶음수량
//    수수료   = 판매가 x 수수료율
//    순이익   = 판매가 - 수수료 - 개당원가 - 택배비
//    마진율   = 순이익 / 판매가
//    원가대비 = 순이익 / 개당원가
//  목표 순이익(T)을 남기려면 얼마에 팔아야 하는가:
//    P - P*f - c - s = T   ->   P = (T + c + s) / (1 - f)
// ------------------------------------------------------------
// ------------------------------------------------------------
// [낱개 신규] 묶음은 파는데 "1개" 옵션이 아직 없는 카탈로그.
//
//  제목 끝에 ", 1개" 를 붙여 등록하면 네이버가 1개 옵션을 새로 만들어 준다.
//  참고할 1개 최저가가 없으니 깎을 대상도 없다 — 목표순이익이 그대로 남는 값을
//  우리가 정해서 올린다.
//      판매가 = (개당원가 + 택배비 + 목표순이익) / (1 - 수수료율)
//  = computeMargin 이 이미 계산해 둔 targetPrice 와 같은 값이다.
// ------------------------------------------------------------
// 리뷰 숫자를 그 행에 적어 둔다. 판매량을 가늠하는 재료다.
//   reviewRecent : 최근 N일(기본 2주) 리뷰 개수 — 리뷰 API 로 하나씩 센 값
//   reviewMore   : 세는 상한에 닿았다 = 실제로는 더 있다 (엑셀에 '+' 로 붙는다)
//   reviewGuess  : API 가 막혀서 화면에 보이는 리뷰 속도로 '추정' 한 값
//   reviewTotal  : 그 카탈로그의 전체 리뷰 수 (오래된 것까지 전부)
//   reviewPerDay : 하루 몇 개 — 기간 기준이 달라도 서로 비교할 수 있게
function putReviewInfo(cat, rv, days) {
    cat.reviewGuess = false;
    cat.reviewMore = false;
    if (!rv) { cat.reviewRecent = ''; cat.reviewTotal = ''; cat.reviewPerDay = ''; return; }
    let n = Number(rv.recentCount);
    if (!isFinite(n) || n < 0) n = '';
    if (rv.mode === 'rate' && Number(rv.projected) > 0) {   // API 실패 → 속도 추정값
        n = Number(rv.projected);
        cat.reviewGuess = true;
    } else if (rv.capped && Number(rv.projected) > 0) {      // 상한에 닿음 → 속도로 환산한 값 ('+' 표시)
        n = Number(rv.projected);
    }
    cat.reviewRecent = n;
    cat.reviewMore = !!rv.capped;
    cat.reviewTotal = Number(rv.total) > 0 ? Number(rv.total) : '';
    cat.reviewPerDay = (n === '' || !days) ? '' : Math.round((n / days) * 10) / 10;
}

function soloTitleOf(title) {
    // 끝의 수량 표기를 떼고 ", 1개" 를 붙인다. 네이버는 제목의 수량을 보고 옵션을 만든다.
    return qtyTitleOf(title, 1);
}

function soloNewPlan(row, unitCost, money) {
    row.soloNew = false; row.soloPrice = ""; row.soloProfit = ""; row.soloRatio = "";

    if (!CONFIG.SOLO_NEW) { row.verdict = "1개 옵션 없음"; return row; }

    const price = Number(row.targetPrice) || 0;
    if (!price || !unitCost) { row.verdict = "1개 옵션 없음"; return row; }

    // [가드 1] 낱개 값이 개당원가에 비해 너무 높으면 아무도 1개로 안 산다.
    //   택배비·목표순이익이 고정이라, 개당원가가 싼 상품일수록 이 비율이 커진다.
    //   (개당 2,000원짜리는 낱개 6,400원 — 5개 10,000원을 두고 1개를 그 값에 살 사람은 없다)
    const ratio = Math.round((price / unitCost) * 100) / 100;
    row.soloRatio = ratio;
    if (ratio > CONFIG.SOLO_NEW_HARD_RATIO) { row.verdict = "낱개 비쌈"; return row; }

    // [가드 2] 1개 값이 가장 작은 묶음의 전체 값보다 비싸면 말이 안 된다.
    let minTotal = Infinity;
    for (const o of (row.cleanOptions || [])) {
        const p = Number(o.price) || 0;
        if (p > 0 && p < minTotal) minTotal = p;
    }
    if (isFinite(minTotal) && price >= minTotal) { row.verdict = "낱개 비쌈"; return row; }

    // 순이익은 일반 행과 똑같은 식으로 계산한다 — 엑셀 숫자만으로 검산해도 맞게.
    const fee = Math.round(price * money.feeRate);
    const profit = price - fee - unitCost - money.shippingFee;

    row.soloNew = true;
    row.soloPrice = price;
    row.soloProfit = profit;
    row.verdict = "낱개신규";
    // 비율이 높으면 버리진 않되, 눈으로 한 번 보라고 표시한다
    if (ratio > CONFIG.SOLO_NEW_MAX_RATIO) row.confidence = "확인필요(낱개 " + ratio + "배)";
    return row;
}

// ------------------------------------------------------------
// [빈 수량 신규] 옵션 사이에 비어 있는 수량 자리를 찾는다.
//
//  낱개신규(1개)와 같은 생각을 2개·3개… 로 넓힌 것이다.
//  카탈로그에 1개·10개만 있으면 2~9개는 **아무도 안 걸어 둔 자리**다.
//  그 자리에 걸면 깎을 상대가 없으니 판매가를 우리가 정한다.
//
//  왜 수량을 늘리면 되나: 택배비는 몇 개를 팔든 한 건만 나간다.
//  개당 차익이 800원인 상품은 1개로 팔면 택배비 3,000원을 혼자 뒤집어쓰지만,
//  6개를 묶으면 차익이 4,800원이 되어 택배비를 내고도 남는다.
//
//  판매가는 "목표순이익이 남는 최소값"과 "옆 옵션들 사이의 자연스러운 값" 중
//  큰 쪽으로 잡는다. 그리고 두 가지를 반드시 지킨다:
//    ① 더 많이 주는 윗 옵션보다 싸야 한다 (아니면 그걸 사지 이걸 안 산다)
//    ② 개당 값이 아랫 옵션의 개당 값보다 싸야 한다 (많이 살수록 싸야 말이 된다)
//  둘 중 하나라도 못 지키면 그 수량은 후보가 아니다.
// ------------------------------------------------------------
function gapQtyPlan(row, unitCost, money, clean) {
    row.gapQty = ""; row.gapPrice = ""; row.gapProfit = "";

    if (!CONFIG.GAP_QTY) return row;
    if (!unitCost || !clean || clean.length < 2) return row;

    const feeRate = money.feeRate;
    const denom = 1 - feeRate;
    if (denom <= 0) return row;

    // 수량 → 총액. 옵션이 있는 수량은 후보에서 빠진다(이미 누가 걸어 둔 자리).
    const priceOf = {};
    const qtys = [];
    for (const o of clean) {
        const q = Number(o.qty) || 0, p = Number(o.price) || 0;
        if (q > 0 && p > 0) { priceOf[q] = p; qtys.push(q); }
    }
    if (qtys.length < 2) return row;
    qtys.sort((a, b) => a - b);

    const minQ = qtys[0], maxQ = qtys[qtys.length - 1];
    const top = Math.min(maxQ - 1, CONFIG.GAP_QTY_MAX);

    // 작은 수량부터 본다 — 적게 묶을수록 사 가는 사람이 많다
    for (let n = minQ + 1; n <= top; n++) {
        if (priceOf[n]) continue;              // 이미 있는 자리

        // 바로 아래/위 옵션
        let lo = 0, hi = 0;
        for (const q of qtys) { if (q < n) lo = q; else { hi = q; break; } }
        if (!lo || !hi) continue;

        // 목표순이익이 그대로 남는 최소 판매가
        const floorPrice = Math.ceil((money.targetProfit + unitCost * n + money.shippingFee) / denom);
        // 아래~위 옵션을 잇는 선 위의 값 (카탈로그 사다리에서 가장 자연스러운 값)
        const lineP = Math.round(priceOf[lo] + (priceOf[hi] - priceOf[lo]) * ((n - lo) / (hi - lo)));

        const price = Math.max(floorPrice, lineP);

        // ① 윗 옵션보다 싸야 한다
        if (price >= priceOf[hi]) continue;
        // ② 개당 값이 아랫 옵션의 개당 값보다 싸야 한다
        if (price / n >= priceOf[lo] / lo) continue;

        // 순이익은 다른 행과 똑같은 식으로 (엑셀 숫자만으로 검산해도 맞게)
        const profit = price - Math.round(price * feeRate) - unitCost * n - money.shippingFee;
        if (profit < money.targetProfit) continue;   // 반올림 때문에 한 끗 모자랄 수 있다

        row.gapQty = n;
        row.gapPrice = price;
        row.gapProfit = profit;
        return row;
    }
    return row;
}

// 제목 끝의 수량 표기를 떼고 ", N개" 를 붙인다 (네이버는 제목의 수량을 보고 옵션을 만든다)
// ------------------------------------------------------------
// [수량 칸 전부] 1~SLOT_MAX 칸을 하나씩 따져 row.slots 에 담는다. (v4.97)
//  그룹상품으로 한 번에 올릴 재료다. 한 칸 = { qty, price, profit, basis, ref }
//    basis '경쟁' : 그 수량 칸에 남이 있다 → 그 값 - 10원. 최소 순이익 미만이면 뺀다.
//    basis '빈칸' : 아무도 없다 → 목표순이익이 그대로 남는 값. 더 큰 묶음보다 비싸면 뺀다.
//  1개 칸은 기존 계산(basePrice / 낱개신규)이 맡으므로 여기서는 2개부터 본다.
//  unitCost 를 바꿔 넣으면 다른 매입 경로(낱개 사입)로도 같은 계산이 된다.
// ------------------------------------------------------------
function slotTarget(unitCost, n, money) {
    const rate = (Number(money.slotRate) || CONFIG.SLOT_PROFIT_RATE) / 100;
    const want = Math.round(unitCost * n * rate);
    return Math.max(CONFIG.SLOT_PROFIT_MIN, Math.min(CONFIG.SLOT_PROFIT_MAX, want));
}

// 칸 계산의 알맹이. 묶음 사입(slotPlan)과 낱개 사입(순방향)이 같이 쓴다.
//   fromQty   : 몇 개 칸부터 볼지 (묶음 사입은 2, 낱개 사입은 1)
//   capBySolo : 빈 칸 값이 "1개 값 × n − 10원" 을 넘지 못하게 막을지 (낱개 사입만)
//               — 1개를 n번 사는 것보다 비싸면 아무도 안 산다. 루크: "표시 최저가 기준으로 n개 칸을 조금 싸게".
function planSlots(unitCost, money, clean, fromQty, capBySolo) {
    const slots = [];
    if (!unitCost || !clean || !clean.length) return slots;
    const feeRate = money.feeRate, ship = money.shippingFee;
    const denom = 1 - feeRate;
    if (denom <= 0) return slots;
    const priceOf = {};
    for (const o of clean) {
        const q = Number(o.qty) || 0, p = Number(o.price) || 0;
        if (q > 0 && p > 0) priceOf[q] = p;
    }
    const qtys = Object.keys(priceOf).map(Number).sort((a, b) => a - b);
    const solo = priceOf[1] || 0;
    const profitAt = (price, n) => price - Math.round(price * feeRate) - unitCost * n - ship;
    for (let n = (fromQty || 2); n <= CONFIG.SLOT_MAX; n++) {
        const target = slotTarget(unitCost, n, money);
        let price, basis, ref = '';
        if (priceOf[n]) {
            price = priceOf[n] - 10;                       // [-10원 규칙]
            basis = '경쟁'; ref = priceOf[n];
            if (profitAt(price, n) < CONFIG.SLOT_PROFIT_MIN) continue;
        } else {
            price = Math.ceil((target + unitCost * n + ship) / denom);
            basis = '빈칸';
            if (capBySolo && solo > 0 && n > 1) {
                const cap = solo * n - 10;
                if (price > cap) price = cap;              // 목표보다 덜 남더라도 1개×n 보다는 싸야 한다
            }
            // 더 많이 주는 묶음보다 비싸면 아무도 안 산다
            let dead = false;
            for (const q of qtys) if (q > n && priceOf[q] <= price) { dead = true; break; }
            if (dead) continue;
        }
        const profit = profitAt(price, n);
        if (profit < CONFIG.SLOT_PROFIT_MIN) continue;
        slots.push({ qty: n, price: price, profit: profit, basis: basis, ref: ref });
    }
    return slots;
}

function slotPlan(row, unitCost, money, clean) {
    row.slots = planSlots(unitCost, money, clean, 2, false);
    return row;
}

function qtyTitleOf(title, n) {
    const t = String(title || '').replace(/\s*,?\s*\d{1,3}\s*개\s*$/, '').trim();
    return t + ', ' + n + '개';
}

function computeMargin(row, money) {
    const shippingFee = money.shippingFee;
    const buyShipFee = Number(money.buyShipFee) || 0;   // 묶음 1건 매입 배송비
    const feeRate = money.feeRate;           // 0.067
    const targetProfit = money.targetProfit; // 1000원

    const r1 = (v) => Math.round(v * 10) / 10;
    const opts = (row.options || []).slice().sort((a, b) => a.qty - b.qty);

    // [이상값 걸러내기]
    const med = (arr) => {
        if (!arr.length) return 0;
        const s = arr.slice().sort((a, b) => a - b);
        const m = Math.floor(s.length / 2);
        return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2;
    };

    // [1차: 가격 역전 해소] 수량이 늘었는데 총액이 줄어드는 건 물리적으로 불가능하다
    // (예: 2개 50,600원인데 3개가 39,600원 = 그 옵션에 다른 구성의 상품이 잘못 매칭된 것).
    // 역전이 발견되면, 관련된 두 옵션 중 "개당단가가 나머지 옵션들의 중앙값에서
    // 더 많이 벗어난 쪽"을 잘못 매칭된 범인으로 보고 제외한다.
    // 역전이 없는 정상 데이터는 절대 건드리지 않는다.
    let work = opts.slice().sort((a, b) => a.qty - b.qty);
    const inversionDrops = [];
    const invUnits = [];   // 제외된 옵션들의 개당단가 (아래로 튄 것인지 판단용)
    const invQtys = [];    // 제외된 옵션들의 수량 (엑셀에 "10개·12개 가격이상"으로 표시용)
    let guard = 0;
    while (guard++ < 10) {
        let vi = -1, vj = -1;
        outer:
        for (let i = 0; i < work.length; i++) {
            for (let j = i + 1; j < work.length; j++) {
                if (work[j].price < work[i].price) { vi = i; vj = j; break outer; }
            }
        }
        if (vi < 0) break;

        const others = work.filter((_, k) => k !== vi && k !== vj).map(o => o.price / o.qty);
        const ref = med(others) || med(work.map(o => o.price / o.qty));
        const di = Math.abs(work[vi].price / work[vi].qty - ref);
        const dj = Math.abs(work[vj].price / work[vj].qty - ref);

        // 역전 쌍에서 싼 쪽(vj: 수량 많은데 총액이 더 싼 옵션)의 개당가가
        // 다수파의 70% 이상 = 정상 범위 안이면 -> 진짜 할인으로 보고 "비싼 슬롯(vi)"을 제외.
        // 정상 범위 밖으로 너무 싸면 -> 가짜 매칭으로 보고 중앙값에서 더 벗어난 쪽을 제외.
        const cheapUnit = work[vj].price / work[vj].qty;
        let dropIdx;
        if (ref > 0 && cheapUnit >= ref * 0.7) {
            dropIdx = vi;                       // 비싼 슬롯이 이상한 것
        } else {
            dropIdx = dj >= di ? vj : vi;       // 더 벗어난 쪽이 이상한 것
        }

        const victim = work[dropIdx];
        const other = work[dropIdx === vj ? vi : vj];
        inversionDrops.push(victim.qty + '개:' + victim.price + ' (↔ ' + other.qty + '개:' + other.price + ' 역전)');
        invUnits.push(victim.price / victim.qty);
        invQtys.push(victim.qty);
        work.splice(dropIdx, 1);
    }

    // [2차: 극단값 제거] 개당단가가 중앙값에서 크게 벗어난 것만 버린다
    const median = med(work.map(o => o.price / o.qty));
    const clean = [];
    let dropped = inversionDrops.length;
    for (const o of work) {
        const u = o.price / o.qty;
        if (median > 0 && (u > median * 2.5 || u < median * 0.35)) { dropped++; continue; }
        clean.push(o);
    }
    row.cleanOptions = clean;
    row.droppedCount = dropped;
    row.medianUnit = median ? Math.round(median) : "";

    // 역전 보정이 있었으면 진단 기록에 남긴다 (무엇을 왜 뺐는지 검증 가능하게)
    // 계산이 상품당 두 번 돌 수 있으므로, 기록은 한 번만 남긴다
    if (inversionDrops.length > 0 && row.title && !row._invLogged) {
        row._invLogged = true;
        row.confidence = "확인필요";   // 보정된 상품은 눈으로 한 번 확인 권장
        S.failures.push({
            cat: { title: '[가격역전 보정] ' + row.title, catalogUrl: row.catalogUrl, collectStatus: '잘못 매칭 의심 옵션 제외' },
            debug: { issues: '제외: ' + inversionDrops.join(' / ') }
        });
        console.warn('[가격역전 보정]', row.title, inversionDrops);
    }

    const priceOf = {};
    for (const o of clean) priceOf[o.qty] = o.price;

    // 기준 판매가: 카탈로그의 "1개 옵션가"에서 10원을 뺀 값.
    // [-10원 규칙] 최저가 자리를 잡으려면 카탈로그 최저가보다 10원 싸게 등록하므로,
    // 판매가와 순이익은 전부 (1개 옵션가 - 10원) 기준으로 계산한다. 매입가는 그대로.
    //
    // ★1개 옵션이 없을 때 "카드에 보이는 최저가"로 때우면 안 된다 (v4.91에서 고침).
    //   그 값은 묶음 가격이다. 3개 30,000원짜리 카탈로그를 "1개를 30,000원에 판다"고
    //   계산해서 순이익이 몇 배로 부풀고, 그대로 '등록가능·신뢰도 높음'이 되어
    //   등록대기 목록까지 올라갔다. ("가격이 안 맞는다"의 주된 원인)
    //   1개 옵션이 없는 카탈로그는 '1개 옵션 없음'으로 따로 빼둔다.
    const onePrice = Number(priceOf[1]) || 0;
    row.noOneOption = !onePrice;
    const basePrice = onePrice > 10 ? onePrice - 10 : onePrice;
    row.basePrice = basePrice;

    const blank = (verdict) => {
        row.bestQty = ""; row.bestTotal = ""; row.unitCost = "";
        row.feeAmount = ""; row.profit = ""; row.marginRate = ""; row.costRate = "";
        row.targetPrice = ""; row.headroom = ""; row.confidence = "";
        row.mainQty = ""; row.mainPrice = ""; row.mainProfit = "";
        row.gapQty = ""; row.gapPrice = ""; row.gapProfit = "";
        row.slots = [];
        row.verdict = verdict;
        return row;
    };

    // 낱개(1개)만 파는 카탈로그 — 카드의 수량 링크로 방문 전에 걸러낸 것 (v4.91).
    // 묶음이 없으니 사입할 것이 없다. 계산할 것도 없다.
    if (row.soloOnly) { row.basePrice = ""; return blank("낱개 전용"); }

    if (clean.length === 0) return blank("옵션 없음");

    // 가장 싼 묶음을 무조건 고른다. 1개보다 비싸도 칸을 비우지 않는다.
    // (비워두면 "못 읽은 것"처럼 보이고, 왜 아닌지도 알 수 없다)
    let best = null;
    for (const o of clean) {
        if (o.qty < 2) continue;
        const unit = o.price / o.qty;
        if (!best || unit < best.unit) best = { qty: o.qty, total: o.price, unit: unit };
    }

    if (!best) return blank("묶음 옵션 없음");

    // [신뢰도] 버리진 않되, 미심쩍으면 표시해서 직접 확인하게 한다
    let suspicious = false;
    if (median > 0 && best.unit < median * 0.75) suspicious = true;
    for (const o of clean) {
        if (o.qty < best.qty && o.price > best.total) suspicious = true;  // 더 적은 수량이 더 비쌈
    }
    // 신뢰도 표시:
    //  제외된 옵션이 최종 매입 옵션보다 "쌌던" 경우에만 가격이상으로 표시한다.
    //  (진짜 초특가라면 놓친 것이고, 잘못 걸린 매물이라면 함정이므로 확인 가치가 있음)
    //  비싼 쪽이 제외된 경우는 어차피 매입 후보가 아니므로 표시하지 않는다.
    //  괄호 안에는 "가격이 이상해서 제외한 옵션명"을 그대로 적는다 - 예: 확인필요(10개·12개 가격이상)
    //  (페이지를 열었을 때 어느 옵션을 봐야 하는지 바로 알 수 있게)
    const lowSideDrops = invUnits.filter(u => u < best.unit).length;
    if (lowSideDrops > 0) {
        const qtys = invQtys.slice().sort((a, b) => a - b);
        const shown = qtys.slice(0, 4).map(q => q + '개').join('·')
            + (qtys.length > 4 ? ' 외' + (qtys.length - 4) : '');
        row.confidence = '확인필요(' + shown + ' 가격이상)';
    } else {
        row.confidence = suspicious ? '확인필요' : '높음';
    }

    // [정수 계산] 엑셀에 "표시된 숫자"만으로 검산해도 1원도 안 틀리게,
    // 개당원가와 수수료를 먼저 반올림한 뒤 정수로 계산한다.
    // 개당원가 = (묶음 총액 + 묶음 1건 매입 배송비) ÷ 수량.
    //   매입 배송비가 빠져 있으면 묶음을 사 올 때 내는 돈이 마진에 안 잡힌다. (v4.94)
    //   기본값은 0 이라, 넣지 않으면 예전과 똑같은 숫자가 나온다.
    const unitCost = Math.round((best.total + buyShipFee) / best.qty);
    const denom = 1 - feeRate;

    // 매입 사실은 1개 판매가를 못 구해도 남긴다 — 낱개 신규를 판단할 재료가 된다 (v4.91)
    row.bestQty = best.qty;
    row.bestTotal = best.total;
    row.unitCost = unitCost;
    // 목표순이익이 그대로 남는 최저 판매가 (1개 기준).
    // 반올림한 개당원가로 계산한다 — 엑셀에 보이는 숫자만으로 검산해도 맞아야 하므로. (v4.92)
    row.targetPrice = denom > 0 ? Math.round((targetProfit + unitCost + shippingFee) / denom) : "";

    // [빈 수량 신규] 옵션 사이에 비어 있는 수량 자리를 찾는다.
    //  1개 옵션이 있든 없든 따진다 — 1개로는 안 되는 상품이 바로 이쪽의 후보다.
    gapQtyPlan(row, unitCost, money, clean);
    // [수량 칸 전부] 2~12개 칸을 전부 계산해 둔다 — 그룹상품으로 올릴 재료 (v4.97)
    slotPlan(row, unitCost, money, clean);

    // 1개 옵션이 없으면 1개 기준 순이익을 낼 수 없다. 숫자를 지어내지 않고 비워 둔다. (v4.91)
    //   묶음수량·매입총액·개당원가·목표판매가는 위에서 채워뒀으니 엑셀에서 눈으로 볼 수 있다.
    if (!basePrice) {
        row.basePrice = "";   // 0원으로 보이면 더 헷갈린다 - 빈칸으로 둔다
        row.feeAmount = ""; row.profit = ""; row.marginRate = ""; row.costRate = "";
        row.headroom = "";
        row.mainQty = ""; row.mainPrice = ""; row.mainProfit = "";
        // 묶음은 파는데 1개 옵션만 없다 → 1개를 새로 만들어 올릴 수 있는지 따진다 (v4.92)
        return soloNewPlan(row, unitCost, money);
    }

    const revenue = basePrice;                   // 1개씩 판다
    const feeAmount = Math.round(revenue * feeRate);
    const profit = revenue - feeAmount - unitCost - shippingFee;

    row.feeAmount = feeAmount;
    row.profit = profit;
    row.marginRate = r1((profit / revenue) * 100);
    row.costRate = r1((profit / unitCost) * 100);
    row.headroom = denom > 0 ? Math.round(basePrice - row.targetPrice) : "";

    // [메인 카탈로그 마진]
    // 검색 1페이지에 노출된 카드는 특정 수량(메인수량) 카탈로그다.
    // "가장 싼 묶음으로 사입해서 메인수량씩 묶어 팔 때"의 마진도 계산한다.
    //   메인 순이익 = 메인수량 판매가 - 수수료 - (개당원가 x 메인수량) - 택배비 1건
    let mainQty = 0;
    const listPrice = Number(row.lowestPrice) || 0;
    if (listPrice > 0) {
        // 카드에 표시된 가격과 일치하는 수량 옵션 찾기
        let bestDiff = Infinity, bestQ = 0;
        for (const o of clean) {
            const d = Math.abs(o.price - listPrice);
            if (d < bestDiff) { bestDiff = d; bestQ = o.qty; }
        }
        if (bestQ && bestDiff <= Math.max(50, listPrice * 0.02)) mainQty = bestQ;
    }
    if (!mainQty) {
        // 예비: 상품명 끝의 ", N개" 표기
        const tm = String(row.title || '').match(/,\s*(\d{1,3})\s*개\s*$/);
        if (tm && priceOf[Number(tm[1])]) mainQty = Number(tm[1]);
    }

    if (mainQty > 1 && priceOf[mainQty]) {
        // [-10원 규칙] 메인 판매가도 카탈로그 최저가보다 10원 싸게
        const mainPrice = priceOf[mainQty] - 10;
        // 정수 계산: 표시된 개당원가 x 수량, 수수료 반올림 후 정수 뺄셈
        const mainProfit = mainPrice - Math.round(mainPrice * feeRate) - unitCost * mainQty - shippingFee;
        row.mainQty = mainQty;
        row.mainPrice = mainPrice;
        row.mainProfit = mainProfit;
    } else {
        // 메인이 1개면 기존 계산과 같으므로 비워둔다
        row.mainQty = ""; row.mainPrice = ""; row.mainProfit = "";
    }

    // 목표 순이익을 넘겼는지로 판정한다 (마진율이 아니라 "얼마 남느냐" 기준)
    // 이 판정은 "1개 기준" 행의 판정이다. 메인수량 기준은 별도의 행으로 추가된다.
    if (profit >= targetProfit) row.verdict = "등록가능";   // 미심쩍은 건 신뢰도 칸이 "확인필요"로 알려준다
    else if (profit > 0) row.verdict = "조금 남음";
    else row.verdict = "PASS";
    return row;
}

// ------------------------------------------------------------
// [자체 검산] 계산된 행이 스스로 앞뒤가 맞는지 다른 방식으로 다시 계산해본다.
//  틀리면 신뢰도를 '확인필요'로 내리고 진단 기록에 남긴다.
// ------------------------------------------------------------
function selfCheckRow(row, money) {
    // [구멍 메우기] 1개 옵션이 없는데 판매가가 채워져 있으면,
    //   묶음 가격을 1개 가격으로 쓴 것이다. 아래 3번 검사는 "1개 옵션이 있을 때만"
    //   돌아서 하필 이 경우를 못 잡았고, 그 구멍으로 순이익이 몇 배로 부풀어
    //   등록대기 목록까지 올라갔다. computeMargin 에서 막았으니 지금은 걸릴 일이
    //   없어야 한다 — 걸리면 그 수리가 되돌아간 것이다. (v4.91)
    if (!row.soloOnly && Number(row.basePrice) > 0
        && !(row.cleanOptions || row.options || []).some(o => Number(o.qty) === 1)) {
        row.confidence = '확인필요';
        S.failures.push({
            cat: { title: '[검산] ' + row.title, catalogUrl: row.catalogUrl, collectStatus: '1개 옵션이 없는데 판매가가 있음' },
            debug: {
                issues: '판매가 ' + row.basePrice + ' / 카드최저가 ' + row.lowestPrice
                    + ' / 수량옵션 ' + (row.options || []).map(o => o.qty + '개:' + o.price).join(', ')
            }
        });
        console.warn('[검산] 1개 옵션 없는데 판매가 있음:', row.title);
    }

    const issues = [];

    // 5) 낱개 신규 행도 같은 식으로 검산한다 (v4.92)
    //    1개 기준 순이익이 없는 행이라 아래 1~4 검사를 못 타므로 여기서 먼저 본다.
    if (row.soloNew && row.soloPrice) {
        const sExpect = row.soloPrice - Math.round(row.soloPrice * money.feeRate)
            - row.unitCost - money.shippingFee;
        if (sExpect !== row.soloProfit) {
            issues.push('낱개신규 순이익 불일치 (검산 ' + sExpect + ' vs 기록 ' + row.soloProfit + ')');
        }
        if (row.soloProfit < money.targetProfit) {
            issues.push('낱개신규가 목표순이익 미달 (' + row.soloProfit + ' < ' + money.targetProfit + ')');
        }
        if (!/,\s*1개$/.test(soloTitleOf(row.title))) {
            issues.push('낱개신규 제목 끝에 ", 1개" 가 안 붙음 (' + soloTitleOf(row.title) + ')');
        }
    }

    // 6) 빈 수량 신규 행도 같은 식으로 검산한다.
    //    값을 우리가 정하는 행이라, 틀려도 비교할 최저가가 없다 — 여기서 못 잡으면 못 잡는다.
    if (row.gapQty && row.gapPrice) {
        const n = row.gapQty;
        const gExpect = row.gapPrice - Math.round(row.gapPrice * money.feeRate)
            - row.unitCost * n - money.shippingFee;
        if (gExpect !== row.gapProfit) {
            issues.push('빈수량신규 순이익 불일치 (검산 ' + gExpect + ' vs 기록 ' + row.gapProfit + ')');
        }
        if (row.gapProfit < money.targetProfit) {
            issues.push('빈수량신규가 목표순이익 미달 (' + row.gapProfit + ' < ' + money.targetProfit + ')');
        }
        // 이미 누가 걸어 둔 수량에 겹쳐 제안하면 안 된다 (빈 자리가 아니다)
        if ((row.cleanOptions || []).some(o => Number(o.qty) === n)) {
            issues.push('빈수량신규인데 그 수량 옵션이 이미 있음 (' + n + '개)');
        }
        if (!new RegExp(',\\s*' + n + '개$').test(qtyTitleOf(row.title, n))) {
            issues.push('빈수량신규 제목 끝에 ", ' + n + '개" 가 안 붙음 (' + qtyTitleOf(row.title, n) + ')');
        }
    }

    // 7) 수량 칸 전부 — 작업지시서로 그대로 올라가는 숫자다. 칸마다 다시 계산해 본다. (v4.97)
    //   빈 칸은 목표순이익이 그대로 남아야 하고(최소 1,000 · 최대 10,000),
    //   경쟁 칸은 그 수량 옵션이 실제로 있고 그 값보다 10원 싸야 한다.
    for (const s of (row.slots || [])) {
        const n = Number(s.qty), price = Number(s.price);
        if (!(n > 1) || !(price > 0) || !(row.unitCost > 0)) { issues.push('수량 칸이 비정상 (' + JSON.stringify(s) + ')'); continue; }
        const expect = price - Math.round(price * money.feeRate) - row.unitCost * n - money.shippingFee;
        if (expect !== s.profit) issues.push(n + '개 칸 순이익 불일치 (검산 ' + expect + ' vs 기록 ' + s.profit + ')');
        if (s.profit < CONFIG.SLOT_PROFIT_MIN) issues.push(n + '개 칸이 최소 순이익 미달 (' + s.profit + ')');
        const has = (row.cleanOptions || []).find(o => Number(o.qty) === n);
        if (s.basis === '경쟁') {
            if (!has) issues.push(n + '개 칸이 경쟁인데 그 옵션이 없음');
            else if (price !== Number(has.price) - 10) issues.push(n + '개 칸 경쟁가 - 10원이 아님 (' + price + ' vs ' + has.price + ')');
        } else if (s.basis === '빈칸') {
            if (has) issues.push(n + '개 칸이 빈칸인데 옵션이 이미 있음');
            // 판매가를 올림(ceil)해서 1원이 더 남을 수 있다
            if (s.profit > CONFIG.SLOT_PROFIT_MAX + 1) issues.push(n + '개 칸 순이익이 상한 초과 (' + s.profit + ')');
        } else issues.push(n + '개 칸 근거가 이상함 (' + s.basis + ')');
    }

    // 8) 낱개 사입 칸 — 실제 매입가(최대할인가 − 적립)로 같은 검산. 빈 칸은 1개×n 보다 비싸면 안 된다. (v4.98)
    if (row.fwdSlots && row.fwdSlots.length) {
        const uc = Number(row.fwdUnitCost);
        if (!(uc > 0) || (Number(row.fwdPay) + (Number(row.fwdShip) || 0) - Number(row.fwdPoints)) !== uc) {
            issues.push('낱개 매입가 불일치 (' + row.fwdPay + ' + ' + (Number(row.fwdShip) || 0) + ' − ' + row.fwdPoints + ' vs ' + uc + ')');
        }
        const one = (row.cleanOptions || []).find(o => Number(o.qty) === 1);
        const solo1 = one ? Number(one.price) : (Number(row.basePrice) > 0 ? Number(row.basePrice) : (Number(row.lowestPrice) || 0));
        for (const s of row.fwdSlots) {
            const n = Number(s.qty), price = Number(s.price);
            if (!(n >= 1) || !(price > 0)) { issues.push('낱개사입 칸이 비정상 (' + JSON.stringify(s) + ')'); continue; }
            const expect = price - Math.round(price * money.feeRate) - uc * n - money.shippingFee;
            if (expect !== s.profit) issues.push('낱개사입 ' + n + '개 칸 순이익 불일치 (검산 ' + expect + ' vs 기록 ' + s.profit + ')');
            if (s.profit < CONFIG.SLOT_PROFIT_MIN) issues.push('낱개사입 ' + n + '개 칸이 최소 순이익 미달 (' + s.profit + ')');
            if (s.basis === '빈칸' && n > 1 && solo1 > 0 && price > solo1 * n - 10) issues.push('낱개사입 ' + n + '개 칸이 1개×' + n + ' 보다 비쌈 (' + price + ')');
        }
    }

    if (row.profit === '' || row.profit === undefined) {
        reportIssues(row, issues);
        return;
    }

    // 1) 개당원가 = 매입총액 / 매입수량 인가
    if (row.bestQty > 0) {
        const unit = Math.round((row.bestTotal + (Number(money.buyShipFee) || 0)) / row.bestQty);
        if (Math.abs(unit - row.unitCost) > 1) {
            issues.push('개당원가 불일치 (계산 ' + unit + ' vs 기록 ' + row.unitCost + ')');
        }
        // 2) 순이익을 표에 보이는 숫자 그대로 다시 계산해서 비교 (정수 검산 - 오차 0이어야 함)
        const expect = row.basePrice - Math.round(row.basePrice * money.feeRate) - row.unitCost - money.shippingFee;
        if (expect !== row.profit) {
            issues.push('순이익 불일치 (검산 ' + expect + ' vs 기록 ' + row.profit + ')');
        }
    }

    // 2-2) 메인 순이익도 같은 방식으로 검산
    if (row.mainProfit !== '' && row.mainProfit !== undefined && row.mainQty > 0 && row.bestQty > 0) {
        const mExpect = row.mainPrice - Math.round(row.mainPrice * money.feeRate) - row.unitCost * row.mainQty - money.shippingFee;
        if (mExpect !== row.mainProfit) {
            issues.push('메인 순이익 불일치 (검산 ' + mExpect + ' vs 기록 ' + row.mainProfit + ')');
        }
    }

    // 3) 판매가 = 1개 옵션가 - 10원 인가 (1개 옵션이 있을 때, -10원 규칙)
    const one = (row.cleanOptions || row.options || []).find(o => o.qty === 1);
    if (one && Number(one.price) - 10 !== Number(row.basePrice)) {
        issues.push('판매가와 1개 옵션가-10원 불일치 (' + one.price + '-10 vs ' + row.basePrice + ')');
    }

    // 4) 옵션 값 자체가 정상인가
    for (const o of (row.options || [])) {
        if (!(o.qty > 0) || !(o.price > 0)) issues.push('옵션 값 이상 (' + o.qty + '개:' + o.price + ')');
    }

    reportIssues(row, issues);
}

// 검산에서 걸린 것들을 한 번에 기록한다. 신뢰도를 내리고 진단 파일에 사유를 남긴다.
function reportIssues(row, issues) {
    if (!issues.length) return;
    // 이미 '확인필요(낱개 2.3배)' 처럼 자세한 사유가 붙어 있으면 덮어쓰지 않는다
    if (String(row.confidence || '').indexOf('확인필요') !== 0) row.confidence = '확인필요';
    S.failures.push({
        cat: { title: '[검산] ' + row.title, catalogUrl: row.catalogUrl, collectStatus: '계산 검증 불일치' },
        debug: { issues: issues.join(' / ') }
    });
    console.warn('[검산 불일치]', row.title, issues);
}
// ------------------------------------------------------------
// 엑셀(.xlsx) 저장 (상품 1개 = 1행)
//  CSV 대신 진짜 엑셀 파일로 저장한다.
//  금액 칸에는 천단위 쉼표 서식(#,##0)이 들어가 있어서
//  열자마자 25,481 처럼 보인다.
// ------------------------------------------------------------
function buildSheetAoa(data) {
    const headers = [
        "판정", "신뢰도", "판매수량", "1건 순이익",
        "2주리뷰", "하루리뷰", "전체리뷰",
        "카탈로그URL", "상품명",
        "매입수량", "매입총액", "개당원가", "판매가",
        "매입옵션URL", "매입처", "매입상태", "검색키워드", "수집시각",
        "낱개사입",                                    // 순방향(네이버 1개 사입)이 왜 안 됐는지 / 됐으면 계산식 (v4.98)
        "사입길"                                       // 이 행이 어느 길인지: 묶음(역방향) / 낱개(순방향) (v4.98, 루크 요청)
    ];
    const numOr = (v) => {
        if (v === '' || v === undefined || v === null) return '';
        const n = Number(v);
        return isNaN(n) ? String(v) : n;
    };
    // 2주리뷰는 숫자로 두되, 더 있는데 상한에서 멈춘 것은 "120+",
    // API 가 막혀 추정한 것은 "34?" 로 적어 눈으로 구분되게 한다.
    const reviewCell = (r) => {
        const n = r.reviewRecent;
        if (n === '' || n === undefined || n === null) return '';
        if (r.reviewMore) return String(n) + '+';
        if (r.reviewGuess) return String(n) + '?';
        return Number(n);
    };
    const aoa = [headers];
    for (const r of data) {
        aoa.push([
            (r.sellVerdict !== undefined ? r.sellVerdict : r.verdict) || '',
            r.confidence || '',
            numOr(r.sellQty !== undefined ? r.sellQty : 1),
            numOr(r.sellProfit !== undefined ? r.sellProfit : r.profit),
            reviewCell(r),                        // 최근 2주 리뷰 (판매량 가늠용)
            numOr(r.reviewPerDay),                // 하루 몇 개
            numOr(r.reviewTotal),                 // 전체 리뷰 (오래된 것까지)
            (r.sellCatalogUrl !== undefined ? r.sellCatalogUrl : r.catalogUrl) || '',   // 그 행의 판매수량 기준 주소
            (r.sellTitle !== undefined ? r.sellTitle : r.title) || '',                  // 그 행의 판매수량 기준 제목
            numOr(r.bestQty),
            numOr(r.bestTotal),
            numOr(r.unitCost),
            numOr(r.sellPrice !== undefined ? r.sellPrice : r.basePrice),
            r.buyUrl || '',                       // 매입할 묶음이 선택된 주소 (여기서 사서 보낸다)
            r.buySeller || '',                    // 그 페이지의 최저가 판매처 (여기서 산다)
            r.buyStatus || '',                    // 판매중단 의심⚠️ 표시 (빈칸 = 문구 미발견)
            r.searchKeyword || '',
            r.scrapedAt || '',                    // 이 가격을 언제 본 것인지
            fwdCell(r),
            isFwdRow(r) ? '낱개(순방향)' : '묶음(역방향)'
        ]);
    }
    return aoa;
}

// 이 행이 낱개 사입(순방향) 행인가 — 판정이 "낱개사입·…" 로 시작한다
function isFwdRow(r) { return String(r.sellVerdict || '').indexOf('낱개사입') === 0; }

// 낱개사입 칸 글자. 낱개사입 행이면 "공식몰 OO · 최저가와 +N원", 1개 행이면 안 된 이유(또는 계산식), 나머지는 빈칸.
function fwdCell(r) {
    const v = String(r.sellVerdict || '');
    if (v.indexOf('낱개사입') === 0) {
        const gap = Number(r.fwdGap);
        return '공식몰 ' + (r.fwdSeller || '') + (isNaN(gap) ? '' : ' · 최저가와 ' + (gap >= 0 ? '+' : '') + gap.toLocaleString('ko-KR') + '원');
    }
    if (Number(r.sellQty !== undefined ? r.sellQty : 1) !== 1) return '';
    if (r.fwdStatus) return r.fwdStatus;
    if (r.fwdNote) return r.fwdNote;
    return '';
}

function makeSheet(data) {
    const aoa = buildSheetAoa(data);
    const ws = XLSX.utils.aoa_to_sheet(aoa);

    // 열 너비. 주소 열들은 링크가 아니라 "제목 글자" 기준으로 좁게 잡는다.
    ws['!cols'] = [
        { wch: 15 }, { wch: 26 }, { wch: 8 }, { wch: 10 },        // 판정/신뢰도/판매수량/순이익
        { wch: 9 }, { wch: 9 }, { wch: 9 },              // 2주리뷰/하루리뷰/전체리뷰
        { wch: 13 },                                     // 카탈로그URL (제목 폭)
        { wch: 52 },                                     // 상품명 (제목이 잘리지 않게 넉넉히)
        { wch: 8 }, { wch: 11 }, { wch: 11 }, { wch: 11 },// 매입수량/총액/원가/판매가
        { wch: 13 }, { wch: 12 }, { wch: 13 }, { wch: 20 }, { wch: 13 }, // 매입옵션URL/매입처/매입상태/키워드("(자동추가)" 포함해 넉넉히)/수집시각
        { wch: 30 },                                     // 낱개사입 (v4.98)
        { wch: 12 }                                      // 사입길
    ];

    // 금액 칸에 천단위 쉼표
    // 리뷰 3열이 끼어들면서 금액 열이 3칸씩 밀렸다 — 열을 더하면 여기도 같이 고쳐야 한다
    const MONEY_COLS = [3, 10, 11, 12];   // 1건 순이익, 매입총액, 개당원가, 판매가
    const range = XLSX.utils.decode_range(ws['!ref']);
    for (let R = 1; R <= range.e.r; R++) {
        for (const C of MONEY_COLS) {
            const cell = ws[XLSX.utils.encode_cell({ r: R, c: C })];
            if (cell && cell.t === 'n') cell.z = '#,##0';
        }
    }

    // 첫 행(제목 줄): 연한 녹색 배경 + 굵게 + 가운데 정렬
    const headerStyle = {
        fill: { patternType: 'solid', fgColor: { rgb: 'E2EFDA' } },
        font: { bold: true },
        alignment: { horizontal: 'center', vertical: 'center' }
    };
    for (let C = 0; C <= range.e.c; C++) {
        const cell = ws[XLSX.utils.encode_cell({ r: 0, c: C })];
        if (cell) cell.s = headerStyle;
    }
    return ws;
}

// 저장해 둔 결과 파일을 디스크에서 지운다(기록 비우기 때 다음사냥감 등 정리용). (v4.90)
async function eraseSavedFile(nameRegex) {
    try {
        const items = await chrome.downloads.search({ filenameRegex: nameRegex, exists: true });
        for (const it of items) {
            try { await chrome.downloads.removeFile(it.id); } catch (_) {}
            try { await chrome.downloads.erase({ id: it.id }); } catch (_) {}
        }
    } catch (_) {}
}

// 시트 여러 장을 가진 엑셀 파일 하나로 저장한다. sheets = [{ name, data }]
async function saveWorkbook(sheets, filename, overwrite) {
    const wb = XLSX.utils.book_new();
    const used = new Set();
    for (const s of sheets) {
        // 시트 이름 규칙: 31자 제한, 몇몇 특수문자 금지, 중복 금지
        let nm = String(s.name).replace(/[\\\/\?\*\[\]:]/g, '_').slice(0, 28) || '시트';
        const base = nm;
        let i = 2;
        while (used.has(nm)) nm = base + '_' + (i++);
        used.add(nm);
        let ws;
        if (s.aoa) {                                   // 단순 목록 시트 ("다음 사냥감", "카테고리 성적")
            ws = XLSX.utils.aoa_to_sheet(s.aoa);
            ws['!cols'] = s.cols || [{ wch: 44 }];
        } else {
            ws = makeSheet(s.data);
        }
        XLSX.utils.book_append_sheet(wb, ws, nm);
    }
    const b64 = XLSX.write(wb, { bookType: 'xlsx', type: 'base64' });

    // 결과는 늘 "다운로드\셀수다결과\" 밑에 종류별 폴더로 모은다
    // (filename 에 "소싱결과/..." 처럼 하위폴더를 붙여 넘긴다)
    // overwrite=true 면 같은 이름을 덮어쓴다(다음사냥감처럼 최신 1개만 남길 때).
    await chrome.downloads.download({
        url: 'data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,' + b64,
        filename: '셀수다결과/' + filename,
        conflictAction: overwrite ? 'overwrite' : 'uniquify'
    });
}

// 못 읽은 카탈로그의 진단 내용을 사람이 읽을 수 있는 텍스트로 만든다
function buildFailReport(failures) {
    const out = [];
    out.push('자체 점검 기록 (누락 의심 / 검산 불일치 / 읽기 실패)');
    out.push('생성: ' + new Date().toLocaleString('ko-KR'));
    out.push('');
    out.push('이 파일을 그대로 보내주시면 원인을 분석할 수 있습니다.');
    out.push('='.repeat(70));

    failures.forEach((f, i) => {
        const c = f.cat || {};
        const d = f.debug || {};
        out.push('');
        out.push('[' + (i + 1) + '] ' + (c.title || '(제목없음)'));
        out.push('주소: ' + (c.catalogUrl || ''));
        out.push('상태: ' + (c.collectStatus || ''));
        if (d.error) out.push('오류: ' + d.error);
        if (d.issues) out.push('검산 결과: ' + d.issues);
        if (d.cardCount !== undefined) out.push('화면 카드 ' + d.cardCount + '개 / 수집 ' + d.collected + '개');
        if (d.skipped && d.skipped.length) {
            out.push('--- 목록에서 걸러진 카드 ---');
            d.skipped.forEach(s => out.push('  ' + s.reason + ' | ' + s.nvMid + ' | ' + s.title));
        }
        if (d.missing && d.missing.length) {
            out.push('--- 페이지에는 있는데 못 담은 카탈로그 (누락 의심) ---');
            d.missing.forEach(s => out.push('  https://search.shopping.naver.com/catalog/' + s.id + ' | ' + s.text));
        }
        if (d.qtyLikeCount !== undefined) {
            out.push('수량처럼 생긴 요소 수: ' + d.qtyLikeCount);
            out.push('"수량" 글자 있음: ' + (d.hasQtyWord ? 'Y' : 'N') + '   펼치기 클릭: ' + (d.expanded || 0) + '회');
        }
        if (d.near) {
            out.push('--- 페이지에서 "수량" 주변 글자 ---');
            out.push(d.near);
        }
        if (d.sampleHtml) {
            out.push('--- 수량 요소 주변 HTML ---');
            out.push(d.sampleHtml);
        }
        out.push('-'.repeat(70));
    });

    return out.join('\n');
}

// ------------------------------------------------------------
// 작업 탭 관리 (탭 1개를 계속 재사용)
// ------------------------------------------------------------
async function ensureWorkerTab() {
    if (S.workerTabId) {
        try {
            await chrome.tabs.get(S.workerTabId);
            return S.workerTabId;
        } catch (e) { S.workerTabId = null; }
    }
    const t = await chrome.tabs.create({ url: 'about:blank', active: true });
    S.workerTabId = t.id;
    return t.id;
}

// ------------------------------------------------------------
// [꼬리물기] 등록가능 상품에서 다음 검색 키워드를 뽑는다
//
//  방법 1) 상품명 명사 조합: 메인키워드 단독 + "수식어 1개 + 메인키워드"
//          - 메인키워드: 소카테고리와 일치하는 상품명 단어 우선,
//            없으면 규격 앞의 마지막 일반 단어
//          - 제외: 브랜드/제조사명(페이지에서 읽은 실제 이름 기준),
//            맨 뒤 구매수량(1개/2박스...), 홍보어, 단위 없는 숫자
//          - 포함: 규격 수치(300ml, 60정, 7% ...), 성분/기능/형태/색상/타깃
//  방법 2) 소카테고리   3) 제조사   4) 브랜드 (영문은 사전으로 한글 변환,
//          낯선 조각이 남으면 지어내지 않고 영문 그대로 검색)
//
//  뽑은 키워드는 "(자동추가)" 꼬리표를 달고 대기줄 맨 뒤에 선다.
//  연료통: 한 번의 실행에서 자동추가는 설정 상한(기본 30개)까지만 검색하고,
//  넘치는 키워드는 버리지 않고 "다음 사냥감" 시트에 기록한다.
//  이미 검색한 키워드(3개월 규칙)와 대기줄 중복은 원형 기준으로 걸러서
//  꼬리가 깊어져도 무한 반복이 없다.
// ------------------------------------------------------------

const RESEARCH_MS_G = 30 * 24 * 60 * 60 * 1000;   // 재검색 제한 1개월 (2026-09-06, 3개월에서 단축)

// 자주 쓰는 영문 조각 -> 한글 표기 (전부 변환될 때만 한글 사용)
const ENG2KOR = {
    BIO: '바이오', PHARMA: '파마', FARMA: '파마', PHARM: '팜', FARM: '팜',
    LABS: '랩스', LAB: '랩', NATURE: '네이처', NATURAL: '내추럴',
    HEALTH: '헬스', CARE: '케어', LIFE: '라이프', PLUS: '플러스',
    GOLD: '골드', GREEN: '그린', PURE: '퓨어', NEW: '뉴', KOREA: '코리아',
    FOODS: '푸드', FOOD: '푸드', NUTRITION: '뉴트리션', NUTRI: '뉴트리',
    VITAMIN: '비타민', VITA: '비타', MEDI: '메디', TECH: '텍', WELL: '웰',
    DAILY: '데일리', ORGANIC: '오가닉', HERB: '허브', POWER: '파워',
    GLOBAL: '글로벌', WORLD: '월드', HOUSE: '하우스', HOME: '홈',
    BABY: '베이비', SKIN: '스킨', DERMA: '더마', NOW: '나우',
    SUPER: '슈퍼', MAX: '맥스', FIRST: '퍼스트', BEST: '베스트',
    ONE: '원', TOP: '탑', STAR: '스타', KING: '킹'
};
const ENG2KOR_KEYS = Object.keys(ENG2KOR).sort((a, b) => b.length - a.length);

function eng2kor(s) {
    s = String(s || '').trim();
    if (!s || /[가-힣]/.test(s)) return s;             // 이미 한글이면 그대로
    const outWords = [];
    for (const w of s.split(/\s+/)) {
        let up = w.toUpperCase().replace(/[^A-Z]/g, '');
        if (!up) return s;                              // 숫자·기호 섞임 -> 원문 그대로
        let kor = '';
        while (up.length) {
            const hit = ENG2KOR_KEYS.find(k => up.indexOf(k) === 0);
            if (!hit) return s;                         // 모르는 조각 -> 영문 그대로 검색
            kor += ENG2KOR[hit];
            up = up.slice(hit.length);
        }
        outWords.push(kor);
    }
    return outWords.join(' ');
}

// 규격 수치(300ml, 60정, 7%, 9배 ...)인지
const SPEC_RE = /^\d+([.,]\d+)?(ml|l|g|mg|kg|cc|%|배|매|cm|mm|정|포|캡슐|환|입|정입|개입|스틱|회분|일분|개월분|p|t)$/i;
const PROMO_WORDS = ['정품', '공식', '본사', '무료배송', '당일발송', '신상', '할인', '인기'];
// 페이지 버튼 문구 등이 키워드로 섞이는 것 최종 차단 (안전망 2겹)
const KW_JUNK_RE = /바로가기|사러가기|더보기|전체보기|리뷰|판매처|쿠폰|혜택|찜하기|로그인|배송|공식인증|최저가/;
// 형태·수식 단어: "무엇인지"가 아니라 "어떤 모양·등급인지"라서 메인키워드가 될 수 없다
// (예: "흑염소진액 스틱 70ml"의 메인은 스틱이 아니라 흑염소진액)
const NOT_MAIN_WORDS = ['스틱', '분말', '가루', '액상', '파우치', '젤리', '앰플', '타블렛', '골드', '프리미엄', '플러스', '오리지널', '스페셜', '대용량', '본품', '세트', '선물세트', '리필'];

function extractTailKeywords(cat) {
    const kws = [];
    const catSmall = String(cat.catSmall || '').trim();
    const brand = String(cat.brand || '').trim();
    const maker = String(cat.maker || '').trim();

    // ----- 방법 1: 상품명 명사 조합 -----
    let name = String(cat.title || '');
    name = name.replace(/,\s*\d+\s*(개|박스|세트)\s*$/, ' ');   // 맨 뒤 구매수량 제거
    name = name.replace(/[()\[\]{},+·]/g, ' ');
    const isBrandWord = (t) => {
        const x = t.toLowerCase();
        for (const b of [brand, maker]) {
            if (!b) continue;
            const bl = b.toLowerCase().replace(/\s+/g, '');
            // 브랜드명과 같거나, 브랜드명의 일부거나("나우"), 브랜드명을 품은 단어("해외나우푸드")면 제외
            if (bl && (bl === x || (x.length >= 2 && bl.indexOf(x) >= 0) || (bl.length >= 2 && x.indexOf(bl) >= 0))) return true;
        }
        return false;
    };
    const tokens = name.split(/\s+/).filter(Boolean)
        .filter(t => !/^[xX×*]$/.test(t))                       // "3g x 30포"의 x
        .filter(t => !/^\d+(개|박스|세트)$/.test(t))            // 남은 구매수량형
        .filter(t => !/^\d+([.,]\d+)?$/.test(t))                // 단위 없는 숫자
        .filter(t => PROMO_WORDS.indexOf(t) < 0)
        .filter(t => !isBrandWord(t));

    // 메인키워드 = 그 상품이 "무엇인지" 말해주는 명사 1개.
    // 띄어 쓴 단어는 각각 한 단어다 - 두 개 이상 붙여서 메인을 만들지 않는다 (2026-09-06 규칙).
    // 판단 순서 (카테고리 우선):
    //  A) 소카테고리와 정확히 같은 단어 (뒤에서부터)
    //  B) 소카테고리를 포함하거나 소카테고리에 포함되는 단어
    //  C) 규격 앞 마지막 단어 중 형태·수식 단어(스틱, 골드 등)가 아닌 것
    //  D) 그래도 없으면 규격 앞 마지막 단어
    const norm = (x) => String(x).replace(/\s+/g, '');
    const catN = norm(catSmall);
    let main = '';
    if (catN.length >= 2) {
        for (let i = tokens.length - 1; i >= 0; i--) {
            const t = tokens[i];
            if (SPEC_RE.test(t)) continue;
            if (norm(t) === catN) { main = t; break; }
        }
        if (!main) {
            for (let i = tokens.length - 1; i >= 0; i--) {
                const t = tokens[i];
                if (SPEC_RE.test(t) || NOT_MAIN_WORDS.indexOf(t) >= 0) continue;
                if (t.indexOf(catSmall) >= 0 || (t.length >= 2 && catSmall.indexOf(t) >= 0)) { main = t; break; }
            }
        }
    }
    if (!main) {
        for (let i = tokens.length - 1; i >= 0; i--) {
            const t = tokens[i];
            if (SPEC_RE.test(t) || NOT_MAIN_WORDS.indexOf(t) >= 0) continue;
            main = t; break;
        }
    }
    if (!main) {
        for (let i = tokens.length - 1; i >= 0; i--) {
            if (!SPEC_RE.test(tokens[i])) { main = tokens[i]; break; }
        }
    }
    if (main && main.length >= 2) {
        kws.push(main);                                          // 첫 항목은 메인키워드 단독
        for (const t of tokens) {
            if (t === main) continue;
            kws.push(t + ' ' + main);                            // 수식어 1개 + 메인키워드
        }
    }

    // ----- 방법 2: 소카테고리 -----
    if (catSmall) kws.push(catSmall);

    // ----- 방법 3~4: 제조사 / 브랜드 -----
    // 이 둘은 "심화" 키워드: 검색 1페이지만이 아니라 카테고리 나무를 끝까지 파고든다
    const deepKws = [];
    const mk = eng2kor(maker);
    const bd = eng2kor(brand);
    if (mk) deepKws.push(mk);
    if (bd && bd.toLowerCase() !== mk.toLowerCase()) deepKws.push(bd);   // 둘이 같으면 1개만

    // 중복 제거 + 쓰레기 문구 차단
    const seen = new Set();
    const clean = (arr) => arr.map(k => k.trim()).filter(k => {
        if (k.length < 2 || seen.has(k) || KW_JUNK_RE.test(k)) return false;
        seen.add(k);
        return true;
    });
    return { normal: clean(kws), deep: clean(deepKws) };
}

// 심화(브랜드/제조사) 키워드는 무거워서 - 카테고리 나무 전체를 훑느라
// 일반 키워드 5~10개 분량 - 한 실행에 이 개수까지만 태운다. 나머지는 사냥감으로.
const DEEP_PER_RUN = 2;

// [사냥감 탑승] 다음 사냥감에서 "한 판 = 최대 10개" 규칙으로 대기줄에 태운다.
// 수동 🎯 버튼과 완전 자동 이어가기가 둘 다 이 함수를 쓴다.
async function boardNextTargets() {
    const st = await chrome.storage.local.get('kwHistory');
    const hist = st.kwHistory || {};
    const known = new Set(S.queue.map(k => bareKw(k)));
    for (const d of S.done) known.add(bareKw(d.keyword));
    if (S.current) known.add(bareKw(S.current.keyword));
    const money = await loadMoney();
    const cap = money.autoLimit > 0 ? money.autoLimit : 10;
    let moved = 0, deepMoved = 0;
    const keep = [];
    for (const k of S.nextTargets) {
        const bk = bareKw(k);
        if (known.has(bk)) continue;
        const t = hist[bk];
        if (t && (Date.now() - t) < RESEARCH_MS_G) continue;
        const deep = isDeepKw(k);
        if (moved >= cap || (deep && deepMoved >= DEEP_PER_RUN)) { keep.push(k); continue; }
        known.add(bk);
        S.queue.push(bk + (deep ? '(자동추가·심화)' : '(자동추가)'));
        if (deep) deepMoved++;
        moved++;
    }
    S.nextTargets = keep;
    if (moved > 0) {
        S.runRound = (S.runRound || 1) + 1;   // 다음 차수의 꼬리 시작
        S.autoAdded = moved;                  // 태운 만큼 이번 판 연료를 쓴 것으로
        S.deepAdded = deepMoved;              // -> 이번 판의 새 발견은 전부 사냥감으로 모인다
        S.runStamp = '';                      // 파일 묶음 표시도 새 차수로 새로 찍는다
    }
    await persist(); broadcast();
    return moved;
}

// 뽑은 키워드를 대기줄에 태운다 (연료 남으면 큐, 넘치면 "다음 사냥감")
async function harvestTailKeywords(cat, money) {
    const found = extractTailKeywords(cat);
    const list = [];
    for (const k of found.normal) list.push({ kw: k, deep: false });
    for (const k of found.deep) list.push({ kw: k, deep: true });
    if (!list.length) return;

    const st = await chrome.storage.local.get('kwHistory');
    const hist = st.kwHistory || {};

    const known = new Set(S.queue.map(k => bareKw(k)));
    for (const d of S.done) known.add(bareKw(d.keyword));
    for (const n of S.nextTargets) known.add(bareKw(n));
    if (S.current) known.add(bareKw(S.current.keyword));

    let queued = 0, parked = 0;
    for (const it of list) {
        const bk = bareKw(it.kw);
        if (known.has(bk)) continue;
        const t = hist[bk];
        if (t && (Date.now() - t) < RESEARCH_MS_G) continue;    // 1개월 안에 이미 검색
        known.add(bk);
        const tag = it.deep ? '(자동추가·심화)' : '(자동추가)';
        if (S.autoAdded < money.autoLimit && (!it.deep || S.deepAdded < DEEP_PER_RUN)) {
            S.queue.push(bk + tag);
            S.autoAdded++;
            if (it.deep) S.deepAdded++;
            queued++;
        } else {
            S.nextTargets.push(bk + tag);   // 꼬리표를 붙여 보관 - 다음에 태울 때도 심화로 돈다
            parked++;
        }
    }
    if (queued || parked) {
        console.log('[꼬리물기]', cat.title, '-> 대기줄 ' + queued + '개, 다음 사냥감 ' + parked + '개');
        await persist();
        broadcast();
    }
}

// ------------------------------------------------------------
// 전체 결과 내보내기 (지금까지 모은 전체)
// ------------------------------------------------------------
async function exportAll() {
    if (!S.results.length) return;
    const money = await loadMoney();

    // 정렬 순위 (2026-09-04 사용자 지정):
    //  ① 등록가능 + 신뢰도 높음
    //  ② 등록가능 + 확인필요
    //  ③ 등록가능 + 판매중단 의심 (신뢰도와 무관하게, 매입 자체가 위험하므로 뒤로)
    //  ④ 새 수량 자리(낱개신규·빈수량신규) + 신뢰도 높음   ⑤ 그 외
    //  ⑥ 조금 남음   ⑦ PASS   ⑧ 낱개 전용/옵션 없음/읽기실패 등 나머지
    const orderOf = new Map();
    for (const r of S.results) {
        if (!orderOf.has(r.searchKeyword)) orderOf.set(r.searchKeyword, orderOf.size);
    }

    // 매입 자체가 위험한 것(판매중단·제품불일치·화장품 비공식)은 뒤로 밀고 등록대기에서 뺀다
    const risky = (r) => {
        const bs = String(r.buyStatus || '');
        return bs.indexOf('판매중단') >= 0 || bs.indexOf('불일치') >= 0 || bs.indexOf('공식몰 아님') >= 0;
    };
    const rankOf = (r) => {
        const v = String((r.sellVerdict !== undefined ? r.sellVerdict : r.verdict) || '');
        const sure = String(r.confidence || '').indexOf('높음') === 0;
        if (v.indexOf('등록가능') === 0) {
            if (risky(r)) return 2;
            return sure ? 0 : 1;
        }
        // 새 수량 자리(낱개신규·빈수량신규) — 등록가능 다음 자리.
        //  참고할 최저가가 없는 값이라, 등록대기에는 상자를 켠 경우에만 들어간다.
        if (v === '낱개신규' || v === '빈수량신규') return (risky(r) || !sure) ? 4 : 3;
        if (v.indexOf('낱개사입') === 0) return 3;   // 실제 매입가로 계산한 칸 — 작업지시서로만 간다 (csv 아님)
        if (v.indexOf('조금') === 0) return 5;
        if (v === 'PASS') return 6;
        return 7;
    };

    // 상품마다 판매행으로 펼친다 (1개 기준 행 + 가능하면 메인수량 기준 행)
    const expanded = expandSellRows(S.results, money.targetProfit);

    const sorted = expanded.slice().sort((a, b) => {
        const ra = rankOf(a), rb = rankOf(b);
        if (ra !== rb) return ra - rb;           // 위의 ①~⑥ 순위 먼저
        const ka = orderOf.get(a.searchKeyword) || 0;
        const kb = orderOf.get(b.searchKeyword) || 0;
        if (ka !== kb) return ka - kb;           // 같은 순위 안에서는 먼저 검색한 키워드끼리 묶음
        return sellProfitOf(b) - sellProfitOf(a);// 같은 키워드 안에서는 순이익순
    });

    // 파일 1개 안에 [전체] 시트 + 키워드별 시트를 넣는다 (키워드 시트도 같은 순위로)
    const kwOrder = [...orderOf.keys()];
    const sheets = [{ name: '전체', data: sorted }];
    // 두 길을 시트로 나눈다 (루크 요청 2026-11-02): 묶음 사입(역방향) / 낱개 사입(순방향). 전체 시트에도 "사입길" 열로 표시된다.
    sheets.push({ name: '묶음사입(역방향)', data: sorted.filter(r => !isFwdRow(r)) });
    sheets.push({ name: '낱개사입(순방향)', data: sorted.filter(r => isFwdRow(r)) });
    for (const kw of kwOrder) {
        const rows = expanded.filter(r => r.searchKeyword === kw)
            .sort((a, b) => {
                const ra = rankOf(a), rb = rankOf(b);
                if (ra !== rb) return ra - rb;
                return sellProfitOf(b) - sellProfitOf(a);
            });
        sheets.push({ name: kw, data: rows });
    }

    // [카테고리 성적] 어느 카테고리에서 실제로 건졌는지 — 다음에 어디를 팔지 고르는 근거 (v4.93)
    sheets.push({
        name: '카테고리 성적',
        aoa: catStatAoa(),
        cols: [{ wch: 40 }, { wch: 7 }, { wch: 10 }, { wch: 12 },
               { wch: 10 }, { wch: 10 }, { wch: 9 }, { wch: 12 }]
    });

    // 소싱결과는 "지금까지 찾은 전부"(S.results)가 계속 누적되는 구조라,
    // 판마다 새 파일을 만들면 앞 파일을 통째로 포함한 중복 superset이 쌓인다.
    // → 최신 전체 1개로 덮어쓴다(다음사냥감과 통일). 특정 시점 보관은 사용자가 직접 복사.
    await saveWorkbook(sheets, '소싱결과/소싱결과_최신.xlsx', true);

    // [학습기록] 이번에 새로 나온 소싱 결과를 강사님 구글시트로 전송(중복 제외) — 실패해도 소싱 계속
    try { await reportSourcingToSheet(sorted); } catch (_) {}

    // [꼬리물기] 상한을 넘겨 이번에 못 돈 키워드는 별도의 "다음사냥감" 파일로.
    // 소싱결과와 같은 차수 표시가 붙어서 "몇 차의 남은 꼬리인지" 바로 보인다.
    // (시트에 넣었더니 맨 끝 탭이라 두 번이나 못 찾음 -> 파일 분리, 2026-09-06)
    if (S.nextTargets.length) {
        const sth = await chrome.storage.local.get('kwHistory');
        const h2 = sth.kwHistory || {};
        S.nextTargets = S.nextTargets.filter(k => {
            const t = h2[bareKw(k)];
            return !(t && (Date.now() - t) < RESEARCH_MS_G);
        });
        if (S.nextTargets.length) {
            const aoa = [['다음에 검색할 키워드 (자동추가 상한을 넘겨 대기 중 - 패널의 「🎯 이어서 찾기」 버튼으로 돌릴 수 있어요)']];
            for (const k of S.nextTargets) aoa.push([k]);
            // 다음사냥감은 "지금 남은 대기줄" 스냅샷이라 최신 1개만 있으면 됨 → 같은 파일에 덮어쓴다
            await saveWorkbook([{ name: '다음 사냥감', aoa: aoa }], '다음사냥감/다음사냥감_최신.xlsx', true);
        }
    }

    // [자동등록 연계] 자동등록 프로그램이 읽어갈 "등록대기목록.csv"를
    // 항상 같은 경로에 최신 내용으로 덮어쓴다.
    //  - 담는 것: 등록가능 + 신뢰도 높음 + 판매중단 의심 없음 (= 정렬 1순위)
    //  - 확인필요/판매중단의심 상품은 자동 등록 사고를 막기 위해 뺀다 (사람이 엑셀 보고 결정)
    //  - 중복 등록 방지는 자동등록 프로그램 자체 기능에 맡긴다 (사용자 확인, 2026-09-09)
    try {
        //  - 새 수량 자리(rank 3)는 「새 수량 자리도 자동등록에 담기」 상자를 켠 경우에만 넣는다
        //    참고할 최저가가 없는 가격이라, 기본은 엑셀로만 보고 사람이 판단한다
        const cand = sorted.filter(r => String(r.sellVerdict || '').indexOf('낱개사입') !== 0
            && (rankOf(r) === 0 || (money.soloNewToCsv && rankOf(r) === 3)));
        if (cand.length) {
            const esc = (v) => '"' + String(v === undefined || v === null ? '' : v).replace(/"/g, '""') + '"';
            const lines = [['상품명', '판매수량', '판매가', '1건순이익', '카탈로그URL', '매입옵션URL', '매입처', '매입수량', '매입총액', '개당원가', '검색키워드', '수집시각'].join(',')];
            for (const r of cand) {
                lines.push([
                    (r.sellTitle !== undefined ? r.sellTitle : r.title) || '',
                    r.sellQty !== undefined ? r.sellQty : 1,
                    r.sellPrice !== undefined ? r.sellPrice : r.basePrice,
                    r.sellProfit !== undefined ? r.sellProfit : r.profit,
                    (r.sellCatalogUrl !== undefined ? r.sellCatalogUrl : r.catalogUrl) || '',
                    r.buyUrl || '',
                    r.buySeller || '',
                    r.bestQty, r.bestTotal, r.unitCost,
                    r.searchKeyword || '',
                    r.scrapedAt || ''
                ].map(esc).join(','));
            }
            await chrome.downloads.download({
                url: 'data:text/csv;charset=utf-8,' + encodeURIComponent('﻿' + lines.join('\r\n')),
                filename: '셀수다결과/등록대기.csv',   // 1개 파일이라 폴더 없이 바로 (등록기가 읽음)
                conflictAction: 'overwrite'
            });
        }
    } catch (e) { console.error('[등록대기] 저장 실패:', e); }

    // [작업지시서] 자동등록이 그대로 복사해 그룹상품으로 올릴 수 있는 자료. (v4.97)
    //  상품 하나에 "이 카탈로그를, 이 매입처에서 사서, 이 수량 칸들에 이 가격으로" 가 전부 들어 있다.
    //  가격·제목·수량은 여기서 끝나므로 자동등록은 AI 를 칸마다 쓰지 않고 그룹당 한 번(상세·사진)만 쓴다.
    //  규격(WORKORDER_SPEC)이 바뀌면 자동등록이 "어디까지 아는지" 로 선을 긋는다.
    try {
        const wo = buildWorkOrder(S.results, money, risky);
        if (wo.상품.length) {
            await chrome.downloads.download({
                url: 'data:application/json;charset=utf-8,' + encodeURIComponent(JSON.stringify(wo, null, 1)),
                filename: '셀수다결과/작업지시서.json',
                conflictAction: 'overwrite'
            });
        }
    } catch (e) { console.error('[작업지시서] 저장 실패:', e); }
}

// 작업지시서 본문. 엑셀·CSV 와 같은 숫자에서 나온다 (따로 계산하지 않는다).
//  risky(r) = 매입 자체가 위험한 것(판매중단·불일치·공식몰 아님) — 빼는 규칙은 등록대기와 같다.
function buildWorkOrder(results, money, risky) {
    const MIN = CONFIG.SLOT_PROFIT_MIN;
    let ver = '';
    try { ver = chrome.runtime.getManifest().version; } catch (_) {}
    const items = [];
    const oneTitleOf = (r) => (r.linkQty === 1 && r.optionTitle) ? r.optionTitle : qtyTitleOf(r.title, 1);
    const mk = (r, buy, slots) => {
        const cid = (String(r.catalogUrl || '').match(/catalog\/(\d+)/) || [])[1] || '';
        return {
            상품명: r.title || '',
            카탈로그URL: r.catalogUrl || '',
            카탈로그번호: cid,
            카테고리: r.catPath || r.catSmall || '',
            신뢰도: r.confidence || '',
            리뷰: { 이주: r.reviewRecent === undefined ? '' : r.reviewRecent, 전체: r.reviewTotal === undefined ? '' : r.reviewTotal },
            매입: buy,
            그룹: { 판매옵션: '수량', 대표수량: slots[0].수량 },
            칸: slots,
            검색키워드: r.searchKeyword || '',
            수집시각: r.scrapedAt || ''
        };
    };
    for (const r of results) {
        if (!r) continue;

        // ── 묶음 사입(역방향): 매입처에서 묶음을 사서 나눠 판다 ──
        if (!risky(r)) {
            const slots = [];
            // 1개 칸: 남이 있으면 기존 계산(basePrice), 없으면 낱개신규. 최소 순이익 미만이면 안 올린다.
            // 주소: 남이 있는 칸은 그 수량을 눌렀을 때의 주소, 빈 칸은 대표 주소 (v5.0)
            if (r.soloNew && Number(r.soloProfit) >= MIN) {
                slots.push({ 수량: 1, 제목: soloTitleOf(r.title), 판매가: r.soloPrice, 순이익: r.soloProfit, 근거: '빈칸', 경쟁가: '', 주소: r.catalogUrl || '' });
            } else if (String(r.verdict) === '등록가능' && Number(r.profit) >= MIN) {
                slots.push({ 수량: 1, 제목: oneTitleOf(r), 판매가: r.basePrice, 순이익: r.profit, 근거: '경쟁', 경쟁가: r.basePrice + 10, 주소: qtyUrlOf(r, 1) });
            }
            for (const s of (r.slots || [])) {
                const isMain = (s.basis === '경쟁' && Number(r.mainQty) === s.qty);
                slots.push({ 수량: s.qty, 제목: isMain ? (r.title || '') : qtyTitleOf(r.title, s.qty),
                             판매가: s.price, 순이익: s.profit, 근거: s.basis, 경쟁가: s.ref || '',
                             주소: s.basis === '경쟁' ? qtyUrlOf(r, s.qty) : (r.catalogUrl || '') });
            }
            if (slots.length) items.push(mk(r, {
                경로: '묶음', 매입처: r.buySeller || '', 옵션URL: r.buyUrl || '',
                수량: r.bestQty, 총액: r.bestTotal, 개당원가: r.unitCost, 상태: r.buyStatus || ''
            }, slots));
        }

        // ── 낱개 사입(순방향): 네이버 1개 최저가를 실제 매입가로 사서 칸에 판다 (규격 1.1) ──
        const f = r.fwdSlots || [];
        if (f.length && !r.notOfficial && !r.fwdStatus) {
            const slots2 = f.map(s => ({
                수량: s.qty, 제목: s.qty === 1 ? oneTitleOf(r) : qtyTitleOf(r.title, s.qty),
                판매가: s.price, 순이익: s.profit, 근거: s.basis, 경쟁가: s.ref || '',
                주소: s.basis === '경쟁' ? qtyUrlOf(r, s.qty) : (r.catalogUrl || '')
            }));
            items.push(mk(r, {
                경로: '낱개', 매입처: r.fwdSeller || '', 옵션URL: r.fwdLink || r.soloBuyLink || '',
                수량: 1, 총액: r.fwdUnitCost, 개당원가: r.fwdUnitCost,
                최대할인가: r.fwdPay, 배송비: Number(r.fwdShip) || 0, 적립합계: r.fwdPoints, 적립: r.fwdPointDetail || {},
                공식몰최저가차이: r.fwdGap === undefined ? '' : r.fwdGap, 상태: ''
            }, slots2));
        }
    }
    return {
        규격: CONFIG.WORKORDER_SPEC,
        소싱버전: ver,
        만든시각: new Date().toISOString(),
        설정: { 수수료퍼센트: Math.round(money.feeRate * 1000) / 10, 택배비: money.shippingFee,
                칸비율퍼센트: Number(money.slotRate) || CONFIG.SLOT_PROFIT_RATE,
                최소순이익: CONFIG.SLOT_PROFIT_MIN, 최대순이익: CONFIG.SLOT_PROFIT_MAX,
                리뷰기준2주: Number(money.reviewMin) || CONFIG.REVIEW_MIN_COUNT },
        상품: items
    };
}

// 진단 파일 저장 (키워드 1개짜리 실행에서도 저장된다)
async function exportFailReport() {
    if (!CONFIG.SAVE_FAIL_REPORT || S.failures.length === 0) return;
    try {
        const text = buildFailReport(S.failures);
        // 진단은 결과 폴더 안의 "진단" 하위 폴더에, 소싱결과와 같은 차수 표시를 달고 저장
        await chrome.downloads.download({
            url: 'data:text/plain;charset=utf-8,' + encodeURIComponent('﻿' + text),
            filename: '셀수다결과/진단/진단_점검기록_' + new Date().toISOString().slice(0, 10)
                + (S.runStamp ? ' (' + S.runStamp + ')' : '') + '.txt',
            conflictAction: 'uniquify'
        });
        // 파일로 내보냈으니 비운다 - 다음 파일에 옛 기록이 중복으로 쌓이지 않게
        S.failures = [];
        await persist();
    } catch (e) { console.error('[진단 파일] 저장 실패:', e); }
}

// ------------------------------------------------------------
// 키워드 1개 처리 (검색 -> 대기 목록에 저장 -> 카탈로그 방문)
//
//  찾은 카탈로그는 방문하기 전에 먼저 "보류 목록(S.pending)"에 저장한다.
//  중지하거나 사고가 나도 목록이 남아 있어서, 다음 시작 때 이어서 본다.
//  같은 키워드를 다시 넣으면 이미 본 카탈로그는 건너뛰고 새 것만 본다.
// ------------------------------------------------------------
async function processKeyword(keyword, money) {
    S.current = { keyword: keyword, phase: '검색 중', catalogIndex: 0, catalogTotal: 0, collected: 0 };
    S.note = '';
    broadcast();

    const tabId = await ensureWorkerTab();

    // ----- 검색 결과에서 카탈로그 목록 -----
    await chrome.tabs.update(tabId, { url: searchUrlBase + encodeURIComponent(bareKw(keyword)) });
    await waitForLoad(tabId, null);
    // 🚫 검색 결과가 차단 화면이면 쉬었다가 다시 연다 (최대 3번)
    for (let bt = 0; bt < 3; bt++) {
        const bk = await isBlockedTab(tabId);
        if (!bk.blocked) { blockRested(); break; }
        const ok = await blockPause(tabId, '검색', bk);
        if (!ok) return;
        await chrome.tabs.update(tabId, { url: searchUrlBase + encodeURIComponent(bareKw(keyword)) });
        await waitForLoad(tabId, null);
    }
    await updateModal(tabId, {
        keyword: keyword, big: '···',
        phase: '카탈로그 목록을 모으는 중', sub: '검색 결과를 훑어보고 있어요'
    });

    // 개수가 더 늘지 않으면 그게 이 페이지의 전부다 - 헛돌지 않고 바로 넘어간다.
    let rows = [];
    let listInfo = null;
    let prevCount = -1;
    for (let attempt = 1; attempt <= 3; attempt++) {
        if (S.stopRequested) break;
        try {
            await chrome.scripting.executeScript({ target: { tabId: tabId }, func: scrollPageToLoadAll });
        } catch (e) { console.debug('[스크롤] 실패:', e.message); }

        try {
            const out = await chrome.scripting.executeScript({ target: { tabId: tabId }, func: extractDataFromPage });
            if (out && out[0] && out[0].result) {
                listInfo = out[0].result;
                rows = listInfo.rows;
            }
        } catch (e) { console.error('[목록 추출] 실패:', e); }

        if (rows.length >= 35) break;                          // 충분히 모였다
        if (rows.length > 0 && rows.length === prevCount) break; // 더 안 늘어난다 = 이게 전부
        prevCount = rows.length;

        if (attempt < 3) await restSleep(rand(1200, 2200));
    }

    // [자체 점검] 카드에서 걸러졌거나 링크 대조에서 빠진 게 있으면 진단 기록에 남긴다
    const hasSkip = listInfo && listInfo.skipped && listInfo.skipped.length > 0;
    const hasMiss = listInfo && listInfo.missing && listInfo.missing.length > 0;
    if (hasSkip || hasMiss) {
        S.failures.push({
            cat: {
                title: '[목록 점검] "' + keyword + '" 카드 ' + listInfo.cardCount + '개 중 ' + rows.length + '개 수집',
                catalogUrl: searchUrlBase + encodeURIComponent(bareKw(keyword)),
                collectStatus: '목록 누락 의심'
            },
            debug: {
                cardCount: listInfo.cardCount,
                collected: rows.length,
                skipped: listInfo.skipped || [],
                missing: listInfo.missing || []
            }
        });
        console.warn('[목록 점검] ' + keyword + ':', listInfo.skipped, listInfo.missing);
    }

    const catalogs = rows.slice(0, CONFIG.MAX_CATALOGS_PER_KEYWORD)
        .map(r => Object.assign({}, r, { searchKeyword: keyword }));

    // 검색이 실제로 됐으면 날짜를 기록한다 (3개월 재검색 제한의 기준)
    if (catalogs.length > 0) await recordKeywordSearched(keyword);

    // 이미 본 것(결과에 있음)과 이미 보류 중인 것은 다시 넣지 않는다.
    // [전역 중복 검사] 키워드가 달라도 같은 카탈로그면 건너뛴다 - 꼬리물기가
    // 만드는 키워드는 1페이지가 서로 많이 겹쳐서, 키워드별로만 걸렀다간
    // 같은 상품을 수십 번 재방문하고 엑셀에도 중복 행이 쌓인다.
    // 카탈로그의 계산 결과는 어느 키워드로 찾았든 같으므로 잃는 정보가 없다.
    const seen = new Set();
    for (const r of S.results) if (r.nvMid) seen.add(r.nvMid);
    for (const p of S.pending) if (p.nvMid) seen.add(p.nvMid);
    const fresh = catalogs.filter(c => !(c.nvMid && seen.has(c.nvMid)));
    const skippedOld = catalogs.length - fresh.length;

    if (fresh.length === 0) {
        S.current = null;
        await updateModal(tabId, {
            keyword: keyword, big: '0',
            phase: '"' + keyword + '" 새로 볼 카탈로그가 없어요',
            sub: skippedOld > 0 ? '이미 조사한 카탈로그 ' + skippedOld + '개는 건너뛰었어요' : '검색 결과가 없어요',
            final: !S.queue.length && !S.pending.length && !isDeepKw(keyword)
        });
        broadcast();
    } else {
        // 방문하기 전에 먼저 저장 - 중지돼도 잃어버리지 않는다
        S.pending.push(...fresh);
        await persist();
        await visitCatalogGroup(keyword, fresh, money, false);
    }

    // [브랜드/제조사 심화] "(자동추가·심화)" 키워드는 1페이지로 끝내지 않고
    // 카테고리 나무를 끝까지 파고든다 (1페이지에 새 상품이 없었어도 실행)
    if (isDeepKw(keyword) && !S.stopRequested) {
        try { await deepScanCategories(tabId, keyword, money); }
        catch (e) { console.error('[심화] 실패:', keyword, e); }
    }
}

// ------------------------------------------------------------
// [브랜드/제조사 심화] 카테고리 나무 훑기 (2026-09-08 사용자 규칙)
//
//  검색 결과의 왼쪽 "카테고리" 패널에서
//   - 상품 1,000개 이상인 카테고리만 들어간다
//   - 화면 표시의 역순(뒤에서 앞)으로 들어간다
//   - 들어가서 뜬 1페이지 카탈로그는 무조건 전부 조사한다
//   - 하위 카테고리도 같은 규칙으로 끝까지 파고든다
//   - 이미 본 카탈로그(전역 nvMid)는 건너뛴다
//  스택(pop) 구조라 "뒤에서 앞" 순서가 저절로 지켜진다.
// ------------------------------------------------------------
const DEEP_MIN_COUNT = 1000;   // 이 개수 이상인 카테고리만 파고든다
const DEEP_MAX_NODES = 25;     // 폭주 방지: 키워드 하나당 최대 방문 카테고리 수

async function deepScanCategories(tabId, keyword, money) {
    const visitedCat = new Set();
    const stack = [];

    const readPanel = async () => {
        // 패널이 iframe 안에 그려지는 경우가 있어 모든 프레임을 뒤진다 (2026-09-08 실측)
        try {
            const out = await chrome.scripting.executeScript({
                target: { tabId: tabId, allFrames: true }, func: readCategoryPanel
            });
            const merged = [];
            for (const fr of (out || [])) {
                if (fr && fr.result && fr.result.length) merged.push(...fr.result);
            }
            return merged;
        } catch (e) { return []; }
    };
    const keyOf = (it) => it.url ? it.url : (it.parentUrl || '') + '|' + it.name + '#' + it.count;
    const pushChildren = (items) => {
        for (const it of items) {                        // 표시 순서대로 쌓으면
            if (it.count < DEEP_MIN_COUNT) continue;     // 꺼낼 때 마지막 것부터 = 뒤에서 앞
            const key = keyOf(it);
            if (visitedCat.has(key)) continue;
            visitedCat.add(key);
            stack.push(it);
        }
    };

    // 검색 1페이지로 돌아가 대분류 목록부터 읽는다
    const rootUrl = searchUrlBase + encodeURIComponent(bareKw(keyword));
    await chrome.tabs.update(tabId, { url: rootUrl });
    await waitForLoad(tabId, null, 15);
    await restSleep(rand(1500, 2500));
    const rootKids = await readPanel();
    for (const it of rootKids) it.parentUrl = rootUrl;
    pushChildren(rootKids);

    // 패널을 못 읽었으면 조용히 넘어가지 말고, 화면 재료를 진단에 남기고 알린다
    if (!stack.length) {
        try {
            const cap = await chrome.scripting.executeScript({
                target: { tabId: tabId, allFrames: true }, func: captureCategoryPanelSample
            });
            const c0 = { catIdLinks: 0, texts: [], candidates: [], classHints: [], sampleHtml: '', frames: 0 };
            for (const fr of (cap || [])) {
                const r = fr && fr.result;
                if (!r) continue;
                c0.frames++;
                c0.catIdLinks += (r.catIdLinks || 0);
                c0.texts = c0.texts.concat(r.texts || []).slice(0, 10);
                c0.candidates = c0.candidates.concat(r.candidates || []).slice(0, 20);
                c0.classHints = c0.classHints.concat(r.classHints || []).slice(0, 25);
                if ((r.sampleHtml || '').length > c0.sampleHtml.length) c0.sampleHtml = r.sampleHtml;
            }
            S.failures.push({
                cat: {
                    title: '[심화 실패] "' + bareKw(keyword) + '" 카테고리 패널을 읽지 못함',
                    catalogUrl: searchUrlBase + encodeURIComponent(bareKw(keyword)),
                    collectStatus: '심화 조사 불가'
                },
                debug: {
                    issues: '프레임 ' + c0.frames + '개, catId 링크 ' + c0.catIdLinks + '개, 링크 글자 예시: ' + c0.texts.join(' | ')
                        + ' / 이름+숫자 후보: ' + c0.candidates.join(' ; ')
                        + ' / class 지도: ' + c0.classHints.join(' ;; '),
                    sampleHtml: c0.sampleHtml
                }
            });
            await persist();
        } catch (e) {}
        await updateModal(tabId, {
            keyword: keyword, big: '⚠️',
            phase: '카테고리 목록을 읽지 못했어요',
            sub: '진단 기록을 남겼어요 - 이 키워드는 1페이지만 조사하고 넘어갑니다'
        });
        await restSleep(2000);
        return;
    }

    let nodes = 0;
    while (stack.length && !S.stopRequested && nodes < DEEP_MAX_NODES) {
        const node = stack.pop();
        nodes++;
        await updateModal(tabId, {
            keyword: keyword, big: '🗂️',
            phase: '카테고리 심화 조사 중 (' + node.name + ')',
            sub: '상품 ' + node.count.toLocaleString() + '개 카테고리 · ' + nodes + '번째'
        });

        if (node.url) {
            // 링크 주소가 있는 카테고리(구형 화면)는 주소로 바로 이동
            await chrome.tabs.update(tabId, { url: node.url });
            await waitForLoad(tabId, null, 15);
            await restSleep(rand(1500, 2500));
        } else {
            // 신형 화면: 부모 페이지로 가서 항목을 직접 눌러 들어간다
            await chrome.tabs.update(tabId, { url: node.parentUrl });
            await waitForLoad(tabId, null, 15);
            await restSleep(rand(1500, 2500));
            let clicked = false;
            try {
                const r = await chrome.scripting.executeScript({
                    target: { tabId: tabId, allFrames: true }, func: clickCategoryItem, args: [node.clickText]
                });
                clicked = (r || []).some(fr => fr && fr.result);
            } catch (e) {}
            if (!clicked) {
                console.warn('[심화] 카테고리 항목 클릭 실패:', node.name);
                continue;
            }
            await restSleep(rand(2000, 3000));   // 화면 갱신 대기
        }
        // 지금 페이지의 실제 주소를 기억해 하위 탐색의 기준으로 쓴다
        let hereUrl = node.url || node.parentUrl;
        try {
            const t = await chrome.tabs.get(tabId);
            if (t && t.url) { node.url = t.url; hereUrl = t.url; }
        } catch (e) {}

        // 하위 카테고리 목록을 페이지를 떠나기 전에 먼저 읽어 쌓는다
        const kids = await readPanel();
        for (const it of kids) it.parentUrl = hereUrl;
        pushChildren(kids);

        // 이 카테고리의 1페이지 카탈로그 수집
        let rows = [];
        for (let attempt = 1; attempt <= 2; attempt++) {
            try {
                await chrome.scripting.executeScript({ target: { tabId: tabId }, func: scrollPageToLoadAll });
            } catch (e) {}
            try {
                const out = await chrome.scripting.executeScript({ target: { tabId: tabId }, func: extractDataFromPage });
                if (out && out[0] && out[0].result) rows = out[0].result.rows;
            } catch (e) {}
            if (rows.length >= 35) break;
            if (attempt < 2) await restSleep(rand(1200, 2200));
        }

        const catalogs = rows.slice(0, CONFIG.MAX_CATALOGS_PER_KEYWORD)
            .map(r => Object.assign({}, r, { searchKeyword: keyword }));
        const seen = new Set();
        for (const r of S.results) if (r.nvMid) seen.add(r.nvMid);
        for (const p of S.pending) if (p.nvMid) seen.add(p.nvMid);
        const fresh = catalogs.filter(c => !(c.nvMid && seen.has(c.nvMid)));

        if (fresh.length) {
            S.pending.push(...fresh);
            await persist();
            await visitCatalogGroup(keyword, fresh, money, false);
        }
        await restSleep(rand(CONFIG.CATALOG_DELAY_MIN, CONFIG.CATALOG_DELAY_MAX));
    }

    if (stack.length && nodes >= DEEP_MAX_NODES) {
        console.warn('[심화] 카테고리가 많아 ' + DEEP_MAX_NODES + '개까지만 조사:', keyword);
    }
}

// ------------------------------------------------------------
// [낱개 사입(순방향)] 카탈로그 하나에 대해: 1개 최저가 판매처 페이지로 가서
//   실제 매입가 = 최대할인가 − 적립(구매·멤버십·현대카드) 을 읽고, 그 원가로 1~12개 칸을 계산한다. (v4.98)
//   스마트스토어/브랜드스토어가 아니면 못 읽는다(다른 몰의 혜택 창은 모른다) — 상태만 적고 넘어간다.
//   화장품은 1개 판매처도 공식이어야 한다 (묶음 사입과 같은 규칙).
//   결과: cat.fwdUnitCost(실제 매입가) · cat.fwdSlots(칸) · cat.fwdStatus(못 한 이유, 빈칸이면 성공)
// ------------------------------------------------------------
// [낱개 사입] 카탈로그 페이지(1개 옵션 상태)에서 공식몰 판매처를 읽어 cat.official 에 둔다. 첫 카탈로그는 진단에 샘플을 남긴다.
async function captureOfficialSeller(tabId, cat, isFirst) {
    try {
        const o = await chrome.scripting.executeScript({ target: { tabId: tabId }, func: readOfficialSellerOnPage, args: [String(cat.brand || cat.maker || '')] });
        const r = o && o[0] && o[0].result;
        cat.official = (r && r.found) ? { name: r.name, price: r.price, ship: Number(r.ship) || 0, maxDiscount: r.maxDiscount, link: r.link } : null;
        if (isFirst) S.failures.push({
            cat: { title: '[공식몰 샘플] ' + cat.title, catalogUrl: cat.catalogUrl, collectStatus: '참고자료' },
            debug: { sampleHtml: (r && r.found ? '찾음: ' + JSON.stringify(cat.official) : '못 찾음') + '\n' + ((r && r.sample) || '') }
        });
    } catch (e) { cat.official = null; }
}

async function forwardBuyStep(tabId, cat, money, isFirst) {
    cat.fwdSlots = []; cat.fwdStatus = '';
    // basePrice 는 "경쟁가 − 10" 이라, 카탈로그의 실제 1개 최저가는 +10 이다 (실측: 15,490 ↔ 15,500)
    const price1 = Number(cat.basePrice) > 0 ? Number(cat.basePrice) + 10 : (Number(cat.lowestPrice) || 0);
    if (!(price1 > 0)) { cat.fwdStatus = '1개 가격 없음'; return; }
    // [공식몰 규칙(루크)] 공식몰이 있어야 하고, 최저가이거나 최저가와 1,000원 미만 차이여야 한다. 사는 곳도 그 공식몰이다.
    const off = cat.official;
    if (!off || !off.link) { cat.fwdStatus = '공식몰 없음'; return; }
    const offTotal = Number(off.price) + (Number(off.ship) || 0);
    // 카탈로그 최저가(배송비 포함)보다 15% 넘게 싼 "공식몰" 은 이 카탈로그 상품이 아니다 (실측: 25kg 다시다에 6,100원짜리 딴 상품 카드)
    if (offTotal < price1 * 0.85) { cat.fwdStatus = '공식몰 카드가 딴 상품(값 ' + offTotal.toLocaleString('ko-KR') + ' vs 최저 ' + price1.toLocaleString('ko-KR') + ' · ' + off.name + ')'; return; }
    const gap = offTotal - price1;   // 카탈로그 최저가는 배송비 포함 → 카드 값도 배송비를 더해 비교
    if (gap >= CONFIG.FORWARD_OFFICIAL_GAP) { cat.fwdStatus = '공식몰이 최저가 아님(+' + gap.toLocaleString('ko-KR') + '원 · ' + off.name + ')'; return; }
    cat.soloBuyLink = off.link;
    cat.soloSeller = off.name;
    cat.fwdGap = gap;

    await updateModal(tabId, {
        keyword: cat.searchKeyword || '', big: '🛍️',
        phase: '낱개 매입가 확인 중',
        sub: (cat.soloSeller ? cat.soloSeller + ' · ' : '') + String(cat.title || '').slice(0, 30)
    });
    await chrome.tabs.update(tabId, { url: cat.soloBuyLink });
    await waitForLoad(tabId, null, 15);
    await restSleep(rand(1500, 2500));
    let url = '';
    try { url = String((await chrome.tabs.get(tabId)).url || ''); } catch (_) {}
    cat.fwdLink = url || cat.soloBuyLink;
    // 🚫 차단 화면이면 쉬었다가 한 번 더 열어 본다
    {
        const bk = await isBlockedTab(tabId);
        if (bk.blocked) {
            const ok = await blockPause(tabId, '공식몰 페이지', bk);
            if (!ok) { cat.fwdStatus = '중지'; return; }
            await chrome.tabs.update(tabId, { url: cat.soloBuyLink });
            await waitForLoad(tabId, null, 15);
            await restSleep(rand(1500, 2500));
            const bk2 = await isBlockedTab(tabId);
            if (bk2.blocked) { cat.fwdStatus = '차단으로 못 읽음'; return; }
        }
    }
    if (!/smartstore\.naver\.com|brand\.naver\.com/.test(url)) {
        cat.fwdStatus = '스마트스토어 아님(' + ((url.match(/^https?:\/\/([^\/]+)/) || [])[1] || '?') + ')';
        return;
    }

    const out = await chrome.scripting.executeScript({ target: { tabId: tabId }, func: readStoreBenefitOnPage });
    const b = out && out[0] && out[0].result;
    // 첫 상품의 혜택 창 글자를 진단에 남긴다 — 네이버가 창을 바꾸면 여기서 보인다
    if (!S.fwdSampleDone && b && (b.modalText || b.ok)) S.failures.push({
        cat: { title: '[적립 모달 샘플] ' + cat.title, catalogUrl: cat.catalogUrl, collectStatus: '참고자료' },
        debug: { sampleHtml: 'url=' + url + '\n읽은 것: ' + JSON.stringify({ ok: b.ok, reason: b.reason, finalPrice: b.finalPrice, shipping: b.shipping, maxDiscount: b.maxDiscount, points: b.points, basis: b.priceBasis })
                 + '\n' + (b.modalText || '') }
    });
    if (b && (b.modalText || b.ok)) S.fwdSampleDone = true;
    if (!b || !b.ok) { cat.fwdStatus = '혜택 창 못 읽음(' + ((b && b.reason) || '?') + ')'; return; }

    // 화장품은 1개 판매처도 공식이어야 한다
    if (/화장품|미용/.test(String(cat.catPath || cat.catSmall || ''))) {
        const seller = String(cat.soloSeller || b.storeName || '');
        const normSN = (x) => String(x || '').toLowerCase().replace(/[\s\-_·.]/g, '');
        const bn = normSN(cat.brand), sn = normSN(seller);
        const official = cat.buyOfficial === true || /공식/.test(seller) || (bn.length >= 2 && sn.includes(bn)) || /brand\.naver\.com/.test(url);
        if (!official) { cat.fwdStatus = '1개 판매처가 공식 아님(화장품)(' + (seller.slice(0, 16) || '판매처미상') + ')'; return; }
    }

    const p = b.points || {};
    const ptSum = (Number(p.basic) || 0) + (Number(p.extra) || 0) + (Number(p.membership) || 0) + (Number(p.card) || 0);
    const payPrice = Number(b.finalPrice) || 0;
    const ship = Number(b.shipping) || 0;
    const unitCost = payPrice + ship - ptSum;
    if (!(unitCost > 0)) { cat.fwdStatus = '매입가 이상(' + payPrice + '+' + ship + '−' + ptSum + ')'; return; }
    cat.fwdPay = payPrice;
    cat.fwdShip = ship;
    cat.fwdPoints = ptSum;
    cat.fwdPointDetail = { 기본적립: Number(p.basic) || 0, 추가구매적립: Number(p.extra) || 0, 멤버십추가적립: Number(p.membership) || 0, 현대카드: Number(p.card) || 0 };
    cat.fwdUnitCost = unitCost;
    cat.fwdSeller = cat.soloSeller || b.storeName || '';
    cat.fwdBasis = b.priceBasis || '';
    cat.fwdNote = '최대할인가 ' + payPrice.toLocaleString('ko-KR') + (ship ? ' + 배송비 ' + ship.toLocaleString('ko-KR') : '') + ' − 적립 ' + ptSum.toLocaleString('ko-KR') + ' = ' + unitCost.toLocaleString('ko-KR');
    const clean = (cat.cleanOptions && cat.cleanOptions.length) ? cat.cleanOptions : [{ qty: 1, price: price1 }];
    cat.fwdSlots = planSlots(unitCost, money, clean, 1, true);
    if (!cat.fwdSlots.length) cat.fwdStatus = '남는 칸 없음';
    console.log('[낱개사입]', cat.title, cat.fwdNote, '· 칸', cat.fwdSlots.map(x => x.qty + '개:' + x.price).join(' '));
}

// ------------------------------------------------------------
// 카탈로그 묶음 방문 (새 키워드든, 이어보기든 여기로 온다)
// ------------------------------------------------------------
async function visitCatalogGroup(keyword, cats, money, resumed) {
    const tabId = await ensureWorkerTab();

    S.current = {
        keyword: keyword,
        phase: resumed ? '못 본 카탈로그 이어보는 중' : '옵션 조사 중',
        catalogIndex: 0, catalogTotal: cats.length, collected: cats.length
    };
    S.note = '';
    broadcast();

    const visitedNow = [];

    for (let i = 0; i < cats.length; i++) {
        if (S.stopRequested) break;
        const cat = cats[i];

        S.current.catalogIndex = i + 1;
        broadcast();
        cat.scrapedAt = new Date().toLocaleString('ko-KR', { month: 'numeric', day: 'numeric', hour: '2-digit', minute: '2-digit' });

        // [수량 링크 샘플] 키워드마다 첫 카탈로그에서 한 번, 카드에서 읽은 수량 목록을
        //   진단 파일에 남긴다. 네이버가 카드 구조를 바꾸면 여기서 바로 보인다. (v4.91)
        if (i === 0) {
            S.failures.push({
                cat: { title: '[카드 수량링크 샘플] "' + keyword + '"', catalogUrl: cat.catalogUrl, collectStatus: '참고자료' },
                debug: {
                    issues: cats.slice(0, 10).map(c =>
                        (c.title || '').slice(0, 24) + ' → [' + ((c.cardQtys || []).join(',') || '못읽음') + ']'
                    ).join(' / ')
                }
            });
        }

        // [낱개 전용 걸러내기] 카드에 수량 링크가 "1개" 하나뿐이면 묶음이 없는 카탈로그다.
        //   묶음으로 사서 낱개로 파는 것이 이 프로그램의 전부이므로, 볼 것이 없다.
        //   페이지를 열기 전에 건너뛴다 — 카탈로그 하나에 10~20초를 아낀다. (v4.91)
        //   ※ 카드에서 수량 링크를 아예 못 읽었으면(빈 배열) 지금처럼 방문한다.
        //     네이버가 카드를 바꿔도 되던 것이 안 되는 일은 없어야 한다.
        const cq = Array.isArray(cat.cardQtys) ? cat.cardQtys : [];
        if (cq.length === 1 && cq[0] === 1) {
            cat.soloOnly = true;
            S.skipN.solo++;
        }
        //   ★낱개 사입(순방향)이 켜져 있으면 낱개 전용도 연다 — 1개를 싸게 사서 2~12개 칸에 파는 길이 있다 (v4.98).
        //     리뷰 기준은 똑같이 걸리므로 안 팔리는 것은 여기서도 걸러진다.
        if (cat.soloOnly && !money.forwardBuy) {
            cat.options = [];
            cat.collectStatus = '낱개 전용(건너뜀)';
            computeMargin(cat, money);
            S.pending = S.pending.filter(p => p !== cat);
            S.results.push(cat);
            visitedNow.push(cat);
            console.log('[낱개전용] 방문 안 함:', cat.title);
            continue;   // 페이지를 안 열었으니 쉬는 시간도 필요 없다
        }

        try {
            const marker = (cat.catalogUrl.match(/\/catalog\/(\d+)/) || [])[1] || null;
            await chrome.tabs.update(tabId, { url: cat.catalogUrl });
            const loaded = await waitForLoad(tabId, marker);
            if (!loaded) throw new Error('로딩 시간 초과: ' + cat.catalogUrl);
            // 🚫 차단 화면이면 쉬었다가 이 카탈로그부터 다시 (같은 자리 3번이면 포기하고 넘어간다)
            {
                const bk = await isBlockedTab(tabId);
                if (bk.blocked) {
                    cat._blockTries = (cat._blockTries || 0) + 1;
                    if (cat._blockTries <= 3) {
                        const ok = await blockPause(tabId, '카탈로그 열기', bk);
                        if (!ok) break;               // 중지
                        i--; continue;                // 같은 카탈로그 다시
                    }
                    throw new Error('차단으로 못 봄: ' + (bk.why || ''));
                }
                blockRested();
            }

            await updateModal(tabId, {
                keyword: keyword,
                big: (i + 1) + ' / ' + cats.length,
                phase: resumed ? '못 본 카탈로그 이어보는 중' : '수량 옵션을 읽는 중',
                sub: cat.title.slice(0, 34) + '\n남은 카탈로그 ' + (cats.length - i - 1) + '개'
            });

            // 페이지를 사람처럼 훑어본다 (지연 로딩도 함께 해결)
            await chrome.scripting.executeScript({
                target: { tabId: tabId }, func: humanScroll,
                args: [CONFIG.READ_PAUSE_MIN, CONFIG.READ_PAUSE_MAX]
            });

            // [리뷰 필터] 최근 N일 리뷰가 기준 미만이면 옵션도 안 읽고 건너뜀 (v4.73)
            //  = 실제로 잘 팔리는(리뷰 활발한) 카탈로그만 남긴다. 못 읽으면 안전하게 건너뜀.
            if (CONFIG.REVIEW_FILTER) {
                let rv = null;
                try {
                    const rvo = await chrome.scripting.executeScript({
                        target: { tabId: tabId }, func: scrapeReviewVelocity,
                        args: [CONFIG.REVIEW_DAYS, money.reviewMin || CONFIG.REVIEW_MIN_COUNT, CONFIG.REVIEW_COUNT_PAGES]
                    });
                    rv = rvo && rvo[0] && rvo[0].result;
                } catch (e) { rv = null; }
                // 첫 카탈로그의 리뷰 영역 구조를 진단 파일에 담는다 (선택자 보정용, 딱 1회)
                if (i === 0 && rv && rv.sample) {
                    S.failures.push({
                        cat: { title: '[리뷰 구조 샘플] ' + cat.title, catalogUrl: cat.catalogUrl, collectStatus: '참고자료' },
                        debug: { sampleHtml: rv.sample, reviewRecent: rv.recentCount, reviewSeen: rv.seen, why: rv.why }
                    });
                    // [낱개 사입 준비] 판매처 목록에서 "쿠폰" 이 붙은 가격 칸이 어떻게 생겼는지도 1회 담는다.
                    //  카탈로그 최저가는 즉시할인가 기준이라 쿠폰가는 로그인한 화면에만 따로 보인다(2025-03~).
                    //  그 값을 읽을 수 있으면 "네이버에서 1개를 싸게 사서 묶음 칸에 판다"는 둘째 매입 경로가 된다.
                    try {
                        const cp = await chrome.scripting.executeScript({ target: { tabId: tabId }, func: sampleCouponCells });
                        const txt = cp && cp[0] && cp[0].result;
                        if (txt) S.failures.push({
                            cat: { title: '[쿠폰가 구조 샘플] ' + cat.title, catalogUrl: cat.catalogUrl, collectStatus: '참고자료' },
                            debug: { sampleHtml: txt }
                        });
                    } catch (_) {}
                }
                if (!rv || !rv.pass) {
                    cat.collectStatus = '리뷰 부족(건너뜀)';
                    cat.skippedByReview = true;
                    putReviewInfo(cat, rv, CONFIG.REVIEW_DAYS);
                    S.skipN.review++;
                    S.pending = S.pending.filter(p => p !== cat);   // 다시 방문 안 하게 대기줄에서 뺀다
                    console.log('[리뷰필터] 건너뜀:', cat.title, '· 최근' + CONFIG.REVIEW_DAYS + '일 리뷰 '
                        + (rv ? rv.recentCount : '?') + '개 (기준 ' + (money.reviewMin || CONFIG.REVIEW_MIN_COUNT) + ')');
                    await restSleep(rand(CONFIG.CATALOG_DELAY_MIN, CONFIG.CATALOG_DELAY_MAX));
                    continue;   // 옵션·마진·꼬리물기 전부 건너뛰고 다음 카탈로그로
                }
                putReviewInfo(cat, rv, CONFIG.REVIEW_DAYS);   // 통과 — 숫자를 남긴다
            }

            const out = await chrome.scripting.executeScript({
                target: { tabId: tabId }, func: scrapeQuantityOptions,
                args: [CONFIG.OPTION_READ_TRIES]
            });
            const result = (out && out[0] && out[0].result) || null;

            if (result && result.options && result.options.length) {
                cat.options = result.options;
                cat.collectStatus = '성공';
                cat.optionGroup = result.groupName;

                // 마진을 먼저 계산해서 "가장 싼 묶음(매입수량)"을 알아낸다
                computeMargin(cat, money);

                // 1) 매입할 묶음 옵션을 눌러 그 상태의 주소를 기록 (여기서 사서 보낸다)
                if (cat.bestQty) {
                    try {
                        const b = await chrome.scripting.executeScript({
                            target: { tabId: tabId },
                            func: clickQtyAndInspect,
                            args: [cat.bestQty, 0]
                        });
                        const gotB = b && b[0] && b[0].result;
                        if (gotB && gotB.url) cat.buyUrl = gotB.url;
                        if (gotB && gotB.sellerName) cat.buySeller = gotB.sellerName;   // 매입할 최저가 판매처
                        if (gotB && gotB.buyLink) cat.buyLinkHref = gotB.buyLink;       // 그 판매처의 상품 페이지
                    } catch (e) { console.debug('[매입 주소] 실패:', e.message); }
                    await restSleep(rand(700, 1300));
                }

                // 2) 가장 적은 수량(보통 1개)을 눌러 판매 기준 주소 + 쿠팡 빠른배송 여부 기록
                const minQty = result.options[0].qty;
                cat.linkQty = minQty;
                try {
                    const u = await chrome.scripting.executeScript({
                        target: { tabId: tabId },
                        func: clickQtyAndInspect,
                        args: [minQty, cat.basePrice || 0]
                    });
                    const got = u && u[0] && u[0].result;
                    if (got && got.url) cat.optionUrl = got.url;
                    if (got && got.catalogTitle) cat.optionTitle = got.catalogTitle;   // 1개 옵션 상태의 제목
                    if (got) cat.coupangFast = !!got.coupangFast;   // 1개 최저가 판매처가 쿠팡 빠른배송인지
                    if (got) cat.buyOfficial = !!got.sellerOfficial;   // 최저가 판매처가 공식몰인지 (화장품 공식몰 필터용)
                    // [낱개 사입] 1개 최저가 판매처의 상품 페이지 — 순방향 단계가 여기로 가서 실제 매입가를 읽는다 (v4.98)
                    if (got && minQty === 1) { cat.soloBuyLink = got.buyLink || ''; cat.soloSeller = got.sellerName || ''; }
                    if (money.forwardBuy && minQty === 1) await captureOfficialSeller(tabId, cat, i === 0);   // 공식몰 판매처 (1개 옵션 상태에서)
                    // 키워드마다 첫 카탈로그 1건만 판매처 구조 샘플을 진단에 담는다
                    // (쿠팡 빠른배송 감지를 정확히 다시 만들기 위한 재료)
                    if (got && got.sellerSample && i === 0) {
                        S.failures.push({
                            cat: { title: '[판매처 구조 샘플] ' + cat.title, catalogUrl: cat.optionUrl || cat.catalogUrl, collectStatus: '참고자료' },
                            debug: { sampleHtml: got.sellerSample }
                        });
                    }
                } catch (e) { console.debug('[옵션 주소] 실패:', e.message); }

                // 3) 옵션이 있는 수량마다 "그 수량이 눌린 상태의 주소" 를 받아 적는다. (v5.0)
                //    ★실측(루크 11-04): 12개·11개·10개 묶음은 각각 주소가 다른데 엑셀엔 전부 대표 주소로 적혔다.
                //    이미 아는 것(매입 묶음 = buyUrl, 가장 적은 수량 = optionUrl)은 다시 안 누른다.
                //    경쟁 칸(남이 있는 칸)만 이 주소가 필요하고, 그 칸은 전부 옵션이 있으므로 옵션 수량 전부를 본다.
                try {
                    const known = {};
                    if (cat.bestQty && cat.buyUrl) known[Number(cat.bestQty)] = cat.buyUrl;
                    if (cat.linkQty && cat.optionUrl) known[Number(cat.linkQty)] = cat.optionUrl;
                    const want = [];
                    for (const o of (cat.cleanOptions || cat.options || [])) {
                        const q = Number(o.qty);
                        if (!(q > 0) || known[q] || want.indexOf(q) !== -1) continue;
                        want.push(q);
                    }
                    cat.qtyUrls = Object.assign({}, known);
                    if (want.length) {
                        const w = await chrome.scripting.executeScript({
                            target: { tabId: tabId },
                            func: clickQtysAndReadUrls,
                            args: [want]
                        });
                        const gotW = (w && w[0] && w[0].result) || {};
                        let n = 0;
                        for (const q of want) if (gotW[q]) { cat.qtyUrls[q] = gotW[q]; n++; }
                        if (n < want.length) console.debug('[수량 칸 주소] 못 받은 수량:', want.filter(q => !gotW[q]).join(','), cat.title);
                        await restSleep(rand(500, 1000));
                    }
                } catch (e) { console.debug('[수량 칸 주소] 실패:', e.message); }

                // [매입처 판매 여부 확인] 등록 후보 상품만, 매입처 상품 페이지를
                // 실제로 열어 판매중단/품절 문구를 확인한다. 표시만 하고 제외하지 않는다.
                //  새 수량 자리(낱개신규·빈수량신규)도 등록 후보다 — 안 넣으면
                //  제품불일치·화장품 공식몰 확인이 그쪽에만 안 돌아간다
                const isCandidate =
                    (cat.profit !== '' && cat.profit !== undefined && cat.profit >= money.targetProfit)
                    || (cat.mainProfit !== '' && cat.mainProfit !== undefined && cat.mainProfit >= money.targetProfit)
                    || (cat.soloProfit !== '' && cat.soloProfit !== undefined && cat.soloProfit >= money.targetProfit)
                    || (cat.gapProfit !== '' && cat.gapProfit !== undefined && cat.gapProfit >= money.targetProfit);

                // [카테고리·제조사·브랜드] 페이지를 떠나기 전에 읽어둔다.
                //   꼬리물기 키워드 재료이면서, 카테고리 성적을 쌓는 재료이기도 하다.
                //   ★v4.93부터 "등록 후보"가 아니어도 전부 읽는다 —
                //     거른 카탈로그까지 세어야 "이 카테고리는 건질 게 많은가"를 알 수 있다.
                //     (이미 열려 있는 페이지를 한 번 더 읽는 것뿐이라 느려지지 않는다)
                {
                    try {
                        const ti = await chrome.scripting.executeScript({
                            target: { tabId: tabId }, func: readTailInfoOnPage
                        });
                        const t0 = ti && ti[0] && ti[0].result;
                        if (t0) {
                            cat.catSmall = t0.category || '';
                            cat.catPath = t0.categoryPath || '';
                            cat.maker = t0.maker || '';
                            cat.brand = t0.brand || '';
                        }
                    } catch (e) { console.debug('[카테고리·브랜드 읽기] 실패:', e.message); }
                    // 💄 화장품 공식몰 필터: 카테고리가 화장품인데 최저가 매입처가 신뢰 채널이 아니면 제외
                    //   (화장품은 공식 입점만 가능 → 비공식 매입은 정가품/리셀 위험)
                    //   신뢰 채널 = ①네이버 공식/브랜드스토어(buyOfficial) ②홈쇼핑 ③쿠팡 로켓배송(쿠팡+빠른배송)
                    //   ※쿠팡인데 빠른배송 아님(마켓플레이스 3자)·랜덤 소형 리셀러 = 비공식
                    if (/화장품|미용/.test(String(cat.catPath || cat.catSmall || ''))) {
                        const _seller = String(cat.buySeller || '');
                        const _homeShop = /GS\s?SHOP|GS샵|롯데홈쇼핑|현대홈쇼핑|CJ\s?온스타일|공영쇼핑|NS\s?홈쇼핑|홈앤쇼핑|신세계\s?(TV|쇼핑)|SSG/i.test(_seller);
                        const _coupangRocket = /쿠팡/.test(_seller) && cat.coupangFast === true;   // 쿠팡+빠른배송 = 로켓배송
                        // 🆕 판매자(스토어)가 '브랜드 본인'으로 보이면 공식으로 인정 (사용자 요청)
                        //    ① 판매자명에 '공식' 글자가 있으면  ② 판매자명에 브랜드명이 들어 있으면
                        //    (오픈마켓·스마트스토어 공통 — 브랜드가 자기 이름으로 파는 것은 정품이라 안전)
                        const _normSN = s => String(s || '').toLowerCase().replace(/[\s\-_·.]/g, '');
                        const _brandN = _normSN(cat.brand), _sellerN = _normSN(_seller);
                        const _sellerIsBrand = /공식/.test(_seller) || (_brandN.length >= 2 && _sellerN.includes(_brandN));
                        const _trusted = (cat.buyOfficial === true) || _homeShop || _coupangRocket || _sellerIsBrand;
                        if (cat.buyOfficial !== undefined && !_trusted) {   // 판매처를 읽었는데 신뢰 채널이 아니면 제외
                            cat.buyStatus = '공식몰 아님(화장품)⚠️(' + (_seller.slice(0, 16) || '판매처미상') + ')';
                            cat.notOfficial = true;
                            console.warn('[공식몰] 화장품 비공식 매입처 → 제외:', cat.title, '/', _seller, '(' + (cat.catPath || cat.catSmall) + ')');
                        }
                    }
                }

                if (CONFIG.CHECK_BUY_LINK && isCandidate && cat.buyLinkHref) {
                    try {
                        await updateModal(tabId, {
                            keyword: keyword, big: '🛒',
                            phase: '매입처 판매 여부 확인 중',
                            sub: (cat.buySeller ? cat.buySeller + ' · ' : '') + cat.title.slice(0, 30)
                        });
                        await chrome.tabs.update(tabId, { url: cat.buyLinkHref });
                        await waitForLoad(tabId, null, 15);
                        await restSleep(rand(1500, 2500));
                        {   // 🚫 차단이면 쉬었다가 한 번 더
                            const bk = await isBlockedTab(tabId);
                            if (bk.blocked) {
                                const ok = await blockPause(tabId, '매입처 페이지', bk);
                                if (!ok) throw new Error('중지');
                                await chrome.tabs.update(tabId, { url: cat.buyLinkHref });
                                await waitForLoad(tabId, null, 15);
                                await restSleep(rand(1500, 2500));
                            }
                        }
                        const sc = await chrome.scripting.executeScript({
                            target: { tabId: tabId }, func: checkSellableOnPage
                        });
                        const r0 = sc && sc[0] && sc[0].result;
                        // 걸린 문구를 함께 적어서, 사용자가 왜 의심인지 바로 알 수 있게 한다
                        cat.buyStatus = (r0 && r0.sellable === false)
                            ? '판매중단 의심⚠️(' + (r0.hit || '') + ')' : '';
                        if (r0 && r0.title) cat.buyPageTitle = r0.title;   // 매입처의 실제 상품명
                        // 🚨 제품 불일치: 매입처 실제 상품이 카탈로그 제품과 완전히 다르면 = 잘못 매칭 → 등록 제외
                        if (!cat.buyStatus && cat.buyPageTitle && _titleMismatch(cat.title, cat.buyPageTitle)) {
                            cat.buyStatus = '제품 불일치 의심⚠️(매입: ' + String(cat.buyPageTitle).replace(/\s+/g, ' ').trim().slice(0, 24) + ')';
                            cat.buyMismatch = true;
                            console.warn('[매입처 확인] 제품 불일치:', cat.title, '≠', cat.buyPageTitle);
                        }
                        if (cat.buyStatus) console.warn('[매입처 확인]', cat.buyStatus, cat.title);
                    } catch (e) {
                        console.debug('[매입처 확인] 실패:', e.message);
                        cat.buyStatus = '';   // 확인 실패는 빈칸 - 틀린 표시를 하지 않는다
                    }
                }
            } else {
                cat.options = [];
                const reason = (result && result.reason) || 'unknown';
                cat.collectStatus = reason === 'no-qty-text' ? '옵션없음(확인)' : '읽기실패(의심)';
                if (cat.soloOnly) cat.collectStatus = '낱개 전용';   // 수량 칸이 없는 것이 정상이다
                if (money.forwardBuy) {
                    // [낱개 사입] 수량 칸이 없어도 "최저가 사러가기" 링크는 있다 — 1개 판매처와 카테고리를 담아 둔다 (v4.98)
                    try {
                        const u1 = await chrome.scripting.executeScript({ target: { tabId: tabId }, func: clickQtyAndInspect, args: [1, Number(cat.lowestPrice) || 0] });
                        const g1 = u1 && u1[0] && u1[0].result;
                        if (g1) { cat.soloBuyLink = g1.buyLink || ''; cat.soloSeller = g1.sellerName || ''; cat.buyOfficial = !!g1.sellerOfficial; }
                    } catch (_) {}
                    await captureOfficialSeller(tabId, cat, i === 0);
                    try {
                        const ti = await chrome.scripting.executeScript({ target: { tabId: tabId }, func: readTailInfoOnPage });
                        const t0 = ti && ti[0] && ti[0].result;
                        if (t0) { cat.catSmall = t0.category || ''; cat.catPath = t0.categoryPath || ''; cat.maker = t0.maker || ''; cat.brand = t0.brand || ''; }
                    } catch (_) {}
                }
                if (reason !== 'no-qty-text' && !cat.soloOnly) {
                    S.failures.push({ cat: { title: cat.title, catalogUrl: cat.catalogUrl, collectStatus: cat.collectStatus }, debug: (result && result.debug) || null });
                }
            }
        } catch (e) {
            console.error('[옵션 추출] 실패:', cat.catalogUrl, e);
            cat.options = [];
            cat.collectStatus = /차단으로 못 봄/.test(String(e && e.message)) ? '차단으로 못 봄' : '페이지 열기 실패';
            S.failures.push({ cat: { title: cat.title, catalogUrl: cat.catalogUrl, collectStatus: cat.collectStatus }, debug: { error: String(e && e.message || e) } });
        }

        // [낱개 사입(순방향)] 네이버 1개 최저가 판매처에서 실제 매입가(최대할인가 − 적립)를 읽어 1~12개 칸을 계산한다 (v4.98)
        if (money.forwardBuy && !S.stopRequested && cat.collectStatus !== '페이지 열기 실패') {
            try { await forwardBuyStep(tabId, cat, money, i === 0); }
            catch (e) { cat.fwdSlots = []; cat.fwdStatus = '읽기 실패(' + String(e && e.message || e).slice(0, 40) + ')'; }
        }

        computeMargin(cat, money);
        selfCheckRow(cat, money);   // 계산이 앞뒤가 맞는지 즉시 검산

        // [요약] 이번 판에서 무엇이 왜 빠졌는지 센다 (상품당 한 번만) (v4.91)
        if (!cat._skipCounted) {
            cat._skipCounted = true;
            bumpCatStat(cat);   // [카테고리 성적] 이 카테고리에서 건졌는지 세어 둔다 (v4.93)
            if (cat.verdict === '낱개신규') S.skipN.soloNew++;
            else if (cat.verdict === '낱개 비쌈') S.skipN.soloPricey++;
            else if (cat.verdict === '1개 옵션 없음') S.skipN.noOne++;
            else if (cat.verdict === '묶음 옵션 없음') S.skipN.bundleless++;
            if (cat.buyMismatch) S.skipN.mismatch++;
            if (cat.notOfficial) S.skipN.notOfficial++;

            // [낱개신규 근거] 한 판에 앞 20건은 어떻게 값을 정했는지 진단에 남긴다.
            //   가격을 우리가 정하는 길이라, 눈으로 확인할 수 있어야 한다. (v4.92)
            if (cat.soloNew && S.skipN.soloNew <= 20) {
                S.failures.push({
                    cat: { title: '[낱개신규] ' + cat.title, catalogUrl: cat.catalogUrl, collectStatus: '참고자료' },
                    debug: {
                        issues: '새 제목 "' + soloTitleOf(cat.title) + '"'
                            + ' / 판매가 ' + cat.soloPrice + ' / 개당원가 ' + cat.unitCost
                            + ' (' + cat.soloRatio + '배) / 순이익 ' + cat.soloProfit
                            + ' / 수량옵션 ' + (cat.options || []).map(o => o.qty + '개:' + o.price).join(', ')
                    }
                });
            }
        }

        // [꼬리물기] 등록가능 상품이면 상품명·카테고리·제조사·브랜드에서
        // 새 검색 키워드를 뽑아 대기줄 맨 뒤에 (자동추가)로 태운다
        if (money.autoLimit > 0 && bestProfitOf(cat) >= money.targetProfit) {
            try { await harvestTailKeywords(cat, money); }
            catch (e) { console.debug('[꼬리물기] 실패:', e.message); }
        }

        // 방문 완료: 보류 목록에서 빼고 결과에 넣는다 (하나씩, 즉시)
        S.pending = S.pending.filter(p => p !== cat);
        S.results.push(cat);
        visitedNow.push(cat);
        if (visitedNow.length % 3 === 0) await persist();

        await restSleep(rand(CONFIG.CATALOG_DELAY_MIN, CONFIG.CATALOG_DELAY_MAX));

        if ((i + 1) % CONFIG.LONG_BREAK_EVERY === 0 && i + 1 < cats.length) {
            const br = Math.round(rand(CONFIG.LONG_BREAK_MIN, CONFIG.LONG_BREAK_MAX) / 1000);
            await updateModal(tabId, {
                keyword: keyword, big: '☕',
                phase: '잠시 쉬는 중 (' + br + '초)',
                sub: '차단을 피하려면 필요한 시간이에요'
            });
            await restSleep(br * 1000);
        }
    }

    // ----- 키워드 집계 (이어본 만큼 합산) -----
    const tally = updateDoneEntry(keyword, visitedNow, money);
    S.current = null;


    const allDone = !S.queue.length && !S.pending.length && !S.stopRequested;
    await updateModal(tabId, {
        keyword: keyword,
        big: tally.hits + ' / ' + tally.total,
        phase: S.stopRequested
            ? '"' + keyword + '" 잠시 멈춤 (못 본 ' + tally.remaining + '개는 저장해뒀어요)'
            : (allDone ? '전체 작업 완료!' : '"' + keyword + '" 완료! 다음으로 갑니다'),
        sub: '수익 가능 카탈로그 ' + tally.hits + '개를 찾았어요',
        final: allDone || S.stopRequested,
        success: allDone
    });

    await persist();
    broadcast();
    return visitedNow.length;
}

// 완료 목록의 키워드 집계를 갱신한다. 이어본 결과는 기존 숫자에 합산된다.
function updateDoneEntry(keyword, visitedNow, money) {
    const hitsNow = visitedNow.filter(c => bestProfitOf(c) >= money.targetProfit).length;
    const remaining = S.pending.filter(p => p.searchKeyword === keyword).length;

    let e = S.done.find(d => d.keyword === keyword && !d.error);
    if (!e) {
        e = { keyword: keyword, total: 0, hits: 0 };
        S.done.push(e);
    }
    e.total += visitedNow.length;
    e.hits += hitsNow;
    S.runHits += hitsNow;   // [통계] 이번 실행 누적 등록가능 수
    e.partial = remaining > 0;

    return { hits: e.hits, total: e.total, remaining: remaining };
}

// 보류 목록이 빌 때까지 키워드 단위로 마저 본다. 하나라도 봤으면 true.
async function drainPending(noteText) {
    let any = false;
    while (!S.stopRequested && S.pending.length > 0) {
        S.note = noteText;
        broadcast();

        const kw = S.pending[0].searchKeyword;
        const cats = S.pending.filter(p => p.searchKeyword === kw);
        const money = await loadMoney();

        let n = 0;
        try {
            n = await visitCatalogGroup(kw, cats, money, true);
        } catch (e) {
            console.error('[이어보기] 실패:', kw, e);
            S.pending = S.pending.filter(p => p.searchKeyword !== kw);
            await persist();
        }
        if (n > 0) any = true;
        if (n === 0 && !S.stopRequested) {
            // 하나도 못 봤는데 멈춘 것도 아니면 같은 자리를 맴돌지 않게 정리
            S.pending = S.pending.filter(p => p.searchKeyword !== kw);
            await persist();
        }

        if (S.pending.length && !S.stopRequested) {
            await restSleep(rand(CONFIG.KEYWORD_DELAY_MIN, CONFIG.KEYWORD_DELAY_MAX));
        }
    }
    return any;
}

// ------------------------------------------------------------
// 대기줄 작업자
//  0) 지난번 중지로 못 본 카탈로그부터 마저 본다
//  1) 그다음 새 키워드들을 처리한다
//  2) 읽기 실패한 카탈로그가 있으면 자동으로 한 번 더 확인한다
//  3) 다 끝나면 CSV를 저장한다
// ------------------------------------------------------------
async function workerLoop() {
    if (S.running) return;
    S.running = true;
    // 🔁 번갈아: 지금 '등록 차례'면 새 판을 시작하지 않는다(겹침 방지).
    //   내 차례(source)가 되면 폴러가 다시 불러준다. 등록이 5분+ 조용하면 내가 이어간다.
    try {
        const _cc = await coordFetch({ side: 'source' });
        if (_cc && _cc.enabled) {
            if (_cc.paused) { S.running = false; coordStartResumePoller(); return; }   // ⏸ 일시정지 → 대기
            const _myTurn = (_cc.turn === 'source') || ((_cc.now - (_cc.reg_hb || 0)) > 300);
            if (!_myTurn) { S.running = false; coordStartResumePoller(); return; }
            await coordFetch({ turn: 'source' });   // 내 차례 확정 → 등록이 대기하도록
        }
    } catch (_) { }
    coordHbStart();   // 🔁 판 도는 동안 '나 소싱 살아있음' 을 알린다
    S.stopRequested = false;
    S.userStopped = false;   // 판이 돌기 시작했다 = 멈춘 상태가 아니다 (다시 켜지면 이어간다)
    S.lastError = '';
    broadcast();

    try {
        // ----- 회원 확인 (반드시 통과해야 시작) -----
        S.note = '회원 확인 중...';
        broadcast();
        const member = await verifyMember();
        if (!member.ok) {
            S.lastError = memberFailText(member);
            S.haltedFor = 'member'; _haltAt = Date.now();   // 5분 뒤에나 다시 본다
            return;
        }

        // ----- 네이버 로그인 확인 -----
        S.note = '네이버 로그인 확인 중...';
        broadcast();
        if (CONFIG.REQUIRE_NAVER_LOGIN) {
            const naverOk = await isNaverLoggedIn();
            if (!naverOk) {
                const opened = await openNaverLoginOnce();
                S.lastError = '네이버 로그인이 필요합니다. ' + (opened ? '로그인 창을 열었습니다.' : '열어 둔 로그인 창에서 로그인해 주세요.') + ' 로그인하면 5분 안에 저절로 이어갑니다.';
                S.haltedFor = 'login'; _haltAt = Date.now();
                return;
            }
            S.haltedFor = '';
        }

        // 이번 실행의 묶음 표시: "몇 차 꼬리인지 + 시작 시각"을 파일명에 넣어
        // 꼬리가 이어져도 순서가 헷갈리지 않게 한다 (이어서 재개하면 기존 표시 유지)
        if (!S.runStamp) {
            const d = new Date();
            S.runStamp = S.runRound + '차 ' + (d.getMonth() + 1) + '월' + d.getDate() + '일 '
                + String(d.getHours()).padStart(2, '0') + '시' + String(d.getMinutes()).padStart(2, '0') + '분';
            // 새 판이 시작됐으니 통계도 새로 센다
            S.runHits = 0;
            S.runActiveMs = 0;
            S.skipN = newSkipN();
        }
        S.sessionStartMs = Date.now();   // 이번 작업 구간의 시작 (통계용)

        let processedAny = false;

        // ----- 0단계: 못 본 카탈로그 이어보기 -----
        if (await drainPending('지난번에 못 본 카탈로그를 이어서 봅니다...')) processedAny = true;

        // ----- 1단계: 새 키워드 처리 -----
        S.note = '검색 페이지 여는 중...';
        broadcast();
        while (!S.stopRequested) {
            const keyword = S.queue.shift();
            if (!keyword) break;

            // 키워드마다 사용기간을 다시 확인 (만료되면 자동 정지)
            const re = await verifyMember();
            if (!re.ok && re.reason !== 'network') {
                S.queue.unshift(keyword);
                S.lastError = memberFailText(re);
                break;
            }

            const money = await loadMoney();
            try {
                await processKeyword(keyword, money);
                processedAny = true;
                // 엑셀은 매번 누적본이므로 중간마다 저장하지 않는다 (파일이 쌓여 헷갈림).
                // 전체 완료 또는 중지 시 아래 마무리 단계에서 한 번만 저장된다 (2026-09-06)
            } catch (e) {
                console.error('[키워드 처리] 실패:', keyword, e);
                S.done.push({ keyword: keyword, total: 0, hits: 0, error: true });
                S.current = null;
                await persist();
                broadcast();
            }

            if (S.queue.length && !S.stopRequested) {
                await restSleep(rand(CONFIG.KEYWORD_DELAY_MIN, CONFIG.KEYWORD_DELAY_MAX));
            }
        }

        // ----- 2단계: 읽기 실패한 카탈로그 자동 재확인 (딱 1회) -----
        // "옵션이 있는 상품은 하나도 빼놓지 않는다"가 목표이므로,
        // 실패한 것들을 자동으로 다시 방문해본다. (재시도 표시를 남겨 무한반복 방지)
        if (!S.stopRequested) {
            const retry = [];
            S.results = S.results.filter(r => {
                const bad = (r.collectStatus === '읽기실패(의심)' || r.collectStatus === '페이지 열기 실패'
                    || r.collectStatus === '미수집') && !r.retried;   // '미수집'은 옛 버전이 남긴 미방문 행
                if (bad) { r.retried = true; retry.push(r); return false; }
                return true;
            });
            if (retry.length > 0) {
                // 완료 집계에서 빼놨다가 재방문 때 다시 합산되게 한다
                for (const r of retry) {
                    const e = S.done.find(d => d.keyword === r.searchKeyword && !d.error);
                    if (e) {
                        e.total = Math.max(0, e.total - 1);
                        e.partial = true;
                    }
                }
                S.pending.push(...retry);
                await persist();
                if (await drainPending('읽기 실패한 카탈로그 ' + retry.length + '개를 다시 확인합니다...')) {
                    processedAny = true;
                }
            }
        }

        // ----- 마무리: 엑셀 파일 1개 저장 ([전체] 시트 + 키워드별 시트) -----
        if (processedAny && S.results.length) {
            try { await exportAll(); } catch (e) {
                console.error('[엑셀] 저장 실패:', e);
                S.lastError = '엑셀 저장에 실패했습니다: ' + String(e && e.message || e);
            }
        }
        // 진단 파일은 결과가 하나도 없어도(전부 건너뜀 등) 저장한다 — 원인분석·리뷰필터 보정용
        if (S.failures.length) {
            await sleep(800);
            await exportFailReport();
        }

        // [완전 자동] 판이 끝났고 사냥감이 남아 있으면, 사람이 🎯를 누른 것처럼
        // 다음 차수(10개 규칙 그대로)를 스스로 태운다 - 사냥감이 마를 때까지 이어진다
        if (!S.stopRequested && !S.queue.length && !S.pending.length && S.nextTargets.length) {
            const mm = await loadMoney();
            if (mm.autoHunt) {
                const cc = await coordFetch({ side: 'source' });   // 🔁 번갈아 모드인가?
                if (cc && cc.enabled) {
                    // 번갈아: 소싱 한 판 끝 → 등록에 차례를 넘기고, 다음 판은 등록이 끝난 뒤에 (폴러가 태움)
                    await coordFetch({ turn: 'register' });
                    coordStartResumePoller();
                    console.log('[번갈아] 소싱 한 판 완료 → 등록 차례로 넘김. 등록이 끝나면 다음 판을 이어갑니다.');
                } else {
                    const moved = await boardNextTargets();
                    if (moved > 0) console.log('[완전자동] 다음 사냥감 ' + moved + '개로 다음 판을 이어갑니다 (' + S.runRound + '차)');
                }
            }
        }

    } catch (e) {
        console.error('[작업자] 오류:', e);
        S.lastError = '오류가 발생했습니다: ' + String(e && e.message || e);
    } finally {
        const wasStopped = S.stopRequested;   // 사용자가 중지한 판은 자동으로 되살리지 않는다
        S.running = false;
        coordHbStop();   // 🔁 판 끝 → 하트비트 중단
        S.stopRequested = false;
        S.current = null;
        S.note = '';
        // [통계] 이번 작업 구간의 시간을 누적 (중지 후 이어가도 실제 작업 시간만 센다)
        if (S.sessionStartMs) {
            S.runActiveMs += Date.now() - S.sessionStartMs;
            S.sessionStartMs = 0;
        }
        // 실행이 완전히 끝났으면(대기줄이 텅 빔) 꼬리물기 연료통을 다시 채우고
        // 실행 묶음 표시도 지운다. 중간에 멈췄다 이으면 둘 다 이어서 쓴다 (한 판 = 한 통).
        // 사냥감까지 다 떨어졌으면 꼬리가 끝난 것이므로 차수도 1로 되돌린다.
        if (S.queue.length === 0 && S.pending.length === 0) {
            S.autoAdded = 0;
            S.deepAdded = 0;
            S.runStamp = '';
            if (S.nextTargets.length === 0) S.runRound = 1;
        }
        if (CONFIG.CLOSE_WORKER_TAB && S.workerTabId && S.queue.length === 0 && S.pending.length === 0) {
            try { await chrome.tabs.remove(S.workerTabId); } catch (e) {}
            S.workerTabId = null;
        }
        if (wasStopped) S.userStopped = true;   // [중지]로 멈춘 판 — 다시 켜져도 저절로 되살리지 않는다
        await persist();
        broadcast();
        // [완전 자동] 위에서 다음 판이 태워졌으면 잠시 숨 고르고 스스로 재시작
        if (!wasStopped && S.queue.length > 0 && !S.haltedFor) {   // 로그인·회원에서 멈춘 것은 3초마다 다시 돌지 않는다
            setTimeout(() => ensureWorker(), 3000);
        }
    }
}

function ensureWorker(force) {
    if (S.running) return;
    if (!force && S.haltedFor && Date.now() - _haltAt < HALT_RETRY_MS) return;   // 로그인·회원 멈춤 중 — 5분에 한 번만
    if (S.queue.length > 0 || S.pending.length > 0) workerLoop();
}

// ── 🔁 소싱↔등록 번갈아 (소싱 쪽) ────────────────────────────────
//  등록 프로그램 백엔드(/coord)가 '차례'를 관리하는 심판. enabled=true 면 차례를 지킨다.
//  등록 백엔드는 기본 포트 8732. 응답이 없거나(등록 안 켜짐) enabled=false 면 평소대로 돈다.
//  한 판이 끝나면 등록에 차례를 넘기고, 등록이 끝나(내 차례가 되면) 다음 판을 태운다.
const COORD_URL = 'http://localhost:8732/coord';
let _coordPoll = null, _coordHbTimer = null;

async function coordFetch(body) {
    try {
        const r = await fetch(COORD_URL, {
            method: 'POST', headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(body || {})
        });
        const j = await r.json();
        if (j && j.ok && j.app === 'selsuda-autoreg') return j;
    } catch (_) { }
    return null;   // 등록 백엔드 없음/응답없음 → 번갈아 제약 없이 평소대로
}

// 판이 도는 동안 '나 소싱, 살아있음' 을 주기적으로 알린다(그래야 등록이 내 차례를 안 뺏는다)
function coordHbStart() {
    if (_coordHbTimer) return;
    _coordHbTimer = setInterval(() => { coordFetch({ side: 'source' }); }, 60 * 1000);
    coordFetch({ side: 'source' });
}
function coordHbStop() { if (_coordHbTimer) { clearInterval(_coordHbTimer); _coordHbTimer = null; } }

// 등록에 차례를 넘긴 뒤, 다시 내 차례가 되면 다음 판을 스스로 태우는 폴러
function coordStopResumePoller() { if (_coordPoll) { clearInterval(_coordPoll); _coordPoll = null; } }
function coordStartResumePoller() {
    if (_coordPoll) return;
    _coordPoll = setInterval(async () => {
        const mm = await loadMoney();
        if (!mm.autoHunt) { coordStopResumePoller(); return; }                 // 완전자동 꺼짐 → 중단
        if (S.running) return;                                                 // 판 도는 중 → 대기
        if (!S.nextTargets.length && !S.queue.length) { coordStopResumePoller(); return; }  // 태울 것 없음
        const c = await coordFetch({ side: 'source' });                        // 살아있음 + 차례 확인
        if (c && c.enabled && c.paused) return;                                // ⏸ 일시정지 중 → 대기
        const free = !c || !c.enabled;                                         // 번갈아 꺼짐/백엔드 없음
        const mine = c && (c.turn === 'source' || (c.now - (c.reg_hb || 0)) > 300);  // 내 차례 or 등록이 5분+ 조용
        if (free || mine) {
            coordStopResumePoller();
            if (!S.queue.length && S.nextTargets.length) await boardNextTargets();
            ensureWorker();
        }
    }, 20 * 1000);
}

// ------------------------------------------------------------
// 📄 실행지시 파일 — 확장 폴더의 `실행지시.json` 을 따른다. 패널 없이 돌릴 때 쓴다 (클로드 점검·예약 실행). (v4.98)
//   { "id": "아무 글자", "키워드": ["닥터지 토너"], "기록무시": true, "기록비우기": false, "새로고침": false, "꼬리물기": true }
//   - 같은 id 는 한 번만 따른다 (chrome.storage.local.fileOrderDone). 다시 돌리려면 id 를 바꾼다.
//   - 기록무시: 1개월 재검색 제한을 그 키워드만 푼다.   기록비우기: 패널의 🧹 와 같다.   새로고침: 확장을 다시 켠다.
//   - 돌고 있는 중이면 끝난 뒤 다음 확인(1분마다) 때 따른다.
//   - 무엇을 했는지 `셀수다결과/진단/실행지시_응답.txt` 에 적는다 (패널이 없어도 결과를 알 수 있게).
// ------------------------------------------------------------
async function fileOrderNote(text) {
    try {
        const line = new Date().toLocaleString('ko-KR') + '  ' + text + '\n';
        await chrome.downloads.download({
            url: 'data:text/plain;charset=utf-8,' + encodeURIComponent('\ufeff' + line),
            filename: '셀수다결과/진단/실행지시_응답.txt', conflictAction: 'overwrite'
        });
    } catch (_) {}
}

let _fileOrderBusy = false;
async function maybeRunFileOrder() {
    if (_fileOrderBusy) return;
    _fileOrderBusy = true;
    try {
        let j = null;
        try {
            const res = await fetch(chrome.runtime.getURL('실행지시.json'), { cache: 'no-store' });
            if (!res.ok) return;
            j = await res.json();
        } catch (_) { return; }
        if (!j || !j.id) return;
        const st = await chrome.storage.local.get(['fileOrderDone', 'fileOrderReloaded']);
        if (st.fileOrderDone === String(j.id)) return;
        if (S.running) return;                                   // 끝난 뒤에 따른다
        // 새로고침이 켜진 지시는 먼저 한 번 새로고침하고(코드를 새로 읽게), 다시 켜지면 그때 키워드를 따른다
        if (j.새로고침 && st.fileOrderReloaded !== String(j.id)) {
            await chrome.storage.local.set({ fileOrderReloaded: String(j.id) });
            await fileOrderNote('[' + j.id + '] ' + appVer() + ' · 새로고침');
            setTimeout(() => chrome.runtime.reload(), 1500);
            return;
        }
        await chrome.storage.local.set({ fileOrderDone: String(j.id) });

        const did = [];
        if (j.탭닫기) {
            // 프로그램이 열어 둔 네이버 로그인 창들을 닫는다 (실측: 수십 개가 쌓였다)
            try {
                const tabs = await chrome.tabs.query({ url: ['https://nid.naver.com/*'] });
                for (const t of tabs) { try { await chrome.tabs.remove(t.id); } catch (_) {} }
                _loginTabId = null;
                did.push('로그인 탭 ' + tabs.length + '개 닫음');
            } catch (e) { did.push('탭 닫기 실패: ' + (e && e.message)); }
        }
        if (j.기록비우기) {
            S.queue = []; S.done = []; S.results = []; S.failures = []; S.pending = []; S.lastError = '';
            S.autoAdded = 0; S.deepAdded = 0; S.nextTargets = []; S.runStamp = ''; S.runRound = 1;
            S.runHits = 0; S.runActiveMs = 0; S.skipN = newSkipN(); S.catStat = {};
            did.push('기록 비움');
        }
        S.noTail = (j.꼬리물기 === false);   // 이 지시에서만 꼬리물기(자동추가 키워드)를 끈다 — 점검용
        const kws = Array.isArray(j.키워드) ? j.키워드.map(x => String(x).trim()).filter(Boolean) : [];
        if (j.기록무시 && kws.length) {
            const sh = await chrome.storage.local.get('kwHistory');
            const h = sh.kwHistory || {};
            for (const k of kws) delete h[bareKw(k)];
            await chrome.storage.local.set({ kwHistory: h });
            did.push('재검색 제한 해제');
        }
        let added = 0;
        for (const k of kws) if (!S.queue.includes(k)) { S.queue.push(k); added++; }
        if (added) did.push('키워드 ' + added + '개 추가: ' + kws.join(', ') + (S.noTail ? ' (꼬리물기 끔)' : ''));
        // (v5.4) "완전자동": 설정의 🤖 완전 자동을 켜거나 끈다. "사냥감": 다음 추적할 키워드에 넣는다(10개씩 판을 나눠 돈다).
        //   새로 깐 확장(저장이 빈 것)에 옛 다음사냥감 파일의 키워드를 되살릴 때 쓴다.
        if (j.완전자동 === true || j.완전자동 === false) {
            const sm = await chrome.storage.local.get('money');
            const m = sm.money || {};
            m.autoHunt = j.완전자동;
            await chrome.storage.local.set({ money: m });
            did.push('완전 자동 ' + (j.완전자동 ? '켬' : '끔'));
        }
        const hunts = Array.isArray(j.사냥감) ? j.사냥감.map(x => String(x).trim()).filter(Boolean) : [];
        if (hunts.length) {
            const known = new Set([].concat(S.queue, S.nextTargets).map(k => bareKw(k)));
            for (const d of S.done) known.add(bareKw(d.keyword));
            if (S.current) known.add(bareKw(S.current.keyword));
            let n = 0;
            for (const k of hunts) { const bk = bareKw(k); if (!bk || known.has(bk)) continue; known.add(bk); S.nextTargets.push(k); n++; }
            did.push('다음 사냥감 ' + n + '개 추가 (전체 ' + S.nextTargets.length + '개)');
            // 쉬고 있고 완전 자동이면 첫 판(10개)을 바로 태운다
            if (n && !added && !S.queue.length && !S.pending.length && !S.running && (await loadMoney()).autoHunt) {
                const moved = await boardNextTargets();
                if (moved) { did.push('첫 판 ' + moved + '개 태움'); added += moved; }
            }
        }
        await persist(); broadcast();
        await fileOrderNote('[' + j.id + '] ' + appVer() + ' · ' + (did.join(' / ') || '할 것 없음'));
        if (added) {
            S.haltedFor = '';
            ensureWorker(true);
            // 30초 뒤 어디까지 갔는지(회원 확인·네이버 로그인에서 막혔는지) 한 줄 더 남긴다
            setTimeout(() => fileOrderNote('[상태 30초] ' + (S.lastError ? '오류: ' + S.lastError : (S.running ? '돌고 있음 · ' + (S.note || (S.current && S.current.phase) || '') : '멈춤'))), 30000);
        }
    } catch (e) {
        try { await fileOrderNote('실패: ' + (e && e.message || e)); } catch (_) {}
    } finally { _fileOrderBusy = false; }
}
try {
    chrome.alarms.create('fileOrder', { periodInMinutes: 1 });
    chrome.alarms.onAlarm.addListener(a => { if (a && a.name === 'fileOrder') { maybeRunFileOrder(); ensureWorker(); } });   // 로그인/회원 멈춤도 5분마다 다시 본다
} catch (e) { /* alarms 권한이 없으면 패널이 열릴 때만 본다 */ }

// ------------------------------------------------------------
// 켜질 때 한 번: 새 버전이 있는지 보고, 폴더가 이미 갈렸으면 스스로 새로고침한다.
//  ★파일 맨 끝에 둡니다 — S·CONFIG 가 다 만들어진 뒤에 돌아야 합니다.
// ------------------------------------------------------------
try {
    clearReloadCountIfArrived();
    maybeCheckUpdate();
    setTimeout(maybeRunFileOrder, 3000);   // 켜지자마자 실행지시 파일도 본다 (S 복원 뒤)
} catch (e) { /* 업데이트 확인 실패가 프로그램을 막으면 안 된다 */ }
