// ============================================================
// [회원 확인 서버]
//
//  소싱 프로그램의 로그인은 자동상품등록 프로그램과 **같은 명단**을 봅니다.
//    원장(명단) : 「멘토루크 컨설팅」 시트의 [잔여] 탭   ← 읽기만 함
//    서버       : 「자동상품등록 회원확인 서버」 (Apps Script 웹앱)
//    쓰는 곳    : [자동등록_회원] (IP·토큰) / [자동등록_기록] (접속·실행 기록)
//
//  그래서 상담 예약 때 쓰는 그 번호·비밀번호로 그대로 들어옵니다.
//  ★가입은 이 프로그램에서 하지 않습니다 — 상담 예약 페이지에서 가입합니다.
//   (명단을 프로그램이 건드리면 꼬이기 때문에 서버가 일부러 막아 둡니다)
//
//  주소를 옮길 일이 생기면 아래 SCRIPT_URL 한 곳만 바꾸면 됩니다.
//  (자동등록 쪽은 1_백엔드_소스/server.py 의 AUTH_URL — 두 프로그램이 같은 주소를 봅니다)
// ============================================================

const APP_CONFIG = {
    SCRIPT_URL: "https://script.google.com/macros/s/AKfycbx8S9Urovgy3pvBCYpYHHLxlb0B5-4DTN32raVJIriLPof4z0Ut0FoS6pZWSpuwKTGlYw/exec"
};

// URL이 제대로 들어 있는지 확인
function isScriptUrlReady() {
    return typeof APP_CONFIG.SCRIPT_URL === 'string'
        && APP_CONFIG.SCRIPT_URL.indexOf('https://script.google.com/') === 0;
}

// ------------------------------------------------------------
// 서버에 요청 보내기
//
//  · 회원확인 서버는 요청을 **JSON 본문**으로 읽습니다(e.postData.contents).
//    예전 소싱 서버는 폼(URLSearchParams)이었습니다 — 섞이면 전부 실패합니다.
//  · Content-Type 을 text/plain 으로 보내야 preflight(OPTIONS) 없이 지나갑니다.
//  · ⚠️ 앱스스크립트 웹앱은 302 로 넘깁니다. 그때 GET 으로 바뀌면 doGet 의 응답
//    ({ok:true, code:'ALIVE'})이 돌아오는데, 이걸 '로그인 성공'으로 오해하면 안 됩니다.
//    실제로 겪은 사고라, 그 모양이 오면 한 번 다시 보냅니다.
// ------------------------------------------------------------
async function apiCall(action, params) {
    const body = JSON.stringify(Object.assign({ action: action }, params || {}));

    async function send() {
        const res = await fetch(APP_CONFIG.SCRIPT_URL, {
            method: 'POST',
            headers: { 'Content-Type': 'text/plain;charset=utf-8' },
            body: body,
            redirect: 'follow'
        });
        return await res.json();
    }

    let j = await send();
    if (j && j.code === 'ALIVE' && action !== 'ping') {
        // doGet 이 대답한 것 = 내 요청이 전달되지 않았다. 한 번만 다시.
        j = await send();
    }
    return j;
}

// ------------------------------------------------------------
// 공인 IP — 서버가 "계정당 컴퓨터 3대"를 세는 데 씁니다.
//  못 알아내면 빈칸을 보냅니다(서버가 IP 등록을 건너뜁니다 — 로그인은 됩니다).
//  같은 PC 에서 자동등록과 소싱을 같이 써도 IP 가 같으므로 한 자리만 씁니다.
// ------------------------------------------------------------
const IP_SERVICES = [
    'https://api.ipify.org',
    'https://checkip.amazonaws.com',
    'https://ipv4.icanhazip.com'
];
let _ipCache = { ip: '', at: 0 };

async function getPublicIp() {
    if (_ipCache.ip && Date.now() - _ipCache.at < 30 * 60 * 1000) return _ipCache.ip;
    for (const url of IP_SERVICES) {
        try {
            const ac = new AbortController();
            const timer = setTimeout(() => ac.abort(), 4000);
            const res = await fetch(url, { signal: ac.signal, cache: 'no-store' });
            clearTimeout(timer);
            const ip = (await res.text()).trim();
            if (/^\d{1,3}(\.\d{1,3}){3}$/.test(ip)) {
                _ipCache = { ip: ip, at: Date.now() };
                return ip;
            }
        } catch (e) { /* 다음 서비스로 */ }
    }
    return '';
}

// 이 프로그램의 버전 (서버 기록에 남겨 어느 버전이 도는지 본다)
function appVer() {
    try { return chrome.runtime.getManifest().version || ''; } catch (e) { return ''; }
}

// ============================================================
// [업데이트] — 자동상품등록과 같은 방식
//
//  깃허브에 올려둔 version.json 이 "최신 버전 번호" 와 "파일 목록(sha256)" 을 들고 있고,
//  윈도우에 깔린 작은 업데이트기(업데이트기.ps1)가 그걸 보고 **바뀐 파일만** 내려받아
//  확장프로그램 폴더에 덮어씁니다. 받는 양은 보통 수백 KB 입니다.
//
//  확장은 파일을 스스로 못 씁니다(크롬이 막습니다). 그래서 역할을 나눕니다.
//    업데이트기 : 파일을 교체하고, 끝나면 `디스크버전.json` 을 남긴다
//    ★파일 이름 앞에 _ 를 붙이면 안 된다 — 크롬이 "_ 로 시작하는 이름은 시스템 예약" 이라며 확장을 아예 못 올린다 (2026-11-01 실측)
//    확장       : 그 파일을 보고 "폴더가 새 버전이 됐다"는 걸 알면 스스로 새로고침
//
//  그래서 수강생은 아무것도 누르지 않습니다.
// ============================================================
const UPDATE_CONFIG = {
    VERSION_URL: 'https://raw.githubusercontent.com/Yoo-Mideum/selsuda-sourcing-releases/main/version.json',
    CHECK_EVERY_MIN: 60      // 너무 자주 묻지 않는다 (깃허브에 부담)
};

// "4.95" vs "4.100" 처럼 자릿수가 달라도 맞게 비교한다 (글자 비교는 틀린다)
function verGreater(a, b) {
    const pa = String(a || '').split('.').map(n => Number(n) || 0);
    const pb = String(b || '').split('.').map(n => Number(n) || 0);
    for (let i = 0; i < Math.max(pa.length, pb.length); i++) {
        const x = pa[i] || 0, y = pb[i] || 0;
        if (x !== y) return x > y;
    }
    return false;
}

// 깃허브의 최신 버전 정보. 못 가져오면 null (인터넷이 없어도 프로그램은 돌아야 한다)
async function fetchLatestVersion() {
    try {
        const ac = new AbortController();
        const timer = setTimeout(() => ac.abort(), 6000);
        const res = await fetch(UPDATE_CONFIG.VERSION_URL, { signal: ac.signal, cache: 'no-store' });
        clearTimeout(timer);
        if (!res.ok) return null;
        const j = await res.json();
        return (j && j.latest) ? j : null;
    } catch (e) { return null; }
}

// 업데이트기가 폴더를 갈아놓고 남긴 쪽지. 없으면 null (= 아직 교체 안 됨)
//  ★이 파일은 설치 때는 없고 업데이트기가 '새로' 만든다. 그래서 캐시에 걸리지 않는다.
// 업데이트기의 진행 쪽지 (업데이트진행.json). 없으면 null = 업데이트기가 이 컴퓨터에서 돈 적이 없거나 옛 업데이트기. (v5.2)
async function fetchUpdateProgress() {
    try {
        const url = chrome.runtime.getURL('업데이트진행.json');
        const res = await fetch(url, { cache: 'no-store' });
        if (!res.ok) return null;
        const j = await res.json();
        return (j && j.stage) ? j : null;
    } catch (e) { return null; }
}

async function fetchDiskVersion() {
    try {
        const url = chrome.runtime.getURL('디스크버전.json');
        const res = await fetch(url, { cache: 'no-store' });
        if (!res.ok) return null;
        const j = await res.json();
        return (j && j.version) ? j : null;
    } catch (e) { return null; }
}
