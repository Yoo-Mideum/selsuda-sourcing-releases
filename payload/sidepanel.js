// ============================================================
// 사이드 패널 (버전 2)
//  - 로그인/회원가입 (구글 시트 연동)
//  - 키워드 대기줄 관리 + 실시간 진행 상황
// ============================================================

const $ = (id) => document.getElementById(id);

// 금액을 1,234 모양으로
const won = (n) => (Number(n) || 0).toLocaleString('ko-KR');

const VIEWS = ['view-loading', 'view-setup', 'view-login', 'view-main'];
function showView(id) {
    for (const v of VIEWS) { const el = $(v); if (el) el.hidden = (v !== id); }
}

function setMsg(id, kind, text) {
    const el = $(id);
    el.className = 'msg' + (kind ? ' ' + kind : '');
    el.innerText = text || '';
}

// ------------------------------------------------------------
// 시작: 세션 확인
// ------------------------------------------------------------
document.addEventListener('DOMContentLoaded', async function () {
    bindAuthEvents();
    bindMainEvents();
    await loadMoneyInputs();

    if (!isScriptUrlReady()) {
        showView('view-setup');
        return;
    }

    const st = await chrome.storage.local.get('member');
    const m = st.member;
    if (m && m.phone && m.token) {
        // 저장된 세션이 있으면 일단 메인을 보여주고, 뒤에서 재확인한다
        enterMain(m);
        try {
            const r = await apiCall('verify', {
                id: m.phone, token: m.token, ip: await getPublicIp(), ver: appVer()
            });
            if (r && r.ok) {
                m.name = r.name || m.name;
                m.nick = r.nick || m.nick;
                m.verifiedAt = Date.now();
                await chrome.storage.local.set({ member: m });
                enterMain(m);
            } else {
                await chrome.storage.local.remove('member');
                showView('view-login');
                setMsg('login-msg', 'err', failText(r));
            }
        } catch (e) {
            // 네트워크 일시 오류면 저장된 세션으로 계속 (작업 시작 시 다시 확인됨)
        }
    } else {
        showView('view-login');
    }
});

// 서버가 돌려준 code 를 사람 말로. (메시지를 주면 그걸 그대로 쓴다 — 서버가 더 정확하다)
function failText(r) {
    if (r && r.message) return String(r.message);
    const map = {
        'NO_USER': '회원 명단에 없는 번호예요. 루크 수강생은 상담 예약 때 쓴 번호, 뷰셀·미니쌤 수강생은 뷰셀 앱에 가입한 번호로 넣어주세요.',
        'NO_PW': '아직 비밀번호가 없어요. 루크 수강생은 상담 예약 페이지에서, 뷰셀·미니쌤 수강생은 뷰셀 앱에서 비밀번호를 먼저 정해주세요.',
        'BAD_PW': '비밀번호가 맞지 않아요.',
        'BLOCKED': '사용이 멈춰 있어요. 강사에게 문의해주세요.',
        'IP_FULL': '등록된 컴퓨터가 3대를 넘었어요. 안 쓰는 것을 지운 뒤 다시 해주세요.',
        'BAD_TOKEN': '다시 로그인해주세요.',
        'NO_SIGNUP': '이 프로그램에서는 가입할 수 없어요. 루크 수강생은 상담 예약 페이지에서, 뷰셀·미니쌤 수강생은 뷰셀 앱에서 가입해주세요.',
        'EMPTY': '전화번호와 비밀번호를 넣어주세요.',
        'BAD_BODY': '요청을 보내지 못했어요. 잠시 후 다시 시도해주세요.'
    };
    return map[(r && r.code) || ''] || '요청에 실패했어요. 잠시 후 다시 시도해주세요.';
}

// ------------------------------------------------------------
// 로그인 / 회원가입
// ------------------------------------------------------------
function bindAuthEvents() {
    $('btn-login').addEventListener('click', doLogin);
    $('login-pw').addEventListener('keydown', (e) => { if (e.key === 'Enter') doLogin(); });

    $('btn-logout').addEventListener('click', async () => {
        await chrome.storage.local.remove('member');
        showView('view-login');
        setMsg('login-msg', '', '');
    });
}

async function doLogin() {
    const phone = $('login-phone').value.replace(/\D/g, '');
    const pw = $('login-pw').value;
    if (phone.length < 10) return setMsg('login-msg', 'err', '전화번호를 확인해주세요 (숫자만).');
    if (!pw) return setMsg('login-msg', 'err', '비밀번호를 입력해주세요.');

    $('btn-login').disabled = true;
    setMsg('login-msg', 'info', '확인하는 중...');
    try {
        // 비밀번호는 원문으로 보낸다 — 사람마다 소금값이 달라서 대조는 서버가 한다.
        //   (상담 예약 때 정한 그 비밀번호와 같은 방식이라 따로 정할 것이 없다)
        const r = await apiCall('login', {
            id: phone, pw: pw, ip: await getPublicIp(), ver: appVer()
        });
        if (r && r.ok) {
            const member = {
                phone: phone,
                token: r.token || '',
                name: r.name || '',
                nick: r.nick || '',
                verifiedAt: Date.now()   // 방금 확인했으니 작업 시작 때 서버에 다시 묻지 않는다
            };
            await chrome.storage.local.set({ member });
            $('login-pw').value = '';    // 화면에 비밀번호를 남겨두지 않는다
            setMsg('login-msg', '', '');
            enterMain(member);
        } else {
            setMsg('login-msg', 'err', failText(r));
        }
    } catch (e) {
        setMsg('login-msg', 'err', '서버에 연결하지 못했어요. 인터넷을 확인해주세요.');
    }
    $('btn-login').disabled = false;
}

function enterMain(member) {
    showView('view-main');
    $('member-name').innerText = member.name || member.nick || '회원';

    // 사용기간(N일) 개념은 없어졌다 — 명단에 있고 멈춰 있지 않으면 쓴다
    const banner = $('days-banner');
    if (banner) banner.hidden = true;

    refreshNaver();
    requestState();
    reportRun(member);   // 🔒 실행 기록(무단복제 억제·증거) — 조용히, 실패해도 무방
}

// ------------------------------------------------------------
// 🔒 실행 기록 / 무단복제 억제·증거 수집
//   로그인·실행 때 인증서버(구글시트)에 '누가·언제·어느 설치본에서' 를 남긴다.
//   같은 설치본은 6시간에 한 번만 보낸다. 신고가 실패해도 프로그램엔 전혀 지장이 없다.
//   ⚠️ 억제책이다. 코드를 뜯어 이 부분을 지우면 안 남는다(근본 방어는 로그인 승인 + 라이선스 계약).
// ------------------------------------------------------------
let _runReported = false;

async function getInstallId() {
    try {
        const s = await chrome.storage.local.get('installId');
        if (s && s.installId) return s.installId;
        const id = (self.crypto && crypto.randomUUID)
            ? crypto.randomUUID()
            : (Date.now().toString(36) + Math.random().toString(36).slice(2));
        await chrome.storage.local.set({ installId: id });
        return id;
    } catch (_) { return ''; }
}

async function reportRun(member) {
    if (_runReported) return;          // 이번 세션에 이미 보냈으면 그만(복귀 시 enterMain 2번 호출 방지)
    _runReported = true;
    try {
        if (!member || !member.phone) return;
        const s = await chrome.storage.local.get('lastRunReport');
        if (s && s.lastRunReport && Date.now() - s.lastRunReport < 6 * 3600 * 1000) return;
        await chrome.storage.local.set({ lastRunReport: Date.now() });
        const ver = appVer();
        const installId = await getInstallId();
        // 비밀번호는 절대 보내지 않는다 — 번호·성함이면 시트에서 식별에 충분.
        //   [자동등록_기록] 탭에 한 줄 남는다.
        await apiCall('usage', {
            id: member.phone,
            product: '소싱 실행 v' + ver,
            result: '설치ID ' + installId + ' / ' + String(navigator.userAgent || '').slice(0, 80)
        });
    } catch (_) { /* 신고 실패는 조용히 넘어간다 — 사용엔 지장 없음 */ }
}

// ------------------------------------------------------------
// 메인 화면
// ------------------------------------------------------------
function bindMainEvents() {
    // 키워드 입력: Enter -> 쉼표로 쪼개 대기줄에 추가
    $('kw-input').addEventListener('keydown', (e) => {
        if (e.key !== 'Enter') return;
        const raw = $('kw-input').value;
        const keywords = raw.split(/[,\n]/).map(k => k.trim()).filter(Boolean);
        if (!keywords.length) return;
        chrome.runtime.sendMessage({ action: 'v2:addKeywords', keywords }, (res) => {
            if (chrome.runtime.lastError) return;
            $('kw-input').value = '';
            // 🔁 완전자동이 켜져 있으면 키워드를 넣는 즉시 자동으로 돌기 시작한다 (세트 모드)
            if ($('set-autohunt').checked) {
                chrome.runtime.sendMessage({ action: 'v2:start' }, () => { if (chrome.runtime.lastError) {} });
            }
            if (res && res.rejected && res.rejected.length) {
                const lines = res.rejected.map(r => r.reason === 'recent'
                    ? '"' + r.keyword + '" 는 최근 1개월 안에 검색했어요 (재검색 가능: ' + r.until + ')\n↓ 아래 목록에서 ✕를 누르면 바로 다시 검색할 수 있어요'
                    : '"' + r.keyword + '" 는 이미 대기줄에 있어요');
                setMsg('kw-msg', 'err', lines.join('\n'));
                // 최근 검색 제한에 걸렸으면 목록을 자동으로 펼쳐서 보여준다
                if (res.rejected.some(r => r.reason === 'recent')) {
                    $('hist-box').open = true;
                    loadHistList();
                }
            } else {
                setMsg('kw-msg', '', '');
            }
        });
    });

    // 이미 검색한 키워드 목록 (펼치면 불러오고, ✕로 제한 해제)
    $('hist-box').addEventListener('toggle', () => {
        if ($('hist-box').open) loadHistList();
    });
    $('hist-clear').addEventListener('click', () => {
        if (!confirm('검색 기록을 전부 삭제할까요?\n모든 키워드의 재검색 제한이 풀려요.')) return;
        chrome.runtime.sendMessage({ action: 'v2:clearHistory' }, () => {
            if (chrome.runtime.lastError) return;
            loadHistList();
            setMsg('kw-msg', 'ok', '검색 기록을 전부 삭제했어요. 어떤 키워드든 바로 검색할 수 있어요.');
        });
    });
    $('hist-list').addEventListener('click', (e) => {
        const btn = e.target.closest('.kw-del');
        if (!btn) return;
        chrome.runtime.sendMessage({ action: 'v2:removeHistory', keyword: btn.dataset.kw }, () => {
            if (chrome.runtime.lastError) return;
            loadHistList();
            setMsg('kw-msg', 'ok', '"' + btn.dataset.kw + '" 제한을 풀었어요. 이제 다시 검색할 수 있어요.');
        });
    });

    // 완료 목록에서 키워드 삭제 (기록 + 1개월 제한 해제)
    $('kw-done').addEventListener('click', (e) => {
        const btn = e.target.closest('.kw-del');
        if (!btn) return;
        if (!confirm('"' + btn.dataset.kw + '" 의 수집 기록을 삭제할까요?\n삭제하면 1개월 제한이 풀려서 바로 다시 검색할 수 있어요.')) return;
        chrome.runtime.sendMessage({ action: 'v2:removeDone', keyword: btn.dataset.kw }, (res) => {
            if (chrome.runtime.lastError) return;
            if (res && !res.ok) setMsg('run-error', 'err', res.error);
        });
    });

    // 대기줄 전체 비우기 (자동추가 키워드는 다음 사냥감으로 돌아간다)
    $('queue-clear').addEventListener('click', () => {
        if (!confirm('대기 중인 키워드를 전부 비울까요?\n자동추가 키워드는 「다음 사냥감」으로 돌아가서 나중에 🎯로 다시 꺼낼 수 있어요.')) return;
        chrome.runtime.sendMessage({ action: 'v2:clearQueue' }, (res) => {
            if (chrome.runtime.lastError) return;
            if (res && res.ok) setMsg('kw-msg', 'ok', '대기줄을 비웠어요.' + (res.parked > 0 ? ' 자동추가 ' + res.parked + '개는 다음 사냥감으로 보관했어요.' : ''));
        });
    });

    // 다음 사냥감 전체 비우기 (새 키워드로 새 출발하고 싶을 때)
    $('next-clear').addEventListener('click', () => {
        if (!confirm('다음 사냥감을 전부 비울까요?\n이미 엑셀 「다음사냥감」 파일로 백업돼 있고, 돈 되는 키워드는 새 꼬리물기에서 다시 발굴돼요.')) return;
        chrome.runtime.sendMessage({ action: 'v2:clearNext' }, (res) => {
            if (chrome.runtime.lastError) return;
            if (res && res.ok) setMsg('kw-msg', 'ok', '다음 사냥감 ' + res.cleared + '개를 비웠어요. 새 키워드로 새 출발!');
        });
    });

    // 다음 사냥감 이어서 찾기 (연료 초과로 못 돈 키워드를 대기줄에 태움)
    $('hunt-next').addEventListener('click', () => {
        const lv = Number($('set-autolimit').value);
        const lim = (isNaN(lv) || lv < 1) ? 10 : lv;
        if (!confirm('다음 사냥감에서 최대 ' + lim + '개만 대기줄에 태울까요?\n(남는 키워드는 계속 「다음 사냥감」에 보관돼요 - 누를 때마다 ' + lim + '개씩 소화)')) return;
        chrome.runtime.sendMessage({ action: 'v2:huntNext' }, (res) => {
            if (chrome.runtime.lastError) return;
            if (res && res.ok) setMsg('kw-msg', 'ok', '다음 사냥감 ' + res.moved + '개를 대기줄에 태웠어요!');
            else if (res) setMsg('kw-msg', 'err', res.error || '이어서 찾기에 실패했어요.');
        });
    });

    // 완료된 키워드 전체 삭제 (수집 기록 + 1개월 제한까지 한 번에)
    $('done-clear').addEventListener('click', () => {
        if (!confirm('완료된 키워드를 전부 삭제할까요?\n수집 기록이 지워지고, 모든 키워드의 1개월 제한도 풀려요.')) return;
        chrome.runtime.sendMessage({ action: 'v2:clearDone' }, (res) => {
            if (chrome.runtime.lastError) return;
            if (res && !res.ok) { setMsg('run-error', 'err', res.error); return; }
            loadHistList();
            setMsg('kw-msg', 'ok', '완료 목록을 비웠어요. 모든 키워드를 바로 다시 검색할 수 있어요.');
        });
    });

    $('btn-start').addEventListener('click', () => {
        chrome.runtime.sendMessage({ action: 'v2:start' }, () => { if (chrome.runtime.lastError) {} });
    });
    $('btn-stop').addEventListener('click', () => {
        chrome.runtime.sendMessage({ action: 'v2:stop' }, () => { if (chrome.runtime.lastError) {} });
    });
    $('btn-csv').addEventListener('click', () => {
        chrome.runtime.sendMessage({ action: 'v2:exportCsv' }, (res) => {
            if (chrome.runtime.lastError) return;
            if (res && !res.ok) setMsg('run-error', 'err', res.error || 'CSV 저장에 실패했어요.');
        });
    });
    // [지금 적용] — 확장을 새로고침해 새 버전으로 갈아탄다.
    //   모아둔 결과가 있으면 서버가 거절하고 무엇을 먼저 하라고 알려준다.
    $('btn-apply-upd').addEventListener('click', () => {
        setMsg('upd-msg', '', '적용하는 중...');
        chrome.runtime.sendMessage({ action: 'v2:applyUpdate' }, (res) => {
            if (chrome.runtime.lastError) return;   // 새로고침되면 여기까지 안 온다 (정상)
            if (res && !res.ok) setMsg('upd-msg', 'err', res.message || res.error || '적용하지 못했어요.');
        });
    });
    $('btn-clear').addEventListener('click', () => {
        if (!confirm('완료 기록 · 수집 결과 · 다음 사냥감을 모두 비울까요?\n(다음사냥감 파일도 지워져요. 소싱결과 엑셀은 그대로 남아요)')) return;
        chrome.runtime.sendMessage({ action: 'v2:clear' }, (res) => {
            if (chrome.runtime.lastError) return;
            if (res && !res.ok) setMsg('run-error', 'err', res.error);
        });
    });

    // 대기줄에서 키워드 삭제 (이벤트 위임)
    $('kw-queue').addEventListener('click', (e) => {
        const btn = e.target.closest('.kw-del');
        if (!btn) return;
        chrome.runtime.sendMessage({ action: 'v2:removeKeyword', keyword: btn.dataset.kw }, () => {
            if (chrome.runtime.lastError) {}
        });
    });

    // 네이버 로그인 칩 클릭 -> 로그인 페이지
    $('naver-chip').addEventListener('click', () => {
        if ($('naver-chip').classList.contains('bad')) {
            chrome.runtime.sendMessage({ action: 'openNaverLogin' }, () => { if (chrome.runtime.lastError) {} });
        }
    });

    // 계산 설정 저장
    for (const id of ['set-ship', 'set-buyship', 'set-fee', 'set-target', 'set-autolimit', 'set-autohunt', 'set-solocsv', 'set-forward']) {
        $(id).addEventListener('change', saveMoneyInputs);
    }
    // 🔁 완전자동을 켜면 [시작]을 따로 안 눌러도 바로 시작한다 (소싱·등록 세트 실행).
    //    돌 게 없으면 background(ensureWorker)가 알아서 가만히 있으니 안전.
    $('set-autohunt').addEventListener('change', () => {
        if ($('set-autohunt').checked) {
            chrome.runtime.sendMessage({ action: 'v2:start' }, () => { if (chrome.runtime.lastError) {} });
        }
    });

    // 진행 상황 수신
    chrome.runtime.onMessage.addListener((request) => {
        if (request.action === 'v2state' && request.state) renderState(request.state);
    });

    // 패널이 다시 열렸을 때를 대비해 주기적으로도 갱신
    setInterval(requestState, 3000);
    // 네이버 로그인 확인은 로컬 쿠키만 보므로 가볍다 - 자주 갱신해도 부담 없음
    setInterval(refreshNaver, 5000);
    // 다른 탭에서 네이버 로그인하고 돌아오면 바로 반영
    document.addEventListener('visibilitychange', () => {
        if (!document.hidden) { refreshNaver(); requestState(); }
    });
}

function requestState() {
    if ($('view-main').hidden) return;
    chrome.runtime.sendMessage({ action: 'v2:getState' }, (res) => {
        if (chrome.runtime.lastError) return;
        if (res && res.state) renderState(res.state);
    });
}

function refreshNaver() {
    if ($('view-main').hidden) return;
    chrome.runtime.sendMessage({ action: 'checkLogin' }, (res) => {
        if (chrome.runtime.lastError) return;
        const chip = $('naver-chip');
        if (res && res.loggedIn) {
            chip.className = 'status-chip ok';
            chip.innerText = '✅ 네이버 로그인됨';
        } else {
            chip.className = 'status-chip bad';
            chip.innerText = '❌ 네이버 로그인 필요 — 눌러서 로그인하기';
        }
    });
}

// ------------------------------------------------------------
// 상태 그리기
// ------------------------------------------------------------
function renderState(s) {
    // 현재 추적 중
    if (s.running) {
        $('now-running').hidden = false;
        $('now-idle').hidden = true;
        if (s.current) {
            $('now-keyword').innerText = s.current.keyword;
            $('now-phase').innerText = s.current.phase || '';
            const total = s.current.catalogTotal || 0;
            const idx = s.current.catalogIndex || 0;
            const pct = total > 0 ? Math.round((idx / total) * 100) : 0;
            $('bar-fill').style.width = pct + '%';
            $('bar-nums').innerText = total > 0
                ? ('카탈로그 ' + idx + ' / ' + total + ' (' + pct + '%)')
                : '카탈로그 목록을 모으는 중...';
        } else {
            // 시작 준비 단계 (회원 확인 / 네이버 확인 / 페이지 열기)
            $('now-keyword').innerText = '준비 중';
            $('now-phase').innerText = s.note || '시작 준비 중...';
            $('bar-fill').style.width = '0%';
            $('bar-nums').innerText = '잠시만요, 토끼가 신발끈을 묶고 있어요 🐇';
        }
        if (s.stopRequested) {
            $('now-phase').innerText = '🛑 중지하는 중... 읽던 카탈로그만 마무리해요';
        }
    } else {
        $('now-running').hidden = true;
        $('now-idle').hidden = false;
        const idleNote = document.querySelector('#now-idle .idle-note');
        if (idleNote) {
            idleNote.innerText = s.pendingCount > 0
                ? '아직 못 본 카탈로그가 ' + s.pendingCount + '개 남아 있어요. ▶ 시작을 누르면 이어서 봐요!'
                : '쉬고 있어요. 키워드를 추가하면 바로 달려가요!';
        }
    }

    // 오류 표시
    if (s.lastError) setMsg('run-error', 'err', s.lastError);
    else setMsg('run-error', '', '');

    // 이번 판 요약 — 무엇을 건지고 무엇을 왜 걸렀는지 한 줄 (규칙이 과한지 바로 보이게)
    const sum = $('run-summary');
    if (sum) {
        const k = s.skipN || {};
        const parts = [];
        if (s.runHits) parts.push('건진 ' + s.runHits);
        if (k.soloNew) parts.push('낱개신규 ' + k.soloNew);
        if (k.review) parts.push('리뷰부족 ' + k.review);
        if (k.solo) parts.push('낱개전용 ' + k.solo);
        if (k.soloPricey) parts.push('낱개비쌈 ' + k.soloPricey);
        if (k.bundleless) parts.push('묶음없음 ' + k.bundleless);
        if (k.noOne) parts.push('1개옵션없음 ' + k.noOne);
        if (k.mismatch) parts.push('제품불일치 ' + k.mismatch);
        if (k.notOfficial) parts.push('공식몰아님 ' + k.notOfficial);
        sum.hidden = parts.length === 0;
        sum.innerText = '이번 판 — ' + parts.join(' · ');
    }

    // 💰 건진 상품 미리보기 — 엑셀을 안 열어도 감이 오게
    // ⬆️ 새 버전 — 업데이트기가 파일을 받아놓으면 "지금 적용" 이 뜬다.
    //    쉬는 중이면 알아서 갈아타므로 이 칸을 볼 일은 드물다 (돌리는 중일 때만 보인다)
    const updCard = $('upd-card'), updText = $('upd-text'), updBtn = $('btn-apply-upd');
    if (updCard && updText) {
        const u = s.update;
        updCard.hidden = !u;
        if (u) {
            updText.innerText = u.ready
                ? ('새 버전 ' + u.latest + ' 을 다 받았어요.'
                   + (u.note ? '\n' + u.note : '')
                   + '\n크롬을 다시 켜면 바뀝니다. 지금 바꾸려면 아래를 누르세요.')
                : ('새 버전 ' + u.latest + ' 이 나왔어요. 받는 중입니다.'
                   + (u.note ? '\n' + u.note : '')
                   + '\n받고 나면 크롬을 다시 켤 때 바뀝니다 — 그냥 쓰셔도 됩니다.');
            if (updBtn) updBtn.hidden = !u.ready;
            // 자동 업데이트가 등록 안 된 PC 를 위한 대비 — 주소가 있을 때만 보여준다
            const link = $('upd-link'), a = $('upd-a');
            if (link && a) {
                if (u.url) { a.href = u.url; link.hidden = false; }
                else link.hidden = true;
            }
        }
    }

    const findCard = $('find-card'), findList = $('find-list');
    if (findCard && findList) {
        const tops = s.topFinds || [];
        findCard.hidden = tops.length === 0;
        findList.innerHTML = '';
        for (const f of tops) {
            const li = document.createElement('li');
            const left = document.createElement('span');
            left.innerText = (f.solo ? '🆕 ' : '') + f.title;
            const right = document.createElement('span');
            right.style.color = 'var(--ink-faint)';
            right.style.fontWeight = '700';
            right.style.whiteSpace = 'nowrap';
            // 리뷰 수를 같이 보여준다 — 마진만 큰 것보다 "팔리는 것" 을 고르게
            right.innerText = won(f.price) + '원 · +' + won(f.profit)
                + (f.rv ? ' · 리뷰 ' + f.rv : '');
            li.appendChild(left);
            li.appendChild(right);
            findList.appendChild(li);
        }
    }

    // 📊 잘 나오는 카테고리 — 어디를 더 팔지 고르는 근거 (건진 수 순)
    const catCard = $('cat-card'), catList = $('cat-list');
    if (catCard && catList) {
        const top = s.catTop || [];
        catCard.hidden = top.length === 0;
        catList.innerHTML = '';
        for (const c of top) {
            const li = document.createElement('li');
            const left = document.createElement('span');
            // 좁은 패널에서 줄이 접히지 않게 끝 두 단계만 보여준다 (전체 경로는 엑셀에 있다)
            const seg = String(c.path).split(' > ');
            left.innerText = seg.length > 2 ? seg.slice(-2).join(' > ') : c.path;
            left.title = c.path;
            const right = document.createElement('span');
            right.style.color = 'var(--ink-faint)';
            right.style.fontWeight = '600';
            right.style.whiteSpace = 'nowrap';
            right.innerText = '건진 ' + c.hit + ' / 방문 ' + c.seen + ' · 묶음 ' + c.bundleRate + '%';
            li.appendChild(left);
            li.appendChild(right);
            catList.appendChild(li);
        }
    }

    // 다음 키워드
    const q = $('kw-queue');
    q.innerHTML = '';
    for (let i = 0; i < s.queue.length; i++) {
        const li = document.createElement('li');
        const left = document.createElement('span');
        const num = document.createElement('span');
        num.className = 'kw-num';
        num.innerText = (i + 1);
        left.appendChild(num);
        left.appendChild(document.createTextNode(s.queue[i]));
        const del = document.createElement('button');
        del.className = 'kw-del';
        del.dataset.kw = s.queue[i];
        del.innerText = '✕';
        li.appendChild(left);
        li.appendChild(del);
        q.appendChild(li);
    }
    $('kw-queue-empty').hidden = s.queue.length > 0;

    // 꼬리물기 연료 게이지 (자동추가를 몇 개나 태웠는지)
    const fuelEl = $('fuel-line');
    if (fuelEl) {
        const lv = Number($('set-autolimit').value);
        const lim = (isNaN(lv) || lv < 0) ? 10 : lv;
        const used = s.autoAdded || 0;
        fuelEl.hidden = (lim === 0 || (used === 0 && s.queue.length === 0));
        fuelEl.innerText = '🔗 꼬리물기 연료 ' + used + ' / ' + lim
            + (used >= lim ? ' — 다 썼어요. 새 키워드는 「다음 사냥감」 시트로 모여요' : '');
    }

    // 다음 사냥감이 남아 있으면 이어서 찾기 버튼 표시
    const hn = $('hunt-next');
    if (hn) {
        hn.hidden = !(s.nextCount > 0);
        if (s.nextCount > 0) hn.innerText = '🎯 다음 사냥감 ' + s.nextCount + '개 이어서 찾기';
    }
    // 대기줄이 있으면 전체 비우기 버튼 표시
    const qc = $('queue-clear');
    if (qc) qc.hidden = s.queue.length === 0;
    // 사냥감이 있으면 사냥감 비우기 버튼도 표시
    const nc = $('next-clear');
    if (nc) nc.hidden = !(s.nextCount > 0);

    // 완료된 키워드: "콜라겐 (4/18)" = (수익 가능 / 전체)
    const d = $('kw-done');
    d.innerHTML = '';
    for (const item of s.done) {
        const li = document.createElement('li');
        const name = document.createElement('span');
        name.innerText = item.keyword;
        const right = document.createElement('span');
        const stat = document.createElement('span');
        stat.className = 'kw-stat' + (item.hits > 0 ? ' hit' : '');
        stat.innerText = item.error ? '오류'
            : '(' + item.hits + '/' + item.total + (item.partial ? '·중단' : '') + ')';
        const del = document.createElement('button');
        del.className = 'kw-del';
        del.dataset.kw = item.keyword;
        del.innerText = '✕';
        right.appendChild(stat);
        right.appendChild(del);
        li.appendChild(name);
        li.appendChild(right);
        d.appendChild(li);
    }
    $('kw-done-empty').hidden = s.done.length > 0;

    // 버튼 상태
    $('btn-start').disabled = s.running || (s.queue.length === 0 && !s.pendingCount);
    $('btn-start').innerText = s.running ? '🏃 작업 중'
        : (s.pendingCount > 0 && s.queue.length === 0 ? '▶ 이어보기' : '▶ 시작');
    $('btn-stop').disabled = !s.running || s.stopRequested;
    $('btn-stop').innerText = (s.running && s.stopRequested) ? '⏳ 중지 중' : '⏹ 중지';
    $('btn-csv').disabled = s.resultCount === 0;
    $('btn-clear').disabled = s.running;
}

// ------------------------------------------------------------
// 이미 검색한 키워드 목록 (최근 1개월, ✕로 제한 해제)
// ------------------------------------------------------------
function loadHistList() {
    chrome.runtime.sendMessage({ action: 'v2:kwHistoryList' }, (res) => {
        if (chrome.runtime.lastError || !res || !res.ok) return;
        const ul = $('hist-list');
        ul.innerHTML = '';
        for (const it of res.items) {
            const li = document.createElement('li');
            const left = document.createElement('span');
            left.appendChild(document.createTextNode(it.keyword + ' '));
            const stat = document.createElement('span');
            stat.className = 'kw-stat';
            stat.innerText = it.when;
            left.appendChild(stat);
            const del = document.createElement('button');
            del.className = 'kw-del';
            del.dataset.kw = it.keyword;
            del.innerText = '✕';
            li.appendChild(left);
            li.appendChild(del);
            ul.appendChild(li);
        }
        $('hist-empty').hidden = res.items.length > 0;
    });
}

// ------------------------------------------------------------
// 계산 설정 (택배비 / 수수료 / 목표순이익)
// ------------------------------------------------------------
async function loadMoneyInputs() {
    const st = await chrome.storage.local.get('money');
    const m = st.money || {};
    if (m.shippingFee !== undefined) $('set-ship').value = m.shippingFee;
    if (m.buyShipFee !== undefined) $('set-buyship').value = m.buyShipFee;
    if (m.feeRate !== undefined) $('set-fee').value = m.feeRate;
    if (m.targetProfit !== undefined) $('set-target').value = m.targetProfit;
    if (m.autoLimit !== undefined) $('set-autolimit').value = m.autoLimit;
    $('set-autohunt').checked = !!m.autoHunt;
    $('set-solocsv').checked = !!m.soloNewToCsv;
    $('set-forward').checked = m.forwardBuy === undefined ? true : !!m.forwardBuy;   // 낱개 사입(순방향) — 기본 켬
    if (m.slotRate !== undefined) $('set-slotrate').value = m.slotRate;
    // 2주 리뷰 기준: 구간 목록(10·30·50·70·100·150…500·600…1000)에서 고른다
    const sel = $('set-reviewmin');
    if (sel && !sel.options.length) {
        const tiers = [10, 30, 50, 70, 100, 150, 200, 250, 300, 350, 400, 450, 500, 600, 700, 800, 900, 1000];
        for (const t of tiers) {
            const o = document.createElement('option');
            o.value = String(t);
            o.textContent = t + '개 이상' + (t >= 1000 ? ' (메가급)' : t >= 100 ? ' (잘 팔림)' : t <= 10 ? ' (그래도 팔림)' : '');
            sel.appendChild(o);
        }
    }
    if (sel) sel.value = String(m.reviewMin || 10);
}

async function saveMoneyInputs() {
    // 꼬리물기 상한은 0도 유효한 값(끄기)이라 || 기본값을 쓰면 안 된다
    const alv = Number($('set-autolimit').value);
    await chrome.storage.local.set({
        money: {
            shippingFee: Number($('set-ship').value) || 3000,
            buyShipFee: Math.max(0, Number($('set-buyship').value) || 0),
            feeRate: Number($('set-fee').value) || 6.7,
            targetProfit: Number($('set-target').value) || 1000,
            autoLimit: (isNaN(alv) || alv < 0) ? 10 : alv,
            autoHunt: $('set-autohunt').checked,
            soloNewToCsv: $('set-solocsv').checked,
            forwardBuy: $('set-forward').checked,
            slotRate: Math.min(100, Math.max(1, Number($('set-slotrate').value) || 15)),
            reviewMin: Number($('set-reviewmin').value) || 10
        }
    });
}
