# 셀수다 자동 소싱 — 업데이트기
#
#   깃허브에 올려둔 version.json 을 보고, 지금 깔린 것과 **다른 파일만** 내려받아
#   확장프로그램 폴더에 덮어씁니다. 다 끝나면 `디스크버전.json` 을 남깁니다.
#   확장은 그 쪽지를 보고 스스로 새로고침합니다 (크롬 재설치·다시 로드 불필요).
#
#   ★확장은 자기 파일을 못 씁니다(크롬이 막습니다). 그래서 이 역할이 따로 있습니다.
#   ★쪽지(디스크버전.json)는 **맨 마지막에** 씁니다.
#     파일 교체가 덜 됐는데 쪽지를 먼저 쓰면, 확장이 새로고침을 반복하게 됩니다.
#
#   예약 작업이 이 파일을 조용히 돌립니다. 손으로 돌려봐도 됩니다.
#     powershell -NoProfile -ExecutionPolicy Bypass -File 업데이트기.ps1

#   시험할 때만 쓰는 두 칸(-VersionUrl, -Root)이 있습니다. 평소에는 비워 둡니다.
param(
    [string]$VersionUrl = '',
    [string]$Root = ''
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8

$VERSION_URL = 'https://raw.githubusercontent.com/Yoo-Mideum/selsuda-sourcing-releases/main/version.json'
if ($VersionUrl) { $VERSION_URL = $VersionUrl }

# ★PowerShell 은 변수 이름의 대소문자를 안 가린다 — $root 를 채우고 나면 $Root(시험용 칸)도 같이 차 버린다.
#   그래서 "시험용 칸이 주어졌는지" 는 채우기 전에 따로 적어 둔다.
$RootWasGiven = -not [string]::IsNullOrEmpty($Root)
$root = if ($Root) { $Root } else { Split-Path -Parent $MyInvocation.MyCommand.Definition }
$ext  = Join-Path $root '확장프로그램'
$log  = Join-Path $root '업데이트기록.txt'
$prog = Join-Path $ext  '업데이트진행.json'   # 진행 쪽지 — 확장이 읽어 패널에 "몇/몇 (몇 %)" 로 보인다 (v5.2)

# 진행 쪽지. 단계마다 덮어쓴다. 확인 / 최신 / 받는중 / 덮어쓰는중 / 대기 / 완료 / 실패
function Note($stage, $here, $latest, $done, $total, $why) {
    try {
        $o = [ordered]@{
            stage  = [string]$stage
            at     = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
            here   = [string]$here
            latest = [string]$latest
            done   = [int]$done
            total  = [int]$total
            why    = [string]$why
        }
        $o | ConvertTo-Json -Depth 3 | Set-Content -LiteralPath $prog -Encoding utf8
    } catch {}
}

function Say($t) {
    $line = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') + '  ' + $t
    Write-Output $line
    try { Add-Content -LiteralPath $log -Value $line -Encoding utf8 } catch {}
}

# 기록이 끝없이 커지지 않게 — 400줄만 남긴다
function TrimLog {
    try {
        if (-not (Test-Path -LiteralPath $log)) { return }
        $lines = Get-Content -LiteralPath $log -Encoding utf8
        if ($lines.Count -gt 400) {
            $lines | Select-Object -Last 200 | Set-Content -LiteralPath $log -Encoding utf8
        }
    } catch {}
}

try {
    if (-not (Test-Path -LiteralPath $ext)) { Say "확장프로그램 폴더가 없습니다: $ext"; exit 1 }

    # ── 0. 업데이트기 자신도 새것으로 ─────────────────────────
    #   확장 파일과 같이 내려온 updater.ps1 이 나와 다르면, 나를 그것으로 갈아끼우고 그것을 돌린다.
    #   (예약 작업은 이 파일(루트)만 알기 때문에, 이렇게 해야 업데이트기도 업데이트된다)
    #   SELSUDA_UPD_CHAIN 은 "한 번만" 이어 달리게 하는 표시 — 혹시 갈아끼우기가 헛돌아도 무한히 되풀이하지 않는다.
    $me = $MyInvocation.MyCommand.Definition
    $newer = Join-Path $ext 'updater.ps1'
    if (-not $RootWasGiven -and -not $env:SELSUDA_UPD_CHAIN -and (Test-Path -LiteralPath $newer)) {
        $h1 = (Get-FileHash -LiteralPath $newer -Algorithm SHA256).Hash
        $h2 = (Get-FileHash -LiteralPath $me -Algorithm SHA256).Hash
        if ($h1 -ne $h2) {
            Copy-Item -LiteralPath $newer -Destination $me -Force
            Say '업데이트기 자체를 새것으로 바꿨습니다 — 새것으로 이어서 돕니다'
            $env:SELSUDA_UPD_CHAIN = '1'
            & powershell -NoProfile -ExecutionPolicy Bypass -File $me
            exit $LASTEXITCODE
        }
    }

    # ── 1. 지금 깔려 있는 버전 ─────────────────────────────────
    $mfPath = Join-Path $ext 'manifest.json'
    if (-not (Test-Path -LiteralPath $mfPath)) { Say 'manifest.json 이 없습니다'; exit 1 }
    # 앞에 BOM(보이지 않는 글자)이 붙어 있으면 ConvertFrom-Json 이 터진다 — 떼고 읽는다
    $mfText = (Get-Content -LiteralPath $mfPath -Raw -Encoding utf8) -replace "^﻿", ''
    $here = ($mfText | ConvertFrom-Json).version
    Note '확인' $here '' 0 0 ''

    # ── 2. 깃허브의 최신 버전 ──────────────────────────────────
    #   캐시를 피하려고 뒤에 시각을 붙인다 (raw.githubusercontent 는 몇 분 캐시한다)
    $url = $VERSION_URL + '?t=' + [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
    try {
        # ★ Invoke-RestMethod 로 바로 받지 않습니다.
        #   version.json 앞에 BOM(보이지 않는 글자 3개)이 붙어 있으면 그대로는 못 읽고,
        #   그러면 업데이트가 조용히 멈춥니다. 글자로 받아 BOM 을 떼고 읽습니다.
        $raw = (Invoke-WebRequest -Uri $url -TimeoutSec 20 -UseBasicParsing).Content
        if ($raw -is [byte[]]) { $raw = [Text.Encoding]::UTF8.GetString($raw) }
        $j = ($raw -replace "^﻿", '') | ConvertFrom-Json
    } catch {
        Say ('최신 버전을 못 읽었습니다 (인터넷/깃허브) — ' + $_.Exception.Message)
        Note '실패' $here '' 0 0 '최신 버전을 못 읽음 (인터넷/깃허브)'; exit 0
    }
    if (-not $j.latest) { Say 'version.json 모양이 이상합니다'; Note '실패' $here '' 0 0 'version.json 모양이 이상함'; exit 0 }

    if ($j.latest -eq $here) { Note '최신' $here $j.latest 0 0 ''; TrimLog; exit 0 }   # 같다 — 조용히 끝낸다 (매시간 돌므로)
    # 깔린 쪽이 더 새 버전이면(시험판을 손으로 넣어 둔 경우) 되돌리지 않는다.
    #   "다르다 = 받는다" 로 두면 시험 중인 4.97 을 깃허브의 4.95 로 덮어써 버린다.
    try {
        if ([version]$here -gt [version]$j.latest) { Say ("깔린 $here 이 깃허브 " + $j.latest + " 보다 새 것 — 그대로 둡니다"); Note '최신' $here $j.latest 0 0 ''; TrimLog; exit 0 }
    } catch { }   # 숫자 모양이 아니면 예전처럼 "다르면 받는다"
    Say ("새 버전 발견: $here → " + $j.latest)

    if (-not $j.patch -or -not $j.patch.files) { Say '받을 파일 목록이 없습니다'; Note '실패' $here $j.latest 0 0 '받을 파일 목록이 없음'; exit 0 }
    $base = [string]$j.patch.base
    $total = @($j.patch.files).Count
    $done = 0
    Note '받는중' $here $j.latest $done $total ''

    # ── 3. 다른 파일만 내려받는다 ──────────────────────────────
    $tmp = Join-Path $env:TEMP ('selsuda_sourcing_' + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $tmp -Force | Out-Null
    $staged = @()          # @{ dst; tmp }
    $skipped = 0

    foreach ($f in $j.patch.files) {
        $rel = [string]$f.path
        $dst = Join-Path $ext $rel
        # 이미 같은 파일이면 건너뛴다 (받는 양을 줄이는 핵심)
        if (Test-Path -LiteralPath $dst) {
            $h = (Get-FileHash -LiteralPath $dst -Algorithm SHA256).Hash.ToLower()
            if ($h -eq ([string]$f.sha256).ToLower()) { $skipped++; $done++; Note '받는중' $here $j.latest $done $total ''; continue }
        }
        $to = Join-Path $tmp ($rel -replace '[\\/]', '_')
        $src = $base + [uri]::EscapeDataString($rel)
        try {
            Invoke-WebRequest -Uri $src -OutFile $to -TimeoutSec 120 -UseBasicParsing
        } catch {
            Say ("받기 실패: $rel — " + $_.Exception.Message)
            Note '실패' $here $j.latest $done $total ("받기 실패: $rel")
            Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue
            exit 1
        }
        # 받은 것이 맞는 파일인지 확인 — 깨진 파일을 덮어쓰면 프로그램이 안 켜진다
        $gh = (Get-FileHash -LiteralPath $to -Algorithm SHA256).Hash.ToLower()
        if ($gh -ne ([string]$f.sha256).ToLower()) {
            Say "받은 파일이 깨졌습니다: $rel"
            Note '실패' $here $j.latest $done $total ("받은 파일이 깨짐: $rel")
            Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue
            exit 1
        }
        $staged += @{ dst = $dst; tmp = $to }
        $done++; Note '받는중' $here $j.latest $done $total ''
    }

    if ($staged.Count -eq 0) {
        # 파일은 다 같은데 버전만 다르다 = version.json 이 먼저 올라간 상태. 다음 차례에 다시 본다.
        Say "받을 파일이 없습니다 (같은 파일 $skipped 개) — 다음에 다시 봅니다"
        Note '대기' $here $j.latest $done $total '파일은 다 같은데 버전만 다름 — 다음에 다시 봄'
        Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue
        TrimLog; exit 0
    }

    # ── 4. 한꺼번에 덮어쓴다 (여기서부터는 빨리 끝나야 한다) ───
    Note '덮어쓰는중' $here $j.latest $done $total ''
    foreach ($s in $staged) {
        $dir = Split-Path -Parent $s.dst
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        Copy-Item -LiteralPath $s.tmp -Destination $s.dst -Force
    }
    Say ("바꾼 파일 " + $staged.Count + "개 (그대로 둔 것 $skipped 개)")

    # ── 5. 맨 마지막에 쪽지 ────────────────────────────────────
    $note = [ordered]@{
        version = [string]$j.latest
        at      = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        changed = @($staged | ForEach-Object { Split-Path -Leaf $_.dst })
        note    = [string]$j.note
    }
    $note | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $ext '디스크버전.json') -Encoding utf8
    Say ("업데이트 완료 → " + $j.latest + " (크롬에서 저절로 바뀝니다)")
    Note '완료' $here $j.latest $done $total ''

    Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue
    TrimLog
    exit 0          # 끝났다는 것을 분명히 알린다 (예약 작업이 성패를 이 숫자로 본다)
}
catch {
    Say ('오류: ' + $_.Exception.Message)
    Note '실패' '' '' 0 0 ('오류: ' + $_.Exception.Message)
    exit 1
}
