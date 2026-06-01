# Mujong 워크스페이스 (D:\testPro)

> 이 폴더는 **STS 워크스페이스** 와 **Inventory 프로젝트 루트** 를 겸하고 있습니다.
> 하위 폴더에 별도 프로젝트들이 있으며 각각 별도 Git 리포로 관리됩니다.

---

## 🗂️ 프로젝트 목록

| 폴더 | 프로젝트 | 포트 | 상태 | CLAUDE.md |
|------|---------|------|------|-----------|
| `.` (D:\testPro) | **Inventory** (재고관리) | 8080 | 운영 중 (Render 배포) | (이 파일) |
| `./alamis-homepage/` | **ARAMIS 회사 홈페이지** | 8081 | 시안 완료, 자료 대기 | `alamis-homepage/CLAUDE.md` |
| `./lumi/` | **Lumi** (AI 학습 도구) | 8082 | **Phase 3 완료** (다중 사용자 + ADMIN/USER + 자료 격리) | `lumi/CLAUDE.md` |
| `./docs/` | 사용설명서 PPT / 캡쳐 등 | - | - | - |

각 프로젝트는 **별도 Git 리포** 입니다. `.gitignore` 에서 서로를 제외합니다.

---

## 📚 Inventory (현재 폴더 기준)

> 아라미스(아버지 회사) 재고관리 시스템. 입고/출고/재고 + 라벨 인쇄 + 엑셀 다운로드.

### 한 줄 정의
판상/나노 알루미나 제조사의 입출고·재고·라벨 인쇄 시스템.

### 기술 스택
- Spring Boot 3.2.5 + JSP + WAR
- PostgreSQL (Neon, `neondb`) + MyBatis
- Apache POI (Excel)
- Render 배포 + UptimeRobot (5분 핑)

### 핵심 도메인
- **Inbound** — 입고 (P-YYYYMMDD-NNN 자동 채번)
- **Outbound + OutboundItem** — 출고 (master-detail, COA 필수)
- **User** — 세션 인증 (BCrypt)
- 메뉴: 입고관리 / 출고관리 / 현재재고조회 / 사용자관리(ADMIN)

### 배포
- https://alamis-pj.onrender.com
- UptimeRobot 5분 핑으로 슬립 방지
- LoginController 가 `redirect=` 파라미터 지원 (캡쳐 자동화용)

### 주요 사용자
- 아버지 (lysabb / 일반 사용자) — 본 운영
- admin / admin1234 — 관리자

### 사용설명서
- `docs/Alamis_재고관리_사용설명서.pptx` — PowerPoint COM 자동화로 생성
- 12 슬라이드 + 실제 화면 캡쳐 + 번호 마크업

---

## 🎯 작업 컨벤션 (모든 프로젝트 공통)

- 🇰🇷 **한국어 응답** — 코드 주석도 한글 OK
- 😄 **친근한 톤** — ㅋㅋ / 이모지 적당히 OK
- 📋 **단계별 정리** — 표 / 헤딩으로 구조화된 답변 선호
- 🎨 **Toss 스타일** — 깔끔하고 모던
- ❌ **인라인 스타일 금지** — 항상 CSS 클래스
- 📱 **반응형 필수** — 모바일/태블릿/데스크탑 다 동작
- 🔒 **비밀 정보** — 항상 `.gitignore` 되는 파일에
- 💬 **결정사항** — 모호하면 `AskUserQuestion` 으로 옵션 제시
- ✏️ **태스크 추적** — 3 단계 이상 작업은 `TaskCreate` 로
- 🪟 **Windows 환경** — PowerShell 사용, BOM 포함 UTF-8 PS1 파일

---

## 🔄 자주 쓰는 PowerShell 스니펫

### PS1 파일 한글 깨짐 방지 (BOM 추가)
```powershell
$path = "..."
$bytes = [System.IO.File]::ReadAllBytes($path)
if ($bytes[0] -ne 0xEF -or $bytes[1] -ne 0xBB -or $bytes[2] -ne 0xBF) {
    $content = [System.Text.Encoding]::UTF8.GetString($bytes)
    [System.IO.File]::WriteAllText($path, $content, (New-Object System.Text.UTF8Encoding $true))
}
```

### Edge headless 스크린샷
```powershell
$edge = "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe"
& $edge --headless=new --disable-gpu --window-size=1600,1000 --screenshot=out.png URL
```

### .NET HttpWebRequest (302 추적, PowerShell 5.1 redirect 우회)
```powershell
$req = [System.Net.HttpWebRequest]::Create($url)
$req.AllowAutoRedirect = $false
# ...
```

---

## 🔑 환경 정보

- **Java**: 17
- **OS**: Windows 11 (PowerShell 5.1)
- **Git**: 깃 글로벌 설정 변경 금지
- **STS**: 같은 워크스페이스에 여러 프로젝트 동시 import 가능

---

## ⚠️ 주의사항 (DRM)

이 PC 에 **Fasoo DRM** 이 설치되어 있어, PowerPoint 가 만든 PNG / PPT 추출물이 자동 암호화됩니다.

- ✅ Edge headless 가 만든 PNG: DRM X (Read 도구로 정상 분석 가능)
- ❌ PowerPoint COM Export PNG: DRM 적용됨 (외부에서 못 읽음)
- ❌ PowerPoint 가 만든 .pptx: DRM 적용됨 (단, 본 PC PowerPoint 로는 열림)
- ➡️ 시각 QA 가 필요한 경우 PowerPoint 출력 대신 Edge headless 캡쳐 활용

---

## 💡 다음 세션 진입 시

각 프로젝트의 `CLAUDE.md` 를 먼저 읽고 작업 시작.

- Lumi 작업 → `D:\testPro\lumi\CLAUDE.md`
- ARAMIS 홈페이지 → `D:\testPro\alamis-homepage\CLAUDE.md`
- Inventory 작업 → 이 파일 + 코드 베이스

---

_마지막 업데이트: Lumi Phase 3 완료 (로그인 + ADMIN/USER + 사용자별 자료 격리 + PDF 뷰어 + Flatpickr). 다음 후보: R2 파일 저장소 / AI 챗봇 / 퀴즈 생성._
