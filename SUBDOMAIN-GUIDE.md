# joocnj.com 서브도메인 구성 가이드

문서번호: HYCNJ-WEB-SUBDOMAIN-2026-001

하나의 저장소·하나의 Vercel 배포로 **앱마다 서브도메인**을 제공하는 구성이다.
새 앱을 추가할 때 서버를 새로 만들 필요 없이 **폴더 하나 + 서브도메인 등록**이면 끝난다.

## 1. 구조 한눈에 보기

```
사용자 → Cloudflare DNS (DNS-only) → Vercel (와일드카드/개별 도메인)
                                        └─ vercel.json 호스트 라우팅
                                             joocnj.com        → / (주식 대시보드)
                                             www.joocnj.com    → joocnj.com 리다이렉트 (301)
                                             japan.joocnj.com  → /japan-travel/   (일본여행 번역기)
                                             apps.joocnj.com   → /apps/           (앱 포털)
                                             {이름}.joocnj.com → /{이름}/         (일반 규칙)
```

- **일반 규칙**: `{이름}.joocnj.com` 요청은 저장소의 `/{이름}/` 폴더로 자동 매핑된다.
  (`vercel.json`의 정규식 rewrite — 코드 수정 없이 폴더만 추가하면 서브도메인이 살아난다)
- **예외 매핑**: 폴더명과 서브도메인이 다르면 개별 규칙을 추가한다. (예: `japan` → `japan-travel`)
- **예약 서브도메인** (테넌트/앱 발급 금지): `www, api, admin, mail, static, assets, dashboard`

## 2. 최초 1회 설정 (Cloudflare + Vercel)

### 2-1. Vercel — 도메인 등록
Vercel 프로젝트 → Settings → Domains 에서 추가:

| 도메인 | 용도 |
|---|---|
| `joocnj.com` | 루트(주식 대시보드) |
| `www.joocnj.com` | apex로 301 리다이렉트 |
| `japan.joocnj.com` | 일본여행 번역기 |
| `apps.joocnj.com` | 앱 포털 |

> 정적 서빙 전제: 프로젝트 설정에서 Framework Preset **Other**, Build Command **없음**,
> Output Directory **`.`(저장소 루트)** 이어야 `/japan-travel/`, `/apps/` 폴더가 그대로 서빙된다.

### 2-2. Cloudflare — DNS 레코드
Cloudflare 대시보드 → joocnj.com → DNS:

| Type | Name | Target | Proxy |
|---|---|---|---|
| CNAME | `japan` | `cname.vercel-dns.com` | **DNS only (회색 구름)** |
| CNAME | `apps` | `cname.vercel-dns.com` | **DNS only (회색 구름)** |
| CNAME | `www` | `cname.vercel-dns.com` | **DNS only (회색 구름)** |
| A | `@` | `76.76.21.21` | **DNS only (회색 구름)** |

⚠️ **절대 규칙: 반드시 DNS-only(회색 구름).** 프록시(주황 구름)를 켜면 Vercel SSL 발급이 깨진다.

### 2-3. 와일드카드가 필요해지면
서브도메인이 많아지면(10개+) `*.joocnj.com` 와일드카드를 Vercel에 등록할 수 있는데,
이 경우 **네임서버를 Vercel로 위임**해야 한다 (Cloudflare DNS 관리 포기).
앱이 몇 개 수준이면 위 표처럼 **개별 CNAME 방식을 권장**한다 — Cloudflare를 그대로 유지할 수 있다.

## 3. 새 앱(서브도메인) 추가 절차 — 3단계

1. 저장소에 `/{앱이름}/index.html` 폴더 생성 (앱이름 = 소문자·숫자·하이픈)
2. Vercel Domains에 `{앱이름}.joocnj.com` 추가
3. Cloudflare에 CNAME `{앱이름}` → `cname.vercel-dns.com` (DNS only)

→ 배포·코드 수정 없이 서브도메인이 즉시 활성화된다 (`vercel.json` 일반 규칙이 자동 매핑).

## 4. 앱 작성 규칙

- **모든 자산 경로는 상대경로**로 쓴다 (`icon.png`, `./sw.js`).
  같은 앱이 `/japan-travel/`(경로 방식)과 `japan.joocnj.com/`(서브도메인 방식) 양쪽에서 동작해야 하기 때문.
- PWA 서비스워커는 `new URL('./', self.location).pathname`으로 마운트 위치를 동적으로 계산한다
  (`japan-travel/sw.js` 참고).
- 예약 서브도메인 이름은 폴더명으로 쓰지 않는다.

## 5. 점검 체크리스트

```
[도메인·SSL]
[ ] Cloudflare 레코드 전부 DNS-only(회색 구름)
[ ] Vercel Domains에 각 서브도메인 등록 + SSL 발급 완료(Valid)
[ ] www.joocnj.com → joocnj.com 301 동작
[ ] japan.joocnj.com 접속 시 번역 앱, joocnj.com 접속 시 대시보드

[앱]
[ ] japan.joocnj.com에서 PWA 설치·오프라인 동작 (manifest/sw 상대경로)
[ ] /japan-travel/ 경로 방식도 병행 동작 (Render Flask 포함)
[ ] apps.joocnj.com 포털에서 각 앱 링크 정상

[모바일]
[ ] iOS 사파리·안드로이드 크롬에서 서브도메인 접속 확인
```

## 6. 관련 파일

| 파일 | 역할 |
|---|---|
| `vercel.json` | 호스트(서브도메인) → 폴더 라우팅 규칙 |
| `apps/index.html` | 앱 포털 (apps.joocnj.com) |
| `japan-travel/` | 일본여행 번역기 (japan.joocnj.com) |
| `app.py` | Render(Flask) 경로 방식 서빙 — `/japan-travel/` |
