# Gemini for Google Docs (Apps Script Sidebar Add-on)

Google Docs에서 사이드바를 통해 프롬프트를 입력하면,
Vertex AI의 Gemini 모델(`streamGenerateContent`)로부터 응답을 받아 문서 하단에 삽입해주는 Apps Script 예제입니다.

---

## 주요 기능

* **Docs UI 메뉴 추가**

  * 문서 열 때 `Gemini` 메뉴가 자동 생성됩니다.
* **OAuth 토큰 발급/캐싱**

  * `Gemini > Authorise` 실행 시 사용자 OAuth 토큰을 가져와 User Cache에 저장합니다.
* **사이드바 프롬프트 입력**

  * `Gemini > Show Sidebar`에서 입력창 사이드바를 띄우고 프롬프트를 전송합니다.
* **Gemini 응답을 문서에 삽입**

  * 모델 응답 텍스트를 현재 문서 **하단에 새 문단으로 append**합니다.

---

## 아키텍처/흐름

1. 사용자가 Google Docs를 엽니다 → `onOpen()` 실행 → 커스텀 메뉴 생성
2. `Gemini > Authorise` 클릭 → `auth()`에서 OAuth 토큰 발급 후 캐시에 저장
3. `Gemini > Show Sidebar` 클릭 → HTML 사이드바 표시
4. 사이드바에서 프롬프트 입력 후 `Send` 클릭
5. `askGemini(prompt)`가 Vertex AI REST endpoint로 요청
6. 응답 파싱 후 `insertText(answer)`로 문서에 삽입

---

## 사전 준비

### 1) Vertex AI / Generative AI API 활성화

* GCP 프로젝트에서 Vertex AI 및 Generative AI 기능이 활성화되어 있어야 합니다. ([Google Cloud Documentation][1])

### 2) 모델 ID 확인 (중요)

현재 코드:

```js
const MODEL_ID = 'gemini-1.0-pro';
```

`gemini-1.0-pro`(alias: `gemini-pro`)는 **2025-04-09에 Vertex AI에서 지원 종료(Shutdown)** 된 모델이므로,
지금은 호출이 실패할 가능성이 큽니다. ([GitHub][2])

**권장 업그레이드 예시**

* `gemini-1.5-pro`
* `gemini-2.0-flash` 등 최신 안정(stable) 모델

최신 모델 명명/라이프사이클은 공식 문서에 따라 바뀔 수 있습니다. ([Google Cloud Documentation][3])

---

## 설정 값

코드 상단 상수로 대상 모델/프로젝트/리전을 지정합니다.

```js
const MODEL_ID = 'gemini-1.0-pro';        // ⚠️ 최신 모델로 교체 권장
const PROJECT_ID = 'agentbuilder-429502';
const REGION = 'us-central1';

const API_ENDPOINT =
  `https://${REGION}-aiplatform.googleapis.com/v1/projects/${PROJECT_ID}/locations/${REGION}/publishers/google/models/${MODEL_ID}:streamGenerateContent`;
```

`streamGenerateContent` REST 엔드포인트 형식은 Vertex AI Gemini API 레퍼런스에 기반합니다. ([Google Cloud][4])

---

## 설치 방법

1. Google Docs 문서 열기
2. **확장 프로그램 → Apps Script**에서 새 프로젝트 생성
3. 아래 파일들을 추가/복사

   * `Code.gs` (스크립트)
   * `index.html` (사이드바 UI)
4. GCP 프로젝트가 올바르게 연결되어 있는지 확인
5. 최초 실행 시 권한 승인

---

## 사용 방법

1. 문서를 다시 열거나 새로고침
2. 상단 메뉴에서 **Gemini > Authorise**

   * 사용자 OAuth 토큰을 발급해서 캐시에 저장합니다.
3. **Gemini > Show Sidebar**
4. 사이드바에 프롬프트 입력 후 **Send**
5. Gemini 응답이 문서 하단에 자동 삽입됩니다.

---

## 코드 설명

### `onOpen()`

* Google Docs UI에 커스텀 메뉴 `Gemini`를 추가합니다.

### `auth()`

* `ScriptApp.getOAuthToken()`으로 현재 사용자 토큰을 가져와 `CacheService.getUserCache()`에 저장합니다.

### `showSidebar()`

* `index.html`을 사이드바로 렌더링합니다.

### `askGemini(prompt)`

* 캐시에서 토큰을 가져오고 없으면 안내 문구 삽입
* Vertex AI `streamGenerateContent` 엔드포인트로 POST 요청
* 응답 JSON을 순회하며 텍스트를 누적 후 문서에 삽입

> 현재 응답 파싱은 단순 루프이며,
> 스트리밍 응답/다중 part를 더 정확히 처리하려면 개선이 필요합니다. ([Google Cloud][4])

### `insertText(text)`

* 문서 body 끝에 새 문단으로 append 합니다.

---

## HTML 사이드바(`index.html`)

* 텍스트 입력 + 버튼으로 구성된 최소 UI
* `google.script.run.askGemini(prompt)`로 서버 함수 호출

---

## 필요한 OAuth Scope (권장)

Apps Script의 `appsscript.json`에 아래 권한이 필요합니다.

* `https://www.googleapis.com/auth/script.external_request`
  (UrlFetchApp로 외부 REST 호출)
* `https://www.googleapis.com/auth/documents` 또는 Docs 관련 기본 권한
  (문서 쓰기)
* `https://www.googleapis.com/auth/cloud-platform`
  (Vertex AI 호출)

정확한 스코프는 조직 보안 정책에 맞게 최소 권한으로 조정하세요. ([Google Cloud Documentation][1])

---

## 주의사항 / 제한

1. **모델 지원 종료**

   * `gemini-1.0-pro`는 이미 종료된 모델이라 호출 실패 시 모델 교체가 필요합니다. ([GitHub][2])

2. **토큰 캐시 만료**

   * User Cache는 만료될 수 있으므로 오류 시 다시 `Authorise` 실행해야 합니다.

3. **응답 삽입 위치**

   * 항상 문서 맨 아래에 append 됩니다.
   * 커서 위치 삽입이 필요하면 `getCursor()` 기반으로 수정해야 합니다.

4. **에러 처리 단순**

   * HTTP 200이 아니면 고정 에러 메시지를 삽입합니다.
   * 응답 body에 있는 실제 에러 메세지 표출 개선 가능.

---

## 트러블슈팅

### “Please generate a new authorisation token…”

* 캐시에 토큰이 없음
  → **Gemini > Authorise** 먼저 실행하세요.

### “Publisher Model not found” / 404

* 모델 ID가 더 이상 유효하지 않음
  → 최신 Gemini 모델로 교체 필요. ([GitHub][2])

### 403 Permission Error

* GCP 프로젝트에서 Vertex AI 권한 부족
  → 사용자/서비스 계정에 `Vertex AI User` 이상의 권한 부여 필요. ([Google Cloud Documentation][1])

---

## 개선 아이디어

* 스트리밍 응답을 **실시간으로 사이드바에 표시**
* 문서 **커서 위치에 삽입**
* 시스템 프롬프트/템플릿 지원
* 모델 파라미터(temperature, topP, maxOutputTokens 등) UI 제공
* 토큰 자동 갱신 또는 만료 안내

---
