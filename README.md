# SAP IBP IPAV2 Master Data Generator

SAP IBP Planning Area **IPAV2** 용 Master Data 및 Key Figure CSV 파일을 자동 생성하는 Node.js 웹 서버입니다.  
회사명을 입력하면 AI가 회사 정보를 분석하여 IBP 데모용 데이터 패키지(ZIP)를 생성합니다.

---

## 시작하기

### 사전 요구사항

- Node.js 18 이상
- SAP AI Core 계정 (Claude 모델 배포)

### 설치

```bash
npm install
```

### 환경변수 설정

프로젝트 루트에 `.env` 파일 생성:

```env
AI_CLIENT_ID=...
AI_CLIENT_SECRET=...
AI_AUTH_URL=https://<tenant>.authentication.<region>.hana.ondemand.com
AI_API_URL=https://api.ai.prod.<region>.aws.ml.hana.ondemand.com
AI_DEPLOYMENT_ID=...
AI_RESOURCE_GROUP=...
AI_INFERENCE_URL=...
```

### 실행

```bash
# 일반 실행
node server.mjs

# 개발 모드 (파일 변경 시 자동 재시작)
npm run dev
```

브라우저에서 `http://localhost:3001` 접속

---

## 기능

### 입력
- 회사명 (영문/한글)
- 고객사 목록 (채널, 국가, 추천 제품군)
- 제품 패밀리 구성 (FG / Semi / Raw Material)
- 생산 거점 (Location)

### 출력 — ZIP 패키지

#### Master Data (17개 생성 파일 + 선택적 REF 파일)
| 파일 | 설명 |
|---|---|
| 010_CUSTOMER_NEW.CSV | 고객 마스터 (CP1CUSTOMER) |
| 020_LOCATION_NEW.CSV | 거점 마스터 (CP1LOCATION) |
| 030_PRODUCT_NEW.CSV | 제품 마스터 (CP1PRODUCT) |
| 040_RESOURCE_NEW.CSV | 자원 마스터 (CP1RESOURCE) |
| 100_UOMCONVERSIONFACTOR_NEW.CSV | UOM 변환 (CP1UOMCONVERSIONFACTOR) |
| 110_CUSTOMERPRODUCT_NEW.CSV | 고객-제품 (CP1CUSTOMERPRODUCT) |
| 120_LOCATIONPRODUCT_SOP_DP_NEW.CSV | 거점-제품 (CP1LOCATIONPRODUCT) |
| 130_RESOURCELOCATION_NEW.CSV | 자원-거점 (CP1RESOURCELOCATION) |
| 135_RESOURCELOCATIONPRODUCT_NEW.CSV | 자원-거점-제품 |
| 140_SOURCECUSTOMER_NEW.CSV | 소싱-고객 (CP1SOURCECUSTOMER) |
| 150_SOURCELOCATION_SOP_DP_NEW.CSV | 소싱-거점 (CP1SOURCELOCATION) |
| 160_SOURCEPRODUCTIONHEADER_NEW.CSV | 생산 소싱 헤더 |
| 170_PRODUCTIONSOURCEITM_NEW.CSV | 생산 소싱 아이템 |
| 180_PRODUCTIONSOURCERESOURCE_NEW.CSV | 생산 소싱 자원 |
| 190_SOURCECUSTGROUP_NEW.CSV | 고객 그룹 소싱 |
| 200_SA_NEW.CSV | 판매 영역 (CP1SA) |
| 205_PLANNINGUNIT_NEW.CSV | 계획 단위 (CP1PLANNINGUNIT) — 관계 파일보다 먼저 업로드 |
| + Static REF files (BP 2608 기준) | LAG, SLC, CURRENCY, UOMTO 등 |

#### KF Data (17개 파일, 미입력 항목은 헤더 전용)
| 파일 | Planning Level | 주요 KF |
|---|---|---|
| IPAV2__WKPRODLOC | W \| Prod \| Loc | 재고, 원가, CO2 |
| IPAV2__WKPRODLOCSRC | W \| Prod \| Loc \| Src | 생산원가 |
| IPAV2__WKPRODLOCCUST | W \| Prod \| Loc \| Cust | 고객 비율 |
| IPAV2__WKPRODLOCSACUST | W \| Prod \| Loc \| SA \| Cust | 수요계획 실적/예측 |
| IPAV2__WKPRODLOCCURR | W \| Prod \| Loc \| Curr | 단위원가 |
| IPAV2__WKRESLOC | W \| Res \| Loc | 자원 용량 |
| IPAV2__WKRESPRODLOCSRC | W \| Res \| Prod \| Loc \| Src | 자원 소비 |
| IPAV2__DAYPRODLOCSACUST | Day \| Prod \| Loc \| SA \| Cust | 납품 실적 |
| IPAV2__MTHCOSTCENTERCURR | M \| CostCenter \| Curr | 재무 예산 |
| IPAV2__MTHPRODSACUSTCURR | M \| Prod \| SA \| Cust \| Curr | 월별 가격 (ZCPRICE) |
| IPAV2__WKPRODSACUSTCURR | W \| Prod \| SA \| Cust \| Curr | 주별 가격 (ZCPRICEW) |

---

## 데이터 생성 기준

- **과거 데이터**: SAP IBP Best Practice 2608 참조 데이터 기반 시드값 사용
- **주별 수요 KF**: 과거 105주(2년 이상) + 미래 26주, 합성 데모 실적
- **월별 KF**: 과거 12개월 + 미래 12개월
- **일별 KF**: 과거 26주
- **SALESAREA 형식**: `{SALESORG}/{CHANNEL_CODE}/00` (예: `1010/10/00`)
- **COSTCENTER**: `CC-{CUSTID}` 형식으로 고객과 1:1 매핑

---

## IBP 업로드 순서

ZIP의 **UPLOAD_ORDER.txt**가 실제 생성 파일 기준의 업로드 순서입니다. 계획 단위 및 UOM/Currency 참조 마스터를 관계 파일보다 먼저 등록하세요.

```
Master Data (순서 중요):
  0. 205_PLANNINGUNIT 및 UOM/Currency 등 REF 마스터
  1. 010_CUSTOMER
  2. 020_LOCATION
  3. 030_PRODUCT
  4. 040_RESOURCE
  5. 060~280 REF files
  6. 100_UOMCONVERSIONFACTOR 및 200_SA
  7. 110_CUSTOMERPRODUCT
  8. 120_LOCATIONPRODUCT
  9. 130~190 관계 파일
 10. 200_SA

KF Data (순서 무관):
  - KF 17개 파일 (헤더 전용 파일은 업로드할 데이터 없음)
```

---

## 기술 스택

- **Runtime**: Node.js (ESM)
- **AI**: SAP AI Core (Claude via Anthropic SDK)
- **ZIP 생성**: adm-zip
- **Excel 파싱**: xlsx
- **Frontend**: Vanilla HTML/JS (public/index.html)

## 모든 고객에 적용되는 IPAV2 주의사항

[공통 생성 규칙](docs/generation-guidelines.md)은 `IPAV2_IBP_MasterData_KF_주의사항.md`의 오류 사례를 반영합니다. 회사명 분석과 Excel 업로드 모두 `generation-rules.mjs`의 동일한 검증을 적용합니다.

- 제품별 생산거점, BOM 순환(1020), 운송 자기 루프(1023), 계획 단위 FK, 한 글자 제한, DBP 필수값 검증.
- Excel 기술 헤더를 우선 인식하고 기존 ID, 지역, Sales Area, UOM, 거점-제품/계획 단위를 보존. 샘플 수 제한으로 실제 제품/고객을 잘라내지 않음.
- **RESTYPE 문자열 ≠ IBP Production Category**: 기본값은 생산 Resource 연결을 만들지 않습니다. 실제 모델에서 확인한 Resource ID/Category만 고급 설정으로 연결하며 Type 1은 차단합니다.
- Capacity KF 이름/의미가 메모와 템플릿에서 다르면 임의로 대체하지 않습니다. tenant 템플릿에서 확인한 입력만 사용하세요.
- ZIP에 `VALIDATION_REPORT.json`, `UPLOAD_ORDER.txt`, `GENERATION_GUIDELINES.md` 포함. 검증 오류는 HTTP 422로 반환합니다. 기술 헤더로 업로드한 실제 ACTUALSQTY는 보존하며 빈 기간을 합성값으로 채우지 않습니다.
- CP1BU/CP1PRODUCTBU 및 DBP Driver/Include in Plan 설정은 tenant 템플릿/모델 확인 후 IBP에서 별도로 수행해야 합니다. 생성된 과거 실적은 실제 고객 실적이 아닌 **합성 데모 데이터**입니다.

UI 다운로드 화면의 **IPAV2 고급 설정**에서 `generationOptions` JSON을 입력할 수 있습니다. 제품별 매핑/UOM은 Excel에서 가져오거나 제품 객체의 `productionLocationIds`, `stockLocationIds`, `uomId`에 지정합니다.

```bash
npm test
```

테스트는 실제 AI 호출 및 SAP 업로드 없이 ZIP 내용과 여러 고객의 Excel 포맷을 검증합니다.
