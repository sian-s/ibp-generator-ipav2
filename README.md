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

#### Master Data (15개 파일)
| 파일 | 설명 |
|---|---|
| 010_CUSTOMER_NEW.CSV | 고객 마스터 (CP1CUSTOMER) |
| 020_LOCATION_NEW.CSV | 거점 마스터 (CP1LOCATION) |
| 030_PRODUCT_NEW.CSV | 제품 마스터 (CP1PRODUCT) |
| 040_RESOURCE_NEW.CSV | 자원 마스터 (CP1PRODUCTIONRESOURCE) |
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
| + Static REF files (BP 2608 기준) | LAG, SLC, CURRENCY, UOMTO 등 |

#### KF Data (11개 파일)
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
- **주별 KF**: 과거 52주 + 미래 26주
- **월별 KF**: 과거 12개월 + 미래 12개월
- **일별 KF**: 과거 26주
- **SALESAREA 형식**: `{SALESORG}/{CHANNEL_CODE}/00` (예: `1010/10/00`)
- **COSTCENTER**: `CC-{CUSTID}` 형식으로 고객과 1:1 매핑

---

## IBP 업로드 순서

```
Master Data (순서 중요):
  1. 010_CUSTOMER
  2. 020_LOCATION
  3. 030_PRODUCT
  4. 040_RESOURCE
  5. 060~280 REF files
  6. 100_UOMCONVERSIONFACTOR
  7. 110_CUSTOMERPRODUCT
  8. 120_LOCATIONPRODUCT
  9. 130~190 관계 파일
 10. 200_SA

KF Data (순서 무관):
  - 위 KF 11개 파일 일괄 업로드
```

---

## 기술 스택

- **Runtime**: Node.js (ESM)
- **AI**: SAP AI Core (Claude via Anthropic SDK)
- **ZIP 생성**: adm-zip
- **Excel 파싱**: xlsx
- **Frontend**: Vanilla HTML/JS (public/index.html)
