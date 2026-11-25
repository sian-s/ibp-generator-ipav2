# IPAV2 공통 데이터 생성·업로드 주의사항

2026-11-25 `IPAV2_IBP_MasterData_KF_주의사항.md`의 LX PoC 오류 사례를 모든 고객에 적용하는 규칙으로 정리했습니다. 고객명 입력과 Excel 업로드 모두 동일한 서버 검증을 통과해야 ZIP이 생성됩니다. 특정 LX 제품명·거점 ID는 공통 규칙에 하드코딩하지 않습니다.

## 자동 적용

- LOCVALID/CUSTVALID/PRDVALID=X. LOCTYPE는 공장 P, 물류센터 DC로 정규화합니다.
- LOCATIONPRODUCT는 제공된 CP1 템플릿의 순서로 출력합니다. PLANNER=1, 등록된 PLUNITID 사용, DC의 MRPTYPE=X. DISTRIBUTIONTYPE/STOCKINGNODETYPE는 한 글자만 허용하고 NAGINDEPNETWORK는 임의 생성하지 않습니다.
- 제품별 생산거점 매핑을 생산 Header, BOM, 관계/KF 파일에도 적용합니다. Excel에서 생산거점이 여러 개인데 제품 매핑이 없으면 생성을 차단합니다. 회사명으로 만든 거점 매핑은 제안이며 업로드 전에 반드시 검토해야 합니다.
- BOM 자품목이 빠졌으면 제품 및 거점-제품에 추가합니다. 자기 참조와 다단계 순환 BOM(1020), 0/음수 계수는 차단합니다. 자품목이 없는 경우 자신을 자품목으로 만들지 않습니다.
- 운송 Source는 공급 거점과 수요 DC가 서로 다르고 해당 제품이 등록된 경우만 생성합니다. 임의의 공장 간 경로나 자기 루프(1023)를 만들지 않습니다.
- Customer Region이 없으면 국가/HQ 기준으로 제안하고 보고서에 경고합니다. Sales Area는 모든 관련 KF와 SA 마스터에 일관되게 채웁니다.
- WKPRODLOCSACUST에는 양수 ZXDISAGGREGATIONFACTOR를 생성합니다. Excel 시트에 PRDID, LOCID, CUSTID, KEYFIGUREDATE, ACTUALSQTY 기술 헤더가 있으면 해당 실적을 보존하고 합성 실적으로 빈 기간을 채우지 않습니다. 실제 실적이 없을 때만 최소 2년(105주)의 **합성 데모 실적**을 생성합니다. 합성값을 실제 실적으로 오인해 운영 시스템에 적재하면 안 됩니다. 미래 예측 및 납품 샘플은 계속 데모값입니다.
- UTF-8 BOM, sep=,, CSV 인용/이스케이프를 사용합니다. 선택 오류·중복 ID·잘못된 값은 HTTP 422로 차단합니다.
- UPLOAD_ORDER.txt는 계획 단위와 참조 마스터를 관계 파일보다 먼저 제시합니다. 헤더 전용 파일도 ZIP에 포함하되, Delete 작업으로 업로드하지 않습니다.

## Resource Type 1 / 에러 1080

`CP1RESOURCE.RESTYPE=ASSEMBLY`와 `CP1RESOURCELOCATION.RESOURCETYPE`/IBP 내부 Production Category는 서로 다른 개념입니다. 문자열 검사만으로 내부 Category를 확인할 수 없습니다.

기본값에서는 CP1PRODUCTIONRESOURCE 및 생산 자원 소비 관련 관계/KF를 **헤더 전용**으로 출력합니다. 기존 시스템의 잘못된 연결을 삭제하거나 모델을 수정하는 작업은 자동으로 수행하지 않습니다.

Planning Model에서 Production Category를 확인한 뒤에만 `generationOptions`로 아래 설정을 전달합니다. 숫자 1을 Production Type으로 지정하면 차단합니다. 확인 목록은 이 생성기가 출력하는 `LOCID ASSEMBLY` 등의 Resource ID여야 합니다. 기존 Excel 설비 ID를 자동으로 다른 설비에 연결하지 않습니다.

```json
{
  "planningUnitId": "CUSTOMER_PLAN",
  "confirmedProductionResourceIds": ["LOC-PLANT ASSEMBLY"],
  "productionResourceTypeCode": "<실제 모델에서 확인한 생산 Category 코드>",
  "capacityProfiles": {
    "LOC-PLANT ASSEMBLY": { "ZSSHIFTNUM": 2, "ZSCAPAPERSHIFT": 8, "ZSEFFICIENCY": 0.85 }
  }
}
```

확인 목록은 실제 출력 Resource ID와 일치해야 합니다. 거점·제품 매핑은 제품 객체의 `productionLocationIds`, 재고 위치는 `stockLocationIds`를 사용합니다. 고급 설정의 `productionLocations: {"제품ID": ["LOC-공장ID"]}`로 생산 매핑을 명시할 수도 있습니다. UOM은 제품의 `uomId`, 실측 변환계수는 `uomConversions`에 지정합니다. 고객 Sales Area는 `salesArea`에 지정합니다. UI의 고급 설정 JSON 또는 analysis/selections.generationOptions에서 전달할 수 있습니다.

## Capacity 계산은 모델 확인 필요

원본 메모의 입력은 ZSSHIFTNUM × ZSCAPAPERSHIFT × ZSEFFICIENCY × ZSRESOURCEWORKDAYS × ZSCAPAMULTIPLIER이며, HANDLING은 운송 Capacity 분기를 사용합니다.

제공된 로컬 템플릿에는 `ZSCAPASUPPLYSTORAGE`, `ZSCAPASUPPLYTRANSPORT`, `ZSRESOURCE`, `ZSSHIFTHRS`가 있습니다. 메모의 `ZSCAPASUPLY*` 철자와 작업일수/대수 KF는 템플릿과 다릅니다. `ZSSHIFTHRS`를 작업일수 대신 쓰거나 `ZSRESOURCE`를 리소스 대수라고 **가정하지 않습니다**. 두 칼럼은 모델 정의를 확인해 capacityProfiles에 지정하기 전까지 비워둡니다. 지원되지 않는 KF는 자동 추가하지 않고 차단합니다. CAPASUPPLY는 계산 KF로 비워둡니다. HANDLING에는 템플릿 운송 입력도 채웁니다. Base Version/Scenario를 확인해야 합니다.

## IBP에서 직접 확인할 것

- CP1BU/CP1PRODUCTBU: 현재 제공된 템플릿에 없으므로 실제 tenant 템플릿을 먼저 확보하고 제품-사업부 연결을 등록해야 합니다. 임의 컬럼을 생성하지 않습니다.
- CP1PLANNINGUNIT: 로컬 템플릿은 PLUNITID만 포함합니다. PLUNITDESCR/UOMID는 실제 템플릿과 모델 정의를 확인한 뒤 등록합니다. PLUNITID를 EA/KG 같은 UOM과 혼동하지 않습니다.
- UOM/Currency/FK 참조 마스터가 먼저 존재해야 합니다. 서로 다른 물리 단위 사이에 임의의 1:1 변환은 생성하지 않습니다.
- DBP Driver의 ROPID/ROPTYPEID/ROPDESCR 등록 및 Include in Plan=1은 IBP에서 수행해야 합니다. Driver 없이 ROAGGSLSQTY 행을 임의 생성하지 않습니다.
- 실제 과거 실적 적재 → Statistical Forecast 실행 → Consensus Demand 확정 → HDEMANDRATIO 및 ROCALCSALESQTY 확인. 신제품 분배계수와 Customer Region/Sales Area를 검토합니다.
- CSV Insert/Update와 Delete를 구분합니다. 생성 ZIP은 기존 잘못된 BOM/Resource/Transport 행을 자동 삭제하지 않습니다.
- 에러 400: 템플릿 컬럼 확인. 010: 길이 제한. 145: 선행 FK. 1020: 순환 BOM. 1023: 자기 루프. 1080: 비생산 Category 연결.

ZIP의 VALIDATION_REPORT.json은 로컬 생성 검증 결과이며 실제 SAP IBP 업로드 성공이나 모델 설정 검증을 보장하지 않습니다.
