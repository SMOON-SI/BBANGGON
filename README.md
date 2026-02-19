# 병원관리대장(.xlsm) 구성 가이드

이 저장소에는 **엑셀 매크로 기반 병원관리대장**을 빠르게 구축하기 위한 VBA 모듈이 포함되어 있습니다.

## 포함 파일
- `hospital_manager_vba.bas`: 시트/테이블 자동 생성 + FORM 저장 매크로 + 검증 로직

## 요구사항
- Excel 데스크톱(Windows)
- 매크로 사용 통합문서(`.xlsm`)

## 설치 방법
1. 새 통합문서를 `Hospital_Manager.xlsm`으로 저장합니다.
2. `Alt+F11` → VBA 편집기 실행
3. `파일 > 파일 가져오기`에서 `hospital_manager_vba.bas`를 가져옵니다.
4. 즉시창(`Ctrl+G`) 또는 매크로 실행창(`Alt+F8`)에서 `InitializeHospitalWorkbook` 실행
   - 아래 시트/테이블이 생성됩니다.
     - `HOSPITAL` / `tblHospital`
     - `SERVICE_LOG` / `tblLog`
     - `ISSUE_MASTER` / `tblIssueMaster`
     - `FORM`

## FORM 버튼 연결
- `FORM` 시트에 도형(버튼)을 추가하고 `SaveServiceLog` 매크로에 연결하세요.
- 병원명 선택 시 병원ID 자동 입력을 위해 `FORM` 시트 코드 창에 아래 이벤트를 추가하세요.

```vb
Private Sub Worksheet_Change(ByVal Target As Range)
    If Not Intersect(Target, Me.Range("B3")) Is Nothing Then
        Application.EnableEvents = False
        UpdateHospitalIdByName
        Application.EnableEvents = True
    End If
End Sub
```

## 동작 요약
- 로그ID 자동 생성: `LOG-yyyymmdd-0001`
- 필수값 검증: 병원명/일자/증상(대/소)/조치(대/소)
- 저장 후 FORM 입력칸 초기화
- 증상/조치 드롭다운은 `ISSUE_MASTER(tblIssueMaster)`의 `사용여부=Y` 행만 반영
- 모든 데이터는 Excel Table(ListObject)로 생성되어 필터/피벗 분석에 적합

## 운영 팁
- `ISSUE_MASTER`에 신규 항목 입력 후 `RefreshIssueDropdowns` 실행하면 FORM 드롭다운이 즉시 반영됩니다.
- `HOSPITAL` 시트에서 병원명을 변경한 경우 FORM의 병원명 드롭다운도 자동 반영됩니다.
