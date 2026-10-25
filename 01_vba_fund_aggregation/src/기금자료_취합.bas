Attribute VB_Name = "기금자료_취합"
'==============================================================================
' [복원본] 기금별 A자료 취합 매크로
'
' ※ 회계법인 근무 중(2026.01~2026.03) 실제 업무에 사용한 원본 VBA가 아닙니다.
'    원본은 회사 노트북에서 작성해 남아 있지 않습니다. 근무 이후 같은 업무를
'    합성 데이터로 다시 실습했으나 그 실습 코드도 남아 있지 않아, 실습 자료의
'    목적과 처리 기준을 바탕으로 이 코드를 다시 작성했습니다.
'    (이 파일은 UTF-8입니다. 엑셀에 가져올 때는 excel_import 폴더의 zip 안 CP949 파일을 쓰세요.)
'
' 사용법: 이 모듈을 매크로 통합문서(.xlsm)에 가져온 뒤 [취합실행]을 실행합니다.
'   1) 폴더를 선택합니다.
'   2) 폴더 안의 .xlsx 파일을 하나씩 열어 계정코드가 대상 코드(기본 111300,
'      정기예금)인 행만 추출합니다.
'   3) 결과를 [통합결과] 시트에 모으고, 파일별 행 수·차변·대변 합계를
'      [검증요약] 시트에 남깁니다. 이 요약을 검증기준 파일과 대조합니다.
'
' 처리 기준
'   - 열 위치가 아니라 머리글 이름(계정코드, 차변, 대변 등)으로 열을 찾습니다.
'     파일마다 열 순서가 달라도 동작하게 하기 위한 조건입니다.
'   - 계정코드는 숫자/문자 저장 형식이 섞여도 같은 값으로 비교합니다.
'   - 임시 파일(~$로 시작)과 매크로 파일 자신은 건너뜁니다.
'   - 필요한 머리글이 없는 파일은 건너뛰고 [검증요약]에 사유를 남깁니다.
'==============================================================================
Option Explicit

Private Const TARGET_CODE As String = "111300"
Private Const RESULT_SHEET As String = "통합결과"
Private Const SUMMARY_SHEET As String = "검증요약"

Public Sub 취합실행()
    Dim folderPath As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "A자료 파일이 들어 있는 폴더를 선택하세요"
        If .Show <> -1 Then Exit Sub
        folderPath = .SelectedItems(1)
    End With

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Dim n As Long
    n = 폴더취합(folderPath, TARGET_CODE, ThisWorkbook)
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "취합 완료: " & n & "행 (계정코드 " & TARGET_CODE & ")", vbInformation
End Sub

' 폴더 안의 모든 xlsx 파일에서 대상 계정코드 행을 모아 target 통합문서에 기록한다.
' 반환값: 추출한 총 행 수
Public Function 폴더취합(ByVal folderPath As String, ByVal code As String, _
                        ByVal target As Workbook) As Long
    Dim wsOut As Worksheet, wsSum As Worksheet
    Set wsOut = 시트준비(target, RESULT_SHEET)
    Set wsSum = 시트준비(target, SUMMARY_SHEET)
    wsSum.Range("A1:F1").Value = Array("파일명", "기금명", "추출 행 수", "차변 합계", "대변 합계", "비고")

    If Right$(folderPath, 1) <> Application.PathSeparator Then
        folderPath = folderPath & Application.PathSeparator
    End If

    Dim fileName As String, outRow As Long, sumRow As Long, total As Long
    Dim headerWritten As Boolean
    outRow = 2: sumRow = 2

    fileName = Dir$(folderPath & "*.xlsx")
    Do While Len(fileName) > 0
        If Left$(fileName, 2) <> "~$" And fileName <> target.Name Then
            Dim wb As Workbook, ws As Worksheet
            Set wb = Workbooks.Open(folderPath & fileName, ReadOnly:=True)
            Set ws = wb.Worksheets(1)

            Dim data As Variant, lastRow As Long, lastCol As Long
            lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
            data = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Value

            Dim cCode As Long, cDr As Long, cCr As Long, cFund As Long
            cCode = 열찾기(data, "계정코드")
            cDr = 열찾기(data, "차변")
            cCr = 열찾기(data, "대변")
            cFund = 열찾기(data, "기금명")

            wsSum.Cells(sumRow, 1).Value = fileName
            If cCode = 0 Or cDr = 0 Or cCr = 0 Then
                wsSum.Cells(sumRow, 6).Value = "필수 머리글(계정코드/차변/대변) 없음 - 건너뜀"
            Else
                If Not headerWritten Then
                    Dim h As Long
                    For h = 1 To lastCol
                        wsOut.Cells(1, h).Value = data(1, h)
                    Next h
                    wsOut.Cells(1, lastCol + 1).Value = "원본파일"
                    headerWritten = True
                End If

                Dim r As Long, c As Long, cnt As Long
                Dim sumDr As Double, sumCr As Double
                cnt = 0: sumDr = 0: sumCr = 0
                For r = 2 To lastRow
                    If Trim$(CStr(data(r, cCode))) = code Then
                        For c = 1 To lastCol
                            wsOut.Cells(outRow, c).Value = data(r, c)
                        Next c
                        wsOut.Cells(outRow, lastCol + 1).Value = fileName
                        sumDr = sumDr + Val(data(r, cDr))
                        sumCr = sumCr + Val(data(r, cCr))
                        outRow = outRow + 1
                        cnt = cnt + 1
                    End If
                Next r

                If cFund > 0 And lastRow >= 2 Then wsSum.Cells(sumRow, 2).Value = data(2, cFund)
                wsSum.Cells(sumRow, 3).Value = cnt
                wsSum.Cells(sumRow, 4).Value = sumDr
                wsSum.Cells(sumRow, 5).Value = sumCr
                total = total + cnt
            End If
            sumRow = sumRow + 1
            wb.Close SaveChanges:=False
        End If
        fileName = Dir$()
    Loop

    wsSum.Cells(sumRow, 1).Value = "합계"
    wsSum.Cells(sumRow, 3).Value = total
    If sumRow > 2 Then
        wsSum.Cells(sumRow, 4).Formula = "=SUM(D2:D" & sumRow - 1 & ")"
        wsSum.Cells(sumRow, 5).Formula = "=SUM(E2:E" & sumRow - 1 & ")"
    End If
    폴더취합 = total
End Function

' 시트가 있으면 비우고, 없으면 만든다.
Private Function 시트준비(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = sheetName
    Else
        ws.Cells.Clear
    End If
    Set 시트준비 = ws
End Function

' 1행 머리글에서 이름이 일치하는 열 번호를 찾는다. 없으면 0.
Private Function 열찾기(ByVal data As Variant, ByVal header As String) As Long
    Dim c As Long
    For c = LBound(data, 2) To UBound(data, 2)
        If Trim$(CStr(data(1, c))) = header Then
            열찾기 = c
            Exit Function
        End If
    Next c
    열찾기 = 0
End Function
