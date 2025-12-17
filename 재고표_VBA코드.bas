Sub 재고표_생성()
'
' 재고표 자동 생성 매크로
' 정리표 시트 → 재고표_출력 시트 변환
'
    Dim ws정리표 As Worksheet
    Dim ws재고표 As Worksheet
    Dim ws양식 As Worksheet

    Dim lastRow As Long
    Dim i As Long, outRow As Long
    Dim 제품코드 As String
    Dim 제품명 As String
    Dim 아주 As Variant
    Dim 잔량 As Variant
    Dim 소비기한 As Date
    Dim 아주날짜 As String
    Dim 잔량날짜 As String

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' 정리표 시트 확인
    On Error Resume Next
    Set ws정리표 = ThisWorkbook.Sheets("정리표")
    On Error GoTo 0

    If ws정리표 Is Nothing Then
        MsgBox "정리표 시트를 찾을 수 없습니다.", vbCritical
        Exit Sub
    End If

    ' 재고표_출력 시트 삭제 후 재생성
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("재고표_출력").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    ' 새 시트 생성
    Set ws재고표 = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws재고표.Name = "재고표_출력"

    ' 헤더 작성
    With ws재고표
        .Cells(1, 1).Value = "제품코드"
        .Cells(1, 2).Value = "제품명"
        .Cells(1, 3).Value = "아주/날짜"
        .Cells(1, 4).Value = "일반/날짜"
        .Cells(1, 5).Value = "잔량/날짜"

        ' 헤더 서식
        .Range("A1:E1").Font.Bold = True
        .Range("A1:E1").Font.Size = 11
        .Range("A1:E1").Font.Name = "맑은 고딕"
        .Range("A1:E1").HorizontalAlignment = xlCenter
        .Range("A1:E1").VerticalAlignment = xlCenter
        .Range("A1:E1").Interior.Color = RGB(217, 217, 217)
        .Range("A1:E1").Borders.LineStyle = xlContinuous

        ' 열 너비 설정 (재고실사양식표 콜라 시트와 동일)
        .Columns("A").ColumnWidth = 6.75    ' 제품코드
        .Columns("B").ColumnWidth = 23.13   ' 제품명
        .Columns("C").ColumnWidth = 19.88   ' 아주/날짜
        .Columns("D").ColumnWidth = 13      ' 일반/날짜
        .Columns("E").ColumnWidth = 13      ' 잔량/날짜

        ' 헤더 행 높이
        .Rows(1).RowHeight = 20
    End With

    ' 정리표 데이터 읽기
    lastRow = ws정리표.Cells(ws정리표.Rows.Count, "A").End(xlUp).Row
    outRow = 2

    ' 정리표의 2행부터 시작 (1행은 빈 행, 2행이 헤더)
    For i = 3 To lastRow
        ' 제품코드 (A열, 정리표 2열 = B열)
        제품코드 = ws정리표.Cells(i, 2).Value  ' 제품코드 컬럼
        If 제품코드 = "" Then GoTo NextRow

        ' 제품명 (Brand Name, 정리표 4열 = D열)
        제품명 = ws정리표.Cells(i, 4).Value

        ' 아주 (정리표 5열 = E열)
        아주 = ws정리표.Cells(i, 5).Value

        ' 잔량 (정리표 7열 = G열)
        잔량 = ws정리표.Cells(i, 7).Value

        ' 소비기한 (정리표 10열 = J열)
        소비기한 = ws정리표.Cells(i, 10).Value

        ' 아주/날짜 생성
        아주날짜 = ""
        If IsNumeric(아주) And 아주 > 0 And IsDate(소비기한) Then
            아주날짜 = CLng(아주) & "(" & Format(소비기한, "mm/dd") & ")"
        End If

        ' 잔량/날짜 생성
        잔량날짜 = ""
        If IsNumeric(잔량) And 잔량 > 0 And IsDate(소비기한) Then
            잔량날짜 = CLng(잔량) & "(" & Format(소비기한, "mm/dd") & ")"
        End If

        ' 아주 또는 잔량이 있는 경우만 출력
        If 아주날짜 <> "" Or 잔량날짜 <> "" Then
            ws재고표.Cells(outRow, 1).Value = 제품코드
            ws재고표.Cells(outRow, 2).Value = 제품명
            ws재고표.Cells(outRow, 3).Value = 아주날짜
            ws재고표.Cells(outRow, 4).Value = ""  ' 일반/날짜는 비워둠
            ws재고표.Cells(outRow, 5).Value = 잔량날짜

            ' 데이터 서식
            With ws재고표.Rows(outRow)
                .Font.Name = "맑은 고딕"
                .RowHeight = 18
            End With

            ' 제품코드 - 가운데 정렬, 10pt
            With ws재고표.Cells(outRow, 1)
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Font.Size = 10
            End With

            ' 제품명 - 왼쪽 정렬, 10pt
            With ws재고표.Cells(outRow, 2)
                .HorizontalAlignment = xlLeft
                .VerticalAlignment = xlCenter
                .Font.Size = 10
            End With

            ' 아주/날짜 - 가운데 정렬, 12pt 굵게
            With ws재고표.Cells(outRow, 3)
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Font.Size = 12
                .Font.Bold = True
            End With

            ' 일반/날짜 - 가운데 정렬, 12pt 굵게
            With ws재고표.Cells(outRow, 4)
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Font.Size = 12
                .Font.Bold = True
            End With

            ' 잔량/날짜 - 가운데 정렬, 12pt 굵게
            With ws재고표.Cells(outRow, 5)
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Font.Size = 12
                .Font.Bold = True
            End With

            outRow = outRow + 1
        End If

NextRow:
    Next i

    ' 전체 테두리
    With ws재고표.Range("A1:E" & (outRow - 1))
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    MsgBox "재고표 생성 완료!" & vbCrLf & vbCrLf & _
           "총 " & (outRow - 2) & "개 제품", vbInformation, "완료"

    ' 재고표_출력 시트로 이동
    ws재고표.Activate

End Sub
