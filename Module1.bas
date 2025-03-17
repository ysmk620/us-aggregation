Attribute VB_Name = "尿沈渣集計マクロ"
Option Explicit


Sub 尿沈渣集計()

    Dim ws As Worksheet
    Dim newSheetNames As Variant
    Dim i As Long, lastRow As Long
    Dim wsMaster As Worksheet
    Dim ws総依頼件数 As Worksheet
    Dim ws未受付 As Worksheet, ws総検査件数 As Worksheet
    Dim ws機器判定件数 As Worksheet, ws目視件数 As Worksheet
    Dim ws時間外機器判定件数 As Worksheet, ws時間外機器判定検査不能件数 As Worksheet
    Dim ws中止件数 As Worksheet, ws尿沈渣集計表 As Worksheet

    'sheet1以外のシートを削除
    Application.DisplayAlerts = False
    For Each ws In ActiveWorkbook.Sheets
        If ws.Index <> 1 Then ws.Delete
    Next ws
    Application.DisplayAlerts = True

    'sheet1の名前を「マスタ」に変更
    Set wsMaster = ActiveWorkbook.Sheets(1)
    wsMaster.Name = "マスタ"

    ' 新しいシートを追加
    newSheetNames = Array("尿沈渣集計表", "総依頼件数", "未受付", "総検査件数", _
                          "0；機器判定件数", "2；目視件数", "3；時間外機器判定件数", _
                          "3（06）；時間外機器判定（検査不能）件数", "中止件数")

    For i = LBound(newSheetNames) To UBound(newSheetNames)
        ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count)).Name = newSheetNames(i)
    Next i

    Set ws総依頼件数 = ActiveWorkbook.Sheets("総依頼件数")
    Set ws未受付 = ActiveWorkbook.Sheets("未受付")
    Set ws総検査件数 = ActiveWorkbook.Sheets("総検査件数")
    Set ws機器判定件数 = ActiveWorkbook.Sheets("0；機器判定件数")
    Set ws目視件数 = ActiveWorkbook.Sheets("2；目視件数")
    Set ws時間外機器判定件数 = ActiveWorkbook.Sheets("3；時間外機器判定件数")
    Set ws時間外機器判定検査不能件数 = ActiveWorkbook.Sheets("3（06）；時間外機器判定（検査不能）件数")
    Set ws中止件数 = ActiveWorkbook.Sheets("中止件数")
    Set ws尿沈渣集計表 = ActiveWorkbook.Sheets("尿沈渣集計表")

    '「マスタ」のシートをコピーして「総依頼件数」にコピー
    wsMaster.UsedRange.Copy Destination:=ws総依頼件数.Range("A1")

    '「総依頼件数」のデータ振り分けと誤入力・入力漏れ修正
    With ws総依頼件数
        lastRow = .Cells(.Rows.Count, "D").End(xlUp).Row
        For i = 3 To lastRow
            '受付番号がない行を「未受付」、ある行を「総検査件数」にコピー
            If .Cells(i, "D").Value = "" Then
               .Rows(i).Copy Destination:=ws未受付.Cells(ws未受付.Rows.Count, 1).End(xlUp).Offset(1)
            Else
                .Rows(i).Copy Destination:=ws総検査件数.Cells(ws総検査件数.Rows.Count, 1).End(xlUp).Offset(1)
            End If
            'X列が"検査不能"の場合、Ｖ列に"検査不能"を入力
            If .Cells(i, "X").Value = "検査不能" Then
                .Cells(i, "V").Value = "検査不能"
                .Cells(i, "V").Interior.Color = vbYellow
            End If
            'V列が"検査不能"かつ検査値が空白セルの場合、結果値にはすべて"3"を入力
            If .Cells(i, "V").Value = "検査不能" And .Cells(i, "AC").Value = "" Then
                .Cells(i, "AC").Value = "3"
                .Cells(i, "AC").Interior.Color = vbYellow
            End If
            '結果値が"1"の場合、すべて"3"を入力
            If .Cells(i, "AC").Value = "1" Then
                .Cells(i, "AC").Value = "3"
                .Cells(i, "AC").Interior.Color = vbYellow
            End If
            
        Next i

     .AutoFilterMode = False
       
        lastRow = .Cells(.Rows.Count, "AC").End(xlUp).Row
      
        ' "0"→「0；機器判定件数」にコピー
        .Range("A1:AC" & lastRow).AutoFilter Field:=29, Criteria1:="0"
        .Rows("2:" & lastRow).SpecialCells(xlCellTypeVisible).Copy _
                Destination:=ws機器判定件数.Cells(ws機器判定件数.Rows.Count, 1).End(xlUp).Offset(1)
    
        ' "2"→「2；目視件数」にコピー
        .Range("A1:AC" & lastRow).AutoFilter Field:=29, Criteria1:="2"
        .Rows("2:" & lastRow).SpecialCells(xlCellTypeVisible).Copy _
                Destination:=ws目視件数.Cells(ws目視件数.Rows.Count, 1).End(xlUp).Offset(1)
    
        ' "3"を抽出してV列で振り分ける
        .Range("A1:AC" & lastRow).AutoFilter Field:=29, Criteria1:="3"
    
        ' V列が空白セル→「3；時間外機器判定件数」にコピー
        .Range("A1:AC" & lastRow).AutoFilter Field:=22, Criteria1:="="
        .Rows("2:" & lastRow).SpecialCells(xlCellTypeVisible).Copy _
            Destination:=ws時間外機器判定件数.Cells(ws時間外機器判定件数.Rows.Count, 1).End(xlUp).Offset(1)
    
        ' V列が空白セルではない→「3（06）；時間外機器判定（検査不能）件数」にコピー
        .Range("A1:AC" & lastRow).AutoFilter Field:=22, Criteria1:="<>"
        .Rows("2:" & lastRow).SpecialCells(xlCellTypeVisible).Copy _
            Destination:=ws時間外機器判定検査不能件数.Cells(ws時間外機器判定検査不能件数.Rows.Count, 1).End(xlUp).Offset(1)
    
        ' 空白セルを抽出してV列またはX列が"検査中止"→「中止件数」にコピー
        .AutoFilterMode = False
        .Range("A1:AC" & lastRow).AutoFilter Field:=29, Criteria1:="="
         For i = 2 To lastRow
            If .Cells(i, "V").Value = "検査中止" Or .Cells(i, "X").Value = "検査中止" Then
                .Rows(i).Copy Destination:=ws中止件数.Cells(ws中止件数.Rows.Count, 1).End(xlUp).Offset(1)
            End If
        Next i
        
        .AutoFilterMode = False
        
    End With
    
    
    '尿沈渣集計表作成
    Dim data As Variant
    data = Array( _
            Array("尿沈渣検査状況", "件数"), _
            Array("0：機器判定済み件数", ""), _
            Array("2：目視済み件数", ""), _
            Array("3：時間外機器判定済み件数", ""), _
            Array("3':時間外機器判定（検査不能）件数", ""), _
            Array("検査中止（量不足など）件数", ""), _
            Array("総検査件数", ""), _
            Array("未受付", ""), _
            Array("総依頼件数", "") _
        )
       
    With ws尿沈渣集計表
        For i = LBound(data) To UBound(data)
            .Cells(i + 1, 1).Value = data(i)(0)
            .Cells(i + 1, 2).Value = data(i)(1)
        Next i

        .Range("A1:B1").Interior.Color = RGB(221, 235, 247)
        .Range("A7:B7").Interior.Color = RGB(255, 242, 204)
        .Range("A9:B9").Interior.Color = RGB(226, 239, 218)
        
        .Range("A1:B1").Font.Bold = True
        .Range("A7:B7").Font.Bold = True
        .Range("A9:B9").Font.Bold = True
         
        .Range("A1:B9").Borders.LineStyle = xlContinuous
  
        .Cells(2, 2) = ws機器判定件数.Cells(ws機器判定件数.Rows.Count, 1).End(xlUp).Row - 1
        .Cells(3, 2) = ws目視件数.Cells(ws目視件数.Rows.Count, 1).End(xlUp).Row - 1
        .Cells(4, 2) = ws時間外機器判定件数.Cells(ws時間外機器判定件数.Rows.Count, 1).End(xlUp).Row - 1
        .Cells(5, 2) = ws時間外機器判定検査不能件数.Cells(ws時間外機器判定検査不能件数.Rows.Count, 1).End(xlUp).Row - 1
        .Cells(6, 2) = ws中止件数.Cells(ws中止件数.Rows.Count, 1).End(xlUp).Row - 1
        .Cells(7, 2) = ws総検査件数.Cells(ws総検査件数.Rows.Count, 1).End(xlUp).Row - 1
        .Cells(8, 2) = ws未受付.Cells(ws未受付.Rows.Count, 1).End(xlUp).Row - 1
        .Cells(9, 2) = ws総依頼件数.Cells(ws総依頼件数.Rows.Count, 1).End(xlUp).Row - 2
        
        .Columns("A:B").AutoFit
    
        .Range("B2:B9").HorizontalAlignment = xlCenter
        .Range("B2:B9").VerticalAlignment = xlCenter
        
    End With
    
    Dim FirstDate As String, lastDate As String
    Dim formattedFirstDate As String
    Dim formattedLastDate As String
    lastRow = ws総依頼件数.Cells(ws総依頼件数.Rows.Count, 3).End(xlUp).Row
    
     '先頭と最終行の8桁の文字列を取得
    FirstDate = ws総依頼件数.Cells(3, 3).Value
    lastDate = ws総依頼件数.Cells(lastRow, 3).Value
    
    '年月日を取得
    If Len(FirstDate) = 8 And Len(lastDate) = 8 Then
        formattedFirstDate = Format(DateSerial(Left(FirstDate, 4), Mid(FirstDate, 5, 2), Right(FirstDate, 2)), "yyyy年mm月dd日")
        formattedLastDate = Format(DateSerial(Left(lastDate, 4), Mid(lastDate, 5, 2), Right(lastDate, 2)), "yyyy年mm月dd日")
    End If
    
    '集計期間を入力
    ws尿沈渣集計表.Range("A11").Value = formattedFirstDate & "～" & formattedLastDate
    
    
    
    MsgBox "すべての作業が完了しました！"
    
End Sub








