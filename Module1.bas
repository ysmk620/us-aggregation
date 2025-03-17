Attribute VB_Name = "�A���ԏW�v�}�N��"
Option Explicit


Sub �A���ԏW�v()

    Dim ws As Worksheet
    Dim newSheetNames As Variant
    Dim i As Long, lastRow As Long
    Dim wsMaster As Worksheet
    Dim ws���˗����� As Worksheet
    Dim ws����t As Worksheet, ws���������� As Worksheet
    Dim ws�@�픻�茏�� As Worksheet, ws�ڎ����� As Worksheet
    Dim ws���ԊO�@�픻�茏�� As Worksheet, ws���ԊO�@�픻�茟���s�\���� As Worksheet
    Dim ws���~���� As Worksheet, ws�A���ԏW�v�\ As Worksheet

    'sheet1�ȊO�̃V�[�g���폜
    Application.DisplayAlerts = False
    For Each ws In ActiveWorkbook.Sheets
        If ws.Index <> 1 Then ws.Delete
    Next ws
    Application.DisplayAlerts = True

    'sheet1�̖��O���u�}�X�^�v�ɕύX
    Set wsMaster = ActiveWorkbook.Sheets(1)
    wsMaster.Name = "�}�X�^"

    ' �V�����V�[�g��ǉ�
    newSheetNames = Array("�A���ԏW�v�\", "���˗�����", "����t", "����������", _
                          "0�G�@�픻�茏��", "2�G�ڎ�����", "3�G���ԊO�@�픻�茏��", _
                          "3�i06�j�G���ԊO�@�픻��i�����s�\�j����", "���~����")

    For i = LBound(newSheetNames) To UBound(newSheetNames)
        ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count)).Name = newSheetNames(i)
    Next i

    Set ws���˗����� = ActiveWorkbook.Sheets("���˗�����")
    Set ws����t = ActiveWorkbook.Sheets("����t")
    Set ws���������� = ActiveWorkbook.Sheets("����������")
    Set ws�@�픻�茏�� = ActiveWorkbook.Sheets("0�G�@�픻�茏��")
    Set ws�ڎ����� = ActiveWorkbook.Sheets("2�G�ڎ�����")
    Set ws���ԊO�@�픻�茏�� = ActiveWorkbook.Sheets("3�G���ԊO�@�픻�茏��")
    Set ws���ԊO�@�픻�茟���s�\���� = ActiveWorkbook.Sheets("3�i06�j�G���ԊO�@�픻��i�����s�\�j����")
    Set ws���~���� = ActiveWorkbook.Sheets("���~����")
    Set ws�A���ԏW�v�\ = ActiveWorkbook.Sheets("�A���ԏW�v�\")

    '�u�}�X�^�v�̃V�[�g���R�s�[���āu���˗������v�ɃR�s�[
    wsMaster.UsedRange.Copy Destination:=ws���˗�����.Range("A1")

    '�u���˗������v�̃f�[�^�U�蕪���ƌ���́E���͘R��C��
    With ws���˗�����
        lastRow = .Cells(.Rows.Count, "D").End(xlUp).Row
        For i = 3 To lastRow
            '��t�ԍ����Ȃ��s���u����t�v�A����s���u�����������v�ɃR�s�[
            If .Cells(i, "D").Value = "" Then
               .Rows(i).Copy Destination:=ws����t.Cells(ws����t.Rows.Count, 1).End(xlUp).Offset(1)
            Else
                .Rows(i).Copy Destination:=ws����������.Cells(ws����������.Rows.Count, 1).End(xlUp).Offset(1)
            End If
            'X��"�����s�\"�̏ꍇ�A�u���"�����s�\"�����
            If .Cells(i, "X").Value = "�����s�\" Then
                .Cells(i, "V").Value = "�����s�\"
                .Cells(i, "V").Interior.Color = vbYellow
            End If
            'V��"�����s�\"�������l���󔒃Z���̏ꍇ�A���ʒl�ɂ͂��ׂ�"3"�����
            If .Cells(i, "V").Value = "�����s�\" And .Cells(i, "AC").Value = "" Then
                .Cells(i, "AC").Value = "3"
                .Cells(i, "AC").Interior.Color = vbYellow
            End If
            '���ʒl��"1"�̏ꍇ�A���ׂ�"3"�����
            If .Cells(i, "AC").Value = "1" Then
                .Cells(i, "AC").Value = "3"
                .Cells(i, "AC").Interior.Color = vbYellow
            End If
            
        Next i

     .AutoFilterMode = False
       
        lastRow = .Cells(.Rows.Count, "AC").End(xlUp).Row
      
        ' "0"���u0�G�@�픻�茏���v�ɃR�s�[
        .Range("A1:AC" & lastRow).AutoFilter Field:=29, Criteria1:="0"
        .Rows("2:" & lastRow).SpecialCells(xlCellTypeVisible).Copy _
                Destination:=ws�@�픻�茏��.Cells(ws�@�픻�茏��.Rows.Count, 1).End(xlUp).Offset(1)
    
        ' "2"���u2�G�ڎ������v�ɃR�s�[
        .Range("A1:AC" & lastRow).AutoFilter Field:=29, Criteria1:="2"
        .Rows("2:" & lastRow).SpecialCells(xlCellTypeVisible).Copy _
                Destination:=ws�ڎ�����.Cells(ws�ڎ�����.Rows.Count, 1).End(xlUp).Offset(1)
    
        ' "3"�𒊏o����V��ŐU�蕪����
        .Range("A1:AC" & lastRow).AutoFilter Field:=29, Criteria1:="3"
    
        ' V�񂪋󔒃Z�����u3�G���ԊO�@�픻�茏���v�ɃR�s�[
        .Range("A1:AC" & lastRow).AutoFilter Field:=22, Criteria1:="="
        .Rows("2:" & lastRow).SpecialCells(xlCellTypeVisible).Copy _
            Destination:=ws���ԊO�@�픻�茏��.Cells(ws���ԊO�@�픻�茏��.Rows.Count, 1).End(xlUp).Offset(1)
    
        ' V�񂪋󔒃Z���ł͂Ȃ����u3�i06�j�G���ԊO�@�픻��i�����s�\�j�����v�ɃR�s�[
        .Range("A1:AC" & lastRow).AutoFilter Field:=22, Criteria1:="<>"
        .Rows("2:" & lastRow).SpecialCells(xlCellTypeVisible).Copy _
            Destination:=ws���ԊO�@�픻�茟���s�\����.Cells(ws���ԊO�@�픻�茟���s�\����.Rows.Count, 1).End(xlUp).Offset(1)
    
        ' �󔒃Z���𒊏o����V��܂���X��"�������~"���u���~�����v�ɃR�s�[
        .AutoFilterMode = False
        .Range("A1:AC" & lastRow).AutoFilter Field:=29, Criteria1:="="
         For i = 2 To lastRow
            If .Cells(i, "V").Value = "�������~" Or .Cells(i, "X").Value = "�������~" Then
                .Rows(i).Copy Destination:=ws���~����.Cells(ws���~����.Rows.Count, 1).End(xlUp).Offset(1)
            End If
        Next i
        
        .AutoFilterMode = False
        
    End With
    
    
    '�A���ԏW�v�\�쐬
    Dim data As Variant
    data = Array( _
            Array("�A���Ԍ�����", "����"), _
            Array("0�F�@�픻��ς݌���", ""), _
            Array("2�F�ڎ��ς݌���", ""), _
            Array("3�F���ԊO�@�픻��ς݌���", ""), _
            Array("3':���ԊO�@�픻��i�����s�\�j����", ""), _
            Array("�������~�i�ʕs���Ȃǁj����", ""), _
            Array("����������", ""), _
            Array("����t", ""), _
            Array("���˗�����", "") _
        )
       
    With ws�A���ԏW�v�\
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
  
        .Cells(2, 2) = ws�@�픻�茏��.Cells(ws�@�픻�茏��.Rows.Count, 1).End(xlUp).Row - 1
        .Cells(3, 2) = ws�ڎ�����.Cells(ws�ڎ�����.Rows.Count, 1).End(xlUp).Row - 1
        .Cells(4, 2) = ws���ԊO�@�픻�茏��.Cells(ws���ԊO�@�픻�茏��.Rows.Count, 1).End(xlUp).Row - 1
        .Cells(5, 2) = ws���ԊO�@�픻�茟���s�\����.Cells(ws���ԊO�@�픻�茟���s�\����.Rows.Count, 1).End(xlUp).Row - 1
        .Cells(6, 2) = ws���~����.Cells(ws���~����.Rows.Count, 1).End(xlUp).Row - 1
        .Cells(7, 2) = ws����������.Cells(ws����������.Rows.Count, 1).End(xlUp).Row - 1
        .Cells(8, 2) = ws����t.Cells(ws����t.Rows.Count, 1).End(xlUp).Row - 1
        .Cells(9, 2) = ws���˗�����.Cells(ws���˗�����.Rows.Count, 1).End(xlUp).Row - 2
        
        .Columns("A:B").AutoFit
    
        .Range("B2:B9").HorizontalAlignment = xlCenter
        .Range("B2:B9").VerticalAlignment = xlCenter
        
    End With
    
    Dim FirstDate As String, lastDate As String
    Dim formattedFirstDate As String
    Dim formattedLastDate As String
    lastRow = ws���˗�����.Cells(ws���˗�����.Rows.Count, 3).End(xlUp).Row
    
     '�擪�ƍŏI�s��8���̕�������擾
    FirstDate = ws���˗�����.Cells(3, 3).Value
    lastDate = ws���˗�����.Cells(lastRow, 3).Value
    
    '�N�������擾
    If Len(FirstDate) = 8 And Len(lastDate) = 8 Then
        formattedFirstDate = Format(DateSerial(Left(FirstDate, 4), Mid(FirstDate, 5, 2), Right(FirstDate, 2)), "yyyy�Nmm��dd��")
        formattedLastDate = Format(DateSerial(Left(lastDate, 4), Mid(lastDate, 5, 2), Right(lastDate, 2)), "yyyy�Nmm��dd��")
    End If
    
    '�W�v���Ԃ����
    ws�A���ԏW�v�\.Range("A11").Value = formattedFirstDate & "�`" & formattedLastDate
    
    
    
    MsgBox "���ׂĂ̍�Ƃ��������܂����I"
    
End Sub








