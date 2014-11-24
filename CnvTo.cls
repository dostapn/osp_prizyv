VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CnvTo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim oExcelApp    As Object
Dim appwd   As Object


Dim oWs             As Object
Dim oWb             As Object
Dim oWa             As Object

Const P_NUM As String = "#NUM#"
Const P_OBL_KOM As String = "$OBL_KOM$"
Const P_DAY As String = "#DAY#"
Const P_MONTH As String = "#MONTH#"
Const P_YEAR As String = "#YEAR#"

''Именной список
Public Function To_F27(ID As Integer)
On Error Resume Next


Dim q As Long
Dim lichn() As String
Dim connt As Boolean
Dim type_vs As String
type_vs = frmCommand.listcommands.ListItems(frmCommand.listcommands.SelectedItem.Index).SubItems(2)
lichn = Split(get_config("print_lichnom"), ";")
For q = 0 To UBound(lichn())
    If lichn(q) = type_vs Then connt = True
Next q






    Set oExcelApp = CreateObject("EXCEL.APPLICATION")
    
        If Err = 429 Then MsgBox "Application Microsoft Excel is not installed!", vbExclamation, strMAIN_TITLE: Screen.MousePointer = vbDefault: Exit Function

    Dim sFile As String
    sFile = sCNV_txtDirShabl & "F27.xls"
    
    If Not ExFile(sFile) Then MsgBox "Файл шаблона" & NL2 & sFile & NL2 & "не найден. Проверьте правильность указания пути шаблонов." & vbNewLine & "Для восстановления шаблонов воспользуйтесь вкладкой `Воссановление` в меню `Опции`", vbCritical, strMAIN_TITLE: oExcelApp.Quit: Screen.MousePointer = vbDefault: Exit Function
    If get_config("view_print") = 1 Then oExcelApp.Visible = False Else oExcelApp.Visible = True
    
    
    oExcelApp.WORkbooks.Open FileName:=sFile, ReadOnly:=True, ignOReReadOnlyRecommended:=True
    
    Set oWb = oExcelApp.ActiveWORkbook
    Set oWs = oExcelApp.Sheets(1)


Dim BUF() As String
Dim pr() As String
Dim cnt As Long
Dim s As Long
Dim cUpk As Long
Dim STROKA As Long
Dim c As Long
Dim t As String
Dim x As Long
Dim LST As Long
Dim CC As Long
Dim dUpk()  As Long
'                                    1                           2                            3                          4                            5                            6                               7                            8                           9                            10                             11                      12                           13                                14
Call mysql.query("SELECT prnik_" & nowBase & ".fam, prnik_" & nowBase & ".name, prnik_" & nowBase & ".otch, prnik_" & nowBase & ".datar, prnik_" & nowBase & ".txtvk, naryad_" & nowBase & ".oblkom, naryad_" & nowBase & ".rodv, naryad_" & nowBase & ".punkt, naryad_" & nowBase & ".vch, naryad_" & nowBase & ".doroga, prnik_" & nowBase & ".vus, prnik_" & nowBase & ".dir, prnik_" & nowBase & ".vod, otpravka_" & nowBase & ".otpravkaid, prnik_" & nowBase & ".vus_p FROM prnik_" & nowBase & ", otpravka_" & nowBase & ", naryad_" & nowBase & " WHERE prnik_" & nowBase & ".otprvid = otpravka_" & nowBase & ".otpravkaid AND otpravka_" & nowBase & ".narid = naryad_" & nowBase & ".narid AND otpravka_" & nowBase & ".otpravkaid = " & CStr(ID) & " ORder by txtvk, fam")
s = 11
cnt = st
            With oWs
            
            
            
Dim j As Long
    j = 1
    Do While j < cnt
.Rows("11:13").Copy
.Rows("11:11").Insert
j = j + 1
    Loop
    
            .cells(s, 2) = 1
            .cells(s, 3) = dat(1, 1)
            .cells(s + 1, 3) = dat(2, 1) & " " & dat(3, 1)
            If connt = True Then .cells(s + 2, 3) = dat(15, 1)
            .cells(s, 4) = CnvDataSqLToWin(dat(4, 1))
            .cells(s, 5) = dat(5, 1)
            .cells(s, 6) = dat(11, 1)
             If CLng(dat(13, 1)) = 1 Then
             .cells(s, 7) = KOD_VOD

             End If
If (dat(11, 1) = "Вод. С" Or dat(11, 1) = "Вод. C") And dat(13, 1) = 1 Then .cells(s, 7) = "837"
If dat(11, 1) = "Вод. Д" And dat(13, 1) = 1 Then .cells(s, 7) = "845"
If (dat(11, 1) = "Вод. E" Or dat(11, 1) = "Вод. Е") And dat(13, 1) = 1 Then .cells(s, 7) = "846"
If dat(11, 1) = "Вод. АПА" And dat(13, 1) = 1 Then .cells(s, 7) = "283"
If dat(11, 1) = "Вод. ЭМ" And dat(13, 1) = 1 Then .cells(s, 7) = "837 ЭМ"
            .cells(s, 9) = dat(12, 1)
            
            s = s + 3

        For CC = 2 To cnt
            
            'cUpk = CLng(VAl(dUpk(CC - 1)))
            'Call GetInfo(cUpk)
            c = c + 1
          '  .cells(s, 2).SELECT
            .cells(s, 2) = CC
            .cells(s, 3) = dat(1, CC)
            .cells(s + 1, 3) = dat(2, CC) & " " & dat(3, CC)
            If connt = True Then .cells(s + 2, 3) = dat(15, CC)
            .cells(s, 4) = CnvDataSqLToWin(dat(4, CC))
            .cells(s, 5) = dat(5, CC)
            .cells(s, 6) = dat(11, CC)
            If CLng(dat(13, CC)) = 1 Then
            .cells(s, 7) = KOD_VOD
End If
If (dat(11, CC) = "Вод. С" Or dat(11, CC) = "Вод. C") And dat(13, CC) = 1 Then .cells(s, 7) = "837"
If dat(11, CC) = "Вод. Д" And dat(13, CC) = 1 Then .cells(s, 7) = "845"
If (dat(11, CC) = "Вод. E" Or dat(11, CC) = "Вод. Е") And dat(13, CC) = 1 Then .cells(s, 7) = "846"
If dat(11, CC) = "Вод. АПА" And dat(13, CC) = 1 Then .cells(s, 7) = "283"
If dat(11, CC) = "Вод. ЭМ" And dat(13, CC) = 1 Then .cells(s, 7) = "837 ЭМ"
            .cells(s, 9) = dat(12, CC)

            s = s + 3

Next CC

  
.cells(4, 9).FormulaR1C1 = CnvDataSqLToWin(dateotp)


.cells(5, 3) = Replace(.cells(5, 3), "##", dat(6, 1))
.cells(6, 3) = Replace(.cells(6, 3), "##########", dat(7, 1))
.cells(7, 3) = Replace(.cells(7, 3), "######", dat(7, 1))
.cells(7, 3) = Replace(.cells(7, 3), "%PUNKT%", dat(8, 1))
.cells(7, 3) = Replace(.cells(7, 3), "#####", dat(9, 1))
.cells(7, 3) = Replace(.cells(7, 3), "$$$$$$$$$$", dat(10, 1))

'.Cells.EntireColumn.AutoFit

If CStr(cnt) Like "*1" Then
    If Not CStr(cnt) Like "*11" Then
        .cells(s + 5, 1) = Replace(.cells(s + 5, 1), "человек", "человека")
    End If
End If

.cells(2, 9) = "Исх. №2/" & dat(14, 1)

End With
''''''''Разметка страницы
Dim CheckRow As Long
Dim countPB As Long
If Int(oWs.cells(9, 10)) > 6 Then oWs.PageSetup.Zoom = 95
If Int(oWs.cells(9, 10)) > 7 Then oWs.PageSetup.Zoom = 90
If Int(oWs.cells(9, 10)) > 8 Then oWs.PageSetup.Zoom = 85
If Int(oWs.cells(9, 10)) > 9 Then oWs.PageSetup.Zoom = 81
CheckRow = Int((Int(oWs.cells(9, 10)) * 3) + 8)
    oWs.ResetAllPageBreaks
    countPB = oWs.HPageBreaks.Count
    If countPB > 1 Then
        If oWs.HPageBreaks(countPB).Location.row > CheckRow Then
        oWs.HPageBreaks.add Before:=oWs.Rows(CheckRow)
        End If
    End If
    
If Int(oWs.cells(9, 10)) > 6 Then oWs.PageSetup.Zoom = 95
If Int(oWs.cells(9, 10)) > 7 Then oWs.PageSetup.Zoom = 90
If Int(oWs.cells(9, 10)) > 8 Then oWs.PageSetup.Zoom = 85
If Int(oWs.cells(9, 10)) > 9 Then oWs.PageSetup.Zoom = 81
'''''''''''''''''''''''''
  If frmCommand.listln.ListItems(9).Checked = True Then
If Int(oWs.cells(9, 10)) <= 10 Then
oExcelApp.Application.Run "F27.xls!cmdPRINT"
oExcelApp.Application.DisplayAlerts = False
        oWb.Close
        oExcelApp.Quit
Else
If Int(oWs.cells(9, 10)) = 11 Or Int(oWs.cells(9, 10)) = 12 Or Int(oWs.cells(9, 10)) = 13 Or Int(oWs.cells(9, 10)) = 14 Or Int(oWs.cells(9, 10)) = 15 Then
oWs.HPageBreaks.add Before:=oWs.Rows(CheckRow)
End If
oExcelApp.Visible = True
Set oExcelApp = Nothing
Screen.MousePointer = vbDefault
End If
Else
oExcelApp.Visible = True

Set oExcelApp = Nothing
Screen.MousePointer = vbDefault
End If
End Function



''Накладные, Справки, Извещения
Public Function naksprav(ID As Integer)
On Error Resume Next


    Set oExcelApp = CreateObject("EXCEL.APPLICATION")
    
        If Err = 429 Then MsgBox "Application Microsoft Excel is not installed!", vbExclamation, strMAIN_TITLE: Screen.MousePointer = vbDefault: Exit Function

    Dim sFile As String
    sFile = sCNV_txtDirShabl & "naksprav.xlsm"
    
    If Not ExFile(sFile) Then MsgBox "Файл шаблона" & NL2 & sFile & NL2 & "не найден. Проверьте правильность указания пути шаблонов." & vbNewLine & "Для восстановления шаблонов воспользуйтесь вкладкой `Воссановление` в меню `Опции`", vbCritical, strMAIN_TITLE: oExcelApp.Quit: Screen.MousePointer = vbDefault: Exit Function
    
 If get_config("view_print") = 1 Then oExcelApp.Visible = False Else oExcelApp.Visible = True

    oExcelApp.WORkbooks.Open FileName:=sFile, ReadOnly:=True, ignOReReadOnlyRecommended:=True
    
    Set oWb = oExcelApp.ActiveWORkbook
    Set oWs = oExcelApp.Sheets(1)




            
Call mysql.query("SELECT otpravka_" & nowBase & ".kolvo, naryad_" & nowBase & ".rodv FROM otpravka_" & nowBase & ", naryad_" & nowBase & " WHERE otpravka_" & nowBase & ".otpravkaid = " & CStr(ID) & " and otpravka_" & nowBase & ".narid = naryad_" & nowBase & ".narid")
datt() = dat()
oExcelApp.Sheets(1).cells(1, 1) = datt(1, 1)
  If frmCommand.listln.ListItems(8).Checked = True Then
If CStr(datt(2, 1)) Like "*МВД*" Then
oExcelApp.Application.Run "naksprav.xlsm!СПН"
End If

If CStr(datt(2, 1)) Like "*ПП ФСО*" Then
oExcelApp.Application.Run "naksprav.xlsm!ФСО"
End If

If Not datt(2, 1) Like "*ПП ФСО*" And Not datt(2, 1) Like "*МВД*" Then
oExcelApp.Application.Run "naksprav.xlsm!Общ"
End If


oExcelApp.Application.DisplayAlerts = False
        oWb.Close
        oExcelApp.Quit
        Else
Call mysql.query("SELECT otpravka_" & nowBase & ".kolvo, naryad_" & nowBase & ".rodv FROM otpravka_" & nowBase & ", naryad_" & nowBase & " WHERE otpravka_" & nowBase & ".otpravkaid = " & CStr(ID) & " and otpravka_" & nowBase & ".narid = naryad_" & nowBase & ".narid")

oExcelApp.Visible = True
End If
End Function

Sub GetInfo(VAlUpk)

    
    Dim x As Long
    Call mysql.query("SELECT otpravka_" & nowBase & ".data, naryad_" & nowBase & ".oblkom, naryad_" & nowBase & ".rodv, naryad_" & nowBase & ".okr, naryad_" & nowBase & ".punkt, naryad_" & nowBase & ".vch, naryad_" & nowBase & ".doroga, prnik_" & nowBase & ".fam, prnik_" & nowBase & ".name, prnik_" & nowBase & ".otch, prnik_" & nowBase & ".datar, prnik_" & nowBase & ".txtvk, naryad_" & nowBase & ".rodv, prnik_" & nowBase & ".servb, prnik_" & nowBase & ".nomvb, prnik_" & nowBase & ".gol, prnik_" & nowBase & ".nog, prnik_" & nowBase & ".fORmaa, prnik_" & nowBase & ".fORmab FROM otpravka_" & nowBase & ", naryad_" & nowBase & ", prnik_" & nowBase & " WHERE prnik_" & nowBase & ".idprnik = '" & VAlUpk & "' AND prnik_" & nowBase & ".otprvid=otpravka_" & nowBase & ".otpravkaid AND otpravka_" & nowBase & ".narid=naryad_" & nowBase & ".narid;")
    For x = 1 To 19
        MAS(x - 1) = dat(x, 1)
    Next x
End Sub

Public Function To_F8(ID As Integer, strOblKom As String)
'Раздаточная ведомость
On Error Resume Next
    Set oExcelApp = CreateObject("EXCEL.APPLICATION")
    If Err = 429 Then MsgBox "Application is Microsoft Excel is not installed!", vbExclamation, strMAIN_TITLE: Screen.MousePointer = vbDefault: Exit Function
    Dim sFile As String
    If Right$(nowBase, 1) = "1" Then
    oExcelApp.WORkbooks.Open FileName:=sCNV_txtDirShabl & "F801.xls", ReadOnly:=True, ignOReReadOnlyRecommended:=True
    Else
    oExcelApp.WORkbooks.Open FileName:=sCNV_txtDirShabl & "F802.xls", ReadOnly:=True, ignOReReadOnlyRecommended:=True
    End If
    
    Set oWb = oExcelApp.ActiveWORkbook
    Set oWs = oExcelApp.Sheets(1)
    
    If get_config("view_print") = 1 Then oExcelApp.Visible = False Else oExcelApp.Visible = True
    
        With oWs
                   
                    Dim BUF() As String
                    Dim pr() As String
                    Dim cnt As Long
                    Dim s As Long
                    Dim cUpk As Long
                    Dim STROKA As Long
                    Dim c As Long
                    Dim t As String
                    Dim x As Long
                    Dim LST As Long
                    Dim CC As Long
                    Dim dUpk()  As Long
                    Call mysql.query("SELECT concat(fam,' ',name,' ',otch) as fam, txtvk, concat(servb,' ',nomvb) FROM prnik_" & nowBase & " WHERE prnik_" & nowBase & ".otprvid = " & ID & " ORder by txtvk,fam")
                
                    s = 20
                    cnt = st
                
                                If Right$(nowBase, 1) = "2" Then .cells(20, 25) = dateotp
                                If Right$(nowBase, 1) = "1" Then .cells(20, 33) = dateotp
Dim j As Long
j = 1
Do While j < cnt
            .Rows("20:20").Copy
            .Rows("20:20").Insert
j = j + 1
    Loop
                                
                                .cells(s, 1) = 1
                                .cells(s, 2) = dat(1, 1)
                                .cells(s, 3) = dat(3, 1)

                               
                                
                                s = s + 1
                                
            
                    For CC = 2 To cnt


                                .cells(s, 1) = CC
                                .cells(s, 2) = dat(1, CC)
                                .cells(s, 3) = dat(3, CC)
                                
                                s = s + 1
                    
                    Next CC
                        
                        
                       

End With


Dim CheckRow As Long
Dim countPB As Long

CheckRow = Int((oWs.cells(18, 28)) + 19)

    oWs.ResetAllPageBreaks
    countPB = oWs.HPageBreaks.Count
    If countPB > 0 Then
        If oWs.HPageBreaks(countPB).Location.row > CheckRow Then
        oWs.HPageBreaks.add Before:=oWs.Rows(CheckRow)
        End If
    End If

  If frmCommand.listln.ListItems(7).Checked = True Then
       oExcelApp.ActiveWindow.SelectedSheets.PrintOut Copies:=2, Collate:=True
       oExcelApp.Application.DisplayAlerts = False
       oWb.Close
       oExcelApp.Quit
Else
oExcelApp.Visible = True
Set oExcelApp = Nothing
Screen.MousePointer = vbDefault
End If
oExcelApp.Visible = True
Set oExcelApp = Nothing
Screen.MousePointer = vbDefault

Exit Function

End Function

Public Function To_vedomostj_krugki(ID As Integer, strOblKom As String, autom As Boolean)
'Кружки ложки
On Error Resume Next

    Set oExcelApp = CreateObject("EXCEL.APPLICATION")
    If Err = 429 Then MsgBox "Application is Microsoft Excel is not installed!", vbExclamation, strMAIN_TITLE: Screen.MousePointer = vbDefault: Exit Function
    
    
        Dim sFile As String
    sFile = sCNV_txtDirShabl & "F82.xls"
    If Not ExFile(sFile) Then MsgBox "Файл шаблона" & NL2 & sFile & NL2 & "не найден. Проверьте правильность указания пути шаблонов." & vbNewLine & "Для восстановления шаблонов воспользуйтесь вкладкой `Воссановление` в меню `Опции`", vbCritical, strMAIN_TITLE: Screen.MousePointer = vbDefault: Exit Function

    
    If get_config("view_print") = 1 Then oExcelApp.Visible = False Else oExcelApp.Visible = True
    
    oExcelApp.WORkbooks.Open FileName:=sCNV_txtDirShabl & "F82.xls", ReadOnly:=True, ignOReReadOnlyRecommended:=True
    
    Set oWb = oExcelApp.ActiveWORkbook
    Set oWs = oExcelApp.Sheets(1)
    
    
    
        With oWs
                   
                    Dim BUF() As String
                    Dim pr() As String
                    Dim cnt As Long
                    Dim s As Long
                    Dim cUpk As Long
                    Dim STROKA As Long
                    Dim c As Long
                    Dim t As String
                    Dim x As Long
                    Dim LST As Long
                    Dim CC As Long
                    Dim dUpk()  As Long
                    Call mysql.query("SELECT idprnik, fam, name, otch, txtvk FROM prnik_" & nowBase & " WHERE prnik_" & nowBase & ".otprvid = " & ID & " ORder by txtvk, fam")
                
                    s = 18
                    cnt = st
                   
                        ReDim dUpk(VAl(st))
                        For x = 1 To cnt
                                dUpk(x - 1) = dat(1, x)
                        Next x
                    
                                cUpk = CLng(VAl(dUpk(0)))
                                Call GetInfo(cUpk)
                                .calls(s, 1).Select
                                .cells(s, 1) = "1"
                                .cells(s, 2) = dat(8, 1) & " " & dat(9, 1) & " " & dat(10, 1)
                                .cells(s, 3) = UCase$(dat(14, 1)) & " " & dat(15, 1)
                                .cells(s, 4) = "1"
                                .cells(s, 8) = "1"
                                .cells(s, 11) = dateotp
                                
                                s = s + 1
                                c = 1
                        
                                c = 1

            If cnt > 1 Then .range("A" & s & ":" & "A" & (s + (cnt - 2))).Select: oExcelApp.Selection.EntireRow.Insert
            
                    For CC = 2 To cnt


                                cUpk = CLng(VAl(dUpk(CC - 1)))
                                Call GetInfo(cUpk)
                                c = c + 1
                              
                                .cells(s, 1) = c
                                .cells(s, 2) = dat(8, 1) & " " & dat(9, 1) & " " & dat(10, 1)
                                .cells(s, 3) = UCase$(dat(14, 1)) & " " & dat(15, 1)
                                .cells(s, 4) = "1"
                                .cells(s, 8) = "1"
                                .cells(s, 11) = dateotp
                                .range("D" & s & ":" & "G" & s).Select: oExcelApp.Selection.Merge
                                .range("H" & s & ":" & "J" & s).Select: oExcelApp.Selection.Merge
                                .range("K" & s & ":" & "L" & s).Select: oExcelApp.Selection.Merge
                                .range("M" & s & ":" & "N" & s).Select: oExcelApp.Selection.Merge
                                .range("O" & s & ":" & "P" & s).Select: oExcelApp.Selection.Merge
                                s = s + 1
                    
                    Next CC
   
                    
                    
                                    For x = 4 To 21
                                            Select Case x
                                                Case 4
                                                    .cells(s, x) = cnt
                                                Case 8
                                                    .cells(s, x) = cnt
                                                
                                            End Select
                                    Next x
                        .cells(4, 2) = strOblKom
End With


Dim CheckRow As Long
Dim countPB As Long

CheckRow = Int((oWs.cells(16, 17)) + 17)

    oWs.ResetAllPageBreaks
    countPB = oWs.HPageBreaks.Count
    If countPB > 0 Then
        If oWs.HPageBreaks(countPB).Location.row > CheckRow Then
        oWs.HPageBreaks.add Before:=oWs.Rows(CheckRow)
        End If
    End If

                                    
  If frmCommand.listln.ListItems(6).Checked = True Then
        oExcelApp.ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
        oExcelApp.Application.DisplayAlerts = False
        oWb.Close
        oExcelApp.Quit
        Else
oExcelApp.Visible = True
Set oExcelApp = Nothing
Screen.MousePointer = vbDefault
    End If


Exit Function

End Function
Public Function to_godr()
''''''''''''''''''''''''''''' Главная - Отчет по годам рождения
On Error Resume Next
   Dim stt As Long
    Set oExcelApp = CreateObject("EXCEL.APPLICATION")
    If Err = 429 Then MsgBox "Application is Microsoft Excel is not installed!", vbExclamation, strMAIN_TITLE: Screen.MousePointer = vbDefault: Exit Function
    Dim sFile As String
    sFile = sCNV_txtDirShabl & "empty.xls"
    If Not ExFile(sFile) Then MsgBox "Файл шаблона" & NL2 & sFile & NL2 & "не найден. Проверьте правильность указания пути шаблонов." & vbNewLine & "Для восстановления шаблонов воспользуйтесь вкладкой `Воссановление` в меню `Опции`", vbCritical, strMAIN_TITLE: Screen.MousePointer = vbDefault: Exit Function
    oExcelApp.Visible = True
    oExcelApp.WORkbooks.Open FileName:=sFile, ReadOnly:=True, ignOReReadOnlyRecommended:=True
    Set oWb = oExcelApp.ActiveWORkbook
    Set oWs = oExcelApp.Sheets(1)
   c = 1
   With oWs
   Call mysql.query("select left(datar,4) from prnik_" & nowBase & " group by left(datar,4) order by left(datar,4)")
   datt() = dat()
   stt = st
   For x = 1 To stt
   .cells(1, x + 1) = datt(1, x)
   Next x
   Call sorting_blok("txtvk", "prnik_" & nowBase)
   For x = 1 To UBound(outm())
    .cells(x + 1, 1) = outm(x)
    For Y = 1 To stt
        Call mysql.query("select count(idprnik) from prnik_" & nowBase & " where datar like '" & datt(1, Y) & "%' and txtvk='" & outm(x) & "'")
        .cells(x + 1, Y + 1) = dat(1, 1)
    Next Y
   Next x
         
            
            
            .range(.cells(1, 1), .cells(x, Y)).Borders(7).LineStyle = 1
            .range(.cells(1, 1), .cells(x, Y)).Borders(8).LineStyle = 1
            .range(.cells(1, 1), .cells(x, Y)).Borders(9).LineStyle = 1
            .range(.cells(1, 1), .cells(x, Y)).Borders(10).LineStyle = 1
            .range(.cells(1, 1), .cells(x, Y)).Borders(11).LineStyle = 1
            .range(.cells(1, 1), .cells(x, Y)).Borders(12).LineStyle = 1
            .cells.EntireColumn.AutoFit
End With
Set oExcelApp = Nothing
End Function
Public Function To_prim()
On Error Resume Next
    Set oExcelApp = CreateObject("EXCEL.APPLICATION")
    If Err = 429 Then MsgBox "Application is Microsoft Excel is not installed!", vbExclamation, strMAIN_TITLE: Screen.MousePointer = vbDefault: Exit Function
    Dim sFile As String
    sFile = sCNV_txtDirShabl & "prim.xls"
    If Not ExFile(sFile) Then MsgBox "Файл шаблона" & NL2 & sFile & NL2 & "не найден. Проверьте правильность указания пути шаблонов." & vbNewLine & "Для восстановления шаблонов воспользуйтесь вкладкой `Воссановление` в меню `Опции`", vbCritical, strMAIN_TITLE: Screen.MousePointer = vbDefault: Exit Function
    oExcelApp.Visible = True
    oExcelApp.WORkbooks.Open FileName:=sFile, ReadOnly:=True, ignOReReadOnlyRecommended:=True
    Set oWb = oExcelApp.ActiveWORkbook
    Set oWs = oExcelApp.Sheets(1)
   c = 1
   With oWs
      Call mysql.query("SELECT okrkom,okrkom_e,punkt,vch,prim FROM naryad_" & nowBase & " ORDER BY okrkom,okrkom_e,oblkom")
                    Dim stt As Long
                    stt = st
                    datt() = dat()
                    For x = 1 To stt
                        If Len(datt(5, x)) > 0 Then
                            .cells(1 + c, 1) = c
                            .cells(1 + c, 2) = datt(1, x) & datt(2, x)
                            .cells(1 + c, 3) = datt(3, x)
                            .cells(1 + c, 4) = datt(4, x)
                            .cells(1 + c, 5) = datt(5, x)
                            c = c + 1
                         End If
                 
                    Next x
            
            
            .range("A1:E" & c).Borders(7).LineStyle = 1
            .range("A1:E" & c).Borders(8).LineStyle = 1
            .range("A1:E" & c).Borders(9).LineStyle = 1
            .range("A1:E" & c).Borders(10).LineStyle = 1
            .range("A1:E" & c).Borders(11).LineStyle = 1
            .range("A1:E" & c).Borders(12).LineStyle = 1
            .cells.EntireColumn.AutoFit
End With
Set oExcelApp = Nothing
End Function
'Личный номер жетоны
Public Function To_vedomostj_lichnomer(ID As Integer, strOblKom As String, autom As Boolean)
On Error Resume Next
Dim x As Long
Dim type_vs As String
Dim lichn() As String
Dim connt As Boolean
type_vs = frmCommand.listcommands.ListItems(frmCommand.listcommands.SelectedItem.Index).SubItems(2)

lichn = Split(get_config("print_lichnom"), ";")
For x = 0 To UBound(lichn())
    If lichn(x) = type_vs Then connt = True
    
Next x

If Not connt = True Then Exit Function


    Set oExcelApp = CreateObject("EXCEL.APPLICATION")
    If Err = 429 Then MsgBox "Application is Microsoft Excel is not installed!", vbExclamation, strMAIN_TITLE: Screen.MousePointer = vbDefault: Exit Function
    
    
        Dim sFile As String
    sFile = sCNV_txtDirShabl & "F-ln.xls"
    If Not ExFile(sFile) Then MsgBox "Файл шаблона" & NL2 & sFile & NL2 & "не найден. Проверьте правильность указания пути шаблонов." & vbNewLine & "Для восстановления шаблонов воспользуйтесь вкладкой `Воссановление` в меню `Опции`", vbCritical, strMAIN_TITLE: Screen.MousePointer = vbDefault: Exit Function
    
    If get_config("view_print") = 1 Then oExcelApp.Visible = False Else oExcelApp.Visible = True
    
    
    oExcelApp.WORkbooks.Open FileName:=sCNV_txtDirShabl & "F-ln.xls", ReadOnly:=True, ignOReReadOnlyRecommended:=True
    
    
    Set oWb = oExcelApp.ActiveWORkbook
    Set oWs = oExcelApp.Sheets(1)

        With oWs
                   
                    Dim BUF() As String
                    Dim pr() As String
                    Dim cnt As Long
                    Dim s As Long
                    Dim cUpk As Long
                    Dim STROKA As Long
                    Dim c As Long
                    Dim t As String
                    
                    Dim LST As Long
                    Dim CC As Long
                    Dim dUpk()  As Long
                    Dim dayprikaz As Long
                    Call mysql.query("Select ln from dayprikaz where `data` = '" & dateotp & "'")
                    dayprikaz = dat(1, 1)
                    Call mysql.query("SELECT fam, name, otch,datar,naryad_" & nowBase & ".oblkom, vus_p  FROM prnik_" & nowBase & ",naryad_" & nowBase & ",otpravka_" & nowBase & " WHERE naryad_" & nowBase & ".narid=otpravka_" & nowBase & ".narid AND otpravka_" & nowBase & ".otpravkaid=" & ID & " AND prnik_" & nowBase & ".otprvid = " & ID & " ORder by txtvk, fam")
                    datt() = dat()
                    s = 17
                    cnt = st
                    
                    .cells(s, 4) = "рядовой"
                    .cells(s, 7) = "Приказ ВК Московской области"
                    .cells(s + 1, 7) = "от " & CnvDataSqLToWin(dateotp) & " г. № " & dayprikaz
                    .cells(s + 2, 7) = "Дело №___, стр. ___"
                   
Dim j As Long
    j = 1
    Do While j < cnt
.Rows("17:19").Copy
.Rows("17:17").Insert
j = j + 1
    Loop
                                .cells(14, 3) = dateotp
                                .cells(14, 6) = "на К-" & datt(5, 1)
                                
                                .cells(s, 1) = 1
                                .cells(s, 2) = datt(1, 1)
                                .cells(s + 1, 2) = datt(2, 1)
                                .cells(s + 2, 2) = datt(3, 1)
                                .cells(s, 5) = datt(6, 1)
                                .cells(s, 3) = Left$(datt(4, 1), 4)
                               
                                
                                s = s + 3
                                
            
                    For CC = 2 To cnt


                                
                                .calls(s, 1).Select
                                .cells(s, 1) = CC
                                .cells(s, 2) = datt(1, CC)
                                .cells(s + 1, 2) = datt(2, CC)
                                .cells(s + 2, 2) = datt(3, CC)
                                .cells(s, 5) = datt(6, CC)
                                .cells(s, 3) = Left$(datt(4, CC), 4)
                                
                                s = s + 3
                    
                    Next CC
   
                 
End With
           ''''''Разбитие на листы
Dim CheckRow As Long
Dim countPB As Long

CheckRow = Int((oWs.cells(16, 10) * 3) + 14)

    oWs.ResetAllPageBreaks
    countPB = oWs.HPageBreaks.Count
    If countPB > 0 Then
        If oWs.HPageBreaks(countPB).Location.row > CheckRow Then
        oWs.HPageBreaks.add Before:=oWs.Rows(CheckRow)
        End If
    End If

  If frmCommand.listln.ListItems(1).Checked = True Then
        oExcelApp.ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
        oExcelApp.Application.DisplayAlerts = False
       oWb.Close
        oExcelApp.Quit
                Else
oExcelApp.Visible = True
Set oExcelApp = Nothing
Screen.MousePointer = vbDefault
    End If


Exit Function

End Function
Public Function To_AKT_Krugki(ID As Integer, strOblKom As String, autom As Boolean)
On Error Resume Next
    Set oExcelApp = CreateObject("EXCEL.APPLICATION")
    
        If Err = 429 Then MsgBox "Application Microsoft Excel is not installed!", vbExclamation, strMAIN_TITLE: Screen.MousePointer = vbDefault: Exit Function

    Dim sFile As String
    sFile = sCNV_txtDirShabl & "F81.xls"
    If Not ExFile(sFile) Then MsgBox "Файл шаблона" & NL2 & sFile & NL2 & "не найден. Проверьте правильность указания пути шаблонов." & vbNewLine & "Для восстановления шаблонов воспользуйтесь вкладкой `Воссановление` в меню `Опции`", vbCritical, strMAIN_TITLE: Screen.MousePointer = vbDefault: Exit Function
    
    If get_config("view_print") = 1 Then oExcelApp.Visible = False Else oExcelApp.Visible = True
    
    
    oExcelApp.WORkbooks.Open FileName:=sFile, ReadOnly:=True, ignOReReadOnlyRecommended:=True
    
    Set oWb = oExcelApp.ActiveWORkbook
    Set oWs = oExcelApp.Sheets(1)


Dim BUF() As String
Dim pr() As String
Dim cnt As Long
Dim s As Long
Dim cUpk As Long
Dim STROKA As Long
Dim c As Long
Dim t As String
Dim x As Long
Dim LST As Long
Dim CC As Long
Dim dUpk()  As Long

Call mysql.query("SELECT naryad_" & nowBase & ".oblkom, naryad_" & nowBase & ".rodv, naryad_" & nowBase & ".punkt, naryad_" & nowBase & ".vch, otpravka_" & nowBase & ".kolvo FROM naryad_" & nowBase & ", otpravka_" & nowBase & " WHERE otpravka_" & nowBase & ".narid = naryad_" & nowBase & ".narid AND otpravka_" & nowBase & ".otpravkaid = " & CStr(ID))
s = 11

cnt = st

With oWs
            .cells(6, 7) = dateotp
            .cells(12, 1) = dateotp
            .cells(22, 2) = dat(5, 1)
            .cells(20, 11) = dat(1, 1)
            .cells(26, 2) = dat(2, 1) & "  " & dat(3, 1)
            .cells(26, 8) = "в/ч  " & dat(4, 1)
            .cells(46, 3) = dateotp
End With
    

    oWb.SaveAs FileName:=base_path & "out\Akt_krugki\" & dat(1, 1) & "-" & CnvDataSqLToWin(dateotp) & "-" & Right$(time, 2) & ".xls", FileFORmat:=-4143
    
  If get_config("auto_print") = 1 Then
        oExcelApp.ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
        oExcelApp.Application.DisplayAlerts = False
         oWb.Close
        oExcelApp.Quit
                Else
oExcelApp.Visible = True
Set oExcelApp = Nothing
Screen.MousePointer = vbDefault
    End If
End Function

Public Function To_List(ID As Integer)
On Error GoTo ErrH
Dim x As Long
Dim bNAR As String
        
        Set oExcelApp = CreateObject("EXCEL.APPLICATION")
        
        If Err = 429 Then MsgBox "Application Microsoft Excel is not installed!", vbExclamation, strMAIN_TITLE: Screen.MousePointer = vbDefault: Exit Function
    
    
    
        Dim sFile As String
    sFile = sCNV_txtDirShabl & "List.xls"
    If Not ExFile(sFile) Then MsgBox "Файл шаблона" & NL2 & sFile & NL2 & "не найден. Проверьте правильность указания пути шаблонов." & vbNewLine & "Для восстановления шаблонов воспользуйтесь вкладкой `Воссановление` в меню `Опции`", vbCritical, strMAIN_TITLE: Screen.MousePointer = vbDefault: Exit Function

    
    
        oExcelApp.WORkbooks.Open FileName:=sFile, ReadOnly:=True, ignOReReadOnlyRecommended:=True

    If get_config("view_print") = 1 Then oExcelApp.Visible = False Else oExcelApp.Visible = True
        
        Set oWb = oExcelApp.ActiveWORkbook
        Set oWs = oExcelApp.Sheets(1)
        
        Call mysql.query("SELECT prnik_" & nowBase & ".fam, prnik_" & nowBase & ".name, prnik_" & nowBase & ".otch, prnik_" & nowBase & ".txtvk, naryad_" & nowBase & ".oblkom FROM prnik_" & nowBase & ", naryad_" & nowBase & ", otpravka_" & nowBase & " WHERE naryad_" & nowBase & ".narid = otpravka_" & nowBase & ".narid AND prnik_" & nowBase & ".otprvid = otpravka_" & nowBase & ".otpravkaid AND otpravka_" & nowBase & ".otpravkaid = " & ID & " ORder by txtvk, fam")
        
                
                With oWs
                    
                    If st > 2 Then
                            .cells(5, 1).Select
                            .range("A5:" & "A" & 5 + (st - 2) - 1).Select
                            oExcelApp.Selection.EntireRow.Insert
                            oExcelApp.CutCopyMode = False
                    End If
                    
                    If st = 1 Then
                            .Rows(5).Select
                            oExcelApp.Selection.EntireRow.Delete
                    End If
                    
                    
                    
                    .cells(4, 1).Select
                    .cells(4, 1) = "1"
                    .cells(4, 2).FormulaR1C1 = dat(1, 1) & " " & dat(2, 1) & " " & dat(3, 1)
                    .cells(4, 3) = dat(4, 1)
                    
        
                
                
                For x = 1 To st - 1
                    .cells(4 + x, 1).Select
                    .cells(4 + x, 1) = x + 1
                    .cells(4 + x, 2) = dat(1, x + 1) & " " & dat(2, x + 1) & " " & dat(3, x + 1)
                    .cells(4 + x, 3) = dat(4, x + 1)
                Next x
                    .cells(2, 2) = dat(5, 1)
                    .cells(2, 3) = CnvDataSqLToWin(dateotp)
                          
End With
    oExcelApp.Visible = True
    'oExcelApp.Quit
   ' oExcelApp.Application.DisplayAlerts = False
    Set oExcelApp = Nothing
    Screen.MousePointer = vbDefault
    Exit Function
ErrH:
MsgBox "Ошибка при генерировании файла Microsoft Exel:" & NL2 & "Выполняемая функция: " & """" & "To_List" & """" & NL2 & "Номер ошибки: " & Err.Number & NL2 & "Описание: " & Err.Description, vbCritical, "Microsoft Excel - Проверочный список"
  Screen.MousePointer = vbDefault
End Function
' Сухой паек
Public Function To_sPayok(ID As Integer, autom As Boolean)
On Error GoTo ErrH
Dim x As Long
Dim bNAR As String
        
        Set oExcelApp = CreateObject("EXCEL.APPLICATION")
        
        If Err = 429 Then MsgBox "Application Microsoft Excel is not installed!", vbExclamation, strMAIN_TITLE: Screen.MousePointer = vbDefault: Exit Function
    
        Dim sFile As String
    sFile = sCNV_txtDirShabl & "sPayok.xls"
    If Not ExFile(sFile) Then MsgBox "Файл шаблона" & NL2 & sFile & NL2 & "не найден. Проверьте правильность указания пути шаблонов." & vbNewLine & "Для восстановления шаблонов воспользуйтесь вкладкой `Воссановление` в меню `Опции`", vbCritical, strMAIN_TITLE: Screen.MousePointer = vbDefault: Exit Function


        oExcelApp.WORkbooks.Open FileName:=sFile, ReadOnly:=True, ignOReReadOnlyRecommended:=True
        
    If get_config("view_print") = 1 Then oExcelApp.Visible = False Else oExcelApp.Visible = True
        
        Set oWb = oExcelApp.ActiveWORkbook
        Set oWs = oExcelApp.Sheets(1)
        
        Call mysql.query("SELECT `fam`, `name`, `otch`, `oblkom`, `rodv`, `punkt`, `vch`, otpravka_" & nowBase & ".`data` FROM `prnik_" & nowBase & "`, `otpravka_" & nowBase & "`, `naryad_" & nowBase & "` WHERE prnik_" & nowBase & ".otprvid = otpravkaid AND otpravka_" & nowBase & ".narid = naryad_" & nowBase & ".narid AND otprvid = " & ID & " ORder by fam")
                
                With oWs
                    
                    If st > 1 Then
                            .Rows("7:7").Copy
                            .range("A8:" & "A" & 8 + (st - 2)).Select
                            oExcelApp.Selection.EntireRow.Insert
                    End If

                    .cells(7, 1).Select
                    .cells(7, 1) = "1."
                    .cells(7, 2) = "ряд."
                    .cells(7, 4).FormulaR1C1 = dat(1, 1) & " " & dat(2, 1) & " " & dat(3, 1)
                    .cells(7, 7) = dat(8, 1)
                
                For x = 2 To st
                    .cells(x + 6, 1).Select
                    .cells(x + 6, 1) = x & "."
                    .cells(x + 6, 2) = "ряд."
                    .cells(x + 6, 4).FormulaR1C1 = dat(1, x) & " " & dat(2, x) & " " & dat(3, x)
                    .cells(x + 6, 7) = dat(8, x)
                Next x
                    .cells(4, 2) = dat(8, 1)
                    .cells(4, 3) = dat(4, 1)
                    .cells(4, 8) = dat(5, 1) & " " & dat(6, 1) & " в/ч " & dat(7, 1)

                    oExcelApp.ActiveWORkbook.RefreshAll
                    On Error Resume Next
                    MkDir sCNV_txtPathResoult & "Payok\"
                    If ExFile(sCNV_txtPathResoult & "Payok\" & "List_sPayok.xls") Then Kill sCNV_txtPathResoult & "Payok\" & "List_sPayok.xls"
                        
                     oWb.SaveAs FileName:=sCNV_txtPathResoult & "Payok\" & "List_sPayok.xls", FileFORmat:=-4143
End With

Dim CheckRow As Long
Dim countPB As Long

CheckRow = Int((oWs.cells(6, 10)) + 6)

    oWs.ResetAllPageBreaks
    countPB = oWs.HPageBreaks.Count
    If countPB > 0 Then
        If oWs.HPageBreaks(countPB).Location.row > CheckRow Then
        oWs.HPageBreaks.add Before:=oWs.Rows(CheckRow)
        End If
    End If
                
  If frmCommand.listln.ListItems(5).Checked = True Then
                    oExcelApp.ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
                    oExcelApp.Application.DisplayAlerts = False
                    oWb.Close
                    oExcelApp.Quit
                     Else
oExcelApp.Visible = True
Set oExcelApp = Nothing
Screen.MousePointer = vbDefault
                End If
    Exit Function
ErrH:
MsgBox "Ошибка при генерировании файла Microsoft Exel:" & NL2 & "Выполняемая функция: " & """" & "To_List" & """" & NL2 & "Номер ошибки: " & Err.Number & NL2 & "Описание: " & Err.Description, vbCritical, "Microsoft Excel - Проверочный список"
End Function

Public Function vishka_r()
On Error GoTo ErrH
Dim x As Long
Dim tVK As String
Dim vv As Long
Dim bNAR As String
Dim sCnt As Long
        
        Set oExcelApp = CreateObject("EXCEL.APPLICATION")
        
        If Err = 429 Then MsgBox "Application Microsoft Excel is not installed!", vbExclamation, strMAIN_TITLE: Screen.MousePointer = vbDefault: Exit Function
    
        Dim sFile As String
    sFile = sCNV_txtDirShabl & "Ray.xls"
    If Not ExFile(sFile) Then MsgBox "Файл шаблона" & NL2 & sFile & NL2 & "не найден. Проверьте правильность указания пути шаблонов." & vbNewLine & "Для восстановления шаблонов воспользуйтесь вкладкой `Воссановление` в меню `Опции`", vbCritical, strMAIN_TITLE: Screen.MousePointer = vbDefault: Exit Function


        oExcelApp.WORkbooks.Open FileName:=sFile, ReadOnly:=True, ignOReReadOnlyRecommended:=True
            If iCNV_chShowObj = 1 Then oExcelApp.Visible = True
        
        Set oWb = oExcelApp.ActiveWORkbook
        Set oWs = oExcelApp.Sheets(1)
          With oWs
          
    
    Call Reg_VK_List
    For x = 0 To UBound(nVK())
    
    
    
                
                
                Call mysql.query("SELECT Count(idprnik) FROM prnik_" & nowBase & " WHERE `otprvid` > '0' AND `txtvk` like '%" & nVK(x) & "%' AND txtobraz like 'Высшее%'")
                            
                .cells(x + 1, 2) = nVK(x)
                .cells(x + 1, 3) = dat(1, 1)
                sCnt = sCnt + CLng(dat(1, 1))
    
 
    
    Next x
    
     .cells(41, 3) = sCnt
      
End With

    If ExFile(sCNV_txtPathResoult & "List_Ray.xls") Then Kill sCNV_txtPathResoult & "List_Ray.xls"
            
                oWb.SaveAs FileName:=sCNV_txtPathResoult & "List_Ray.xls", FileFORmat:=-4143

       Set oExcelApp = Nothing
       Screen.MousePointer = vbDefault
    Exit Function
ErrH:
MsgBox "Ошибка при генерировании файла Microsoft Exel:" & NL2 & "Выполняемая функция: " & """" & "To_List" & """" & NL2 & "Номер ошибки: " & Err.Number & NL2 & "Описание: " & Err.Description, vbCritical, "Microsoft Excel - Проверочный список"
End Function
Public Function rod()
On Error Resume Next
Dim x As Long
Dim tVK As String
Dim vv As Long
Dim bNAR As String
Dim sCnt As Long
Dim sumx, sumy As Long
        
        Set oExcelApp = CreateObject("EXCEL.APPLICATION")
        
        If Err = 429 Then MsgBox "Application Microsoft Excel is not installed!", vbExclamation, strMAIN_TITLE: Screen.MousePointer = vbDefault: Exit Function
    
        Dim sFile As String
    sFile = sCNV_txtDirShabl & "empty.xls"
    If Not ExFile(sFile) Then MsgBox "Файл шаблона" & NL2 & sFile & NL2 & "не найден. Проверьте правильность указания пути шаблонов." & vbNewLine & "Для восстановления шаблонов воспользуйтесь вкладкой `Воссановление` в меню `Опции`", vbCritical, strMAIN_TITLE: Screen.MousePointer = vbDefault: Exit Function


        oExcelApp.WORkbooks.Open FileName:=sFile, ReadOnly:=True, ignOReReadOnlyRecommended:=True
            If iCNV_chShowObj = 1 Then oExcelApp.Visible = True
        
        Set oWb = oExcelApp.ActiveWORkbook
        Set oWs = oExcelApp.Sheets(1)
          With oWs
          Dim rx As Long
          
    
    
    Call sorting_blok("txtvk", "prnik_" & nowBase & "")
    Dim outmm() As String
    ReDim outmm(UBound(outm())) As String
    outmm() = outm()
    Call sorting_blok("rodv", "naryad_" & nowBase & "")
        For x = 1 To UBound(outmm())
            .cells(x + 1, 1) = outmm(x)
            sumx = 0
            For rx = 1 To UBound(outm())
                               .cells(1, rx + 1) = outm(rx)
                Call mysql.query("SELECT count(idprnik) FROM prnik_" & nowBase & ",otpravka_" & nowBase & ",naryad_" & nowBase & " WHERE naryad_" & nowBase & ".narid = otpravka_" & nowBase & ".narid AND prnik_" & nowBase & ".otprvid = otpravka_" & nowBase & ".otpravkaid AND naryad_" & nowBase & ".rodv = '" & outm(rx) & "' AND prnik_" & nowBase & ".txtvk like '%" & outmm(x) & "%'")
                
                .cells(x + 1, rx + 1) = dat(1, 1)
                sumx = sumx + VAl((dat(1, 1)))
       
        Next rx
               .cells(x + 1, rx + 1) = sumx
           
    Next x
    
       .range(.cells(1, 1), .cells(x, rx)).Borders(7).LineStyle = 1
       .range(.cells(1, 1), .cells(x, rx)).Borders(8).LineStyle = 1
       .range(.cells(1, 1), .cells(x, rx)).Borders(9).LineStyle = 1
       .range(.cells(1, 1), .cells(x, rx)).Borders(10).LineStyle = 1
       .range(.cells(1, 1), .cells(x, rx)).Borders(11).LineStyle = 1
       .range(.cells(1, 1), .cells(x, rx)).Borders(12).LineStyle = 1
    
      .cells.EntireColumn.AutoFit
      
End With

       Set oExcelApp = Nothing
       Screen.MousePointer = vbDefault
    Exit Function

End Function
Public Function stat_date()
On Error Resume Next
Dim x As Long
Dim stt As Long
Dim sty As Long
Dim tVK As String
Dim vv As Long
Dim bNAR As String
Dim rx As Long
        Set oExcelApp = CreateObject("EXCEL.APPLICATION")
        If Err = 429 Then MsgBox "Application Microsoft Excel is not installed!", vbExclamation, strMAIN_TITLE: Screen.MousePointer = vbDefault: Exit Function
        Dim sFile As String
    sFile = sCNV_txtDirShabl & "dayotp.xls"
    If Not ExFile(sFile) Then MsgBox "Файл шаблона" & NL2 & sFile & NL2 & "не найден. Проверьте правильность указания пути шаблонов." & vbNewLine & "Для восстановления шаблонов воспользуйтесь вкладкой `Воссановление` в меню `Опции`", vbCritical, strMAIN_TITLE: Screen.MousePointer = vbDefault: Exit Function
        oExcelApp.WORkbooks.Open FileName:=sFile, ReadOnly:=True, ignOReReadOnlyRecommended:=True
            If iCNV_chShowObj = 1 Then oExcelApp.Visible = True
        Set oWb = oExcelApp.ActiveWORkbook
        Set oWs = oExcelApp.Sheets(1)
          With oWs
          .cells(1, 2) = "Количество граждан привозимых на ОСП по дням"
                    .range("B1:H1").Select: oExcelApp.Selection.Merge
               Call sorting_blok("dataosp", "prnik_" & nowBase & "")
        ReDim date_otp(UBound(outm()))
        For x = 1 To UBound(outm())
        date_otp(x) = outm(x)
        .cells(2, x + 1) = outm(x)
        Next x
        
   .range("A4" & ":" & "A" & UBound(nVK()) + 2).Select: oExcelApp.Selection.EntireRow.Insert
    Call Reg_VK_List
    For vv = 0 To UBound(nVK())
    .cells(vv + 3, 1) = nVK(vv)
        For rx = 1 To UBound(outm())
            mysql.query ("SELECT count(idprnik) FROM prnik_" & nowBase & " WHERE dataosp='" & date_otp(rx) & "' AND txtvk = '" & nVK(vv) & "'")
           If dat(1, 1) = "0" Then dat(1, 1) = "-"
            .cells(vv + 3, rx + 1) = dat(1, 1)
            
        
        Next rx
    Next vv
        For x = 0 To UBound(outm())
            mysql.query ("SELECT count(idprnik) FROM prnik_" & nowBase & " WHERE dataosp='" & date_otp(x) & "'")
            .cells(UBound(nVK()) + 4, x + 1) = dat(1, 1)
        Next x
    
       .range(.cells(3, 1), .cells(UBound(nVK()) + 2, UBound(outm()) + 1)).Borders(7).LineStyle = 1
       .range(.cells(3, 1), .cells(UBound(nVK()) + 2, UBound(outm()) + 1)).Borders(8).LineStyle = 1
       .range(.cells(3, 1), .cells(UBound(nVK()) + 2, UBound(outm()) + 1)).Borders(9).LineStyle = 1
       .range(.cells(3, 1), .cells(UBound(nVK()) + 2, UBound(outm()) + 1)).Borders(11).LineStyle = 1
       .range(.cells(3, 1), .cells(UBound(nVK()) + 2, UBound(outm()) + 1)).Borders(12).LineStyle = 1
        .cells(UBound(nVK()) + 4, 1) = "Итого"
        .cells.EntireColumn.AutoFit
        
        
        
End With

       Set oExcelApp = Nothing
       Screen.MousePointer = vbDefault
    Exit Function

End Function
Public Function stat_otpravka()
On Error Resume Next
Dim x As Long
Dim stt As Long
Dim sty As Long
Dim tVK As String
Dim vv As Long
Dim bNAR As String
Dim rx As Long
stt = 0
  ReDim date_otp(60)

        
        Set oExcelApp = CreateObject("EXCEL.APPLICATION")
        
        If Err = 429 Then MsgBox "Application Microsoft Excel is not installed!", vbExclamation, strMAIN_TITLE: Screen.MousePointer = vbDefault: Exit Function
    
        Dim sFile As String
    sFile = sCNV_txtDirShabl & "dayotp.xls"
    If Not ExFile(sFile) Then MsgBox "Файл шаблона" & NL2 & sFile & NL2 & "не найден. Проверьте правильность указания пути шаблонов." & vbNewLine & "Для восстановления шаблонов воспользуйтесь вкладкой `Воссановление` в меню `Опции`", vbCritical, strMAIN_TITLE: Screen.MousePointer = vbDefault: Exit Function


        oExcelApp.WORkbooks.Open FileName:=sFile, ReadOnly:=True, ignOReReadOnlyRecommended:=True
            If iCNV_chShowObj = 1 Then oExcelApp.Visible = True
        
        Set oWb = oExcelApp.ActiveWORkbook
        Set oWs = oExcelApp.Sheets(1)
          With oWs
              .cells(1, 2) = "Количество граждан отправляемых с ОСП по дням"
                    .range("B1:H1").Select: oExcelApp.Selection.Merge
        Call mysql.query("SELECT `data` FROM otpravka_" & nowBase & " group by `data` ORder by data ASC")
        stt = st
        ReDim date_otp(st)
        For x = 1 To st
        date_otp(x) = dat(1, x)
        .cells(2, x + 1) = dat(1, x)
        Next x
   
     
          
    .range("A4" & ":" & "A" & UBound(nVK()) + 2).Select: oExcelApp.Selection.EntireRow.Insert
    Call Reg_VK_List
    For vv = 0 To UBound(nVK())
            .cells(vv + 3, 1) = nVK(vv)
        For rx = 1 To stt
            mysql.query ("SELECT count(idprnik) FROM prnik_" & nowBase & ",otpravka_" & nowBase & " WHERE prnik_" & nowBase & ".otprvid=otpravka_" & nowBase & ".otpravkaid AND prnik_" & nowBase & ".txtvk = '" & nVK(vv) & "' AND otpravka_" & nowBase & ".data='" & date_otp(rx) & "'")
            If dat(1, 1) = "0" Then dat(1, 1) = "-"
            .cells(vv + 3, rx + 1) = dat(1, 1)
          Next rx
    Next vv
            
                
                   
        For x = 0 To stt
            mysql.query ("SELECT count(idprnik) FROM prnik_" & nowBase & ",otpravka_" & nowBase & " WHERE prnik_" & nowBase & ".otprvid=otpravka_" & nowBase & ".otpravkaid AND otpravka_" & nowBase & ".data='" & date_otp(x) & "'")
            .cells(UBound(nVK()) + 4, x + 1) = dat(1, 1)
        Next x
    .cells.EntireColumn.AutoFit
       .range(.cells(3, 1), .cells(UBound(nVK()) + 2, stt + 1)).Borders(7).LineStyle = 1
       .range(.cells(3, 1), .cells(UBound(nVK()) + 2, stt + 1)).Borders(8).LineStyle = 1
       .range(.cells(3, 1), .cells(UBound(nVK()) + 2, stt + 1)).Borders(9).LineStyle = 1
       .range(.cells(3, 1), .cells(UBound(nVK()) + 2, stt + 1)).Borders(11).LineStyle = 1
       .range(.cells(3, 1), .cells(UBound(nVK()) + 2, stt + 1)).Borders(12).LineStyle = 1
        .cells(UBound(nVK()) + 4, 1) = "Итого"
 
    .cells.EntireColumn.AutoFit

    
     
      
End With

       Set oExcelApp = Nothing
       Screen.MousePointer = vbDefault
    Exit Function

End Function
Public Function To_Ray()
'''''''''' Главная - Отчет за день - Районка
On Error GoTo ErrH
Dim x As Long
Dim tVK As String
Dim vv As Long
Dim bNAR As String
Dim sCnt As Long
        
        Set oExcelApp = CreateObject("EXCEL.APPLICATION")
        
        If Err = 429 Then MsgBox "Application Microsoft Excel is not installed!", vbExclamation, strMAIN_TITLE: Screen.MousePointer = vbDefault: Exit Function
    
        Dim sFile As String
    sFile = sCNV_txtDirShabl & "Ray.xls"
    If Not ExFile(sFile) Then MsgBox "Файл шаблона" & NL2 & sFile & NL2 & "не найден. Проверьте правильность указания пути шаблонов." & vbNewLine & "Для восстановления шаблонов воспользуйтесь вкладкой `Воссановление` в меню `Опции`", vbCritical, strMAIN_TITLE: Screen.MousePointer = vbDefault: Exit Function


        oExcelApp.WORkbooks.Open FileName:=sFile, ReadOnly:=True, ignOReReadOnlyRecommended:=True
            If iCNV_chShowObj = 1 Then oExcelApp.Visible = True
        
        Set oWb = oExcelApp.ActiveWORkbook
        Set oWs = oExcelApp.Sheets(1)
          With oWs
          

Call mysql.query("SELECT txtvk, Count(idprnik) FROM prnik_" & nowBase & " group by txtvk")
    For vv = 1 To st

               sCnt = sCnt + dat(2, vv)
                .cells(vv + 1, 1) = dat(1, vv)
                .cells(vv + 1, 2) = dat(2, vv)
                .cells(st + 1, 1) = "Сумма"
                .cells(st + 1, 2) = sCnt
    
    Next vv
      
End With

       Set oExcelApp = Nothing
       Screen.MousePointer = vbDefault
    Exit Function
ErrH:
MsgBox "Ошибка при генерировании файла Microsoft Exel:" & NL2 & "Выполняемая функция: " & """" & "To_List" & """" & NL2 & "Номер ошибки: " & Err.Number & NL2 & "Описание: " & Err.Description, vbCritical, "Microsoft Excel - Проверочный список"
End Function

Public Function To_AKT(ID As Integer, autom As Boolean)
On Error Resume Next

    Set oExcelApp = CreateObject("EXCEL.APPLICATION")
    
        If Err = 429 Then MsgBox "Application Microsoft Excel is not installed!", vbExclamation, strMAIN_TITLE: Screen.MousePointer = vbDefault: Exit Function

    Dim sFile As String
    If Right$(nowBase, 1) = "1" Then sFile = sCNV_txtDirShabl & "akt_v.xls"
    If Right$(nowBase, 1) = "2" Then sFile = sCNV_txtDirShabl & "akt_o.xls"
    If Not ExFile(sFile) Then MsgBox "Файл шаблона" & NL2 & sFile & NL2 & "не найден. Проверьте правильность указания пути шаблонов." & vbNewLine & "Для восстановления шаблонов воспользуйтесь вкладкой `Воссановление` в меню `Опции`", vbCritical, strMAIN_TITLE: Screen.MousePointer = vbDefault: Exit Function
    
    If get_config("view_print") = 1 Then oExcelApp.Visible = False Else oExcelApp.Visible = True
    
    oExcelApp.WORkbooks.Open FileName:=sFile, ReadOnly:=True, ignOReReadOnlyRecommended:=True
    
    Set oWb = oExcelApp.ActiveWORkbook
    Set oWs = oExcelApp.Sheets(1)


Dim BUF() As String
Dim pr() As String
Dim cnt As Long
Dim s As Long
Dim cUpk As Long
Dim STROKA As Long
Dim c As Long
Dim t As String
Dim x As Long
Dim LST As Long
Dim CC As Long
Dim dUpk()  As Long

Call mysql.query("SELECT naryad_" & nowBase & ".oblkom, naryad_" & nowBase & ".rodv, naryad_" & nowBase & ".punkt, naryad_" & nowBase & ".vch, otpravka_" & nowBase & ".kolvo FROM naryad_" & nowBase & ", otpravka_" & nowBase & " WHERE otpravka_" & nowBase & ".narid = naryad_" & nowBase & ".narid AND otpravka_" & nowBase & ".otpravkaid = " & CStr(ID))
s = 11


cnt = st

With oWs
            .cells(6, 7) = dateotp
            .cells(12, 1) = dateotp
            .cells(22, 2) = dat(5, 1)
            .cells(20, 11) = dat(1, 1)
            .cells(26, 2) = dat(2, 1) & "  " & dat(3, 1)
            .cells(26, 8) = "в/ч  " & dat(4, 1)
            .cells(59, 3) = dateotp
End With
    

    
    MkDir sCNV_txtPathResoult & "Atk\"
    
  If frmCommand.listln.ListItems(2).Checked = True Then
    oExcelApp.ActiveWindow.SelectedSheets.PrintOut Copies:=3, Collate:=True
    oExcelApp.Application.DisplayAlerts = False
    oWb.Close
    oExcelApp.Quit
            Else
oExcelApp.Visible = True
Set oExcelApp = Nothing
Screen.MousePointer = vbDefault
     End If
    
End Function
Public Function To_rapORt(ID As Integer, autom As Boolean)
'''''''''''''''''''''' Рапорт старшего команды
On Error Resume Next

    Set oExcelApp = CreateObject("EXCEL.APPLICATION")
    
        If Err = 429 Then MsgBox "Application Microsoft Excel is not installed!", vbExclamation, strMAIN_TITLE: Screen.MousePointer = vbDefault: Exit Function

    Dim sFile As String
    sFile = sCNV_txtDirShabl & "rapORt.xls"
    If Not ExFile(sFile) Then MsgBox "Файл шаблона" & NL2 & sFile & NL2 & "не найден. Проверьте правильность указания пути шаблонов." & vbNewLine & "Для восстановления шаблонов воспользуйтесь вкладкой `Воссановление` в меню `Опции`", vbCritical, strMAIN_TITLE: Screen.MousePointer = vbDefault: Exit Function
    
    
    If get_config("view_print") = 1 Then oExcelApp.Visible = False Else oExcelApp.Visible = True
    
    
    oExcelApp.WORkbooks.Open FileName:=sFile, ReadOnly:=True, ignOReReadOnlyRecommended:=True
    
    Set oWb = oExcelApp.ActiveWORkbook
    Set oWs = oExcelApp.Sheets(1)


Dim BUF() As String
Dim pr() As String
Dim cnt As Long
Dim s As Long
Dim cUpk As Long
Dim STROKA As Long
Dim c As Long
Dim t As String
Dim x As Long
Dim LST As Long
Dim CC As Long
Dim dUpk()  As Long

Call mysql.query("SELECT naryad_" & nowBase & ".oblkom, naryad_" & nowBase & ".rodv, naryad_" & nowBase & ".punkt, naryad_" & nowBase & ".vch, otpravka_" & nowBase & ".kolvo FROM naryad_" & nowBase & ", otpravka_" & nowBase & " WHERE otpravka_" & nowBase & ".narid = naryad_" & nowBase & ".narid AND otpravka_" & nowBase & ".otpravkaid = " & CStr(ID))
s = 11

cnt = st

With oWs
                       
            .cells(14, 6) = dat(1, 1)
            
            .cells(14, 8) = dat(4, 1)
            .cells(33, 1) = dateotp
End With
    

  If frmCommand.listln.ListItems(4).Checked = True Then
  oExcelApp.Application.DisplayAlerts = False
  oExcelApp.ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True

oWb.Close
oExcelApp.Quit
        Else
oExcelApp.Visible = True
Set oExcelApp = Nothing
Screen.MousePointer = vbDefault
End If
    
End Function
Public Function To_knotp()
''''''''''''''''''''' Главная - Книга учета команд
On Error Resume Next
    Set oExcelApp = CreateObject("EXCEL.APPLICATION")
    
        If Err = 429 Then MsgBox "Application Microsoft Excel is not installed!", vbExclamation, strMAIN_TITLE: Screen.MousePointer = vbDefault: Exit Function

    Dim sFile As String
    sFile = sCNV_txtDirShabl & "knotp.xls"
    If Not ExFile(sFile) Then MsgBox "Файл шаблона" & NL2 & sFile & NL2 & "не найден. Проверьте правильность указания пути шаблонов." & vbNewLine & "Для восстановления шаблонов воспользуйтесь вкладкой `Воссановление` в меню `Опции`", vbCritical, strMAIN_TITLE: Screen.MousePointer = vbDefault: Exit Function
    
    If iCNV_chShowObj = 1 Then oExcelApp.Visible = True
    
    
    oExcelApp.WORkbooks.Open FileName:=sFile, ReadOnly:=True, ignOReReadOnlyRecommended:=True
    
    Set oWb = oExcelApp.ActiveWORkbook
    Set oWs = oExcelApp.Sheets(1)


Dim BUF() As String
Dim pr() As String
Dim cnt As Long
Dim s As Long
Dim cUpk As Long
Dim STROKA As Long
Dim c As Long
Dim t As String
Dim x As Long
Dim LST As Long
Dim CC As Long
Dim dUpk()  As Long

Call mysql.query("SELECT otpravka_" & nowBase & ".otpravkaid, naryad_" & nowBase & ".oblkom, naryad_" & nowBase & ".rodv, naryad_" & nowBase & ".punkt, naryad_" & nowBase & ".vch, otpravka_" & nowBase & ".kolvo, otpravka_" & nowBase & ".data FROM otpravka_" & nowBase & ", naryad_" & nowBase & " WHERE otpravka_" & nowBase & ".narid=naryad_" & nowBase & ".narid ORDER BY otpravka_" & nowBase & ".data, otpravka_" & nowBase & ".otpravkaid ASC")
'Screen.MousePointer = vbHourglass
s = 2
cnt = st
With oWs
    '.Cells(1, 11).SELECT: oExcelApp.SELECTion.NumberFORmat = "0.00%"
    '.Cells(1, 11) = "=СЧЕТ('A2':'A9999')/" & cnt
 If cnt > 1 Then .range("A" & s & ":" & "A" & (s + (cnt - 2))).Select: oExcelApp.Selection.EntireRow.Insert
  

            For x = 1 To cnt
                For c = 1 To 7
                 .cells(x + 1, c) = dat(c, x)
                 Next c
                 '.range("A" & x & ":" & "I" & x).BORders(9).LineStyle = 1
                 
            Next x
    '.range("A" & x & ":" & "I" & x).BORders(9).LineStyle = 1
    .cells.EntireColumn.AutoFit
End With
    

    
  
    
  
    'oExcelApp.SendKeys "^p"
  
    
    Set oExcelApp = Nothing
    
    Screen.MousePointer = vbDefault
    
End Function

Public Function To_400(ID As Integer, autom As Boolean)
''''''''''''''''''''''''''''''''''''''''''' 400 приказ
On Error Resume Next

    Set oExcelApp = CreateObject("EXCEL.APPLICATION")
    
        If Err = 429 Then MsgBox "Application Microsoft Excel is not installed!", vbExclamation, strMAIN_TITLE: Screen.MousePointer = vbDefault: Exit Function

    Dim sFile As String
    sFile = sCNV_txtDirShabl & "400.xls"
    If Not ExFile(sFile) Then MsgBox "Файл шаблона" & NL2 & sFile & NL2 & "не найден. Проверьте правильность указания пути шаблонов." & vbNewLine & "Для восстановления шаблонов воспользуйтесь вкладкой `Воссановление` в меню `Опции`", vbCritical, strMAIN_TITLE: Screen.MousePointer = vbDefault: Exit Function
    
    If get_config("view_print") = 1 Then oExcelApp.Visible = False Else oExcelApp.Visible = True
    
    
    oExcelApp.WORkbooks.Open FileName:=sFile, ReadOnly:=True, ignOReReadOnlyRecommended:=True
    
    Set oWb = oExcelApp.ActiveWORkbook
    Set oWs = oExcelApp.Sheets(1)


Dim BUF() As String
Dim pr() As String
Dim cnt As Long
Dim s As Long
Dim cUpk As Long
Dim STROKA As Long
Dim c As Long
Dim t As String
Dim x As Long
Dim LST As Long
Dim CC As Long
Dim dUpk()  As Long
Dim strt As String
Dim strd As String
Dim strl As Integer
Call mysql.query("SELECT naryad_" & nowBase & ".oblkom, naryad_" & nowBase & ".rodv, naryad_" & nowBase & ".punkt, naryad_" & nowBase & ".doroga,  naryad_" & nowBase & ".vch, otpravka_" & nowBase & ".kolvo FROM naryad_" & nowBase & ", otpravka_" & nowBase & " WHERE otpravka_" & nowBase & ".narid = naryad_" & nowBase & ".narid AND otpravka_" & nowBase & ".otpravkaid = " & CStr(ID))
s = 11


cnt = st

With oWs
                       
            .cells(15, 4) = dat(1, 1)
            .cells(16, 9) = dat(2, 1) & "        " & dat(3, 1)
            strt = dat(4, 1)
            strl = Len(strt)
            strd = Left$(strt, strl - 2)
            strt = strd & "ой"
            .cells(13, 8) = dat(1, 1)
            .cells(17, 1) = strt
            .cells(23, 11) = dat(6, 1)
            .cells(24, 9) = "в/ч        " & dat(5, 1)
          
            .cells(26, 3) = dateotp
            .cells(30, 3) = dateotp
            .cells(59, 2) = dateotp
            .cells(7, 7) = dateotp
            
            
End With
    

    
  
    
   If frmCommand.listln.ListItems(3).Checked = True Then
        oExcelApp.ActiveWindow.SelectedSheets.PrintOut Copies:=2, Collate:=True
        oExcelApp.Application.DisplayAlerts = False
        oWb.Close
        oExcelApp.Quit
                Else
oExcelApp.Visible = True
Set oExcelApp = Nothing
Screen.MousePointer = vbDefault
    End If
    
End Function
Public Function To_ListsVK(vk As String)
On Error Resume Next

Dim c As Long
Dim x As Long
Dim Y As Long
Dim sFile As String
Dim sPathResoult As String
Dim cend As Long
Dim allpr As Boolean

With oWs
Set oExcelApp = CreateObject("EXCEL.APPLICATION")

If vk = "Все" Then
cend = UBound(nVK())
progress.ProgressBar1.Max = cend
allpr = True
Else
cend = "0"
nVK(Y) = vk
allpr = False
End If
Unload frmPrintListsVK
progress.Show

For Y = 0 To cend
                
                

                
                sFile = sCNV_txtDirShabl & "ListsVK.xls"
                sPathResoult = sCNV_txtPathResoult
                If Not ExFile(sFile) Then MsgBox "Файл шаблона" & NL2 & sFile & NL2 & "не найден. Проверьте правильность указания пути шаблонов." & vbNewLine & "Для восстановления шаблонов воспользуйтесь вкладкой `Воссановление` в меню `Опции`", vbCritical, strMAIN_TITLE: Screen.MousePointer = vbDefault: Exit Function
                
                'Call mysql.query("SELECT prnik_" & nowBase & ".txtvk, prnik_" & nowBase & ".fam, prnik_" & nowBase & ".name, prnik_" & nowBase & ".otch, prnik_" & nowBase & ".datar, otpravka_" & nowBase & ".data, naryad_" & nowBase & ".rodv, naryad_" & nowBase & ".punkt, naryad_" & nowBase & ".vch, prnik_" & nowBase & ".dir FROM prnik_" & nowBase & ", otpravka_" & nowBase & ", naryad_" & nowBase & " WHERE prnik_" & nowBase & ".otprvid = otpravka_" & nowBase & ".otpravkaid AND otpravka_" & nowBase & ".narid = naryad_" & nowBase & ".narid AND prnik_" & nowBase & ".txtvk like '%" & nVK(Y) & "%' ORder by fam")
               
                
    If get_config("view_print") = 1 Then oExcelApp.Visible = False Else oExcelApp.Visible = True
    
                oExcelApp.WORkbooks.Open FileName:=sFile, ReadOnly:=True, ignOReReadOnlyRecommended:=True
                
                Set oWb = oExcelApp.ActiveWORkbook
                Set oWs = oExcelApp.Sheets(1)
                
                If Err = 429 Then MsgBox "Application Microsoft Excel is not installed!", vbExclamation, strMAIN_TITLE: Screen.MousePointer = vbDefault: Exit Function
                
                Call mysql.query("SELECT CONCAT(prnik_" & nowBase & ".fam, ' ',prnik_" & nowBase & ".name, ' ',prnik_" & nowBase & ".otch) as fam, prnik_" & nowBase & ".datar, otpravka_" & nowBase & ".data, CONCAT(naryad_" & nowBase & ".rodv, ' ', naryad_" & nowBase & ".punkt, ' ', naryad_" & nowBase & ".vch), prnik_" & nowBase & ".dir  FROM prnik_" & nowBase & ", otpravka_" & nowBase & ", naryad_" & nowBase & " WHERE otpravka_" & nowBase & ".otpravkaid = prnik_" & nowBase & ".otprvid and otpravka_" & nowBase & ".narid = naryad_" & nowBase & ".narid and prnik_" & nowBase & ".txtvk like '%" & nVK(Y) & "%' order by fam")
               If allpr = False Then progress.ProgressBar1.Max = st
               
                               
                oExcelApp.Rows("5:5").Select
                oExcelApp.Selection.AutoFill Destination:=oExcelApp.Rows("5:" & st + 4), type:=0
               
                oWs.cells(3, 1) = nVK(Y)
                    For x = 1 To st
                        oWs.cells(x + 4, 1) = CStr(x)
                        oWs.cells(x + 4, 2) = dat(1, x)
                        oWs.cells(x + 4, 3) = dat(2, x)
                        oWs.cells(x + 4, 4) = dat(3, x)
                        oWs.cells(x + 4, 5) = dat(4, x)
                        oWs.cells(x + 4, 6) = dat(5, x)
                        If allpr = False Then progress.ProgressBar1 = x
                    Next x
               
                oWs.cells.EntireColumn.AutoFit
                oExcelApp.Visible = True
If allpr = True Then progress.ProgressBar1 = Y

Next Y
Unload progress
End With

Screen.MousePointer = vbDefault
End Function

Public Function ResoultSearch(LW As ListView, blShowRes As Boolean, Optional sFile As String, Optional strExcelCaption As String = "Результаты поиска")
On Error Resume Next

Dim tmp As String
Dim x As Long
Dim Y As Long
Dim sPathResoult As String
Dim FrF As Long
Dim tExcel As Boolean

'Screen.MousePointer = vbHourglass

FrF = FreeFile

sPathResoult = sCNV_txtPathResoult

Set oExcelApp = CreateObject("EXCEL.APPLICATION")

        If Err = 429 Then
            MsgBox "Application Microsoft Excel is not installed!", vbExclamation, strMAIN_TITLE: Screen.MousePointer = vbDefault
            If Len(Trim$((sFile))) > 0 Then sFile = sPathResoult & sFile & ".txt" Else sFile = sPathResoult & "ResoultSearch.txt"
        Else
            If Len(Trim$((sFile))) > 0 Then sFile = sPathResoult & sFile & ".xls" Else sFile = sPathResoult & "ResoultSearch.xls"
            tExcel = True
        End If
    
If ExFile(sFile) Then Kill sFile

Open sFile For Binary Lock Read Write As #FrF
        For x = 1 To LW.ListItems.Count
            tmp = LW.ListItems(x).Text
                For Y = 1 To LW.ColumnHeaders.Count - 1
                    tmp = tmp & vbTab & LW.ListItems(x).SubItems(Y)
                Next Y
                 tmp = tmp & vbNewLine
                 Put #FrF, , tmp
        Next x
Close #FrF

If tExcel Then
    If blShowRes Then
        oExcelApp.Caption = strExcelCaption
        oExcelApp.Visible = True
        oExcelApp.WORkbooks.Open FileName:=sFile, ReadOnly:=True, ignOReReadOnlyRecommended:=True
        Set oWb = oExcelApp.ActiveWORkbook
        Set oWs = oExcelApp.Sheets(1)
        oWs.cells.EntireColumn.AutoFit
        
    End If
Else
    If blShowRes Then Call ShellExecute(0, "open", sFile, vbNullString, sPathResoult, 3)
End If
Set oExcelApp = Nothing
Screen.MousePointer = vbDefault

Beep

End Function
Public Function obraz_fam(type_p As Long)
On Error Resume Next
    Set oExcelApp = CreateObject("EXCEL.APPLICATION")
    If Err = 429 Then MsgBox "Application is Microsoft Excel is not installed!", vbExclamation, strMAIN_TITLE: Screen.MousePointer = vbDefault: Exit Function
    Dim sFile As String
    sFile = sCNV_txtDirShabl & "obraz_fam.xls"
    If Not ExFile(sFile) Then MsgBox "Файл шаблона" & NL2 & sFile & NL2 & "не найден. Проверьте правильность указания пути шаблонов." & vbNewLine & "Для восстановления шаблонов воспользуйтесь вкладкой `Воссановление` в меню `Опции`", vbCritical, strMAIN_TITLE: Screen.MousePointer = vbDefault: Exit Function
    oExcelApp.WORkbooks.Open FileName:=sFile, ReadOnly:=True, ignOReReadOnlyRecommended:=False
    Set oWb = oExcelApp.ActiveWORkbook
    Set oWs = oExcelApp.Sheets(1)
     progress.Show
    oExcelApp.Visible = False
        With oWs
                    Dim cnt As Long
                    Dim s As Long
                    Dim c As Long
                    Dim t As String
                    Call obraz_in
                    .cells(1, 1) = "Статистика по фамильно : " & obraz(type_p)
                    .range("A1:G1").Select: oExcelApp.Selection.Merge
                    Call mysql.query("SELECT prnik_" & nowBase & ".idprnik, prnik_" & nowBase & ".fam, prnik_" & nowBase & ".name, prnik_" & nowBase & ".otch, prnik_" & nowBase & ".txtvk, naryad_" & nowBase & ".rodv, naryad_" & nowBase & ".vch FROM otpravka_" & nowBase & ", naryad_" & nowBase & ", prnik_" & nowBase & " WHERE  prnik_" & nowBase & ".txtobraz='" & obraz(type_p) & "' and prnik_" & nowBase & ".otprvid = otpravka_" & nowBase & ".otpravkaid AND otpravka_" & nowBase & ".narid = naryad_" & nowBase & ".narid ORDER by `txtvk`,`fam`")
                              
                    cnt = st
                    progress.ProgressBar1.Max = st
                    .range("A" & s + 5 & ":" & "A" & (s + 2 + cnt)).Select: oExcelApp.Selection.EntireRow.Insert
                        For s = 1 To cnt
                                .cells(s + 3, 1) = s
                                .cells(s + 3, 3) = dat(2, s)
                                .cells(s + 3, 4) = dat(3, s)
                                .cells(s + 3, 5) = dat(4, s)
                                .cells(s + 3, 2) = dat(5, s)
                                .cells(s + 3, 6) = dat(6, s)
                                .cells(s + 3, 7) = dat(7, s)
                                
                         progress.ProgressBar1 = s
                         Next s
                         Unload progress
                         
            oExcelApp.Visible = True
.cells.EntireColumn.AutoFit
End With
Set oExcelApp = Nothing
End Function
Public Function obraz_vk(type_p As Long)
On Error Resume Next
    Set oExcelApp = CreateObject("EXCEL.APPLICATION")
    If Err = 429 Then MsgBox "Application is Microsoft Excel is not installed!", vbExclamation, strMAIN_TITLE: Screen.MousePointer = vbDefault: Exit Function
    Dim sFile As String
    sFile = sCNV_txtDirShabl & "empty.xls"
    If Not ExFile(sFile) Then MsgBox "Файл шаблона" & NL2 & sFile & NL2 & "не найден. Проверьте правильность указания пути шаблонов." & vbNewLine & "Для восстановления шаблонов воспользуйтесь вкладкой `Воссановление` в меню `Опции`", vbCritical, strMAIN_TITLE: Screen.MousePointer = vbDefault: Exit Function
    oExcelApp.WORkbooks.Open FileName:=sFile, ReadOnly:=True, ignOReReadOnlyRecommended:=False
    Set oWb = oExcelApp.ActiveWORkbook
    Set oWs = oExcelApp.Sheets(1)
    progress.Show
    
    oExcelApp.Visible = False
        With oWs
                    Dim cnt As Long
                    Dim s As Long
                    Dim c As Long
                    Dim t As String
                    Call obraz_in
                    .cells(1, 1) = "Статистика по военкоматам : " & obraz(type_p)
                    .range("A1:G1").Select: oExcelApp.Selection.Merge
                    Call obraz_in
                    Call mysql.query("SELECT `txtvk`,count(idprnik) from prnik_" & nowBase & " where txtobraz='" & obraz(type_p) & "' group by txtvk order by txtvk")
                            
                    
                    progress.ProgressBar1.Max = st
                        For s = 1 To st
                                .cells(s + 2, 1) = s
                                .cells(s + 2, 2) = dat(1, s)
                                .cells(s + 2, 3) = dat(2, s)
                                                                
                         progress.ProgressBar1 = s
                         Next s
                         Unload progress
                         
            oExcelApp.Visible = True
.cells.EntireColumn.AutoFit

.range("A3:C" & s + 1).Borders(7).LineStyle = 1
.range("A3:C" & s + 1).Borders(8).LineStyle = 1
.range("A3:C" & s + 1).Borders(9).LineStyle = 1
.range("A3:C" & s + 1).Borders(10).LineStyle = 1
.range("A3:C" & s + 1).Borders(11).LineStyle = 1
.range("A3:C" & s + 1).Borders(12).LineStyle = 1
End With
Set oExcelApp = Nothing
End Function
Public Function TO_naryad_search()
On Error Resume Next

    Set oExcelApp = CreateObject("EXCEL.APPLICATION")
    If Err = 429 Then MsgBox "Application is Microsoft Excel is not installed!", vbExclamation, strMAIN_TITLE: Screen.MousePointer = vbDefault: Exit Function
    
    
    Dim sFile As String
    sFile = sCNV_txtDirShabl & "naryad.xls"
    If Not ExFile(sFile) Then MsgBox "Файл шаблона" & NL2 & sFile & NL2 & "не найден. Проверьте правильность указания пути шаблонов." & vbNewLine & "Для восстановления шаблонов воспользуйтесь вкладкой `Воссановление` в меню `Опции`", vbCritical, strMAIN_TITLE: Screen.MousePointer = vbDefault: Exit Function

    
    
    oExcelApp.WORkbooks.Open FileName:=sFile, ReadOnly:=True, ignOReReadOnlyRecommended:=False
    
    Set oWb = oExcelApp.ActiveWORkbook
    Set oWs = oExcelApp.Sheets(1)
    
    oExcelApp.Visible = True
        Dim x, Y As Long
        Dim tmp As String
        With oWs
              '.cells(2, 10) = procent_otp

        For x = 1 To frmNaryad.listres.ListItems.Count
               For Y = 1 To frmNaryad.listres.Columns.Count
                    .cells(x + 2, Y) = frmNaryad.listres.ListItems(x).SubItems(Y + 1).Caption
                Next Y

        Next x

       .range(.cells(1, 1), .cells(frmNaryad.listres.ListItems.Count + 2, 13)).Borders(7).LineStyle = 1
       .range(.cells(1, 1), .cells(frmNaryad.listres.ListItems.Count + 2, 13)).Borders(8).LineStyle = 1
       .range(.cells(1, 1), .cells(frmNaryad.listres.ListItems.Count + 2, 13)).Borders(9).LineStyle = 1
       .range(.cells(1, 1), .cells(frmNaryad.listres.ListItems.Count + 2, 13)).Borders(10).LineStyle = 1
       .range(.cells(1, 1), .cells(frmNaryad.listres.ListItems.Count + 2, 13)).Borders(11).LineStyle = 1
       .range(.cells(1, 1), .cells(frmNaryad.listres.ListItems.Count + 2, 13)).Borders(12).LineStyle = 1
           .cells.EntireColumn.AutoFit

          
End With
Set oExcelApp = Nothing
Screen.MousePointer = vbDefault

Exit Function
End Function

Public Function To_Listv()
On Error Resume Next

Dim c As Long
Dim x As Long
Dim cCnt As Long
Dim sFile As String
Dim sPathResoult As String
Dim cn As Long
Dim cn0 As Long
Call Reg_VK_List

'Screen.MousePointer = vbHourglass
With oWs
Set oExcelApp = CreateObject("EXCEL.APPLICATION")



                
                
            progress.Show
                sFile = sCNV_txtDirShabl & "Listv.xls"
                sPathResoult = sCNV_txtPathResoult
                             
                Call mysql.query("SELECT prnik_" & nowBase & ".txtvk, prnik_" & nowBase & ".fam, prnik_" & nowBase & ".name, prnik_" & nowBase & ".otch, prnik_" & nowBase & ".datar, otpravka_" & nowBase & ".data, naryad_" & nowBase & ".rodv, naryad_" & nowBase & ".punkt, naryad_" & nowBase & ".vch, prnik_" & nowBase & ".dir FROM prnik_" & nowBase & ", otpravka_" & nowBase & ", naryad_" & nowBase & " WHERE prnik_" & nowBase & ".otprvid = otpravka_" & nowBase & ".otpravkaid AND otpravka_" & nowBase & ".narid = naryad_" & nowBase & ".narid AND prnik_" & nowBase & ".txtvk like '%' ORder by txtvk, fam")
                
                If st = 0 Then GoTo vk_skip
                
                oExcelApp.Visible = False
            
                oExcelApp.WORkbooks.Open FileName:=sFile, ReadOnly:=True, ignOReReadOnlyRecommended:=True
                
                Set oWb = oExcelApp.ActiveWORkbook
                Set oWs = oExcelApp.Sheets(1)
                
                If Err = 429 Then MsgBox "Application Microsoft Excel is not installed!", vbExclamation, strMAIN_TITLE: Screen.MousePointer = vbDefault: Exit Function
                
                oExcelApp.Rows("5:5").Select
                oExcelApp.Selection.AutoFill Destination:=oExcelApp.Rows("5:" & st + 4), type:=0
                progress.ProgressBar1.Max = st
                
                    For x = 1 To st
                         oWs.cells(x + 4, 1) = CStr(x)
                        oWs.cells(x + 4, 2) = dat(2, x) & " " & dat(3, x) & " " & dat(4, x)
                        oWs.cells(x + 4, 3) = dat(1, x)
                        oWs.cells(x + 4, 4) = dat(5, x)
                        oWs.cells(x + 4, 5) = dat(6, x)
                        oWs.cells(x + 4, 6) = dat(7, x) & " " & dat(8, x) & " в/ч " & dat(9, x)
                        oWs.cells(x + 4, 7) = dat(11, x)
                    progress.ProgressBar1 = x
                    Next x
                cCnt = cCnt + st
                oWs.cells.EntireColumn.AutoFit
               Unload progress
               oExcelApp.Visible = True
               ' oWb.Close
               
vk_skip:
Set oExcelApp = Nothing
       
'
End With

Screen.MousePointer = vbDefault
End Function
Public Function To_List_gen_pred(frmDate As String, todate As String)
On Error Resume Next
frmDate = CnvDataWinToSql(frmDate)
todate = CnvDataWinToSql(todate)
Dim c As Long
Dim x As Long
Dim cCnt As Long
Dim sFile As String
Dim sPathResoult As String
Dim cn As Long
Dim cn0 As Long
Call Reg_VK_List

'Screen.MousePointer = vbHourglass
With oWs
Set oExcelApp = CreateObject("EXCEL.APPLICATION")



                Unload frmPred
                 progress.Show

                sFile = sCNV_txtDirShabl & "Listv_doc.xls"
                sPathResoult = sCNV_txtPathResoult
                      
                                       ' as                        1           2           3           4           5               6           7               8           9           10              11
                'Call mysql.query("SELECT prnik_" & nowBase & ".txtvk, prnik_" & nowBase & ".fam, prnik_" & nowBase & ".name, prnik_" & nowBase & ".otch, otpravka_" & nowBase & ".data, naryad_" & nowBase & ".oblkom ,naryad_" & nowBase & ".rodv, naryad_" & nowBase & ".punkt, naryad_" & nowBase & ".vch, prnik_" & nowBase & ".rpred, prnik_" & nowBase & ".datar FROM prnik_" & nowBase & ", otpravka_" & nowBase & ", naryad_" & nowBase & " WHERE prnik_" & nowBase & ".otprvid = otpravka_" & nowBase & ".otpravkaid AND otpravka_" & nowBase & ".narid = naryad_" & nowBase & ".narid AND prnik_" & nowBase & ".dataosp >= '" & frmDate & "' AND prnik_" & nowBase & ".dataosp <= '" & todate & "'  ORder by txtvk, fam")
                Call mysql.query("SELECT prnik_" & nowBase & ".txtvk, CONCAT(prnik_" & nowBase & ".fam, ' ',prnik_" & nowBase & ".name, ' ',prnik_" & nowBase & ".otch) as fam, prnik_" & nowBase & ".datar, otpravka_" & nowBase & ".data,naryad_" & nowBase & ".oblkom,prnik_" & nowBase & ".rpred, CONCAT(naryad_" & nowBase & ".rodv, ' ', naryad_" & nowBase & ".punkt, ' ', naryad_" & nowBase & ".vch)  FROM prnik_" & nowBase & ", otpravka_" & nowBase & ", naryad_" & nowBase & " WHERE otpravka_" & nowBase & ".otpravkaid = prnik_" & nowBase & ".otprvid and otpravka_" & nowBase & ".narid = naryad_" & nowBase & ".narid AND prnik_" & nowBase & ".dataosp >= '" & frmDate & "' AND prnik_" & nowBase & ".dataosp <= '" & todate & "'  Order by txtvk, fam")
                progress.ProgressBar1.Max = st
                If st = 0 Then GoTo vk_skip
                
                oExcelApp.Visible = False
                oExcelApp.WORkbooks.Open FileName:=sFile, ReadOnly:=True, ignOReReadOnlyRecommended:=True
                
                Set oWb = oExcelApp.ActiveWORkbook
                Set oWs = oExcelApp.Sheets(1)
                
                If Err = 429 Then MsgBox "Application Microsoft Excel is not installed!", vbExclamation, strMAIN_TITLE: Screen.MousePointer = vbDefault: Exit Function
                
                oExcelApp.Rows("5:5").Select
                oExcelApp.Selection.AutoFill Destination:=oExcelApp.Rows("5:" & st + 4), type:=0
                
                    For x = 1 To st
                         oWs.cells(x + 4, 1) = CStr(x)
                        oWs.cells(x + 4, 2) = dat(1, x)
                        oWs.cells(x + 4, 3) = dat(2, x)
                        oWs.cells(x + 4, 4) = dat(3, x)
                        oWs.cells(x + 4, 5) = dat(4, x)
                        oWs.cells(x + 4, 6) = dat(5, x)
                        oWs.cells(x + 4, 7) = dat(6, x)
                        oWs.cells(x + 4, 8) = dat(7, x)
                        progress.ProgressBar1 = x
                    Next x
                cCnt = cCnt + st
                oWs.cells.EntireColumn.AutoFit
                Unload progress
                oExcelApp.Visible = True
               
               
vk_skip:

Set oExcelApp = Nothing
oExcelApp.Quit
Unload progress

End With

Screen.MousePointer = vbDefault
End Function
Public Function To_marw(ID As Integer)
On Error Resume Next
'Screen.MousePointer = vbHourglass
    Set oExcelApp = CreateObject("EXCEL.APPLICATION")
    
        If Err = 429 Then MsgBox "Application Microsoft Excel is not installed!", vbExclamation, strMAIN_TITLE: Screen.MousePointer = vbDefault: Exit Function

    Dim sFile As String
    sFile = sCNV_txtDirShabl & "marw.xls"
    If Not ExFile(sFile) Then MsgBox "Файл шаблона" & NL2 & sFile & NL2 & "не найден. Проверьте правильность указания пути шаблонов." & vbNewLine & "Для восстановления шаблонов воспользуйтесь вкладкой `Воссановление` в меню `Опции`", vbCritical, strMAIN_TITLE: Screen.MousePointer = vbDefault: Exit Function
    
    If get_config("view_print") = 1 Then oExcelApp.Visible = False Else oExcelApp.Visible = True
    
    
    oExcelApp.WORkbooks.Open FileName:=sFile, ReadOnly:=True, ignOReReadOnlyRecommended:=True
    
    Set oWb = oExcelApp.ActiveWORkbook
    Set oWs = oExcelApp.Sheets(1)


Dim BUF() As String
Dim pr() As String
Dim cnt As Long
Dim s As Long
Dim cUpk As Long
Dim STROKA As Long
Dim c As Long
Dim t As String
Dim x As Long
Dim LST As Long
Dim CC As Long
Dim dUpk()  As Long

Call mysql.query("SELECT naryad_" & nowBase & ".oblkom, naryad_" & nowBase & ".rodv, naryad_" & nowBase & ".punkt, naryad_" & nowBase & ".vch, otpravka_" & nowBase & ".kolvo FROM naryad_" & nowBase & ", otpravka_" & nowBase & " WHERE otpravka_" & nowBase & ".narid = naryad_" & nowBase & ".narid AND otpravka_" & nowBase & ".otpravkaid = " & CStr(ID))
s = 11
'Screen.MousePointer = vbHourglass

cnt = st

With oWs
            .cells(8, 7) = dat(5, 1)
            .cells(17, 2) = "ст. Железнодорожная - ст. " & dat(3, 1)
            .cells(14, 2) = dat(2, 1) & " " & dat(3, 1) & " в/ч " & dat(4, 1)
            .cells(39, 2) = dateotp
            .cells(47, 2) = dateotp
            .cells(46, 9) = dateotp
End With
    

    
  
    
oExcelApp.Visible = True
    oExcelApp.SendKeys "^p"
  
    
    Set oExcelApp = Nothing
    
    Screen.MousePointer = vbDefault
    
End Function

Public Function To_naryad_rz(type_s As String)
On Error Resume Next
Dim x As Long
Dim lastX As Long
Dim stt As Long
Dim datt2() As String

    Set oExcelApp = CreateObject("EXCEL.APPLICATION")
        
        If Err = 429 Then MsgBox "Application Microsoft Excel is not installed!", vbExclamation, strMAIN_TITLE: Screen.MousePointer = vbDefault: Exit Function
    
        Dim sFile As String
    sFile = sCNV_txtDirShabl & "naryad.xls"
    If Not ExFile(sFile) Then MsgBox "Файл шаблона" & NL2 & sFile & NL2 & "не найден. Проверьте правильность указания пути шаблонов." & vbNewLine & "Для восстановления шаблонов воспользуйтесь вкладкой `Воссановление` в меню `Опции`", vbCritical, strMAIN_TITLE: Screen.MousePointer = vbDefault: Exit Function
    oExcelApp.WORkbooks.Open FileName:=sFile, ReadOnly:=True, ignOReReadOnlyRecommended:=True
    Set oWb = oExcelApp.ActiveWORkbook
    Set oWs = oExcelApp.Sheets(1)
    oExcelApp.Visible = True
        
        If type_s = "1" Then Call mysql.query("SELECT `narid` as narid1,concat(okrkom,okrkom_e),`oblkom`,rodv,`okr`,`punkt`,`doroga`,`vch`,`vrp`,`kolvo`, (select sum(kolvo) from naryad_srezki_" & nowBase & " where narid = narid1), `datanar`,`dolg`, prim FROM `naryad_" & nowBase & "` WHERE `dolg` <> '0' ORDER by datanar,oblkom,okrkom,okrkom_e ASC")
        If type_s = "2" Then Call mysql.query("SELECT `narid` as narid1,concat(okrkom,okrkom_e),`oblkom`,rodv,`okr`,`punkt`,`doroga`,`vch`,`vrp`,`kolvo`, (select sum(kolvo) from naryad_srezki_" & nowBase & " where narid = narid1), `datanar`,`dolg`, prim FROM `naryad_" & nowBase & "` WHERE `type` = '1' ORDER by oblkom,okrkom,okrkom_e ASC")
        If type_s = "3" Then Call mysql.query("SELECT `narid` as narid1,concat(okrkom,okrkom_e),`oblkom`,rodv,`okr`,`punkt`,`doroga`,`vch`,`vrp`,`kolvo`, (select sum(kolvo) from naryad_srezki_" & nowBase & " where narid = narid1), `datanar`,`dolg`, prim FROM `naryad_" & nowBase & "` WHERE `type` = '2' ORDER by oblkom,okrkom,okrkom_e ASC")
        If type_s = "4" Then Call mysql.query("SELECT `narid` as narid1,concat(okrkom,okrkom_e),`oblkom`,rodv,`okr`,`punkt`,`doroga`,`vch`,`vrp`,`kolvo`, (select sum(kolvo) from naryad_srezki_" & nowBase & " where narid = narid1), `datanar`,`dolg`, prim FROM `naryad_" & nowBase & "` WHERE `vch` like '%ЦНП%' ORDER by oblkom,okrkom,okrkom_e ASC")
        datt2() = dat()
        stt = st
        lastX = "3"
        With oWs
                  .cells(2, 10) = procent_otp
          For x = 1 To stt

            .cells(lastX, 1) = datt2(2, x)
            .cells(lastX, 2) = datt2(3, x)
            .cells(lastX, 3) = datt2(4, x)
            .cells(lastX, 4) = datt2(5, x)
            .cells(lastX, 5) = datt2(6, x)
            .cells(lastX, 6) = datt2(7, x)
            .cells(lastX, 7) = datt2(8, x)
            .cells(lastX, 8) = datt2(9, x)
            .cells(lastX, 9) = datt2(10, x)
            If datt2(11, x) = "" Then datt2(11, x) = "0"
            .cells(lastX, 10) = datt2(11, x)
            .cells(lastX, 11) = Int(datt2(10, x)) - Int(datt2(11, x))
            .cells(lastX, 12) = datt2(12, x)
            .cells(lastX, 13) = datt2(13, x)
            
            lastX = lastX + 1
            If Len(datt2(14, x)) > 2 Then
            
                                        .range("I" & lastX).AddComment
                            .range("I" & lastX).comment.Visible = False
                            .range("I" & lastX).comment.Text Text:="Примечание:" & Chr(10) & datt(14, Y)
            End If
            
            Next x
       
       .range(.cells(1, 1), .cells(lastX - 1, 13)).Borders(7).LineStyle = 1
       .range(.cells(1, 1), .cells(lastX - 1, 13)).Borders(8).LineStyle = 1
       .range(.cells(1, 1), .cells(lastX - 1, 13)).Borders(9).LineStyle = 1
       .range(.cells(1, 1), .cells(lastX - 1, 13)).Borders(10).LineStyle = 1
       .range(.cells(1, 1), .cells(lastX - 1, 13)).Borders(11).LineStyle = 1
       .range(.cells(1, 1), .cells(lastX - 1, 13)).Borders(12).LineStyle = 1
            .cells.EntireColumn.AutoFit
       End With
         Set oExcelApp = Nothing
        End Function

Public Function To_naryad_srezki(todate As String, type_s As Long)
'''''''''''''''''''''''''''Наряд - Срезки
On Error Resume Next
Dim x, Y, z As Long
Dim lastX As Long
Dim podtr As Boolean
Dim stt As Long
Dim tmpt As Long
Dim narid As Long


    Set oExcelApp = CreateObject("EXCEL.APPLICATION")
        
        If Err = 429 Then MsgBox "Application Microsoft Excel is not installed!", vbExclamation, strMAIN_TITLE: Screen.MousePointer = vbDefault: Exit Function
    
        Dim sFile As String
        sFile = sCNV_txtDirShabl & "naryad_srezki.xls"
        If Not ExFile(sFile) Then MsgBox "Файл шаблона" & NL2 & sFile & NL2 & "не найден. Проверьте правильность указания пути шаблонов." & vbNewLine & "Для восстановления шаблонов воспользуйтесь вкладкой `Воссановление` в меню `Опции`", vbCritical, strMAIN_TITLE: Screen.MousePointer = vbDefault: Exit Function
        oExcelApp.WORkbooks.Open FileName:=sFile, ReadOnly:=True, ignOReReadOnlyRecommended:=True
        Set oWb = oExcelApp.ActiveWORkbook
        Set oWs = oExcelApp.Sheets(1)
        oExcelApp.Visible = True
        
If type_s = "1" Then
Call mysql.query("SELECT naryad_" & nowBase & ".narid as fio ,CONCAT(okrkom,okrkom_e), oblkom, rodv,okr,punkt,doroga,vch,vrp, naryad_" & nowBase & ".kolvo, (select sum(naryad_srezki_" & nowBase & ".kolvo) from naryad_srezki_" & nowBase & " where naryad_srezki_" & nowBase & ".narid = fio), (naryad_" & nowBase & ".kolvo - (select sum(naryad_srezki_" & nowBase & ".kolvo) from naryad_srezki_" & nowBase & " where naryad_srezki_" & nowBase & ".narid = fio)), naryad_" & nowBase & ".datanar, sum(naryad_srezki_" & nowBase & ".kolvo)  from naryad_" & nowBase & ", naryad_srezki_" & nowBase & " where naryad_srezki_" & nowBase & ".narid = naryad_" & nowBase & ".narid and naryad_srezki_" & nowBase & ".data = '" & CnvDataWinToSql(todate) & "' group by naryad_srezki_" & nowBase & ".narid ORDER by okrkom,okrkom_e ASC")
z = st
        
        datt() = dat()
        With oWs
        .cells(4, 10) = otp(CnvDataWinToSql(todate))
        .cells(4, 9) = nar
        .cells(2, 8) = todate
         lastX = 6
        For Y = 1 To z
                        .cells(lastX, 1) = datt(2, Y)
                        .cells(lastX, 2) = datt(3, Y)
                        .cells(lastX, 3) = datt(4, Y)
                        .cells(lastX, 4) = datt(5, Y)
                        .cells(lastX, 5) = datt(6, Y)
                        .cells(lastX, 6) = datt(7, Y)
                        .cells(lastX, 7) = datt(8, Y)
                        .cells(lastX, 8) = datt(9, Y)
                        .cells(lastX, 9) = datt(10, Y)
                        .cells(lastX, 10) = datt(11, Y)
                        .cells(lastX, 11) = datt(12, Y)
                        .cells(lastX, 12) = datt(13, Y)
                        .cells(lastX, 13) = datt(14, Y)
                        
                            .cells(1, 1) = "Срезки"
                            .range("A1:C1").Select: oExcelApp.Selection.Merge
 lastX = lastX + 1
Next Y
End With
End If
                    
                        
If type_s = "3" Then
Call mysql.query("SELECT naryad_" & nowBase & ".narid as fio ,CONCAT(okrkom,okrkom_e), oblkom, rodv,okr,punkt,doroga,vch,vrp, naryad_" & nowBase & ".kolvo, (select sum(naryad_srezki_" & nowBase & ".kolvo) from naryad_srezki_" & nowBase & " where naryad_srezki_" & nowBase & ".narid = fio), (naryad_" & nowBase & ".kolvo - (select sum(kolvo) from naryad_srezki_" & nowBase & " where narid = fio)), naryad_" & nowBase & ".datanar, (select sum(naryad_srezki_" & nowBase & ".kolvo) from naryad_srezki_" & nowBase & " where naryad_srezki_" & nowBase & ".data = '" & CnvDataWinToSql(todate) & "' and naryad_srezki_" & nowBase & ".narid = fio)  from naryad_" & nowBase & ", naryad_srezki_" & nowBase & " where naryad_srezki_" & nowBase & ".narid = naryad_" & nowBase & ".narid and (naryad_" & nowBase & ".datanar = '" & CnvDataWinToSql(todate) & "' or  naryad_srezki_" & nowBase & ".data = '" & CnvDataWinToSql(todate) & "') group by naryad_" & nowBase & ".narid ORDER by okrkom,okrkom_e ASC")
z = st
        
        datt() = dat()
        With oWs
        .cells(4, 10) = otp(CnvDataWinToSql(todate))
        .cells(4, 9) = nar
        .cells(2, 8) = todate
         lastX = 6
        For Y = 1 To z
                        .cells(lastX, 1) = datt(2, Y)
                        .cells(lastX, 2) = datt(3, Y)
                        .cells(lastX, 3) = datt(4, Y)
                        .cells(lastX, 4) = datt(5, Y)
                        .cells(lastX, 5) = datt(6, Y)
                        .cells(lastX, 6) = datt(7, Y)
                        .cells(lastX, 7) = datt(8, Y)
                        .cells(lastX, 8) = datt(9, Y)
                        .cells(lastX, 9) = datt(10, Y)
                        .cells(lastX, 10) = datt(11, Y)
                        .cells(lastX, 11) = datt(12, Y)
                        .cells(lastX, 12) = datt(13, Y)
                        .cells(lastX, 13) = datt(14, Y)
                        
                            .cells(1, 1) = "Срезки + По плану"
                            .range("A1:C1").Select: oExcelApp.Selection.Merge
 lastX = lastX + 1
Next Y
End With
End If
                      
                      
                      
                      
       With oWs
       .range(.cells(6, 1), .cells(lastX - 1, 13)).Borders(7).LineStyle = 1
       .range(.cells(6, 1), .cells(lastX - 1, 13)).Borders(8).LineStyle = 1
       .range(.cells(6, 1), .cells(lastX - 1, 13)).Borders(9).LineStyle = 1
       .range(.cells(6, 1), .cells(lastX - 1, 13)).Borders(10).LineStyle = 1
       .range(.cells(6, 1), .cells(lastX - 1, 13)).Borders(11).LineStyle = 1
       .range(.cells(6, 1), .cells(lastX - 1, 13)).Borders(12).LineStyle = 1
            .cells.EntireColumn.AutoFit
       End With
         Set oExcelApp = Nothing
        End Function
Public Function To_naryad_all(append As String)
On Error Resume Next

Dim x, Y, z, lastX, r As Long
Dim datt() As String
Dim datt2() As String
Dim datt3() As String
Dim stt, sttt As Long


 Set oExcelApp = CreateObject("EXCEL.APPLICATION")
        
        If Err = 429 Then MsgBox "Application Microsoft Excel is not installed!", vbExclamation, strMAIN_TITLE: Screen.MousePointer = vbDefault: Exit Function
    
        Dim sFile As String
    sFile = sCNV_txtDirShabl & "naryad.xls"
    If Not ExFile(sFile) Then MsgBox "Файл шаблона" & NL2 & sFile & NL2 & "не найден. Проверьте правильность указания пути шаблонов." & vbNewLine & "Для восстановления шаблонов воспользуйтесь вкладкой `Воссановление` в меню `Опции`", vbCritical, strMAIN_TITLE: Screen.MousePointer = vbDefault: Exit Function
    oExcelApp.WORkbooks.Open FileName:=sFile, ReadOnly:=True, ignOReReadOnlyRecommended:=True
    Set oWb = oExcelApp.ActiveWORkbook
    Set oWs = oExcelApp.Sheets(1)
    oExcelApp.Visible = True
        lastX = "4"
        With oWs



Call mysql.query("select okr from naryad_" & nowBase & " where narid > '0' " & append & " group by okr ORDER by okr")
stt = st
datt() = dat()


.cells(3, 1) = "Всего"
    .range("A3:H3").Select: oExcelApp.Selection.Merge
          
            Call mysql.query("select sum(kolvo) from naryad_srezki_" & nowBase)
            .cells(3, 10) = dat(1, 1)
            Call mysql.query("select sum(kolvo) from naryad_" & nowBase)
            .cells(3, 9) = dat(1, 1)
            .cells(3, 11) = .cells(3, 9) - .cells(3, 10)
            Call mysql.query("select sum(dolg) from naryad_" & nowBase)
            .cells(3, 13) = dat(1, 1)
            .cells(2, 10) = procent_otp

For Y = 1 To stt

        .cells(lastX, 1) = datt(1, Y)
        .range("A" & lastX & ":M" & lastX & "").Select: oExcelApp.Selection.Merge
        .range("A" & lastX & ":M" & lastX & "").Interior.Color = 5287936
        lastX = lastX + 1
            Call mysql.query("SELECT rodv FROM naryad_" & nowBase & " WHERE okr='" & datt(1, Y) & "' " & append & " group by rodv ORDER by rodv ASC")
            sttt = st
            datt3() = dat()
            For z = 1 To sttt
        .cells(lastX, 1) = datt3(1, z)
        .range("A" & lastX & ":M" & lastX & "").Select: oExcelApp.Selection.Merge
        .range("A" & lastX & ":M" & lastX & "").Interior.Color = 65535
        lastX = lastX + 1
                

            Call mysql.query("SELECT `narid` as narid1,concat(okrkom,okrkom_e),`oblkom`,rodv,`okr`,`punkt`,`doroga`,`vch`,`vrp`,`kolvo`, (select sum(kolvo) from naryad_srezki_" & nowBase & " where narid = narid1), `datanar`,`dolg`, prim FROM naryad_" & nowBase & " WHERE okr='" & datt(1, Y) & "' AND rodv='" & datt3(1, z) & "' " & append & " ORder by okrkom,okrkom_e,oblkom ASC")
          datt2() = dat()
          
          For x = 1 To st

            .cells(lastX, 1) = datt2(2, x)
            .cells(lastX, 2) = datt2(3, x)
            .cells(lastX, 3) = datt2(4, x)
            .cells(lastX, 4) = datt2(5, x)
            .cells(lastX, 5) = datt2(6, x)
            .cells(lastX, 6) = datt2(7, x)
            .cells(lastX, 7) = datt2(8, x)
            .cells(lastX, 8) = datt2(9, x)
            .cells(lastX, 9) = datt2(10, x)
            If datt2(11, x) = "" Then datt2(11, x) = "0"
            .cells(lastX, 10) = datt2(11, x)
            .cells(lastX, 11) = Int(datt2(10, x)) - Int(datt2(11, x))
            .cells(lastX, 12) = datt2(12, x)
            .cells(lastX, 13) = datt2(13, x)
            
            lastX = lastX + 1
            If Len(datt2(14, x)) > 2 Then
            
                                        .range("I" & lastX).AddComment
                            .range("I" & lastX).comment.Visible = False
                            .range("I" & lastX).comment.Text Text:="Примечание:" & Chr(10) & datt(14, Y)
            End If
            
            Next x
            
            Next z
Next Y

       .range(.cells(1, 1), .cells(lastX - 1, 13)).Borders(7).LineStyle = 1
       .range(.cells(1, 1), .cells(lastX - 1, 13)).Borders(8).LineStyle = 1
       .range(.cells(1, 1), .cells(lastX - 1, 13)).Borders(9).LineStyle = 1
       .range(.cells(1, 1), .cells(lastX - 1, 13)).Borders(10).LineStyle = 1
       .range(.cells(1, 1), .cells(lastX - 1, 13)).Borders(11).LineStyle = 1
       .range(.cells(1, 1), .cells(lastX - 1, 13)).Borders(12).LineStyle = 1




End With




Set oExcelApp = Nothing
Screen.MousePointer = vbDefault

Exit Function
ErrH:
        MsgBox "Ошибка при генерировании файла Microsoft Exel:" & NL2 & "Выполняемая функция: " & """" & "To_F36" & """" & NL2 & "Номер ошибки: " & Err.Number & NL2 & "Описание: " & Err.Description, vbCritical, "Microsoft Excel - Форма №36"
        
End Function
Public Function To_naryad_date()

On Error Resume Next
Dim lastX As Long
Dim podrod() As String

    Set oExcelApp = CreateObject("EXCEL.APPLICATION")
        
        If Err = 429 Then MsgBox "Application Microsoft Excel is not installed!", vbExclamation, strMAIN_TITLE: Screen.MousePointer = vbDefault: Exit Function
    
        Dim sFile As String
    sFile = sCNV_txtDirShabl & "naryad_dni.xls"
    If Not ExFile(sFile) Then MsgBox "Файл шаблона" & NL2 & sFile & NL2 & "не найден. Проверьте правильность указания пути шаблонов." & vbNewLine & "Для восстановления шаблонов воспользуйтесь вкладкой `Воссановление` в меню `Опции`", vbCritical, strMAIN_TITLE: Screen.MousePointer = vbDefault: Exit Function
    oExcelApp.WORkbooks.Open FileName:=sFile, ReadOnly:=True, ignOReReadOnlyRecommended:=True
    Set oWb = oExcelApp.ActiveWORkbook
    Set oWs = oExcelApp.Sheets(1)
    oExcelApp.Visible = True
        lastX = "2"
        
        With oWs
        Call sorting_blok("datanar", "naryad_" & nowBase)
        
        For x = 1 To UBound(outm())
        Call mysql.query("SELECT okrkom,okrkom_e,oblkom,rodv,okr,punkt,doroga,vch,vrp,kolvo,narid,dolg,prim,type FROM naryad_" & nowBase & " WHERE `datanar` = '" & outm(x) & "' ORder by okrkom,okrkom_e,oblkom ASC")
            podrod() = dat()
            
            .range("A" & lastX & ":I" & lastX).Select
            oExcelApp.Selection.Merge
            oExcelApp.Selection.NumberFormat = "[$-F800]dddd, mmmm dd, yyyy"
            .cells(lastX, 1) = outm(x)
            
            For Y = 1 To st
                        lastX = lastX + 1
                        .cells(lastX, 1) = podrod(1, Y) & podrod(2, Y)
                        .cells(lastX, 2) = podrod(3, Y)
                        .cells(lastX, 3) = podrod(4, Y)
                        .cells(lastX, 4) = podrod(5, Y)
                        .cells(lastX, 5) = podrod(6, Y)
                        .cells(lastX, 6) = podrod(7, Y)
                        .cells(lastX, 7) = podrod(8, Y)
                        .cells(lastX, 8) = podrod(9, Y)
                        .cells(lastX, 9) = podrod(10, Y)

            Next Y
            lastX = lastX + 1
        Next x
        
       .range(.cells(1, 1), .cells(lastX - 1, 9)).Borders(7).LineStyle = 1
       .range(.cells(1, 1), .cells(lastX - 1, 9)).Borders(8).LineStyle = 1
       .range(.cells(1, 1), .cells(lastX - 1, 9)).Borders(9).LineStyle = 1
       .range(.cells(1, 1), .cells(lastX - 1, 9)).Borders(10).LineStyle = 1
       .range(.cells(1, 1), .cells(lastX - 1, 9)).Borders(11).LineStyle = 1
       .range(.cells(1, 1), .cells(lastX - 1, 9)).Borders(12).LineStyle = 1
  
        .cells.EntireColumn.AutoFit
        
        
        
        End With




Set oExcelApp = Nothing
Screen.MousePointer = vbDefault

Exit Function
ErrH:
        MsgBox "Ошибка при генерировании файла Microsoft Exel:" & NL2 & "Выполняемая функция: " & """" & "To_F36" & """" & NL2 & "Номер ошибки: " & Err.Number & NL2 & "Описание: " & Err.Description, vbCritical, "Microsoft Excel - Форма №36"
        
        
End Function
Public Function To_naryad_date_d()

On Error Resume Next
Dim lastX As Long
Dim ddate As String
Dim podrod() As String

ddate = InputBox("По состоянию на", "Наряд", Date)
    Set oExcelApp = CreateObject("EXCEL.APPLICATION")
        
        If Err = 429 Then MsgBox "Application Microsoft Excel is not installed!", vbExclamation, strMAIN_TITLE: Screen.MousePointer = vbDefault: Exit Function
    
    
        Dim sFile As String
    sFile = sCNV_txtDirShabl & "naryad_dni.xls"
    If Not ExFile(sFile) Then MsgBox "Файл шаблона" & NL2 & sFile & NL2 & "не найден. Проверьте правильность указания пути шаблонов." & vbNewLine & "Для восстановления шаблонов воспользуйтесь вкладкой `Воссановление` в меню `Опции`", vbCritical, strMAIN_TITLE: Screen.MousePointer = vbDefault: Exit Function
    oExcelApp.WORkbooks.Open FileName:=sFile, ReadOnly:=True, ignOReReadOnlyRecommended:=True
    Set oWb = oExcelApp.ActiveWORkbook
    Set oWs = oExcelApp.Sheets(1)
    oExcelApp.Visible = True
        lastX = "3"
        
        With oWs
        
        Call mysql.query("SELECT `narid` as narid1,concat(okrkom,okrkom_e),`oblkom`,rodv,`okr`,`punkt`,`doroga`,`vch`,`vrp`,`kolvo`, (select sum(kolvo) from naryad_srezki_" & nowBase & " where narid = narid1), `datanar`,`dolg`  FROM naryad_" & nowBase & " WHERE `datanar` = '" & CnvDataWinToSql(ddate) & "' ORder by okrkom,okrkom_e,oblkom ASC")
          .cells(2, 1) = dat(12, 1)
          
          .range("A2:I2").Select
            oExcelApp.Selection.Merge
            oExcelApp.Selection.NumberFormat = "[$-F800]dddd, mmmm dd, yyyy"
          
          For x = 1 To st

            .cells(lastX, 1) = dat(2, x)
            .cells(lastX, 2) = dat(3, x)
            .cells(lastX, 3) = dat(4, x)
            .cells(lastX, 4) = dat(5, x)
            .cells(lastX, 5) = dat(6, x)
            .cells(lastX, 6) = dat(7, x)
            .cells(lastX, 7) = dat(8, x)
            .cells(lastX, 8) = dat(9, x)
            .cells(lastX, 9) = dat(10, x)
            
            lastX = lastX + 1
            Next x
        
       .range(.cells(1, 1), .cells((lastX - 1), 9)).Borders(7).LineStyle = 1
       .range(.cells(1, 1), .cells((lastX - 1), 9)).Borders(8).LineStyle = 1
       .range(.cells(1, 1), .cells((lastX - 1), 9)).Borders(9).LineStyle = 1
       .range(.cells(1, 1), .cells((lastX - 1), 9)).Borders(10).LineStyle = 1
       .range(.cells(1, 1), .cells((lastX - 1), 9)).Borders(11).LineStyle = 1
       .range(.cells(1, 1), .cells((lastX - 1), 9)).Borders(12).LineStyle = 1
  
        .cells.EntireColumn.AutoFit
        
        
        
        End With




Set oExcelApp = Nothing
Screen.MousePointer = vbDefault

Exit Function
ErrH:
        MsgBox "Ошибка при генерировании файла Microsoft Exel:" & NL2 & "Выполняемая функция: " & """" & "To_F36" & """" & NL2 & "Номер ошибки: " & Err.Number & NL2 & "Описание: " & Err.Description, vbCritical, "Microsoft Excel - Форма №36"
        
        
End Function

Public Function To_naryad_ostatki()

On Error Resume Next
Dim lastX As Long
Dim otp As Long
Dim print_c As Boolean


    Set oExcelApp = CreateObject("EXCEL.APPLICATION")
        
        If Err = 429 Then MsgBox "Application Microsoft Excel is not installed!", vbExclamation, strMAIN_TITLE: Screen.MousePointer = vbDefault: Exit Function
    
        Dim sFile As String
        Dim stt As Long
    sFile = sCNV_txtDirShabl & "naryad.xls"
    If Not ExFile(sFile) Then MsgBox "Файл шаблона" & NL2 & sFile & NL2 & "не найден. Проверьте правильность указания пути шаблонов." & vbNewLine & "Для восстановления шаблонов воспользуйтесь вкладкой `Воссановление` в меню `Опции`", vbCritical, strMAIN_TITLE: Screen.MousePointer = vbDefault: Exit Function
    oExcelApp.WORkbooks.Open FileName:=sFile, ReadOnly:=True, ignOReReadOnlyRecommended:=True
    Set oWb = oExcelApp.ActiveWORkbook
    Set oWs = oExcelApp.Sheets(1)
    oExcelApp.Visible = True
        lastX = "3"
        
        With oWs
        .cells(2, 10) = procent_otp
    Call mysql.query("SELECT okrkom,okrkom_e,oblkom,rodv,okr,punkt,doroga,vch,vrp,kolvo,datanar,narid as narid1,dolg,prim,type, (select sum(kolvo) from naryad_srezki_" & nowBase & " where narid = narid1) FROM naryad_" & nowBase & " ORDER by okrkom,okrkom_e,oblkom ASC")
                datt() = dat()
                stt = st
                For x = 1 To stt
                If datt(10, x) <> datt(16, x) Or datt(13, x) <> "0" Then
                        
                        
                        .cells(lastX, 1) = datt(1, x) & datt(2, x)
                        .cells(lastX, 2) = datt(3, x)
                        .cells(lastX, 3) = datt(4, x)
                        .cells(lastX, 4) = datt(5, x)
                        .cells(lastX, 5) = datt(6, x)
                        .cells(lastX, 6) = datt(7, x)
                        .cells(lastX, 7) = datt(8, x)
                        .cells(lastX, 8) = datt(9, x)
                        .cells(lastX, 9) = datt(10, x)
                        
                        If datt(13, x) = "0" Then
                            .cells(lastX, 13) = ""
                        Else
                            .cells(lastX, 13) = datt(13, x)
                        End If
                        
                  
                       
                        .cells(lastX, 10) = datt(16, x)
                        .cells(lastX, 11) = .cells(lastX, 9) - .cells(lastX, 10)
                        
                        If Len(datt(14, x)) > 2 Then
                            
                            .range("I" & lastX).AddComment
                            .range("I" & lastX).comment.Visible = False
                            .range("I" & lastX).comment.Text Text:="Примечание:" & Chr(10) & datt(14, x)
                            
                        
                        End If
                        If datt(15, x) = 1 Then
                        .Rows(lastX & ":" & lastX).Select
                        oExcelApp.Selection.Font.Color = -16776961
                        End If
                        .cells(lastX, 12) = CnvDataSqLToWin(datt(11, x))
                        
                        lastX = lastX + 1
                End If
                Next x
        
       .range(.cells(1, 1), .cells(lastX - 1, 13)).Select
       oExcelApp.Selection.Font.size = "12"
       .range(.cells(1, 1), .cells(lastX - 1, 13)).Borders(7).LineStyle = 1
       .range(.cells(1, 1), .cells(lastX - 1, 13)).Borders(8).LineStyle = 1
       .range(.cells(1, 1), .cells(lastX - 1, 13)).Borders(9).LineStyle = 1
       .range(.cells(1, 1), .cells(lastX - 1, 13)).Borders(10).LineStyle = 1
       .range(.cells(1, 1), .cells(lastX - 1, 13)).Borders(11).LineStyle = 1
       .range(.cells(1, 1), .cells(lastX - 1, 13)).Borders(12).LineStyle = 1
  
        .cells.EntireColumn.AutoFit
        
        
        
        End With




Set oExcelApp = Nothing
Screen.MousePointer = vbDefault

Exit Function
ErrH:
        MsgBox "Ошибка при генерировании файла Microsoft Exel:" & NL2 & "Выполняемая функция: " & """" & "To_F36" & """" & NL2 & "Номер ошибки: " & Err.Number & NL2 & "Описание: " & Err.Description, vbCritical, "Microsoft Excel - Форма №36"
        
        
End Function
