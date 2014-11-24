VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPRNIK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Function Set_Vozvrat(cUpk As Long) As Boolean

Dim tmpSTR As String
Dim X As Long

Call mysql.query("select min(idprnik) from vozvrat_" & nowBase & "")
tmpSTR = "'" & val(DAT(1, 1)) - 1

Call mysql.query("select * from `prnik_" & nowBase & "` where `idprnik` = " & cUpk)

 If CLng(val(DAT(24, 1))) = 4 Then MsgBox "Простите, призывник " & NL2 & DAT(2, 1) & " " & DAT(3, 1) & " " & DAT(4, 1) & " " & DAT(5, 1) & NL2 & " находится в убывшей команде. Трогать его нельзя!", vbExclamation, strMAIN_TITLE: Exit Function

 If CLng(val(DAT(24, 1))) = 1 Then
'    MsgBox "Простите, призывник " & NL2 & DAT(1, 1) & " " & DAT(2, 1) & " " & DAT(3, 1) & " " & DAT(4, 1) & NL2 & " находится в убывшей команде. Трогать его нельзя!", vbExclamation, strMAIN_TITLE: Exit Sub
        If MsgBox("Призывник заблокирован. Для внесения изменений в данные необходим пароль." & NL2 & "Хотите внести призывника в таблицу возвратов?", vbExclamation + vbOKCancel, strMAIN_TITLE) = vbOK Then
                Call MessageBeep(40)
                If Not InputBox("Password:", "Обычная блокировка") = "5915" Then
                    Exit Function
                End If
        End If
End If

    For X = 1 To 18
        tmpSTR = tmpSTR & "','" & DAT(X + 1, 1)
    Next X
 tmpSTR = tmpSTR & "','"
 Call MessageBeep(40)
 
 Dim t As String
 t = Trim$(InputBox("Введите текст примечания:", "ометка возвратов", "ВОЗВРАТ"))
 tmpSTR = tmpSTR & t & "'"

If Len(Trim$(t)) = 0 Then MsgBox "Операция прервана пользователем!", vbExclamation, strMAIN_TITLE: Exit Function


Call mysql.query("INSERT INTO `vozvra_" & nowBase & "t` ( `idprnik` , `txtvk` , `fam` , `name` , `otch` , `rkom` , `rpred` , `datar` , `txtobraz` , `txtspec` , `txtsem` , `txtsud` , `servb` , `nomvb` , `gol` , `nog` , `formaa` , `formab` , `dataosp`, `prim`) VALUES (" & tmpSTR & ")")
Call mysql.query("delete from `prnik_" & nowBase & "` where `idprnik` = " & cUpk)
Dim inf  As String
'inf = tmpSTR(2) & " " & tmpSTR(3) & " " & tmpSTR(4) & " " & tmpSTR(1)
'Call log_sql(lgn, "Поместил в ВОЗВРАТ призывника " & inf & " военкомат")

End Function


Function Delete_FromCommand(cUpk As Long) As Boolean

On Error Resume Next

    Dim X As Long
    
    
If cUpk < 1 Then MsgBox "Возможность перемещения возвратов в базу отключена.", vbExclamation, strMAIN_TITLE: Exit Function

Call mysql.query("select `lock` from prnik_" & nowBase & " where idprnik = " & cUpk)


If CLng(DAT(1, 1)) = 4 Then MsgBox "Простите, призывник находится в убывшей команде. Трогать его нельзя!", vbExclamation, strMAIN_TITLE: Exit Function

If CLng(DAT(1, 1)) = 1 Then
        If MsgBox("Призывник заблокирован. Для удаления его из команды необходимо указать пароль." & NL2 & "Хотите удалить призывника из команды?", vbExclamation + vbOKCancel, strMAIN_TITLE) = vbOK Then
            Call MessageBeep(40)
            If Not InputBox("Password:", "Обычная блокировка") = "5915" Then Delete_FromCommand = False: Exit Function
        Else
            Delete_FromCommand = False: Exit Function
        End If
End If

    If MsgBox("Подтверждаете удаление призывника из команды?", vbYesNo + vbQuestion, strMAIN_TITLE) = vbYes Then
        Dim OTPRv_ID As Long
        Dim tColVo As Long

            Call mysql.query("select `otpravkaid`, `kolvo` from otpravka_" & nowBase & ", prnik_" & nowBase & " where otpravkaid = otprvid and idprnik = " & cUpk)

            OTPRv_ID = DAT(1, 1)
            tColVo = DAT(2, 1)
            
            Call mysql.query("UPDATE `otpravka_" & nowBase & "` SET `kolvo` = '" & tColVo - 1 & "' WHERE `otpravkaid` = '" & OTPRv_ID & "'")
            Call mysql.query("UPDATE `prnik_" & nowBase & "` SET `otprvid` = '0' WHERE `idprnik` = '" & cUpk & "'")
            Call mysql.query("UPDATE `prnik_" & nowBase & "` SET `vus` = '' WHERE `idprnik` = '" & cUpk & "'")
            Call mysql.query("Select naryad_" & nowBase & ".oblkom from otpravka_" & nowBase & ", naryad_" & nowBase & " where otpravka_" & nowBase & ".otpravkaid ='" & OTPRv_ID & "' and naryad_" & nowBase & ".narid = otpravka_" & nowBase & ".narid ")
            Dim komm As String
            Dim inf As String
            komm = DAT(1, 1)
            Call mysql.query("select fam, name, otch, txtvk from prnik_" & nowBase & " where idprnik ='" & cUpk & "'")
            inf = DAT(1, 1) & " " & DAT(2, 1) & " " & DAT(3, 1) & " " & DAT(4, 1)
            Call log_sql(lgn, "Удалил из команды " & komm & " призывника " & inf & " военкомат")

    End If
    
End Function
