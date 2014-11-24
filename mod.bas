Attribute VB_Name = "modMyFunc"
Option Explicit
Public Cnv As New CnvTo
Public LF As ListItem
Public INI As New clsINI
Public x, Y, z, c, zz As Long
Public GetColumn As Boolean
Public infIMG As Image
Public expupk As Integer
Public nowBase As String
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private stt As Long
Public datt() As String

Public sql_com As String
Public errf As Boolean
Public lock_tp As Integer
Public append As String


Public st                      As Long
Public STB                   As Long


Public mysql                As cMysql
'Public MREC                As ADODB.Recordset
Public nVK() As String

Public kolVK As Long
Public MAS(50)          As String
Public CRITICAL_OPER As String

'Public VAlue    As String
'Public DB_NAME      As String
'Public ME_PATH      As String
Public dat()            As String
Public sCOL()           As String



Public vrod(20)        As String
Public komend(22)         As String
Public INP_UPK       As String
Public lgn As String
Public acl As String

Public dateotp As String
Public table As String
Public log_type_act(2) As String
Public log_act(8) As String

'Public selTable As String
'Public bStop_Oper As BookmarkEnum
Public B_QUERY_SHOWERR As Boolean
Public p_com_id As Integer
Public p_com_name As String
'***************** Config`s CONSTANTS **************************************************

Public H_NAME As String
Public U_NAME As String
Public C_PASSWORD As String
Public D_BASE As String
Public C_PORT As Long
Public base_path As String
Public doc_out As String
Public iAPP_HideTest As Integer

Public iCNV_chShowObj As Integer
Public iCNV_chAutoPrint As Integer
Public iCNV_Copy  As Integer
Public sCNV_txtPathResoult  As String
Public sCNV_txtPattern  As String
Public sCNV_txtDirShabl  As String
Public outm() As String
Public iREP_chFIO As Integer
Public iREP_chVK As Integer
Public iREP_chUPK As Integer
Public iREP_chKom As Integer
Public iREP_chSend As Integer
Public sREP_txtSORt  As String
Public id_mk As Long
Public txtHost As String
Public txtUser As String
Public txtpass As String
Public txtDB As String
Public txtPORt As String
Public obraz(5) As String
Public iCOMM_chFillSelComm As Integer

Public sDIRS_FileName As String

'*************************************************************************************

Global Const NL2 = vbNewLine & vbNewLine
Global Const strMAIN_TITLE As String = "База Приывзников"
Global Const KOD_VOD As String * 3 = "837"
Global Const MAX_P As Long = 500
Global Const str_DELIMITER As String = "|"
Global Const strPRNIK_IN_SEND_COMMAND As String = "Этот призывник в списке отправленных комманд. Изменение каких-либо параметров невозможно!"
Global Const strNO_ADMIN As String = "Нет доступа к использованию данной функции. Необходимо являться членом группы администраторов, чтобы использовать эти возможности!"

Public Const HWND_TOPMOST = -1&
Public Const HWND_TOP = 0&
Public Const HWND_BOTTOM = 1&

Public Const SWP_NOSIZE = &H1&
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOZORDER = &H4
Public Const SWP_NOREDRAW = &H8
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_SHOWWINDOW = &H40
Public Const SWP_HIDEWINDOW = &H80
Public Const SWP_NOCOPYBITS = &H100
Public Const SWP_NOOWNERZORDER = &H200
Public Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Public Const SWP_NOREPOSITION = SWP_NOOWNERZORDER

Public Const CFG_FILENAME As String = "Priz.ini"
Public VRP_REPLACE_VOD() As String
Global Const strCHAR_VOD As String = "V"
Public Const VRP_COUNT As Long = 8
Public Const VRP_VOD_STR As String = "ВОД"
Public Const LWA_COLORKEY = &H1
Public Const LWA_ALPHA = &H2
Public Const GWL_EXSTYLE = (-20)
Public Const WS_EX_LAYERED = &H80000

Private Declare Sub InitCommonControls Lib "Comctl32.dll" ()
'/***********************
Public Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectORy As String, ByVal nShowCmd As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function GetInputState Lib "user32" () As Long
Declare Function SHGetPathFROMIDList Lib "shell32.dll" Alias "SHGetPathFROMIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
'Declare Function SHBrowseFORFolder Lib "shell32.dll" Alias "SHBrowseFORFolderA" (lpBrowseInfo As BROWSEINFO) As Long
'Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long


Public Sub SetTransparent(hWnd As Long, Layered As Byte)

    Dim ret As Long
        ret = GetWindowLong(hWnd, GWL_EXSTYLE)
        ret = ret Or WS_EX_LAYERED
        SetWindowLong hWnd, GWL_EXSTYLE, ret
        SetLayeredWindowAttributes hWnd, 0, Layered, LWA_ALPHA

End Sub
Public Function GetClientWidth(hWnd As Long) As Long
Dim lpRect As RECT
Dim lReturn As Long
    lReturn = GetClientRect(hWnd, lpRect)
    GetClientWidth = lpRect.Right - lpRect.Left
End Function
Public Function ReSizeColumnHeaders_new(ListView As vbalListViewCtl)
Dim HeadersWidth As Long
Dim ClientWidth As Long
 For x = 1 To ListView.Columns.Count
        HeadersWidth = HeadersWidth + ListView.Columns.Item(x).Width
    Next x
    ClientWidth = GetClientWidth(ListView.hWnd) * Screen.TwipsPerPixelX
 For x = 1 To ListView.Columns.Count
   ListView.Columns.Item(x).Width = (ClientWidth * ListView.Columns.Item(x).Width) \ HeadersWidth
    Next x
    End Function


Public Function ReSizeColumnHeaders(ListView As MSComctlLib.ListView)

Dim HeadersWidth As Long
Dim ClientWidth As Long
Dim ColumnHeader As MSComctlLib.ColumnHeader
    For Each ColumnHeader In ListView.ColumnHeaders
        HeadersWidth = HeadersWidth + ColumnHeader.Width
    Next
    ClientWidth = GetClientWidth(ListView.hWnd) * Screen.TwipsPerPixelX
    For Each ColumnHeader In ListView.ColumnHeaders
        ColumnHeader.Width = (ClientWidth * ColumnHeader.Width) \ HeadersWidth
    Next
End Function

Sub Main()

    
End Sub

Public Sub INIT_VRP_LIST()

    ReDim VRP_REPLACE_VOD(1 To VRP_COUNT)
    
    VRP_REPLACE_VOD(1) = "ЛИН"
    VRP_REPLACE_VOD(2) = "СЕР"
    VRP_REPLACE_VOD(3) = "ВОД"
    VRP_REPLACE_VOD(4) = "ЗАС"
    VRP_REPLACE_VOD(5) = "ССОА"
    VRP_REPLACE_VOD(6) = "АПА-80"
    VRP_REPLACE_VOD(7) = "СПС"
    VRP_REPLACE_VOD(8) = "СУД"
    
End Sub

Public Function CnvDataSqLToWin(sqlData As String) As String
On Error Resume Next
Dim t() As String

    t = Split(sqlData, "-")
    
    CnvDataSqLToWin = t(2) & "." & t(1) & "." & t(0)

End Function


Public Function CnvDataWinToSql(sqlData As String) As String
On Error Resume Next
Dim t() As String

    t = Split(sqlData, ".")
    
    CnvDataWinToSql = Year("01.01." & t(2)) & "-" & Format$(t(1), "00") & "-" & Format(t(0), "00")

End Function

Sub MovePrnikToBase(nUpk As Long)
On Error Resume Next
Dim newDate As String
    newDate = CnvDataWinToSql(Date)
        'Check
        Call mysql.query("SELECT * FROM prnik_" & nowBase & " WHERE idprnik='" & nUpk & "'")
        ''''
        Dim newUpk As Integer
        If st > 0 Then
        Call mysql.query("SELECT max(idprnik) FROM prnik_" & nowBase & "")
        newUpk = dat(1, 1) + 1
        Else
        newUpk = nUpk
        End If
        ''''''
               
        Call mysql.query("SELECT * FROM `delprnik_" & nowBase & "` WHERE idprnik=" & nUpk)
        Call mysql.query("INSERT INTO `prnik_" & nowBase & "` VALUES ('" & newUpk & "','" & dat(2, 1) & "','" & dat(3, 1) & "','" & dat(4, 1) & "','" & dat(5, 1) & "','" & dat(6, 1) & "','" & dat(7, 1) & "','" & dat(8, 1) & "','" & dat(9, 1) & "','" & dat(10, 1) & "','" & dat(11, 1) & "','" & dat(12, 1) & "','" & dat(13, 1) & "','" & dat(14, 1) & "','" & dat(15, 1) & "','" & dat(16, 1) & "','" & dat(17, 1) & "','" & dat(18, 1) & "','" & newDate & "','" & dat(20, 1) & "','" & dat(21, 1) & "','" & dat(22, 1) & "','" & dat(23, 1) & "','" & dat(24, 1) & "','" & dat(25, 1) & "') ")
        Dim inf As String
        inf = dat(3, 1) & " " & dat(4, 1) & " " & dat(5, 1) & " " & dat(2, 1)
        Call mysql.query("DELETE FROM `delprnik_" & nowBase & "` WHERE idprnik=" & nUpk)
        
        
        
        Call mysql.query("SELECT * FROM `prnik_del_" & nowBase & "` WHERE idprnik=" & nUpk)
        
        Dim in_com As String
        Dim x As Long
            in_com = "'" & newUpk & "'" & ","
        For x = 2 To UBound(dat()) - 1
            in_com = in_com & "'" & dat(x, 1) & "'" & ","
        Next x
            in_com = in_com & "'" & dat(UBound(dat()), 1) & "'"
            Call mysql.query("insert into prnik_" & nowBase & " VAlues(" & in_com & ")")
            Call mysql.query("DELETE FROM prnik_del_" & nowBase & " WHERE idprnik='" & nUpk & "'")
        
        
        Call log_sql("0", "4", nUpk, "")
End Sub


Public Sub READ_CONFIG()
        Dim doc_out As String
        doc_out = "\\srv\baza\out\"
        INI.FileName = "\\srv\baza\Priz.ini"
        H_NAME = INI.VAlue("con", "host")
        C_PORT = VAl(INI.VAlue("con", "pORt"))
        U_NAME = INI.VAlue("con", "user")
        C_PASSWORD = INI.VAlue("con", "pass")
        D_BASE = INI.VAlue("con", "db")
        '----------------------------------------
        base_path = INI.VAlue("main", "basepath")
        sCNV_txtPathResoult = base_path & INI.VAlue("main", "PathResoult")
        sCNV_txtDirShabl = base_path & INI.VAlue("main", "DirShabl")
        iCNV_chShowObj = INI.VAlue("main", "ShowObj")
        iCNV_chAutoPrint = INI.VAlue("main", "AutoPrint")
        iCNV_Copy = VAl(INI.VAlue("main", "Copyes"))
        sCNV_txtPattern = INI.VAlue("main", "Pattern")
        
        sDIRS_FileName = INI.VAlue("DIRS", "Dir")
        iREP_chFIO = INI.VAlue("REPORT", "FIO")
        iREP_chVK = INI.VAlue("REPORT", "VK")
        iREP_chUPK = INI.VAlue("REPORT", "UPK")
        iREP_chKom = INI.VAlue("REPORT", "Kom")
        iREP_chSend = INI.VAlue("REPORT", "Send")
        sREP_txtSORt = INI.VAlue("REPORT", "SORt")
    
        iCOMM_chFillSelComm = INI.VAlue("LISTCOMM", "FillSELECTed")
        Exit Sub

End Sub

Public Sub SAVE_CONFIG()

        INI.FileName = CFG_FILENAME
        INI.VAlue("Z_CON", "H_NAME") = H_NAME
        INI.VAlue("Z_CON", "U_NAME") = U_NAME
        INI.VAlue("Z_CON", "PASS") = C_PASSWORD
        INI.VAlue("Z_CON", "DB_NAME") = D_BASE
        INI.VAlue("Z_CON", "C_PORT") = C_PORT
        INI.UPDATEFile
        
End Sub


Public Function ExFile(strPathName As String) As Boolean
    On Error Resume Next
    
    Dim res As Long
    
    Err = 0
     res = FileLen(strPathName)
     
     If res > 0 Then ExFile = True: Exit Function
     
     If Not Err = 0 Then
         ExFile = False
     Else
         ExFile = True
     End If
     
End Function

Public Function Lock_Comm(id_COM_OTPR As Long, kLock As Long)
Call mysql.query("UPDATE `prnik_" & nowBase & "` SET `lock` = `" & kLock & "` WHERE `otprvid` = '" & id_COM_OTPR & "'")
Call mysql.query("UPDATE `otpravka_" & nowBase & "` SET `kLock` = `" & kLock & "` WHERE `otpravkaid` = '" & id_COM_OTPR & "'")
End Function

Public Function UnLock_Comm(id_COM_OTPR As Long, Optional kLock As Long = 0)
    Call mysql.query("UPDATE `prnik_" & nowBase & "` SET `lock` = `" & kLock & "` WHERE `otprvid` = " & id_COM_OTPR)
End Function

Public Function log_sql(type_act, act As Long, to_id, argw As String)
On Error Resume Next
Dim idnew As Long
Dim str As String
Call log_types
Call mysql.query("SELECT max(id) FROM logs_" & nowBase & " ")
If dat(1, 1) = "" Then
    idnew = "1"
Else
    idnew = dat(1, 1) + 1
End If
    
str = "'" & idnew & "','" & type_act & "','" & act & "','" & to_id & "','" & argw & "','" & CnvDataWinToSql(Date) & "','" & time & "','" & lgn & "'"

Call mysql.query("Insert into logs_" & nowBase & " VAlues (" & str & ")")
End Function
Public Function get_access()
On Error Resume Next
Call mysql.query("SELECT access FROM users WHERE name='" & lgn & "'")
acl = dat(1, 1)
End Function
Public Function sorting_blok(arg_w As String, arg_v As String)
On Error Resume Next
Call mysql.query("SELECT `" & arg_w & "` FROM `" & arg_v & "` GROUP BY `" & arg_w & "` ORDER by `" & arg_w & "`")
    ReDim outm(st)
        For x = 1 To st
             outm(x) = dat(1, x)
        Next x
    End Function
Public Function get_VAlue(Name As String) As String
Call mysql.query("SELECT VAlue FROM config WHERE name='" & Name & "'")
get_VAlue = dat(1, 1)
End Function
Public Function get_config(Name As String) As String
Call mysql.query("SELECT " & Name & " FROM bases WHERE VAl='" & nowBase & "'")
get_config = dat(1, 1)
End Function



Public Function get_name_base(VAl As String) As String
    Call mysql.query("SELECT name FROM bases WHERE VAl='" & VAl & "'")
        get_name_base = dat(1, 1)
End Function
Public Function get_fio(login As String) As String
Call mysql.query("SELECT fio FROM users WHERE name='" & login & "'")
get_fio = dat(1, 1)
End Function
Public Function log_get_fio(ID As Long) As String
Call mysql.query("SELECT fam,name,otch,txtvk FROM prnik_" & nowBase & " WHERE idprnik='" & ID & "'")
If st > 0 Then
log_get_fio = dat(1, 1) & " " & dat(2, 1) & " " & dat(3, 1) & " (" & dat(4, 1) & " ОВК)"
Else
log_get_fio = "Удален ранее из базы"
End If
End Function
Public Function log_get_kom(ID As Long) As String
Call mysql.query("SELECT oblkom FROM naryad_" & nowBase & ", otpravka_" & nowBase & " WHERE naryad_" & nowBase & ".narid=otpravka_" & nowBase & ".narid AND otpravkaid='" & ID & "'")
If st > 0 Then
log_get_kom = "Областная команда - " & dat(1, 1)
Else
log_get_kom = "Удалена ранее из базы"
End If
End Function
Public Function get_info_prnik(ID As String) As String
Call mysql.query("SELECT naryad_" & nowBase & ".rodv,naryad_" & nowBase & ".punkt,naryad_" & nowBase & ".vch,otpravka_" & nowBase & ".data FROM naryad_" & nowBase & ",otpravka_" & nowBase & ",prnik_" & nowBase & " WHERE  otpravka_" & nowBase & ".narid=naryad_" & nowBase & ".narid AND otpravka_" & nowBase & ".otpravkaid=prnik_" & nowBase & ".otprvid AND prnik_" & nowBase & ".idprnik='" & ID & "'")
If st > 0 Then get_info_prnik = dat(1, 1) & " " & dat(2, 1) & " в/ч " & dat(3, 1) & ". Отправлен " & CnvDataSqLToWin(dat(4, 1)): 'Call mysql.free_result
End Function

Public Function procent_otp() As Double
On Error Resume Next
Dim otp, nar As Long
Call mysql.query("select sum(kolvo) from naryad_srezki_" & nowBase)
otp = dat(1, 1)
Call mysql.query("select sum(kolvo) from naryad_" & nowBase)
nar = dat(1, 1)
procent_otp = otp / nar
End Function
Public Function nar() As Long
On Error Resume Next
Call mysql.query("select sum(kolvo) from naryad_" & nowBase)
nar = dat(1, 1)
End Function
Public Function otp(data As String) As Long
On Error Resume Next
Call mysql.query("select sum(kolvo) from naryad_srezki_" & nowBase & " where data <= '" & data & "'")
otp = dat(1, 1)
End Function

Public Sub Reg_VK_List()
nVK() = Split(get_config("vk"), ";")
End Sub
Public Sub log_types()

log_type_act(0) = "Призывник"
log_type_act(1) = "Команда"
log_type_act(2) = "Наряд"
log_act(0) = "Добавлен в команду"
log_act(1) = "Удален из команды"
log_act(2) = "Изменен"
log_act(3) = "Удален из базы"
log_act(4) = "Востановлен в базу"
log_act(5) = "Заблокировал"
log_act(6) = "Разблокирован"
log_act(7) = "Добавилена команда"
log_act(8) = "Удалена команда"
End Sub

Public Sub obraz_in()
obraz(0) = "Высшее"
obraz(1) = "Незаконченное высшее"
obraz(2) = "Средне-специальное"
obraz(3) = "Среднее"
obraz(4) = "Неполное среднее"
obraz(5) = "9 классов и ниже"
End Sub
