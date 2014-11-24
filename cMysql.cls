VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMysql"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
                             

Option Explicit
Private myConHANDle As Long
Private myRec As API_MYSQL
Private myRec_res As API_MYSQL_RES
Private myRec_field As API_MYSQL_FIELD
Private myRec_rows As API_MYSQL_ROWS
Private mUser As String
Private mPasswORd As String
Private mHost As String
Private mDb As String
Private mUnix_socket As String
Private mClientFlag As Long
Private mPORt As Long

Private Sub Class_Initialize()
On Error Resume Next
    myConHANDle = API_mysql_init(myRec)
    If myConHANDle Then
      CopyMemory myRec, ByVal myConHANDle, LenB(myRec)
    Else
      MsgBox "Непредвиденная ошибка при инициализации класса cMySql:" & NL2 & Err.Description, vbCritical, ""
    End If
End Sub

Private Sub Class_Terminate()
    If myConHANDle <> 0 Then Call API_mysql_close(myRec)
End Sub
Public Function real_connect(Optional host As String = vbNullString, Optional user As String = vbNullString, Optional passwd As String = vbNullString, Optional DB As String = vbNullString, Optional Port As Long = MYSQL_PORT, Optional Unix_Socket As String = vbNullString, Optional clientflag As Long = 0) As Boolean


   If Len(host) > 0 Then mHost = host
   If Len(user) > 0 Then mUser = user
   If Len(passwd) > 0 Then mPasswORd = passwd
   If Len(DB) > 0 Then mDb = DB
   If Len(Unix_Socket) > 0 Then mUnix_socket = Unix_Socket
   If Port <> MYSQL_PORT Then mPORt = Port
   If clientflag <> 0 Then mClientFlag = clientflag

ret:
   On Error GoTo ErrH
   myConHANDle = API_mysql_real_connect(myRec, StrPtr(StrConv(mHost, vbFromUnicode)), StrPtr(StrConv(mUser, vbFromUnicode)), StrPtr(StrConv(mPasswORd, vbFromUnicode)), StrPtr(StrConv(mDb, vbFromUnicode)), Port, StrPtr(StrConv(mUnix_socket, vbFromUnicode)), mClientFlag)
    If myConHANDle Then
      CopyMemory myRec, ByVal myConHANDle, LenB(myRec)
      real_connect = True
    Else
      If MsgBox("Возникла ошибка при подключении к `" & host & "` - [" & mDb & "]." & NL2 & "Номер ошибки: " & API_mysql_errno(myRec) & NL2 & "Описание: " & ptr2str(API_mysql_error(myRec)) & NL2 & "Хотите повторить попытку подключения?", vbExclamation + vbRetryCancel, "") = vbRetry Then GoTo ret Else End
    End If
    Exit Function
    
ErrH:
    
    If Not Err = 0 Then If MsgBox("Возникла ошибка при подключении к `" & host & "` - [" & mDb & "]." & NL2 & "Номер ошибки: " & Err.Number & NL2 & "Описание: " & Err.Description & NL2 & "Хотите повторить попытку подключения?", vbExclamation + vbRetryCancel, Err.Source) = vbRetry Then GoTo ret Else End
    
End Function


Public Function list_tables(Optional wild As String = "%") As Long

On Error GoTo ErrH
Dim ret As Long
  ret = API_mysql_list_tables(myRec, StrPtr(StrConv(wild, vbFromUnicode)))
  If ret Then
    CopyMemory myRec_res, ByVal ret, LenB(myRec_res)
     list_tables = ConvertResult(myRec_res)
  Else
    list_tables = False
    MsgBox "Ошибка выполнения запроса MySQL Query" & NL2 & "Номер ошибки: " & API_mysql_errno(myRec) & NL2 & "Описание: " & ptr2str(API_mysql_error(myRec)), vbExclamation, strMAIN_TITLE
  End If

Exit Function
ErrH:
        MsgBox "Ошибка выполнения запроса MySQL Query" & NL2 & "Номер ошибки: " & Err.Number & NL2 & "Описание: " & Err.Description, vbExclamation, strMAIN_TITLE
End Function

Private Function ptr2str(ByVal ptr As Long) As String
On Error Resume Next
    If ptr = 0 Then Exit Function    'yeah..should never happen though
    Dim test As Byte
    Dim sout() As Byte
    Dim cchars As Long
     cchars = lstrlen(ptr)
     If cchars = 0 Then Exit Function
     ReDim sout(1 To cchars)                   'byte array
     sout = Space$(cchars)
     CopyMemory sout(1), ByVal (ptr), cchars
    ptr2str = StripNull(StrConv(sout, vbUnicode))  'convert byte array (lpstr) to unicode
End Function
Private Function StripNull(sName As String) As String
   Dim F As Long
   F = InStr(sName, vbNullChar)
   If F > 0 Then
    StripNull = Trim$(Left$(sName, F - 1))
   Else
    StripNull = Trim$(sName)
   End If
End Function
Public Function affected_rows() As Long
   affected_rows = convert642l(API_mysql_affected_rows(myRec))
End Function
Public Sub create_db(DB As String)
Dim ret As Long
   ret = API_mysql_create_db(myRec, StrPtr(StrConv(DB, vbFromUnicode)))
   If ret <> 0 Then Err.Raise API_mysql_errno(myRec), "cMysql::create_db", ptr2str(API_mysql_error(myRec))
End Sub
Public Sub debug_push(DBUG_PUSH As String) ' does a DBUG_PUSH with the given string.
  Call API_mysql_debug(StrPtr(StrConv(DBUG_PUSH, vbFromUnicode)))
End Sub
Public Sub drop_db(DB As String)
Dim ret As Long
   ret = API_mysql_drop_db(myRec, StrPtr(StrConv(DB, vbFromUnicode)))
   If ret <> 0 Then Err.Raise API_mysql_errno(myRec), "cMysql::drop_db", ptr2str(API_mysql_error(myRec))
End Sub
Public Sub dump_debug_info() 'fORces server to dump debug info
   Dim ret As Long
   ret = API_mysql_dump_debug_info(myRec)
   If ret <> 0 Then Err.Raise API_mysql_errno(myRec), "cMysql::dump_debug_info", ptr2str(API_mysql_error(myRec))
End Sub
Public Function eof() As Boolean
   eof = API_mysql_eof(myRec_res)
End Function
Public Function err_no() As Long
  err_no = API_mysql_errno(myRec)
End Function
Public Function err_description() As String
  err_description = ptr2str(API_mysql_error(myRec))
End Function
Public Sub escape_string(to_ As String, FROM_ As String, FROM_length As Long)
Dim ret As Long
  ret = API_mysql_escape_string(myRec, StrPtr(StrConv(to_, vbFromUnicode)), StrPtr(StrConv(FROM_, vbFromUnicode)), FROM_length)
End Sub
Public Sub free_result()
  'free the memeORy
  Dim ret As Long
  ret = API_mysql_free_result(myRec_res)
  If ret <> 0 Then Err.Raise API_mysql_errno(myRec), "cMysql::free_result", ptr2str(API_mysql_error(myRec))
End Sub
Public Function get_client_info() As String
Dim ret As Long
 ret = API_mysql_get_client_info
 If ret <> 0 Then get_client_info = ptr2str(ret) Else Err.Raise API_mysql_errno(myRec), "cMysql::get_client_info", ptr2str(API_mysql_error(myRec))
End Function
Public Function get_host_info() As String
Dim ret As Long
 ret = API_mysql_get_host_info(myRec)
 If ret <> 0 Then get_host_info = ptr2str(ret) Else Err.Raise API_mysql_errno(myRec), "cMysql::get_host_info", ptr2str(API_mysql_error(myRec))
End Function
Public Function get_proto_info() As String
Dim ret As Long
 ret = API_mysql_get_proto_info(myRec)
 If ret <> 0 Then get_proto_info = ptr2str(ret) Else Err.Raise API_mysql_errno(myRec), "cMysql::get_proto_info", ptr2str(API_mysql_error(myRec))
End Function
Public Function get_server_info()
Dim ret As Long
 ret = API_mysql_get_server_info(myRec)
 If ret <> 0 Then get_server_info = ptr2str(ret) Else Err.Raise API_mysql_errno(myRec), "cMysql::get_server_info", ptr2str(API_mysql_error(myRec))
End Function
Public Function info() As String
Dim ret As Long
 ret = API_mysql_info(myRec)
 If ret <> 0 Then info = ptr2str(ret) Else Err.Raise API_mysql_errno(myRec), "cMysql::info", ptr2str(API_mysql_error(myRec))
End Function
Public Function Insert_ID() As Long
  Insert_ID = convert642l(API_mysql_insert_id(myRec))
End Function
Public Sub Kill(PID As Long)
    Dim ret As Long
    ret = API_mysql_kill(myRec, PID)
    If ret <> 0 Then Err.Raise API_mysql_errno(myRec), "cMysql::kill", ptr2str(API_mysql_error(myRec))
End Sub

Public Function num_fields() As Long
  num_fields = API_mysql_num_fields(myRec_res)
End Function
Public Function num_rows() As Long
  num_rows = convert642l(API_mysql_num_rows(myRec_res))
End Function

Private Sub data_seek(row As Long)  'Seeks to an arbitrary row in a query result set.
  Dim a As Double
  CopyMemory a, convertl264(row).bytes(1), 8
   Call API_mysql_data_seek(myRec_res, a)
End Sub

Public Sub ping()
  Dim ret As Long
  ret = API_mysql_ping(myRec)
  If ret <> 0 Then Err.Raise API_mysql_errno(myRec), "cMysql::ping", ptr2str(API_mysql_error(myRec))
End Sub
Public Function query(SQL As String) As Long
Dim ret As Long
Debug.Print SQL
On Error GoTo ErrH
            ret = API_mysql_query(myRec, StrPtr(StrConv(SQL, vbFromUnicode)))
            
            If ret = 0 Then
            
                    ret = API_mysql_store_result(myRec)
                      If ret Then
                            CopyMemory myRec_res, ByVal ret, LenB(myRec_res)
                             query = ConvertResult(myRec_res)
                        Else
                        query = 2
                        Exit Function
                      End If
                       query = True
                      'MsgBox "Функция StORe_Result возвратила нулевое значение!", vbInfORmation, strMAIN_TITLE
             Else
                      
                        query = False
                       If Not B_QUERY_SHOWERR Then MsgBox "Ошибка выполнения запроса MySQL Query" & NL2 & "Номер ошибки: " & API_mysql_errno(myRec) & NL2 & "Описание: " & ptr2str(API_mysql_error(myRec)), vbExclamation, strMAIN_TITLE
             
            End If
             Exit Function
             
ErrH:
        MsgBox "Ошибка выполнения запроса MySQL Query" & NL2 & "Номер ошибки: " & Err.Number & NL2 & "Описание: " & Err.Description, vbExclamation, strMAIN_TITLE
             
             
End Function

Public Sub Refresh(refresh_options As API_refresh_options)
Dim ret As Long
  ret = API_mysql_refresh(myRec, refresh_options)
  If ret <> 0 Then Err.Raise API_mysql_errno(myRec), "cMysql::refresh", ptr2str(API_mysql_error(myRec))
End Sub
Public Sub SELECT_db(DB As String)
Dim ret As Long
  ret = API_mysql_SELECT_db(myRec, StrPtr(StrConv(DB, vbFromUnicode)))
  If ret <> 0 Then Err.Raise API_mysql_errno(myRec), "cMysql::SELECT_db", ptr2str(API_mysql_error(myRec))
End Sub
Public Sub shutdown()
Dim ret As Long
  ret = API_mysql_shutdown(myRec)
  If ret <> 0 Then Err.Raise API_mysql_errno(myRec), "cMysql::shutdown", ptr2str(API_mysql_error(myRec))
End Sub
Public Function stat() As String
Dim ret As Long
 ret = API_mysql_stat(myRec)
 If ret <> 0 Then stat = ptr2str(ret) Else Err.Raise API_mysql_errno(myRec), "cMysql::stat", ptr2str(API_mysql_error(myRec))
End Function
Public Function thread_id() As Long
Dim ret As Long
 ret = API_mysql_thread_id(myRec)
 If ret <> 0 Then thread_id = ret Else Err.Raise API_mysql_errno(myRec), "cMysql::thread_id", ptr2str(API_mysql_error(myRec))
End Function
Public Function thread_safe() As Long
thread_safe = API_mysql_thread_safe
End Function

'##############################################################################'
'right now i automatically use these to stORe AND convert into recORdsets
'  Public Sub stORe_result()
'  Public Sub use_result()
'##############################################################################'

'##############################################################################'

Private Function ConvertResult(m As API_MYSQL_RES) As Boolean

On Error Resume Next

Dim m_fieldcount As Long
Dim m_rowcount As Long
Dim i As Long
Dim j As Long
Dim ret As Long
Dim PickUp() As Long
    Dim test As Byte
    Dim sout() As Byte
    Dim cchars As Long
  m_fieldcount = myRec_res.field_count
  m_rowcount = convert642l(myRec_res.row_count)

       
          st = m_rowcount
         STB = m_fieldcount
         ReDim PickUp(1 To m_fieldcount)
         
         If GetColumn Then
                ReDim sCOL(1 To STB)
                 For i = 1 To m_fieldcount
                   ret = API_mysql_fetch_field(myRec_res)
                         If ret Then
                              CopyMemory myRec_field, ByVal ret, LenB(myRec_field)
                           sCOL(i) = ptr2str(myRec_field.Name)
                    End If
                 Next
          End If
          
          ReDim DAT(m_fieldcount, m_rowcount)
          
            For j = 1 To m_rowcount
            
               ret = API_mysql_fetch_row(myRec_res)
               
               If ret Then
                CopyMemory PickUp(1), ByVal ret, SIZE_OF_CHAR * m_fieldcount
                    For i = 1 To m_fieldcount
                                If PickUp(i) = 0 Then GoTo Nxti
                                cchars = lstrlen(PickUp(i))
                                If cchars = 0 Then GoTo Nxti
                                ReDim sout(1 To cchars)
                                sout = String(cchars, " ")
                                CopyMemory sout(1), ByVal (PickUp(i)), cchars
                                DAT(i, j) = StripNull(StrConv(sout, vbUnicode))
Nxti:
                                
                   Next i
               End If
            Next j
ConvertResult = True

ErrH:

End Function


Private Function convert642l(a As API_myulonglong) As Long
Dim res As Long
  CopyMemory res, a.bytes(1), 4
  convert642l = res
End Function

Private Function convertl264(l2convert As Long) As API_myulonglong
Dim a As API_myulonglong
  CopyMemory a.bytes(1), l2convert, 4
  convertl264 = a
End Function


Public Property Let user(ByVal sUserName As String)
    mUser = sUserName
End Property
Public Property Get user() As String
   user = mUser
End Property
Public Property Let PasswORd(ByVal sPasswORd As String)
   mPasswORd = sPasswORd
End Property
Public Property Get PasswORd() As String
  PasswORd = mPasswORd
End Property
Public Property Let host(ByVal sHost As String)
   mHost = sHost
End Property
Public Property Get host() As String
  host = mHost
End Property
Public Property Let Port(ByVal sPORt As Long)
  mPORt = sPORt
End Property
Public Property Get Port() As Long
  Port = mPORt
End Property
Public Property Let DB(ByVal sDB As String)
  mDb = sDB
End Property
Public Property Get DB() As String
  DB = mDb
End Property
Public Property Let ClientFlags(ByVal sClientFlags As Long)
 mClientFlag = sClientFlags
End Property
Public Property Get ClientFlags() As Long
 ClientFlags = mClientFlag
End Property
Public Property Let Unix_Socket(ByVal sUnix_socket As String)
  mUnix_socket = sUnix_socket
End Property
Public Property Get Unix_Socket() As String
  Unix_Socket = mUnix_socket
End Property

