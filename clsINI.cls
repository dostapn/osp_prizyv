VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsINI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Класс для работы с INI-файлами"

Option Explicit


Private mFilename As String

Private AllSections() As tSection

Private Type tVAlue
    Name As String
    cntVAlue As String
End Type
Private Type tSection
    Name As String
    VAlues() As tVAlue
End Type
Public Function RemoveSection(ByVal Section As String) As Boolean
Attribute RemoveSection.VB_Description = "Удаляет секцию"
Dim i&, uSN$, a&
uSN = UCase$(Section)
' Просмотр всех секций
For i = 1 To UBound(AllSections)
    If UCase$(AllSections(i).Name) = uSN Then
        ' Нужная секция есть, сдвигаем параметры "вниз"
        For a = i To UBound(AllSections) - 1
            AllSections(a) = AllSections(a + 1)
        Next a
        ' Удаляем
        ReDim Preserve AllSections(UBound(AllSections) - 1)
        RemoveSection = True
        Exit Function
    End If
Next i
End Function
Public Property Get SectionExists(ByVal Section As String) As Boolean
Attribute SectionExists.VB_Description = "Возвращает True, если заданная секция уже существует"
Dim i&, uSN$
uSN = UCase$(Section)
' Просто пробежимся по вссем секциям и посмотрим, есть ли заданная
For i = 1 To UBound(AllSections)
    If UCase$(AllSections(i).Name) = uSN Then
        ' Есть!
        SectionExists = True
        Exit Property
    End If
Next i
End Property
Public Function UPDATEFile() As Boolean
Attribute UPDATEFile.VB_Description = "Сохраняет файл"
' Имя файла не задано - куда сохранять-то? :)
If Len(mFilename) = 0 Then Exit Function
Dim i&, a&
Open mFilename For Output Access Write Lock Write As #1
For i = 1 To UBound(AllSections)
    ' Запись всех секций
    With AllSections(i)
        Print #1, "[" & .Name & "]"
        ' Запись всех параметров в секции
        For a = 1 To UBound(.VAlues)
            Print #1, .VAlues(a).Name & "=" & .VAlues(a).cntVAlue
        Next a
    End With
    Print #1, vbNullString
Next i
Close #1
UPDATEFile = True
End Function
Public Property Get VAlueExists(ByVal Section As String, ByVal VAlueName As String) As Boolean
Attribute VAlueExists.VB_Description = "Возращает True, если заданный параметр в секции уже существует"
Dim i&, uSN$, uVN$, a&
uSN = UCase$(Section)
uVN = UCase$(VAlueName)
' Просмотр всех секций
For i = 1 To UBound(AllSections)
    If UCase$(AllSections(i).Name) = uSN Then
        ' Нужная секция, теперь ищем в ней параметр
        For a = 1 To UBound(AllSections(i).VAlues)
            If AllSections(i).VAlues(a).Name = uVN Then
                ' Искомый параметр
                VAlueExists = True
                Exit Property
            End If
        Next a
        Exit Property
    End If
Next i
End Property
Public Property Get VAlueName(ByVal Section As String, ByVal Index As Long) As String
Attribute VAlueName.VB_Description = "Возвращает имя сторки по ее номеру"
' Примечание: если параметр или секция найдены не будут, то свойство
' вернет vbNullChar - символ с нулевым кодом
Dim uSN$, i&
uSN = UCase$(Section)
' Просмотр всех секций
For i = 1 To UBound(AllSections)
    If UCase$(AllSections(i).Name) = uSN Then
    
        If Index > 0 And Index <= UBound(AllSections(i).VAlues) Then
            ' Нужный параметр
            VAlueName = AllSections(i).VAlues(Index).Name
            Exit Property
        Else
            ' Индекс вне границ диапазона...
            VAlueName = vbNullString
        End If
        Exit Property
    End If
Next i
VAlueName = vbNullString
End Property
Public Function RemoveVAlue(ByVal Section As String, ByVal VAlueName As String) As Boolean
Attribute RemoveVAlue.VB_Description = "Удаляет значение из секции"
Dim i&, a&, uVN$, uSN$, c&
uSN = UCase$(Section)
uVN = UCase$(VAlueName)
' Просмор всех секций
For i = 1 To UBound(AllSections)
    If UCase$(AllSections(i).Name) = uSN Then
        ' Нужная секция есть
        For a = 1 To UBound(AllSections(i).VAlues)
            If UCase$(AllSections(i).VAlues(a).Name) = uVN Then
                ' Нужный параметр есть
                With AllSections(i)
                    ' Сдвиг параметров "вниз" на один
                    For c = a To UBound(.VAlues) - 1
                        .VAlues(a) = .VAlues(a + 1)
                    Next c
                    ' Удаление параметра
                    ReDim Preserve .VAlues(UBound(.VAlues) - 1)
                    RemoveVAlue = True
                    Exit Function
                End With
            End If
        Next
        Exit Function
    End If
Next i
End Function
Public Function AddVAlue(ByVal Section As String, ByVal VAlueName As String, Optional ByVal lVAlue As String = vbNullString) As Boolean
Attribute AddVAlue.VB_Description = "Добавляет строку (значение) в секцию"
Dim uSN$, i&, uVN$, a&
uSN = UCase$(Section)
uVN = UCase$(VAlueName)
' Просмотрим-ка все секции...
For i = 1 To UBound(AllSections)
    If UCase$(AllSections(i).Name) = uSN Then
add_VAlue:
        ' Есть нужная секция!
        For a = 1 To UBound(AllSections(i).VAlues)
            If UCase$(AllSections(i).VAlues(a).Name) = uVN Then
                ' Такая строка уже есть
                Exit Function
            End If
        Next a
        ' Добавляем новый параметр
        With AllSections(i)
            ReDim Preserve .VAlues(UBound(.VAlues) + 1)
            .VAlues(UBound(.VAlues)).Name = VAlueName
            .VAlues(UBound(.VAlues)).cntVAlue = lVAlue
        End With
        AddVAlue = True
        Exit Function
    End If
Next i
' Секции такой нет, добавляем...
AddSection Section
GoTo add_VAlue
End Function
Public Property Get VAluesCount(ByVal Section As String) As Long
Attribute VAluesCount.VB_Description = "Возвращает количество строк в секции"
' Примечание: если секция не найдена, свойство вернет
' значение -1.
Dim uSN$, i&
uSN = UCase$(Section)
' Просмотр всех секций
For i = 1 To UBound(AllSections)
    If UCase$(AllSections(i).Name) = uSN Then
        ' Нужная секция
        VAluesCount = UBound(AllSections(i).VAlues)
        Exit Property
    End If
Next i
VAluesCount = -1
End Property
Public Property Let VAlue(ByVal Section As String, ByVal VAlueName As String, ByVal vData As String)
Dim i&, a&, uVN$, uSN$
uSN = UCase$(Section)
uVN = UCase$(VAlueName)
' Просмотр всех секций
For i = 1 To UBound(AllSections)
    If UCase$(AllSections(i).Name) = uSN Then
        ' Просмотр параметров в секции
        For a = 1 To UBound(AllSections(i).VAlues)
            If UCase$(AllSections(i).VAlues(a).Name) = uVN Then
                ' Нашли параметр, обновляем его значение...
                AllSections(i).VAlues(a).cntVAlue = vData
                Exit Property
            End If
        Next
        ' Не нашли параметр, добавляем новый...
        With AllSections(i)
            ReDim Preserve .VAlues(UBound(.VAlues) + 1)
            .VAlues(UBound(.VAlues)).cntVAlue = vData
            .VAlues(UBound(.VAlues)).Name = VAlueName
        End With
        Exit Property
    End If
Next i
' Не нашли соотвествующую секцию, добавляем новую...
ReDim Preserve AllSections(UBound(AllSections) + 1)
With AllSections(UBound(AllSections))
    .Name = Section
    ReDim .VAlues(1)
    .VAlues(1).Name = VAlueName
    .VAlues(1).cntVAlue = vData
End With
End Property
Public Property Get VAlue(ByVal Section As String, ByVal VAlueName As String) As String
Attribute VAlue.VB_Description = "Устанавливает/возвращает значение строки"
Attribute VAlue.VB_UserMemId = 0
' Примечание: если параметр найден не будет, то свойство
' вернет vbNullChar - символ с нулевым кодом
Dim i&, a&, uVN$, uSN$
uSN = UCase$(Section)
uVN = UCase$(VAlueName)
' Просмотр всех секций
For i = 1 To UBound(AllSections)
    If UCase$(AllSections(i).Name) = uSN Then
        ' Просмотр всех параметров в секции
        For a = 1 To UBound(AllSections(i).VAlues)
            If UCase$(AllSections(i).VAlues(a).Name) = uVN Then
                ' Нашли нужный параметр
                VAlue = AllSections(i).VAlues(a).cntVAlue
                Exit Property
            End If
        Next
        VAlue = vbNullString
        Exit Property
    End If
Next i
VAlue = vbNullString
End Property
Public Property Get SectionName(ByVal Index As Long) As String
Attribute SectionName.VB_Description = "Возвращает имя секции по ее номеру"
' Здесь все просто: проверка - попадает ли индекс в границы
' массива и возврат имени секции.
' Примечание: если секция найдена не будет, то свойство
' вернет vbNullChar - символ с нулевым кодом
If Index > 0 And Index <= UBound(AllSections) Then
    SectionName = AllSections(Index).Name
Else
    SectionName = vbNullString
End If
End Property
Public Function AddSection(ByVal Section As String) As Boolean
Attribute AddSection.VB_Description = "Добавляет секцию в файл"
Dim uSN$, i&
uSN = UCase$(Section)
For i = 1 To UBound(AllSections)
    If UCase$(AllSections(i).Name) = uSN Then
        ' Уже такая секция есть...
        Exit Function
    End If
Next i
' Добавляем элемент к массиву секций
ReDim Preserve AllSections(UBound(AllSections) + 1)
With AllSections(UBound(AllSections))
    .Name = Section
    ReDim .VAlues(0)
End With
AddSection = True
End Function
Public Property Get SectionsCount() As Long
Attribute SectionsCount.VB_Description = "Возвращает количество секций в файле"
SectionsCount = UBound(AllSections)
End Property
Public Property Let FileName(ByVal vData As String)
Attribute FileName.VB_Description = "Имя INI-файла"
Dim s&, v&, l$, t&
mFilename = vData
ReDim AllSections(0)
' Файл в наличии?
If Len(dir(mFilename)) = 0 Then
    ' Если нет, то...
    Exit Property
End If
' Считываем файл
Open mFilename For Input Access Read Lock Write As #1
Do While Not eof(1)
    Line Input #1, l
    ' Комментарии есть?
    t = InStr(1, l, ";")
    l = Trim(l)
    ' Откомментирована вся строка
    If t = 1 Or Len(l) = 0 Then GoTo skip_line
    ' Комментарии в конце строки
    If t > 1 Then l = Left$(l, t - 1)
    If Left$(l, 1) = "[" Then
        ' Нашли секцию
        t = InStr(1, l, "]")
        If t < 3 Then GoTo skip_line
        l = Mid$(l, 2, t - 2)
        ' Секций стало на одну больше
        s = s + 1
        ' Строк в новой секции пока нет
        v = 0
        ReDim Preserve AllSections(s)
        AllSections(s).Name = l
        ReDim AllSections(s).VAlues(0)
    Else
        ' Нашли параметр в секцию
        ' Если ранеее не было найдено ни одной секции - пропускаем
        If s = 0 Then GoTo skip_line
        ' Разделяем строку на название параметра и на его значение
        t = InStr(1, l, "=")
        If t < 2 Then GoTo skip_line
        v = v + 1
        ' Добавляем параметр
        With AllSections(s)
            ReDim Preserve .VAlues(v)
            .VAlues(v).Name = Left$(l, t - 1)
            If t < Len(l) Then .VAlues(v).cntVAlue = Mid$(l, t + 1, Len(l) - t)
        End With
    End If
skip_line:
Loop
Close #1
End Property
Public Property Get FileName() As String
FileName = mFilename
End Property
Private Sub Class_Initialize()
ReDim AllSections(0)
End Sub
Private Sub Class_Terminate()
ReDim AllSections(0)
End Sub
