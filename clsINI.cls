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
Attribute VB_Description = "����� ��� ������ � INI-�������"

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
Attribute RemoveSection.VB_Description = "������� ������"
Dim i&, uSN$, a&
uSN = UCase$(Section)
' �������� ���� ������
For i = 1 To UBound(AllSections)
    If UCase$(AllSections(i).Name) = uSN Then
        ' ������ ������ ����, �������� ��������� "����"
        For a = i To UBound(AllSections) - 1
            AllSections(a) = AllSections(a + 1)
        Next a
        ' �������
        ReDim Preserve AllSections(UBound(AllSections) - 1)
        RemoveSection = True
        Exit Function
    End If
Next i
End Function
Public Property Get SectionExists(ByVal Section As String) As Boolean
Attribute SectionExists.VB_Description = "���������� True, ���� �������� ������ ��� ����������"
Dim i&, uSN$
uSN = UCase$(Section)
' ������ ���������� �� ����� ������� � ���������, ���� �� ��������
For i = 1 To UBound(AllSections)
    If UCase$(AllSections(i).Name) = uSN Then
        ' ����!
        SectionExists = True
        Exit Property
    End If
Next i
End Property
Public Function UPDATEFile() As Boolean
Attribute UPDATEFile.VB_Description = "��������� ����"
' ��� ����� �� ������ - ���� ���������-��? :)
If Len(mFilename) = 0 Then Exit Function
Dim i&, a&
Open mFilename For Output Access Write Lock Write As #1
For i = 1 To UBound(AllSections)
    ' ������ ���� ������
    With AllSections(i)
        Print #1, "[" & .Name & "]"
        ' ������ ���� ���������� � ������
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
Attribute VAlueExists.VB_Description = "��������� True, ���� �������� �������� � ������ ��� ����������"
Dim i&, uSN$, uVN$, a&
uSN = UCase$(Section)
uVN = UCase$(VAlueName)
' �������� ���� ������
For i = 1 To UBound(AllSections)
    If UCase$(AllSections(i).Name) = uSN Then
        ' ������ ������, ������ ���� � ��� ��������
        For a = 1 To UBound(AllSections(i).VAlues)
            If AllSections(i).VAlues(a).Name = uVN Then
                ' ������� ��������
                VAlueExists = True
                Exit Property
            End If
        Next a
        Exit Property
    End If
Next i
End Property
Public Property Get VAlueName(ByVal Section As String, ByVal Index As Long) As String
Attribute VAlueName.VB_Description = "���������� ��� ������ �� �� ������"
' ����������: ���� �������� ��� ������ ������� �� �����, �� ��������
' ������ vbNullChar - ������ � ������� �����
Dim uSN$, i&
uSN = UCase$(Section)
' �������� ���� ������
For i = 1 To UBound(AllSections)
    If UCase$(AllSections(i).Name) = uSN Then
    
        If Index > 0 And Index <= UBound(AllSections(i).VAlues) Then
            ' ������ ��������
            VAlueName = AllSections(i).VAlues(Index).Name
            Exit Property
        Else
            ' ������ ��� ������ ���������...
            VAlueName = vbNullString
        End If
        Exit Property
    End If
Next i
VAlueName = vbNullString
End Property
Public Function RemoveVAlue(ByVal Section As String, ByVal VAlueName As String) As Boolean
Attribute RemoveVAlue.VB_Description = "������� �������� �� ������"
Dim i&, a&, uVN$, uSN$, c&
uSN = UCase$(Section)
uVN = UCase$(VAlueName)
' ������� ���� ������
For i = 1 To UBound(AllSections)
    If UCase$(AllSections(i).Name) = uSN Then
        ' ������ ������ ����
        For a = 1 To UBound(AllSections(i).VAlues)
            If UCase$(AllSections(i).VAlues(a).Name) = uVN Then
                ' ������ �������� ����
                With AllSections(i)
                    ' ����� ���������� "����" �� ����
                    For c = a To UBound(.VAlues) - 1
                        .VAlues(a) = .VAlues(a + 1)
                    Next c
                    ' �������� ���������
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
Attribute AddVAlue.VB_Description = "��������� ������ (��������) � ������"
Dim uSN$, i&, uVN$, a&
uSN = UCase$(Section)
uVN = UCase$(VAlueName)
' ����������-�� ��� ������...
For i = 1 To UBound(AllSections)
    If UCase$(AllSections(i).Name) = uSN Then
add_VAlue:
        ' ���� ������ ������!
        For a = 1 To UBound(AllSections(i).VAlues)
            If UCase$(AllSections(i).VAlues(a).Name) = uVN Then
                ' ����� ������ ��� ����
                Exit Function
            End If
        Next a
        ' ��������� ����� ��������
        With AllSections(i)
            ReDim Preserve .VAlues(UBound(.VAlues) + 1)
            .VAlues(UBound(.VAlues)).Name = VAlueName
            .VAlues(UBound(.VAlues)).cntVAlue = lVAlue
        End With
        AddVAlue = True
        Exit Function
    End If
Next i
' ������ ����� ���, ���������...
AddSection Section
GoTo add_VAlue
End Function
Public Property Get VAluesCount(ByVal Section As String) As Long
Attribute VAluesCount.VB_Description = "���������� ���������� ����� � ������"
' ����������: ���� ������ �� �������, �������� ������
' �������� -1.
Dim uSN$, i&
uSN = UCase$(Section)
' �������� ���� ������
For i = 1 To UBound(AllSections)
    If UCase$(AllSections(i).Name) = uSN Then
        ' ������ ������
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
' �������� ���� ������
For i = 1 To UBound(AllSections)
    If UCase$(AllSections(i).Name) = uSN Then
        ' �������� ���������� � ������
        For a = 1 To UBound(AllSections(i).VAlues)
            If UCase$(AllSections(i).VAlues(a).Name) = uVN Then
                ' ����� ��������, ��������� ��� ��������...
                AllSections(i).VAlues(a).cntVAlue = vData
                Exit Property
            End If
        Next
        ' �� ����� ��������, ��������� �����...
        With AllSections(i)
            ReDim Preserve .VAlues(UBound(.VAlues) + 1)
            .VAlues(UBound(.VAlues)).cntVAlue = vData
            .VAlues(UBound(.VAlues)).Name = VAlueName
        End With
        Exit Property
    End If
Next i
' �� ����� �������������� ������, ��������� �����...
ReDim Preserve AllSections(UBound(AllSections) + 1)
With AllSections(UBound(AllSections))
    .Name = Section
    ReDim .VAlues(1)
    .VAlues(1).Name = VAlueName
    .VAlues(1).cntVAlue = vData
End With
End Property
Public Property Get VAlue(ByVal Section As String, ByVal VAlueName As String) As String
Attribute VAlue.VB_Description = "�������������/���������� �������� ������"
Attribute VAlue.VB_UserMemId = 0
' ����������: ���� �������� ������ �� �����, �� ��������
' ������ vbNullChar - ������ � ������� �����
Dim i&, a&, uVN$, uSN$
uSN = UCase$(Section)
uVN = UCase$(VAlueName)
' �������� ���� ������
For i = 1 To UBound(AllSections)
    If UCase$(AllSections(i).Name) = uSN Then
        ' �������� ���� ���������� � ������
        For a = 1 To UBound(AllSections(i).VAlues)
            If UCase$(AllSections(i).VAlues(a).Name) = uVN Then
                ' ����� ������ ��������
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
Attribute SectionName.VB_Description = "���������� ��� ������ �� �� ������"
' ����� ��� ������: �������� - �������� �� ������ � �������
' ������� � ������� ����� ������.
' ����������: ���� ������ ������� �� �����, �� ��������
' ������ vbNullChar - ������ � ������� �����
If Index > 0 And Index <= UBound(AllSections) Then
    SectionName = AllSections(Index).Name
Else
    SectionName = vbNullString
End If
End Property
Public Function AddSection(ByVal Section As String) As Boolean
Attribute AddSection.VB_Description = "��������� ������ � ����"
Dim uSN$, i&
uSN = UCase$(Section)
For i = 1 To UBound(AllSections)
    If UCase$(AllSections(i).Name) = uSN Then
        ' ��� ����� ������ ����...
        Exit Function
    End If
Next i
' ��������� ������� � ������� ������
ReDim Preserve AllSections(UBound(AllSections) + 1)
With AllSections(UBound(AllSections))
    .Name = Section
    ReDim .VAlues(0)
End With
AddSection = True
End Function
Public Property Get SectionsCount() As Long
Attribute SectionsCount.VB_Description = "���������� ���������� ������ � �����"
SectionsCount = UBound(AllSections)
End Property
Public Property Let FileName(ByVal vData As String)
Attribute FileName.VB_Description = "��� INI-�����"
Dim s&, v&, l$, t&
mFilename = vData
ReDim AllSections(0)
' ���� � �������?
If Len(dir(mFilename)) = 0 Then
    ' ���� ���, ��...
    Exit Property
End If
' ��������� ����
Open mFilename For Input Access Read Lock Write As #1
Do While Not eof(1)
    Line Input #1, l
    ' ����������� ����?
    t = InStr(1, l, ";")
    l = Trim(l)
    ' ���������������� ��� ������
    If t = 1 Or Len(l) = 0 Then GoTo skip_line
    ' ����������� � ����� ������
    If t > 1 Then l = Left$(l, t - 1)
    If Left$(l, 1) = "[" Then
        ' ����� ������
        t = InStr(1, l, "]")
        If t < 3 Then GoTo skip_line
        l = Mid$(l, 2, t - 2)
        ' ������ ����� �� ���� ������
        s = s + 1
        ' ����� � ����� ������ ���� ���
        v = 0
        ReDim Preserve AllSections(s)
        AllSections(s).Name = l
        ReDim AllSections(s).VAlues(0)
    Else
        ' ����� �������� � ������
        ' ���� ������ �� ���� ������� �� ����� ������ - ����������
        If s = 0 Then GoTo skip_line
        ' ��������� ������ �� �������� ��������� � �� ��� ��������
        t = InStr(1, l, "=")
        If t < 2 Then GoTo skip_line
        v = v + 1
        ' ��������� ��������
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
