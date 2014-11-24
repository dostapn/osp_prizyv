VERSION 5.00
Begin VB.Form frmPrintListsVK 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Списки отправленных"
   ClientHeight    =   1110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3270
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPrintListsVK.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1110
   ScaleWidth      =   3270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox lstvk 
      Height          =   315
      Left            =   360
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Генерировать"
      Height          =   345
      Left            =   885
      TabIndex        =   0
      Top             =   600
      Width           =   1515
   End
End
Attribute VB_Name = "frmPrintListsVK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim x As Long

Private Sub cmdGenerate_Click()
   
    Call Cnv.To_ListsVK(lstvk.Text)
    
End Sub

Private Sub Form_Load()
Call Reg_VK_List
    lstvk.AddItem ("Все")
For x = 0 To UBound(nVK())
    lstvk.AddItem (nVK(x))
    Next x
End Sub

