VERSION 5.00
Begin VB.Form Uvoz 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Inport files"
   ClientHeight    =   4650
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7005
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   7005
   ShowInTaskbar   =   0   'False
   Begin VB.FileListBox File1 
      Height          =   3210
      Left            =   2640
      TabIndex        =   5
      Top             =   120
      Width           =   4215
   End
   Begin VB.DirListBox Dir1 
      Height          =   3870
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   2415
   End
   Begin VB.DriveListBox Drive1 
      Height          =   360
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "Inport"
      Height          =   375
      Left            =   3960
      TabIndex        =   0
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   3600
      Width           =   4215
   End
End
Attribute VB_Name = "Uvoz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
Unload Uvoz
Set Uvoz = Nothing
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
On Error Resume Next
Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
Label1.Caption = File1.FileName
End Sub

Private Sub File1_DblClick()
OKButton_Click
End Sub

Public Sub OKButton_Click()
On Error Resume Next
Label1.Caption = File1.FileName

Dim Tres As String, Bes As String
Bes = Right$(Label1.Caption, 3)
If Bes = "peg" Then Bes = "mpeg"
Tres = Filmovi.IDbr
Kill App.Path & "\Clips\" & Tres & "." & "*"


If Bes = "jpg" Or Bes = "jpe" Or Bes = "bmp" Or Bes = "gif" Then
If Right$(File1.Path, 1) <> "\" Then
FileCopy File1.Path & "\" & File1.FileName, App.Path & "\Pictures\" & Tres & "." & Bes
Else
FileCopy File1.Path & File1.FileName, App.Path & "\Pictures\" & Tres & "." & Bes
End If
Else
FileCopy File1.Path & "\" & File1.FileName, App.Path & "\Clips\" & Tres & "." & Bes
End If


Unload Uvoz
Set Uvoz = Nothing
End Sub
