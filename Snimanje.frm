VERSION 5.00
Begin VB.Form Snimanje 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Save"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6345
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   6345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      BackColor       =   &H008080FF&
      Caption         =   "Snimi Folder"
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   960
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Novi Folder"
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   960
      Width           =   1095
   End
   Begin VB.DirListBox Dir1 
      Height          =   1665
      Left            =   1560
      TabIndex        =   6
      Top             =   240
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   3960
      TabIndex        =   4
      Text            =   "Grafik1"
      Top             =   2040
      Width           =   1815
   End
   Begin VB.FileListBox File1 
      Height          =   1650
      Left            =   3960
      TabIndex        =   3
      Top             =   240
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SAVE"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1920
      Width           =   1095
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Height          =   1575
      Left            =   120
      TabIndex        =   11
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   ".bmp"
      Height          =   375
      Left            =   5760
      TabIndex        =   10
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "Fajlovi:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   9
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Folderi:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   8
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Diskovi:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Unesite ime fajla:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   5
      Top             =   2040
      Width           =   2295
   End
End
Attribute VB_Name = "Snimanje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
If Trim$(Text1.Text) <> "" Then
Dim ImeSlike As String, gwq As Integer
For gwq = 0 To File1.ListCount
If File1.List(gwq) = Trim$(Text1.Text) & ".bmp" Then
If MsgBox("Fajl sa istim imenom vec postoji. Zelite li da ga presnimite?", vbYesNo + vbExclamation, "UPOZORENJE!") = vbNo Then GoTo 323
End If
Next gwq
ImeSlike = Dir1.Path & "\" & Trim$(Text1.Text) & ".bmp"
SavePicture Clipboard.GetData, ImeSlike
MsgBox "Grafik je snimljen u:" & ImeSlike
Unload Snimanje
Set Snimanje = Nothing
323:
Else
MsgBox "Niste upisali naziv fajla."
End If
End Sub

Private Sub Command2_Click()
Unload Snimanje
Set Snimanje = Nothing
End Sub

Private Sub Command3_Click()
Text1.Text = " "
Command1.Enabled = False
Label1.Caption = "Unesite ime foldera:"
Label5.Caption = ""
Command4.Visible = True

End Sub

Private Sub Command4_Click()
Dim wee As Integer
If Trim$(Text1.Text) <> "" Then
For wee = 0 To Dir1.ListCount
If Trim$(Text1.Text) = Dir1.List(wee) Then MsgBox "Folder vec postoji. Odaberite drugo ime.": Exit Sub
Next wee
MkDir Dir1.Path & "\" & Trim$(Text1.Text)
Else
MsgBox "Niste napisali ime foldera."
End If
Command4.Visible = False
Command1.Enabled = True
Text1.Text = "Grafik1"
Label5.Caption = ".bmp"
Label1.Caption = "Unesite ime fajla:"
Dir1.Refresh
End Sub

Private Sub Drive1_Change()
   Dir1.Path = Drive1.Drive
End Sub
Private Sub Dir1_Change()
   File1.Path = Dir1.Path
End Sub


