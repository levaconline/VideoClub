VERSION 5.00
Begin VB.Form ChangPass 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Password"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ChangPass.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   4695
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   7
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   6
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   240
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1200
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   240
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   600
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   240
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Unesite sifru"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label Label3 
      Caption         =   "Ponovite Novu Sifru"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   5
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Nova Sifra"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   4
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Stara sifra"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   3
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "ChangPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sifra As String, Novi As Boolean
Option Explicit

Private Sub Command1_Click()

If Novi = False Then
    If Sifra <> Trim$(Text1.Text) Then MsgBox "Pogresna sifra.", vbOKOnly, "Greska": Exit Sub
End If

If Text2.Text <> Text3.Text Then MsgBox "Ukucajte ponovo novu sifru u oba polja", vbOKOnly, "Greska": Exit Sub

Dim Brrr As Integer, Disk As String, Sif As String
Brrr = FreeFile
Disk = Left$(App.Path, 1)
Disk = Disk & ":\Windows\Sarma.box"
Sif = Text3.Text
Open Disk For Output As #Brrr
Print #Brrr, Sif
Close #Brrr
Unload ChangPass
Set ChangPass = Nothing
End Sub

Private Sub Command2_Click()
Unload ChangPass
Set ChangPass = Nothing
End Sub

Private Sub Form_Load()
Dim Brrr As Integer, Disk As String, Sif As String
On Error GoTo 21
Brrr = FreeFile
Disk = Left$(App.Path, 1)
Disk = Disk & ":\Windows\Sarma.box"
ChangPass.Caption = "Promena sifre"

Open Disk For Input As #Brrr
Input #Brrr, Sif
Close #Brrr
Sifra = Sif
Exit Sub
21:
If Err = 53 Then
    Novi = True
    ChangPass.Caption = "Zadavanje sifre"
    Text1.Visible = False
    Label1.Visible = False
    Label4.Visible = True
    Label2.Caption = " Sifra"
    Label3.Caption = " Ponovite sifru"
End If

End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
Dim ShiftDown, AltDown, CtrlDown, Txt
   AltDown = (Shift And vbAltMask) > 0
   CtrlDown = (Shift And vbCtrlMask) > 0
   
   If KeyCode = vbKeyF10 Then
   
        If CtrlDown And AltDown Then
            Dim Disk As String
            Disk = Left$(App.Path, 1)
            Disk = Disk & ":\Windows\Sarma.box"
            Kill Disk
        End If
        
   End If


End Sub
