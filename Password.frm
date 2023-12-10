VERSION 5.00
Begin VB.Form Password 
   Caption         =   "Provera identiteta"
   ClientHeight    =   1140
   ClientLeft      =   7140
   ClientTop       =   2325
   ClientWidth     =   3705
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   1140
   ScaleWidth      =   3705
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Promena"
      Height          =   255
      Left            =   2640
      TabIndex        =   4
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      Height          =   255
      Left            =   2640
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   2640
      TabIndex        =   2
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   240
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Unesite sifru:"
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
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "Password"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Unload Password
Set Password = Nothing

End Sub

Private Sub Command2_Click()
Dim Brrr As Integer, Disk As String, Sif As String
Brrr = FreeFile
Disk = Left$(App.Path, 1)
Disk = Disk & ":\Windows\Sarma.box"

Open Disk For Input As #Brrr
Input #Brrr, Sif
Close #Brrr

    If Trim$(Text1.Text) <> Sif Then
        MsgBox "Pogresna sifra!", vbOKOnly, "Greska"
    Else
        Unload Password
        Set Password = Nothing
        If Glavni.Direkt = True Then
        Glavni.Direkt = False
        Load Direct: Direct.SetFocus
        Else
        Load Smak: Smak.Show vbModal
        End If
        
    End If

End Sub

Private Sub Command3_Click()
Load ChangPass
ChangPass.Show vbModal
End Sub

Private Sub Form_Activate()
Dim Brrr As Integer, Disk As String, Sif As String
On Error GoTo 22
Brrr = FreeFile
Disk = Left$(App.Path, 1)
Disk = Disk & ":\Windows\Sarma.box"

Open Disk For Input As #Brrr
Input #Brrr, Sif
Close #Brrr
Exit Sub

22:
    If Err = 53 Then
    
        Unload Password
        Set Password = Nothing
        Load ChangPass: ChangPass.Show vbModal
    End If

End Sub

