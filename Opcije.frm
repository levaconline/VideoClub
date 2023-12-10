VERSION 5.00
Begin VB.Form Opcije 
   Caption         =   " Setings"
   ClientHeight    =   3045
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5145
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
   Icon            =   "Opcije.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3045
   ScaleWidth      =   5145
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   " About Club "
      Height          =   2415
      Left            =   2400
      TabIndex        =   11
      Top             =   120
      Width           =   2655
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   120
         TabIndex        =   23
         Top             =   2040
         Width           =   1695
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   120
         TabIndex        =   21
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   120
         TabIndex        =   19
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   120
         TabIndex        =   17
         ToolTipText     =   "Telephone number in the Video Club"
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   120
         TabIndex        =   15
         ToolTipText     =   "Video Club Address"
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   14
         ToolTipText     =   "Video Club Name"
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label9 
         Caption         =   "E-mail"
         Height          =   255
         Left            =   1920
         TabIndex        =   24
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "Web"
         Height          =   255
         Left            =   1920
         TabIndex        =   22
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "City"
         Height          =   255
         Left            =   1920
         TabIndex        =   20
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Phone"
         Height          =   255
         Left            =   1920
         TabIndex        =   18
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "Address"
         Height          =   255
         Left            =   1920
         TabIndex        =   16
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Name"
         Height          =   255
         Left            =   1920
         TabIndex        =   13
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Rent Value"
      Height          =   1095
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   2175
      Begin VB.CheckBox Check1 
         Caption         =   "Include Sunday"
         Height          =   255
         Left            =   360
         TabIndex        =   12
         ToolTipText     =   "Include / exclude Sunday in calculations."
         Top             =   720
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   360
         TabIndex        =   9
         ToolTipText     =   "Rent Price"
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1440
         TabIndex        =   10
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Membership Period"
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   2175
      Begin VB.OptionButton Option2 
         Caption         =   "Option2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Value           =   -1  'True
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   255
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   360
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   " Forever"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " months"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1320
         TabIndex        =   6
         Top             =   360
         Width           =   735
      End
   End
End
Attribute VB_Name = "Opcije"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Command1_Click()
Unload Opcije
Set Opcije = Nothing

End Sub

Private Sub Command2_Click()
Dim Fio As Integer

If Val(Text3.Text) = 0 Then MsgBox "How match is rent price?": Exit Sub
If (Option1.Value = True) And (Val(Text2.Text) = 0) Then MsgBox "Membership period is not valid.": Exit Sub
If Text1.Text = "" Or Text1.Text = " " Then MsgBox "What's Name of Your Video Club?": Exit Sub
Glavni.RV = Text3.Text
If Check1.Value = 1 Then
Glavni.InSund = True
Else
Glavni.InSund = False
End If
Glavni.MemPer = Val(Text2.Text)
Glavni.NamClu = Text1.Text
Glavni.AdsClu = Text4.Text
Glavni.TelClu = Text5.Text
Glavni.Grad = Text6.Text
Glavni.WWW = Text7.Text
Glavni.Mail = Text8.Text
Fio = FreeFile
If Right$(App.Path, 1) <> "\" Then
Open App.Path & "\Info.aca" For Output As #Fio
Else
Open App.Path & "Info.aca" For Output As #Fio
End If

Write #Fio, Glavni.RV, Glavni.InSund, Glavni.MemPer, Glavni.NamClu, Glavni.AdsClu, Glavni.TelClu, Glavni.Grad, Glavni.WWW, Glavni.Mail
Close #Fio
Unload Opcije
Set Opcije = Nothing

End Sub

Private Sub Form_Load()
 Text3.Text = Glavni.RV
If Glavni.InSund = True Then
Check1.Value = 1
Else
Check1.Value = 0
End If
If Glavni.MemPer > 0 Then
Text2.Text = Glavni.MemPer
Option1.Value = True
Else
Text2.Text = ""
Text2.Enabled = False
End If
Text1.Text = Glavni.NamClu
Text4.Text = Glavni.AdsClu
Text5.Text = Glavni.TelClu
Text6.Text = Glavni.Grad
Text7.Text = Glavni.WWW
Text8.Text = Glavni.Mail
Label4.Caption = " " & Chr$(163) & "/day"
End Sub

Private Sub Option1_Click()
Text2.Enabled = True
Label2.Enabled = True
Label3.Enabled = False
End Sub

Private Sub Option2_Click()
Text2.Text = ""
Text2.Enabled = False
Label2.Enabled = False
Label3.Enabled = True
End Sub
