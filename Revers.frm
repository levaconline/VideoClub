VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Revers 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Video Club"
   ClientHeight    =   4500
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   3255
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Revers.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   3255
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView ListView1 
      Height          =   1935
      Left            =   0
      TabIndex        =   13
      Top             =   2520
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   3413
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Title"
         Object.Width           =   2927
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Amount"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Dat1"
         Object.Width           =   35
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Dat2"
         Object.Width           =   35
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Period"
         Object.Width           =   1234
      EndProperty
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   1920
      ItemData        =   "Revers.frx":0442
      Left            =   0
      List            =   "Revers.frx":0444
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   2520
      Width           =   3255
   End
   Begin VB.Shape Shape1 
      Height          =   4455
      Left            =   0
      Top             =   0
      Width           =   3255
   End
   Begin VB.Line Line5 
      X1              =   0
      X2              =   3240
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line4 
      X1              =   0
      X2              =   3240
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   3240
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   3240
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   3240
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Payment:"
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label11 
      BackColor       =   &H00E0E0E0&
      Caption         =   " X  cena"
      Height          =   255
      Left            =   1800
      TabIndex        =   11
      ToolTipText     =   "Cena rentiranja po danu"
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label10 
      BackColor       =   &H00E0E0E0&
      Caption         =   " Br. dana: "
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Balance:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   2040
      Width           =   3255
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Amount:"
      Height          =   255
      Left            =   1920
      TabIndex        =   8
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E0E0E0&
      Caption         =   " Phone:"
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   600
      Width           =   3255
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Address:"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   360
      Width           =   3255
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   3255
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   " Date:"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      ToolTipText     =   "Trenutni datum"
      Top             =   840
      Width           =   3255
   End
   Begin VB.Label Label2 
      BackColor       =   &H00808080&
      Caption         =   " Rent:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   2280
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   " Member: "
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   3255
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   " ID:"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   1320
      Width           =   3255
   End
   Begin VB.Menu opp 
      Caption         =   "Options"
      Begin VB.Menu rre 
         Caption         =   "&Print"
      End
      Begin VB.Menu wgr 
         Caption         =   "-"
      End
      Begin VB.Menu hhoo 
         Caption         =   "OK"
      End
   End
End
Attribute VB_Name = "Revers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Initialize()
Dim Boopp As Integer, Cezar As String, Duz As Integer, Ruz As Integer, Dfgtty As String
Label3.Caption = " Date: " & Now
Label1 = " Member: " & Pocetna.adoClanovi.Recordset.Fields(1) & " " & Pocetna.adoClanovi.Recordset.Fields(3)
Label2 = " " & Pocetna.Rev
Label4.Caption = " Phone: " & Glavni.TelClu
Label5.Caption = "Video Club " & Glavni.NamClu
Label6.Caption = " Address: " & Glavni.AdsClu
Label7.Caption = " ID: " & Pocetna.adoClanovi.Recordset.Fields(0)
Label8.Caption = " Amount: " & Pocetna.Text3.Text
List1.Clear

If Pocetna.Rev = "Rent:" Then    ' Ako se izdaje film
ListView1.Visible = False
Label9.Caption = ""
Label10.Caption = ""
Label11.Caption = ""
Label12.Caption = ""

Else                              ' Ako se film vraca
ListView1.Visible = True
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Revers
Set Revers = Nothing

End Sub

Private Sub hhoo_Click()
Unload Revers
Set Revers = Nothing

End Sub

Private Sub rre_Click()
Revers.PrintForm

End Sub
