VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form Grafikoni 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Chart"
   ClientHeight    =   7605
   ClientLeft      =   60
   ClientTop       =   930
   ClientWidth     =   11880
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   7605
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   7455
      Left            =   0
      OleObjectBlob   =   "Graf.frx":0000
      TabIndex        =   0
      ToolTipText     =   "To close Chart click CLOSE in menu Chart."
      Top             =   0
      Width           =   11895
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin VB.Menu grafik 
      Caption         =   "CHART"
      Begin VB.Menu sev 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu cop 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu dd 
         Caption         =   "-"
      End
      Begin VB.Menu prnt 
         Caption         =   "Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu fgt 
         Caption         =   "-"
      End
      Begin VB.Menu clos 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu izgle 
      Caption         =   "VISAGE"
      Begin VB.Menu linn 
         Caption         =   "Lines 2d"
      End
      Begin VB.Menu stubb 
         Caption         =   "Columnnade 2d"
      End
      Begin VB.Menu trak 
         Caption         =   "Ribbons 3d"
      End
      Begin VB.Menu stub 
         Caption         =   "Columnnade 3d"
      End
      Begin VB.Menu kruu 
         Caption         =   "Ring"
      End
   End
End
Attribute VB_Name = "Grafikoni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Tip As Integer 'Tip grafika
Option Explicit

Private Sub clos_Click()
Unload Grafikoni
Set Grafikoni = Nothing
Statistika.Smer = 0

End Sub
Private Sub cop_Click()
MSChart1.EditCopy
End Sub

Public Sub Form_Initialize()

Glavni.Toolbar1.Enabled = False

MSChart1.chartType = Tip
   
Dim zar As Integer, col As Integer, fes As String, ce As Integer
Dim Sif As Series
Dim serX As Series, a As Integer
Dim Nam As String, ar As Variant

Dim Sew(1 To 10) As String
Dim Coli(1 To 10) As Long

For a = 1 To 10
Sew(a) = Statistika.ListView1.ListItems(a)
If Sew(a) = "" Then Sew(a) = " "
Coli(a) = Statistika.ListView1.ListItems(a).SubItems(1)
If Coli(a) = 0 Then Coli(a) = 0
Next a

' Prve sreije sadrze labele po X osi.
   MSChart1.chartType = Tip
   Dim arrData(11, 1 To 2)
   arrData(1, 1) = 0
   arrData(2, 1) = Sew(1)
   arrData(3, 1) = Sew(2)
   arrData(4, 1) = Sew(3)
   arrData(5, 1) = Sew(4)
   arrData(6, 1) = Sew(5)
   arrData(7, 1) = Sew(6)
   arrData(8, 1) = Sew(7)
   arrData(9, 1) = Sew(8)
   arrData(10, 1) = Sew(9)
   arrData(11, 1) = Sew(10)
   
   arrData(1, 2) = "Sume"
   arrData(2, 2) = Coli(1)
   arrData(3, 2) = Coli(2)
   arrData(4, 2) = Coli(3)
   arrData(5, 2) = Coli(4)
   arrData(6, 2) = Coli(5)
   arrData(7, 2) = Coli(6)
   arrData(8, 2) = Coli(7)
   arrData(9, 2) = Coli(8)
   arrData(10, 2) = Coli(9)
   arrData(11, 2) = Coli(10)

   MSChart1.ChartData = arrData
   
   MSChart1.Title = "TOP LISTA za period od " & Statistika.Dy1 & " do " & Statistika.Dy2
    

End Sub



Private Sub Form_Unload(Cancel As Integer)
Glavni.Toolbar1.Enabled = True

End Sub

Private Sub kruu_Click()
Tip = 14
Form_Initialize
End Sub

Private Sub linn_Click()
Tip = 3
Form_Initialize

End Sub

Private Sub prnt_Click()
Me.PrintForm

End Sub

Private Sub sev_Click()
MSChart1.EditCopy

Load Snimanje
Snimanje.Show vbModal

End Sub

Private Sub stub_Click()
Tip = 0
Form_Initialize

End Sub

Private Sub stubb_Click()
Tip = 1
Form_Initialize

End Sub

Private Sub trak_Click()
Tip = 2
Form_Initialize

End Sub
