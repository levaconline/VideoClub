VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Filmovi 
   BackColor       =   &H00C0C0C0&
   Caption         =   " Movies"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Filmovi.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command12 
      Caption         =   "Inport"
      Height          =   255
      Left            =   4440
      TabIndex        =   39
      Top             =   5520
      Width           =   855
   End
   Begin MSComctlLib.ListView LiV52 
      Height          =   495
      Left            =   360
      TabIndex        =   38
      ToolTipText     =   "Double click to close"
      Top             =   6960
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   49152
      BackColor       =   0
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Member"
         Object.Width           =   5116
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ID"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "RentDate"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Amount "
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "AmountDays"
         Object.Width           =   2893
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "RevertDate"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00E0E0E0&
      Caption         =   "+"
      Height          =   255
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "On rent"
      Top             =   6240
      Width           =   255
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5400
      Picture         =   "Filmovi.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Play"
      Top             =   5520
      Width           =   495
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Search"
      Height          =   255
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Upisite naslov koji trzite"
      Top             =   1920
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Word"
      Height          =   255
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   6480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   32
      Text            =   "Filmovi.frx":0C14
      ToolTipText     =   "Mesto za opis filma"
      Top             =   5280
      Width           =   5055
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Print"
      Height          =   255
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   840
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Clip?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   4560
      TabIndex        =   30
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00E0E0E0&
      Caption         =   "+"
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
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Edit Liste zanrovq"
      Top             =   4080
      Width           =   255
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Picture "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   3975
      Left            =   6480
      TabIndex        =   27
      Top             =   1200
      Width           =   5055
      Begin VB.OLE OLE1 
         AutoActivate    =   1  'GetFocus
         Class           =   "mplayer"
         Height          =   3615
         Left            =   120
         SizeMode        =   2  'AutoSize
         TabIndex        =   36
         Top             =   240
         Visible         =   0   'False
         Width           =   4815
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   3660
         Left            =   120
         Picture         =   "Filmovi.frx":0C21
         Top             =   240
         Width           =   4860
      End
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   2280
      TabIndex        =   25
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Next"
      Height          =   255
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Previous"
      Height          =   255
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Save"
      Height          =   255
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   2280
      TabIndex        =   7
      Top             =   5280
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   360
      ItemData        =   "Filmovi.frx":5041
      Left            =   2280
      List            =   "Filmovi.frx":505D
      Sorted          =   -1  'True
      TabIndex        =   12
      Text            =   "Combo1"
      ToolTipText     =   "Select Zanr "
      Top             =   4080
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Remove"
      Height          =   255
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Remove movies"
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Add New"
      Height          =   255
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "New movies"
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   2280
      TabIndex        =   6
      Top             =   4680
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   3480
      Width           =   3615
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   2880
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   2280
      Width           =   3615
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   6480
      TabIndex        =   34
      Top             =   720
      Width           =   1935
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00808080&
      FillColor       =   &H00404040&
      Height          =   495
      Left            =   4440
      Top             =   6120
      Width           =   1455
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   2880
      TabIndex        =   26
      Top             =   720
      Width           =   3255
   End
   Begin VB.Label Label16 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   360
      TabIndex        =   24
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Broj Izdatih:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   720
      TabIndex        =   21
      Top             =   6360
      Width           =   1455
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   2280
      TabIndex        =   20
      Top             =   6240
      Width           =   1095
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Broj Kopija:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   960
      TabIndex        =   19
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "min."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3480
      TabIndex        =   14
      Top             =   5400
      Width           =   375
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Running Time:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   600
      TabIndex        =   13
      Top             =   5400
      Width           =   1575
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00808080&
      Height          =   5295
      Left            =   480
      Top             =   1440
      Width           =   5535
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   5535
      Left            =   360
      Top             =   1320
      Width           =   5775
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "First Published:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   480
      TabIndex        =   11
      Top             =   4800
      Width           =   1695
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Zanr:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1200
      TabIndex        =   10
      Top             =   4200
      Width           =   975
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Actors:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1080
      TabIndex        =   9
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Directed By:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   600
      TabIndex        =   8
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Naslov:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1200
      TabIndex        =   2
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   2280
      TabIndex        =   17
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ID:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "M O V I E S"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Left            =   3840
      TabIndex        =   0
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "Filmovi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Snim As Integer  'Tip snimanja
Public EhoF As Integer  'Kontrola snimanja podataka
Public ImF As String    'Ime filma za proveru promene imena
Public IDbr As Long     'ID broj filma
Option Explicit

Private Sub Check1_Click()
If Check1.Value <> 1 Then
Command10.Enabled = False
Else
Command10.Enabled = True
End If
End Sub

Private Sub Command1_Click()
Load Sablon: Sablon.Show vbModal
End Sub

Private Sub Command10_Click()

On Error Resume Next
If OLE1.Visible = False Then
Command10.ToolTipText = "Hide"
Frame2.Caption = " Video Clip "
OLE1.Visible = True
OLE1.AutoActivate = 2
OLE1.DoVerb -3
OLE1.CreateEmbed App.Path & "\Clips\" & Label3.Caption & ".avi", "mplayer"
OLE1.CreateEmbed App.Path & "\Clips\" & Label3.Caption & ".asf", "mplayer"
OLE1.CreateEmbed App.Path & "\Clips\" & Label3.Caption & ".mpeg", "mplayer"
OLE1.CreateEmbed App.Path & "\Clips\" & Label3.Caption & ".mp3", "mplayer"
OLE1.CreateEmbed App.Path & "\Clips\" & Label3.Caption & ".wav", "mplayer"
OLE1.CreateEmbed App.Path & "\Clips\" & Label3.Caption & ".mid", "mplayer"

Else
Frame2.Caption = " Picture "
Command10.ToolTipText = "Show"
OLE1.Close
OLE1.Visible = False
OLE1.Delete
End If

End Sub

Private Sub Command11_Click()
On Error Resume Next
Frame2.Caption = " Picture "
Pocetna.adoFilmovi.Refresh
Pocetna.adoFilmovi.Recordset.MoveFirst
While Not Pocetna.adoFilmovi.Recordset.EOF
If UCase$(Pocetna.adoFilmovi.Recordset.Fields(1)) >= UCase$(Trim$(Text1.Text)) Then
Form_Activate
Exit Sub
End If
Pocetna.adoFilmovi.Recordset.MoveNext
Wend
Pocetna.adoFilmovi.Recordset.MoveLast
MsgBox "Trazeni naslov nije pronadjen", vbOKOnly, "Obavestenje"
End Sub

Private Sub Command12_Click()
IDbr = Label3.Caption
Load Uvoz
Uvoz.Show vbModal
End Sub

Private Sub Command2_Click()
Dim obWor As Word.Application
Dim obList As Word.Document
Set obWor = New Word.Application
obWor.Visible = True

Set obList = obWor.Documents.Add(App.Path & "\Data\Film.dot")
obList.Activate
obList.Bookmarks("ImeKluba").Range.Text = Glavni.NamClu
obList.Bookmarks("Adresa").Range.Text = Glavni.AdsClu
obList.Bookmarks("Grad").Range.Text = Glavni.Grad
obList.Bookmarks("Telefon").Range.Text = Glavni.TelClu
obList.Bookmarks("Mail").Range.Text = Glavni.Mail
obList.Bookmarks("Web").Range.Text = Glavni.WWW

obList.Bookmarks("Naslov").Range.Text = Text1.Text
obList.Bookmarks("Reziser").Range.Text = Text2.Text
obList.Bookmarks("Uloga").Range.Text = Text3.Text
obList.Bookmarks("Zanr").Range.Text = Combo1.Text
obList.Bookmarks("Trajanje").Range.Text = Text8.Text
obList.Bookmarks("Godina").Range.Text = Text5.Text
obList.Bookmarks("Tekst").Range.Text = Text6.Text

On Error Resume Next
Clipboard.Clear
Clipboard.SetData LoadPicture(App.Path & "\Pictures\" & Label3.Caption & ".jpg")
obList.Bookmarks("Slika").Range.Paste
Set obWor = Nothing

End Sub
Private Sub Command3_Click()
On Error Resume Next
Snim = 1
OLE1.Visible = False
OLE1.Close
OLE1.Delete
Frame2.Caption = " Picture "
Text6.Visible = True
Command5.Caption = "Save New"
Command5.BackColor = &HFFC0C0
Text1.Text = " "
Text2.Text = " "
Text3.Text = " "
Combo1.Text = " "
Text5.Text = " "
Text6.Text = " "
Text7.Text = 1
Text8.Text = 0
Text6.Visible = True
Check1.Value = 0
Label14.Caption = 0
End Sub

Private Sub Command4_Click()
'Brisanje filma iz svih tabela
On Error Resume Next
If MsgBox("Da li potvrdjujete brisnje filma iz baze?", vbOKCancel, "Brisanje") = vbCancel Then Exit Sub
Pocetna.adoFilmovi.Recordset.Delete
Pocetna.adoFilmovi.Refresh
End Sub

Public Sub Command5_Click()
Dim Film As clsFilmovi, fil As Integer, Zanr As String, Com1 As Integer
Dim Sima As String
If Trim$(Text1.Text) = "" Then MsgBox "What is the movie name?": Exit Sub
Command5.Caption = "Save"
Command5.BackColor = &HC0C0C0
Set Film = New clsFilmovi

Film.Ime = Trim$(Text1.Text)

If Trim$(Text2.Text) = "" Then
Film.Rezija = "-"
Else
Film.Rezija = Trim$(Text2.Text)
End If

If Trim$(Text3.Text) = "" Then
Film.GlUloga = "-"
Else
Film.GlUloga = Trim$(Text3.Text)
End If

If Trim$(Combo1.Text) = "" Then
Film.Zanr = "-"
Else
Film.Zanr = Trim$(Combo1.Text)
End If

If Trim$(Text5.Text) = "" Then
Film.Godina = "0"
Else
Film.Godina = Text5.Text
End If

If Check1.Value = 1 Then
Film.Slika = True
Else
Film.Slika = False
End If


If Trim$(Text8.Text) = "" Then
Film.Trajanje = "0"
Else
Film.Trajanje = Text8.Text
End If

If Trim$(Text7.Text) = "" Then
Film.BrKopi = 1
Else
Film.BrKopi = Text7.Text
End If

If Text6.Text = "" Then
Film.Opis = "-"
Else
Film.Opis = Text6.Text
End If

Sima = Combo1.Text
Combo1.Text = "Horror"
For Com1 = 0 To Combo1.ListCount - 1
If UCase$(Film.Zanr) = UCase$(Combo1.List(Com1)) Then GoTo 23
Next Com1
Combo1.AddItem Sima

' Upis zanrova u fajl

 fil = FreeFile
Open App.Path & "\zanr.vic" For Append As #fil
If Trim$(Film.Zanr) <> "" Then
Print #fil, Film.Zanr
End If

Close #fil
23:

'Raskrsnica
If Snim = 1 Then
    Film.SnomanjeNew
    Pocetna.Command6_Click
Else
    Film.Snimanje
    If ImF <> Film.Ime Then
        If MsgBox("Promenjeno je ime filma! Zelite li da ime bude zamenjeno novim u tabelama u arhivi. (u koliko je film vec rentiran pritisnite Yes)", vbYesNo, "Upozorenje!") = vbYes Then
            Izmena.IzmF Film.Ime, ImF
        End If
    End If
End If
Set Film = Nothing
Form_Activate
End Sub

Private Sub Command6_Click()
'Kretanje unazad
Frame2.Caption = " Picture "
Snim = 0
On Error Resume Next
OLE1.Visible = False
OLE1.Close
OLE1.Delete
Command5.BackColor = &HC0C0C0
Pocetna.adoFilmovi.Recordset.MovePrevious
If Pocetna.adoFilmovi.Recordset.BOF Then Pocetna.adoFilmovi.Recordset.MoveLast
Form_Activate
End Sub

Private Sub Command7_Click()
'Kretanje u napred
On Error Resume Next
Frame2.Caption = " Picture "
Snim = 0
OLE1.Visible = False
OLE1.Close
OLE1.Delete
Command5.BackColor = &HC0C0C0
On Error Resume Next
Pocetna.adoFilmovi.Recordset.MoveNext
If Pocetna.adoFilmovi.Recordset.EOF Then Pocetna.adoFilmovi.Recordset.MoveFirst
Form_Activate
End Sub

Private Sub Command8_Click()
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
LiV52.ListItems.Clear
With LiV52
.Left = 1
.Top = 1
.Width = 11800
.Height = 7450
.Visible = True
End With
Dim Jop As ListItem, Bid As String, Zuc As String
Bid = Text1.Text 'Pocetna.adoIzdato.Recordset.Fields("Film")
Pocetna.adoIzdato.Refresh
Pocetna.adoIzdato.RecordSource = "select * from izdato where film='" & Bid & "'"
Pocetna.adoIzdato.Refresh
While Not Pocetna.adoIzdato.Recordset.EOF
    
    Pocetna.adoClanovi.Recordset.MoveFirst
    While Not Pocetna.adoClanovi.Recordset.EOF
        If Pocetna.adoClanovi.Recordset.Fields("ID") = Pocetna.adoIzdato.Recordset.Fields("IDClana") Then
        Zuc = Pocetna.adoClanovi.Recordset.Fields(1) & "  " & Pocetna.adoClanovi.Recordset.Fields(3)
        GoTo 44
        End If
    Pocetna.adoClanovi.Recordset.MoveNext
    Wend
44:
Set Jop = LiV52.ListItems.Add(, , Zuc)
Jop.SubItems(1) = Pocetna.adoIzdato.Recordset.Fields("IDClana")
Jop.SubItems(2) = Pocetna.adoIzdato.Recordset.Fields("DatumRentir")
Jop.SubItems(3) = Pocetna.adoIzdato.Recordset.Fields("Uplaceno")
Jop.SubItems(4) = CInt(Pocetna.adoIzdato.Recordset.Fields("Uplaceno") / Glavni.RV)
Jop.SubItems(5) = Format(DateAdd("d", Jop.SubItems(4), Jop.SubItems(2)), "dd-mm-yy")
Pocetna.adoIzdato.Recordset.MoveNext
Wend
Pocetna.adoIzdato.RecordSource = "select * from izdato"

End Sub

Private Sub Command9_Click()
Dim Horus As Double, dren As String
On Error Resume Next
MsgBox "Ovde mozete da izmenite spisak zanrova. Posle unosa novog zanra pritisnite ENTER"
dren = App.Path & "\zanr.vic"
Horus = Shell("Notepad" & " " & dren, vbNormalFocus)

End Sub

Public Sub Form_Activate()
Dim Hoock As String
On Error Resume Next
If Check1.Value <> 1 Then
Command10.Enabled = False
Else
Command10.Enabled = True
End If
Label3.Caption = Pocetna.adoFilmovi.Recordset.Fields(0)
Text1.Text = Pocetna.adoFilmovi.Recordset.Fields(1)
Text2.Text = Pocetna.adoFilmovi.Recordset.Fields(2)
Text3.Text = Pocetna.adoFilmovi.Recordset.Fields(3)
Text5.Text = Pocetna.adoFilmovi.Recordset.Fields(5)
Text8.Text = Pocetna.adoFilmovi.Recordset.Fields(6)
Text6.Text = Pocetna.adoFilmovi.Recordset.Fields(10)
If Pocetna.adoFilmovi.Recordset.Fields(7) = True Then
Check1.Value = 1
Else
Check1.Value = 0
End If
Text7.Text = Pocetna.adoFilmovi.Recordset.Fields(8)
Label14.Caption = Pocetna.adoFilmovi.Recordset.Fields(9)

'Hoock = Pocetna.adoFilmovi.Recordset.Fields(7)
'Image1.Picture = LoadPicture(App.Path & "\" & Hoock)
'//////////////////////////////////
Dim Hamburg As Integer, Sir As Single
On Error GoTo 33
Image1.Stretch = False
Image1.Visible = False
Hamburg = UCase$(Pocetna.adoFilmovi.Recordset.Fields(0))
Image1.Picture = LoadPicture(App.Path & "\Pictures\" & Hamburg & ".jpg")
Sir = Image1.Width / Image1.Height
If Sir > 1 Then
    If Image1.Height > 3600 Then
        Image1.Height = 3600
        Image1.Width = 3600 * Sir
    End If
    If Image1.Width > 4860 Then
        Image1.Width = 4860
        Image1.Height = 4860 / Sir
    End If
Else
    If Image1.Width > 4860 Then
        Image1.Width = 4860
        Image1.Height = 4860 / Sir
        End If
    If Image1.Height > 3600 Then
        Image1.Height = 3600
        Image1.Width = 3600 * Sir
    End If
End If
Image1.Stretch = True
Image1.Visible = True
GoTo 22
33:
   
'MsgBox "Nema dalje", vbOKOnly, "Obavestenje"
22:


'//////////////////////////////////
Dim fil As Integer, Los As Integer, K As String

Combo1.Clear
fil = FreeFile
Open App.Path & "\zanr.vic" For Input As #fil
While Not EOF(fil)
Input #fil, K
If Trim$(K) <> "" Then
Combo1.AddItem (K)
End If
Wend
Close #fil
Combo1.Text = Pocetna.adoFilmovi.Recordset.Fields(4)
ImF = Text1.Text

End Sub

Private Sub Form_Unload(Cancel As Integer)
If EhoF = 1 Then
If MsgBox("Save changes?", vbYesNo + vbDefaultButton1, "Not saved") = vbYes Then
Command5_Click
End If
End If
EhoF = 0

End Sub

Private Sub LiV52_DblClick()
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
With LiV52
    .Visible = False
    .Left = 240
    .Top = 7800
    .Width = 1935
    .Height = 615
    
End With

End Sub

Private Sub Text1_GotFocus()
ImF = Text1.Text

End Sub
