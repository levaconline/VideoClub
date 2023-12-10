VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Clanstvo 
   BackColor       =   &H00C0C0C0&
   Caption         =   " Members"
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
   Icon            =   "Clanstvo.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2895
      Left            =   6000
      TabIndex        =   38
      Top             =   2520
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   5106
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      Enabled         =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Members"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Word"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0C0C0&
      Caption         =   ">>"
      Height          =   255
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Next"
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0C0&
      Caption         =   "<<"
      Height          =   255
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Previous"
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0C0&
      Caption         =   "AddNew"
      Height          =   255
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Remove"
      Height          =   255
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   840
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Searching"
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
      Height          =   975
      Left            =   6000
      TabIndex        =   20
      Top             =   1440
      Width           =   5415
      Begin VB.ComboBox Combo1 
         Height          =   330
         ItemData        =   "Clanstvo.frx":030A
         Left            =   120
         List            =   "Clanstvo.frx":0329
         TabIndex        =   37
         Text            =   "Name"
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox Text10 
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
         Left            =   1560
         TabIndex        =   21
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   255
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "SAVE"
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox Text9 
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
      Height          =   420
      Left            =   1920
      TabIndex        =   18
      Top             =   6480
      Width           =   2175
   End
   Begin VB.TextBox Text8 
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
      Height          =   420
      Left            =   1920
      TabIndex        =   16
      Top             =   5880
      Width           =   2175
   End
   Begin VB.TextBox Text7 
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
      Height          =   420
      Left            =   1920
      TabIndex        =   15
      Top             =   5280
      Width           =   3375
   End
   Begin VB.TextBox Text6 
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
      Height          =   420
      Left            =   1920
      TabIndex        =   14
      Top             =   4680
      Width           =   3375
   End
   Begin VB.TextBox Text5 
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
      Height          =   420
      Left            =   1920
      TabIndex        =   13
      Top             =   4080
      Width           =   3375
   End
   Begin VB.TextBox Text4 
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
      Height          =   420
      Left            =   1920
      TabIndex        =   8
      Top             =   3360
      Width           =   3375
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
      Height          =   420
      Left            =   1920
      TabIndex        =   7
      Top             =   2760
      Width           =   3375
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
      Height          =   420
      Left            =   1920
      TabIndex        =   6
      Top             =   2160
      Width           =   3375
   End
   Begin VB.TextBox Text1 
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
      Height          =   420
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Membership"
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
      Height          =   1455
      Left            =   6000
      TabIndex        =   23
      Top             =   5520
      Width           =   5415
      Begin VB.CommandButton Command6 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Produzi clanstvo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   960
         Width           =   2415
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Produzenje Clanstva: "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   375
         Left            =   240
         TabIndex        =   28
         Top             =   960
         Width           =   2415
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Start membership: "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   375
         Left            =   240
         TabIndex        =   25
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label12 
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
         Left            =   2760
         TabIndex        =   24
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.Label Label13 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   6000
      TabIndex        =   35
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label Label18 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   2280
      TabIndex        =   33
      Top             =   720
      Width           =   3375
   End
   Begin VB.Label Label17 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   240
      TabIndex        =   30
      Top             =   720
      Width           =   2055
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00808080&
      Height          =   6135
      Left            =   240
      Top             =   1200
      Width           =   5415
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Document:"
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
      Height          =   375
      Left            =   480
      TabIndex        =   17
      Top             =   6600
      Width           =   1335
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Telephone:"
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
      Height          =   375
      Left            =   480
      TabIndex        =   12
      Top             =   5880
      Width           =   1335
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "State:"
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
      Height          =   375
      Left            =   840
      TabIndex        =   11
      Top             =   5280
      Width           =   975
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "City:"
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
      Height          =   375
      Left            =   960
      TabIndex        =   10
      Top             =   4680
      Width           =   855
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
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
      Height          =   375
      Left            =   600
      TabIndex        =   9
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Surname:"
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
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Midle:"
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
      Height          =   495
      Left            =   600
      TabIndex        =   3
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
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
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   2280
      Width           =   1215
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
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "M E M B E R S"
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
      Height          =   495
      Left            =   3480
      TabIndex        =   0
      Top             =   120
      Width           =   4695
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      FillColor       =   &H00004000&
      Height          =   5895
      Left            =   360
      Top             =   1320
      Width           =   5175
   End
End
Attribute VB_Name = "Clanstvo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Eho As Integer  'Kontrola snimanja podataka (da li je zaboravljeno snimanje izmena?)
Public Krik As Integer 'Kontrola, novi clan, ili izmene
Public Sabloni As Integer
Option Explicit

Public Sub Command1_Click()
Dim Clan As clsClan
Set Clan = New clsClan
Clan.Ime = Trim$(Text2.Text)
If Clan.Ime = "" Then MsgBox "Nije upisano ime clana.", vbOKOnly, "Greska": Exit Sub
Clan.SredIme = Trim$(Text3.Text)
If Clan.SredIme = "" Then Clan.SredIme = " "
Clan.Prezme = Trim$(Text4.Text)
If Clan.Prezme = "" Then MsgBox "Nije upisano prezime clana.", vbOKOnly, "Greska": Exit Sub
Clan.Ads = Text5.Text
If Clan.Ads = "" Then Clan.Ads = " "
Clan.Mesto = Text6.Text
If Clan.Mesto = "" Then Clan.Mesto = " "
Clan.Drzava = Text7.Text
If Clan.Drzava = "" Then Clan.Drzava = " "
Clan.Telefon = Text8.Text
If Clan.Telefon = "" Then Clan.Telefon = " "
Clan.Dokument = Text9.Text
If Clan.Dokument = "" Then Clan.Dokument = " "
Clan.StartClanstva = Label12.Caption
If Krik = 1 Then
        Clan.SnimanjeNew
        Pocetna.Command5_Click
        
    Else
        Clan.Snimanje
End If
Eho = 0

Form_Activate
Command1.BackColor = &HC0C0C0
Command1.Caption = "Save"
Command1.Enabled = False
Command4.Enabled = True
Command3.Enabled = True

Set Clan = Nothing
End Sub

Private Sub Command2_Click()
If Text10.Text = "" Then Exit Sub

Dim Pop As String, Rty As String, C As Integer
'Pronaci kategoriju iz Text1 u bazi podataka i rezultate prikazati u List1
Pop = Trim$(Combo1.Text)
If Pop = "Name" Then Pop = "Ime"
If Pop = "MidName" Then Pop = "Srednje Ime"
If Pop = "Surname" Then Pop = "Prezime"
If Pop = "Address" Then Pop = "Adresa"
If Pop = "City" Then Pop = "Grad"
If Pop = "State" Then Pop = "Drzava"
If Pop = "Phono" Then Pop = "Telefon"
If Pop = "Document" Then Pop = "Dokument"


If Pop = "Kategorije" Then MsgBox "Odaberite kategoriju", vbOKOnly, "Greska": Exit Sub
Rty = Trim$(Text10.Text)
C = Len(Rty)
Pocetna.adoClanovi.RecordSource = "select * from clanovi where [" & Pop & "] >= '" & Rty & "' order by [" & Pop & "]"        ' like 's*' " ''" & Rty & "'"
Pocetna.adoClanovi.Refresh

'Pronaci clana u listi i podatke ucitati u polja
'Pocetna.adoClanovi.Recordset.MoveFirst
'While Not Pocetna.adoClanovi.Recordset.EOF
'If UCase$(Pocetna.adoClanovi.Recordset.Fields(Pop)) = UCase$(Trim$(Text10.Text)) Then
Form_Activate
'Exit Sub
'Else
'Pocetna.adoClanovi.Recordset.MoveNext
'End If
'Wend
'MsgBox "Clan nije pronadjen", vbOKOnly, "Obavestenje"
'Pocetna.adoClanovi.Recordset.MoveLast
End Sub

Private Sub Command3_Click()
'Ucitati u listu clanove iz baze podataka
Dim obWor As Word.Application
Dim obList As Word.Document
Set obWor = New Word.Application
obWor.Visible = True

Set obList = obWor.Documents.Add(App.Path & "\Data\Clan.dot")
obList.Activate
obList.Bookmarks("ImeKluba").Range.Text = Glavni.NamClu
obList.Bookmarks("Adresa").Range.Text = Glavni.AdsClu
obList.Bookmarks("Grad").Range.Text = Glavni.Grad
obList.Bookmarks("Telefon").Range.Text = Glavni.TelClu
obList.Bookmarks("Mail").Range.Text = Glavni.Mail
obList.Bookmarks("Web").Range.Text = Glavni.WWW

obList.Bookmarks("ID").Range.Text = Text1.Text
obList.Bookmarks("ImeCl").Range.Text = Text2.Text
obList.Bookmarks("SredIme").Range.Text = Text3.Text
obList.Bookmarks("Prezime").Range.Text = Text4.Text
obList.Bookmarks("Ads").Range.Text = Text5.Text
obList.Bookmarks("MesCl").Range.Text = Text6.Text
obList.Bookmarks("Zemlja").Range.Text = Text7.Text
obList.Bookmarks("TelCl").Range.Text = Text8.Text
obList.Bookmarks("Doc").Range.Text = Text9.Text

Set obWor = Nothing


End Sub

Private Sub Command4_Click()
Sabloni = 1
Load Sablon: Sablon.Show vbModal
'Print
End Sub

Private Sub Command5_Click()
' Provera, da li se clan razduzio
Pocetna.adoIzdato.Refresh
While Not Pocetna.adoIzdato.Recordset.EOF
If Pocetna.adoIzdato.Recordset.Fields(1) = Pocetna.adoClanovi.Recordset.Fields(0) Then MsgBox "Clan nije vratio sve filmove. Ne moze biti izbrisan.", vbOKOnly, "Obavestenje": Exit Sub
Pocetna.adoIzdato.Recordset.MoveNext
Wend

' Brisac
If MsgBox("Potvrdite brisanje podataka o clanu.", vbYesNo + vbDefaultButton2, "BRISANJE PODATAKA!") = vbNo Then Exit Sub
If Glavni.Smer1 = 1 Then
Pocetna.adoClanovi.Recordset.Delete
Else
Pocetna.adoClanovi.Recordset.Delete
End If
End Sub

Private Sub Command6_Click()
If MsgBox("Zelite da produzite clanstvo?", vbYesNo, "Upit") = vbYes Then
     Pocetna.adoClanovi.Recordset.Fields(9) = Format(Now, "dd-mm-yy")
     Pocetna.adoClanovi.Recordset.Update
     Label12.Caption = Format(Now, "dd-mm-yy")
     MsgBox "Clanstvo je produzeno", vbOKOnly, "Obavestenje"
End If
End Sub

Private Sub Command7_Click()
Dim Ggin As Integer, Mmes As Integer, Roda As Integer ' Roda = mesec + period clanstva

Roda = Month(Now) + Glavni.MemPer
Label12.Caption = Format(Now, "dd-mm-yy")

    If Roda > 12 Then
        Ggin = Year(Now) + Roda \ 12
        Mmes = Roda Mod 12
    Else
        Ggin = Year(Now)
        Mmes = Roda
    End If

Text1.Text = " "
Text2.Text = " "
Text3.Text = " "
Text4.Text = " "
Text5.Text = " "
Text6.Text = " "
Text7.Text = " "
Text8.Text = " "
Text9.Text = " "
Command4.Enabled = False
Command3.Enabled = False
Command6.Enabled = False
Command5.Enabled = False
Command1.Caption = "Save New"
Command1.BackColor = &HFFC0C0
Text2.Text = " "
Text3.Text = " "
Text4.Text = " "
Text5.Text = " "
Text6.Text = " "
Text7.Text = " "
Text8.Text = " "
Text9.Text = " "
Krik = 1
End Sub

Private Sub Command8_Click()
On Error Resume Next
Pocetna.adoClanovi.Recordset.MovePrevious
If Pocetna.adoClanovi.Recordset.BOF Then Pocetna.adoClanovi.Recordset.MoveLast
Form_Activate
Command1.BackColor = &HC0C0C0
Command1.Caption = "Save"
Command4.Enabled = True
Command3.Enabled = True
Eho = 0
Krik = 0

End Sub

Private Sub Command9_Click()
On Error Resume Next
Pocetna.adoClanovi.Recordset.MoveNext
If Pocetna.adoClanovi.Recordset.EOF Then Pocetna.adoClanovi.Recordset.MoveFirst
Form_Activate
Command1.BackColor = &HC0C0C0
Command1.Caption = "Save"
Command4.Enabled = True
Command3.Enabled = True
Krik = 0
Eho = 0
End Sub

Private Sub DataGrid1_SelChange(Cancel As Integer)
Form_Activate
End Sub

Public Sub Form_Activate()
On Error Resume Next
Set DataGrid1.DataSource = Pocetna.adoClanovi
Command5.Enabled = True
Text1.Text = Pocetna.adoClanovi.Recordset.Fields(0)
Text2.Text = Pocetna.adoClanovi.Recordset.Fields(1)
Text3.Text = Pocetna.adoClanovi.Recordset.Fields(2)
Text4.Text = Pocetna.adoClanovi.Recordset.Fields(3)
Text5.Text = Pocetna.adoClanovi.Recordset.Fields(4)
Text6.Text = Pocetna.adoClanovi.Recordset.Fields(5)
Text7.Text = Pocetna.adoClanovi.Recordset.Fields(6)
Text8.Text = Pocetna.adoClanovi.Recordset.Fields(7)
Text9.Text = Pocetna.adoClanovi.Recordset.Fields(8)
Label12.Caption = Pocetna.adoClanovi.Recordset.Fields(9)

If Glavni.MemPer = 0 Then
Command6.Enabled = False: Label16.Enabled = False
Else
Command6.Enabled = True: Label16.Enabled = True
End If

End Sub


Private Sub Form_LostFocus()
Command1.BackColor = &HC0C0C0
Command1.Caption = "Save"
Krik = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)
If Eho = 1 Then
If MsgBox("Save changes im member's data?", vbYesNo + vbDefaultButton1, "Not saved") = vbYes Then
Command1_Click
End If
End If
Eho = 0
End Sub

Private Sub Text2_GotFocus()
Eho = 1
Command1.Enabled = True
End Sub

Private Sub Text3_GotFocus()
Eho = 1
Command1.Enabled = True
End Sub

Private Sub Text4_GotFocus()
Eho = 1
Command1.Enabled = True
End Sub

Private Sub Text5_GotFocus()
Eho = 1
Command1.Enabled = True
End Sub

Private Sub Text6_GotFocus()
Eho = 1
Command1.Enabled = True
End Sub

Private Sub Text7_GotFocus()
Eho = 1
Command1.Enabled = True
End Sub

Private Sub Text8_GotFocus()
Eho = 1
Command1.Enabled = True
End Sub

Private Sub Text9_GotFocus()
Eho = 1
Command1.Enabled = True
End Sub
