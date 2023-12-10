VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Sablon 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Opis"
   ClientHeight    =   8340
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   7230
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Slike.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8340
   ScaleWidth      =   7230
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   2280
      ScaleHeight     =   241
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   316
      TabIndex        =   1
      Top             =   1320
      Width           =   4740
   End
   Begin RichTextLib.RichTextBox Sablon1 
      Height          =   8295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   14631
      _Version        =   393217
      BackColor       =   16777215
      Enabled         =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"Slike.frx":030A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu fif 
      Caption         =   "File"
      Begin VB.Menu prn 
         Caption         =   "&Print"
      End
      Begin VB.Menu frf 
         Caption         =   "-"
      End
      Begin VB.Menu hit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Sablon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub Popuna(Polje As String, Sadrzaj As String)
Sablon1.Find "#" & Polje & "#"
Sablon1.SelText = Sadrzaj

End Sub

Private Sub Form_Activate()
Select Case Clanstvo.Sabloni
Case 0
Sablon1.Text = ""
Sablon1.LoadFile (App.Path & "\Data\ImeKluba.rtf")
Popuna "Film", Filmovi.Text1.Text
Popuna "Zanr", Filmovi.Combo1.Text
Popuna "Rezija", Filmovi.Text2.Text
Popuna "GlUloga", Filmovi.Text3.Text
Popuna "Godina", Filmovi.Text5.Text
Popuna "Trajanje", Filmovi.Text8.Text
Popuna "Text", Pocetna.adoFilmovi.Recordset.Fields(10)
On Error Resume Next
Picture1.Visible = True
Picture1.Picture = LoadPicture(App.Path & "\Pictures\" & Filmovi.Label3.Caption & ".jpg")

Case 1
Sablon1.Text = ""
Sablon1.LoadFile (App.Path & "\Data\Clanovi.rtf")
Picture1.Visible = False
Popuna "ime", Clanstvo.Text2.Text
Popuna "srime", Clanstvo.Text3.Text
Popuna "prezime", Clanstvo.Text4.Text
Popuna "ads", Clanstvo.Text5.Text
Popuna "mestoc", Clanstvo.Text6.Text
Popuna "drzava", Clanstvo.Text7.Text
Popuna "telefon", Clanstvo.Text8.Text
Popuna "dokument", Clanstvo.Text9.Text
Clanstvo.Sabloni = 0
End Select

Popuna "ImeKluba", Glavni.NamClu
Popuna "Adresa", Glavni.AdsClu
Popuna "Mesto", Glavni.Grad
Popuna "Tel", "Pho. " & Glavni.TelClu
Popuna "Mail", Glavni.Mail
Popuna "Web", Glavni.WWW


End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Sablon: Set Sablon = Nothing
End Sub

Private Sub hit_Click()
Unload Sablon: Set Sablon = Nothing
End Sub

Private Sub prn_Click()
Sablon.PrintForm

End Sub

