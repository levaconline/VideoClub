VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm Glavni 
   BackColor       =   &H8000000C&
   Caption         =   " Video Club"
   ClientHeight    =   8310
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2582
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2E5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":37AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":40A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":49EE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "hom"
            Object.ToolTipText     =   "Home"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "mem"
            Object.ToolTipText     =   "Members"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "mov"
            Object.ToolTipText     =   "Movies"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "sta"
            Object.ToolTipText     =   "Statistic"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ggt"
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "direkt"
            Object.ToolTipText     =   "Direkno u bazu"
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin VB.Menu werw 
      Caption         =   "&File"
      Begin VB.Menu ert 
         Caption         =   "-"
      End
      Begin VB.Menu Exi 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu tol 
      Caption         =   "&Tools"
      Begin VB.Menu clrb 
         Caption         =   "Clear Data"
      End
      Begin VB.Menu lofe 
         Caption         =   "Load Data"
      End
      Begin VB.Menu phio 
         Caption         =   "-"
      End
      Begin VB.Menu Rep 
         Caption         =   "&Report"
         Begin VB.Menu mov 
            Caption         =   "Movies"
         End
         Begin VB.Menu mem 
            Caption         =   "Members"
         End
      End
      Begin VB.Menu rer 
         Caption         =   "-"
      End
      Begin VB.Menu psw 
         Caption         =   "Password"
      End
      Begin VB.Menu opt 
         Caption         =   "Setings"
      End
   End
   Begin VB.Menu hlp 
      Caption         =   "Help"
      Begin VB.Menu hel 
         Caption         =   "Help"
      End
      Begin VB.Menu Abba 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Glavni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public RV As Currency, InSund As Boolean, MemPer As Integer, NamClu As String, AdsClu As String, TelClu As String 'Podaci za Podesavanje: Cena Iznajmljivanja, Da li se obracunava i nedelja, Rok trajanja clanstva, Ime kluba, adresa kluba, telefon kluba
Public Grad As String, WWW As String, Mail As String
Public Smer1 As Integer  'Usmerivac novih ili starih clanova
Public Direkt As Boolean
Option Explicit

Private Sub Abba_Click()
Load AboutO
AboutO.Show vbModal
End Sub


Private Sub clrb_Click()
Direkt = False
Load Password: Password.Show vbModal
End Sub

Private Sub Exi_Click()

Unload Direct
Set Direct = Nothing

Unload Clanstvo
Set Clanstvo = Nothing

Unload Filmovi
Set Filmovi = Nothing

Unload Statistika
Set Statistika = Nothing

Unload Glavni
Set Glavni = Nothing

End
End Sub

Private Sub MDIForm_Activate()
Glavni.Caption = NamClu
Load Pocetna

End Sub

Private Sub MDIForm_Load()
Dim Fio As Integer, Difi As String
Difi = Dir$(App.Path & "\")
While Difi <> ""
If Difi = "Info.aca" Then
GoTo 12
End If
 Difi = Dir$
Wend
MsgBox "Welcome In Video Club! In first time you mast to define setings...", vbOKOnly, "WELCOME"
Load Opcije: Opcije.Show vbModal

12:
Fio = FreeFile
If Right$(App.Path, 1) <> "\" Then
Open App.Path & "\Info.aca" For Input As #Fio
Else
Open App.Path & "Info.aca" For Input As #Fio
End If

Input #Fio, RV, InSund, MemPer, NamClu, AdsClu, TelClu, Grad, WWW, Mail
Close #Fio
'////////////////////
Difi = Dir$(App.Path & "\")
While Difi <> ""
If Difi = "zanr.vic" Then
GoTo 121
End If
 Difi = Dir$
Wend
Open App.Path & "\zanr.vic" For Output As #Fio
Print #Fio, "Horor"
Close #Fio
121:
End Sub

Private Sub MDIForm_Resize()
On Error Resume Next

If Me.Width > 12000 Then Me.Width = 12000
If Me.Height > 9000 Then Me.Height = 9000
End Sub

Private Sub mem_Click()
'Set DRC.DataSource = Pocetna.adoClanovi
DRC.Show

End Sub

Private Sub mov_Click()
'Set DataReport1h.DataSource = Pocetna.adoFilmovi
DataReport1h.Show
End Sub

Private Sub opt_Click()
Load Opcije
Opcije.Show vbModal
End Sub

Private Sub psw_Click()
Load ChangPass
ChangPass.Show vbModal
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "hom"
Load Pocetna: Pocetna.SetFocus
Case "mem"
Smer1 = 0: Unload Clanstvo: Load Clanstvo: Clanstvo.SetFocus
Case "mov"
Load Filmovi: Filmovi.SetFocus
Case "sta"
Load Statistika: Statistika.SetFocus
Case "direkt"
Direkt = True
Load Password: Password.Show vbModal
End Select
End Sub
