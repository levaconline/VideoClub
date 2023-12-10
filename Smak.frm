VERSION 5.00
Begin VB.Form Smak 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " DataBase Clearning"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2985
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Smak.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   2985
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Cancel"
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00000000&
      Height          =   615
      Left            =   2280
      Picture         =   "Smak.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Brisanje odabranih podataka"
      Top             =   240
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Caption         =   " Delete what? "
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Selektujte stavke koje zelite da budu izbrisane iz baze podataka."
      Top             =   120
      Width           =   1935
      Begin VB.CheckBox Check3 
         Caption         =   " All Archives"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Brisanje svih podataka iz arhive"
         Top             =   960
         Width           =   1455
      End
      Begin VB.CheckBox Check2 
         Caption         =   " All Members"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Brisanje svih podataka o clanovima kluba. Ukljucuje i brisanje spiska trenutnih zaduzenja."
         Top             =   600
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         Caption         =   "  All Movies"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "Brisanje svih filmova iz baze podataka. Ne podrazumeva brisanje trenutnih zaduzenja."
         Top             =   240
         Width           =   1335
      End
   End
End
Attribute VB_Name = "Smak"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
If MsgBox("Potvrdjujete li brisanje odabranih podataka?", vbYesNo + vbDefaultButton2, "Provera") = vbYes Then

    If Check1.Value = 1 Then
        'Pocetna.adoFilmovi.RecordSource = "delete from Filmovi"
        Pocetna.adoFilmovi.Refresh
        While Not Pocetna.adoFilmovi.Recordset.EOF
        Pocetna.adoFilmovi.Recordset.Delete
        Pocetna.adoFilmovi.Recordset.MoveNext
        Wend
        'If Pocetna.adoFilmovi.Recordset.EOF Then Pocetna.adoFilmovi.Recordset.MoveLast
'Pocetna.adoFilmovi.Recordset.Update
        
    End If
    If Check2.Value = 1 Then
    Pocetna.adoClanovi.Refresh
    While Not Pocetna.adoClanovi.Recordset.EOF
    Pocetna.adoClanovi.Recordset.Delete
    Pocetna.adoClanovi.Recordset.MoveNext
    Wend
        'Pocetna.adoClanovi.RecordSource = "delete from Clanovi"
'Pocetna.adoClanovi.Recordset.Update
    Pocetna.adoIzdato.Refresh
    While Not Pocetna.adoIzdato.Recordset.EOF
    Pocetna.adoIzdato.Recordset.Delete
    Pocetna.adoIzdato.Recordset.MoveNext
    Wend
        'Pocetna.adoIzdato.RecordSource = "delete from Izdato"
'Pocetna.adoIzdato.Recordset.Update
        
    End If
        If Check3.Value = 1 Then
       ' Pocetna.adoArhiva.RecordSource = "delte from Arhiva"
        Pocetna.adoArhiva.Refresh
        While Not Pocetna.adoArhiva.Recordset.EOF
        Pocetna.adoArhiva.Recordset.Delete
        Pocetna.adoArhiva.Recordset.MoveNext
        Wend
'Pocetna.adoArhiva.Refresh
        
    End If
Unload Smak
Set Smak = Nothing
End If
End Sub

Private Sub Command2_Click()
Unload Smak
Set Smak = Nothing
End Sub
