VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFilmovi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Ime As String
Public Rezija As String
Public GlUloga As String
Public Zanr As String
Public Godina As Integer
Public Trajanje As Integer
Public Slika As Boolean
Public BrKopi As Integer       'Da li je film u klubu
Public Opis As String
Option Explicit

Public Sub SnomanjeNew()
Pocetna.adoFilmovi.Refresh
With Pocetna.adoFilmovi.Recordset
    .AddNew
    .Fields(1) = Ime
    .Fields(2) = Rezija
    .Fields(3) = GlUloga
    .Fields(4) = Zanr
    .Fields(5) = Godina
    .Fields(6) = Trajanje
    .Fields(7) = Slika
    .Fields(8) = BrKopi
    .Fields(9) = 0
    .Fields(10) = Opis
    .Update
End With
Filmovi.Snim = 0
End Sub

Public Sub Snimanje()
On Error Resume Next
With Pocetna.adoFilmovi.Recordset
    .Fields(1) = Ime
    .Fields(2) = Rezija
    .Fields(3) = GlUloga
    .Fields(4) = Zanr
    .Fields(5) = Godina
    .Fields(6) = Trajanje
    .Fields(7) = Slika
    .Fields(8) = BrKopi
    .Fields(10) = Opis
    .Update
End With

End Sub
