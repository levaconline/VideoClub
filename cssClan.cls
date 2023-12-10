VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsClan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Ime As String
Public Prezme As String
Public SredIme As String
Public Ads As String
Public Mesto As String
Public Drzava As String
Public Telefon As String
Public StartClanstva As String
Public Dokument


Public Sub SnimanjeNew()
Pocetna.adoClanovi.Refresh
With Pocetna.adoClanovi.Recordset
    .AddNew
    .Fields(1) = Ime
    .Fields(3) = Prezme
    .Fields(2) = SredIme
    .Fields(4) = Ads
    .Fields(5) = Mesto
    .Fields(6) = Drzava
    .Fields(7) = Telefon
    .Fields(8) = Dokument
    .Fields(9) = StartClanstva
    .Update
End With
Clanstvo.Krik = 0
End Sub
Public Sub Snimanje()
On Error Resume Next
With Pocetna.adoClanovi.Recordset
    .Fields(1) = Ime
    .Fields(3) = Prezme
    .Fields(2) = SredIme
    .Fields(4) = Ads
    .Fields(5) = Mesto
    .Fields(6) = Drzava
    .Fields(7) = Telefon
    .Fields(8) = Dokument
    .Fields(9) = StartClanstva
    .Update
End With
End Sub
