Attribute VB_Name = "Izmena"
Option Explicit

Public Function IzmF(NIme As String, SIme As String)

'Zamena u Arhivi
'****************
Pocetna.adoArhiva.RecordSource = "select * from arhiva"
Pocetna.adoArhiva.Refresh
While Not Pocetna.adoArhiva.Recordset.EOF
If Pocetna.adoArhiva.Recordset.Fields("Film") = SIme Then
Pocetna.adoArhiva.Recordset.Fields("Film") = NIme
Pocetna.adoArhiva.Recordset.Update
End If
Pocetna.adoArhiva.Recordset.MoveNext
Wend

'Zamena u Izdato
'******************
Pocetna.adoIzdato.RecordSource = "select * from Izdato"
Pocetna.adoIzdato.Refresh
While Not Pocetna.adoIzdato.Recordset.EOF
If Pocetna.adoIzdato.Recordset.Fields("Film") = SIme Then
Pocetna.adoIzdato.Recordset.Fields("Film") = NIme
Pocetna.adoIzdato.Recordset.Update
End If
Pocetna.adoIzdato.Recordset.MoveNext
Wend

End Function
