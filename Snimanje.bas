Attribute VB_Name = "mdDatumi"
Public D As Integer
Option Explicit

Function Dani(Datum As Date) As Integer
Dim G As Integer, M As MonthConstants, Dar As String, Mn As Integer, Prestup As Integer
Dim Petljanje As Integer, Razd As Integer, Dol As Integer

Dar = WeekdayName(Weekday(Datum))
G = Year(Now) - Year(Datum)
Mn = Month(Datum)

If G = 0 Then      'Ista godina
    M = Month(Now) - Month(Datum)

    If M = 0 Then  'Isti mesec u istoj godini
        Dani = Day(Now) - Day(Datum)
    Else           ' Razlicit mesec u istoj godini
        For Petljanje = Mn To (Month(Now) - 1)
        
            If Petljanje = 1 Then Dani = Dani + 31
            If Petljanje = 2 Then
            If (Year(Now) - 2000) / 4 = (Year(Now) - 2000) \ 4 Then
            Dani = Dani + 29
            Else
            Dani = Dani + 28
            End If
            End If
            If Petljanje = 3 Then Dani = Dani + 31
            If Petljanje = 4 Then Dani = Dani + 30
            If Petljanje = 5 Then Dani = Dani + 31
            If Petljanje = 6 Then Dani = Dani + 30
            If Petljanje = 7 Then Dani = Dani + 31
            If Petljanje = 8 Then Dani = Dani + 31
            If Petljanje = 9 Then Dani = Dani + 30
            If Petljanje = 10 Then Dani = Dani + 31
            If Petljanje = 10 Then Dani = Dani + 30
            If Petljanje = 12 Then Dani = Dani + 31
            
        Next Petljanje
        
        Dani = Dani - Day(Datum) + Day(Now)

    End If
Else        'Razlicita godina
    'If G = 1 Then 'Godina za godinom
'Broj dana u prvoj godini
        For Petljanje = Mn To 12
        
            If Petljanje = 1 Then Dani = Dani + 31
            If Petljanje = 2 Then
            If (Year(Datum) - 2000) / 4 = (Year(Datum) - 2000) \ 4 Then
            Dani = Dani + 29
            Else
            Dani = Dani + 28
            End If
            End If
            If Petljanje = 3 Then Dani = Dani + 31
            If Petljanje = 4 Then Dani = Dani + 30
            If Petljanje = 5 Then Dani = Dani + 31
            If Petljanje = 6 Then Dani = Dani + 30
            If Petljanje = 7 Then Dani = Dani + 31
            If Petljanje = 8 Then Dani = Dani + 31
            If Petljanje = 9 Then Dani = Dani + 30
            If Petljanje = 10 Then Dani = Dani + 31
            If Petljanje = 10 Then Dani = Dani + 30
            If Petljanje = 12 Then Dani = Dani + 31
            
        Next Petljanje
        Dani = Dani - Day(Datum)
        
'+Broj dana u zadnjoj godini
        For Petljanje = 1 To (Month(Now) - 1)
        
            If Petljanje = 1 Then Dani = Dani + 31
            If Petljanje = 2 Then
            If (Year(Now) - 2000) / 4 = (Year(Now) - 2000) \ 4 Then
            Dani = Dani + 29
            Else
            Dani = Dani + 28
            End If
            End If
            If Petljanje = 3 Then Dani = Dani + 31
            If Petljanje = 4 Then Dani = Dani + 30
            If Petljanje = 5 Then Dani = Dani + 31
            If Petljanje = 6 Then Dani = Dani + 30
            If Petljanje = 7 Then Dani = Dani + 31
            If Petljanje = 8 Then Dani = Dani + 31
            If Petljanje = 9 Then Dani = Dani + 30
            If Petljanje = 10 Then Dani = Dani + 31
            If Petljanje = 10 Then Dani = Dani + 30
            If Petljanje = 12 Then Dani = Dani + 31
            
        Next Petljanje
            Dani = Dani + Day(Now)
    
    'End If
    
    If G > 1 Then 'Neverovatna mgucnost: Nije godina za godinom
        Razd = Year(Now)
        Prestup = (G + ((Razd - 1996) Mod 4)) / 4 'Broj prestupnih godina
        If (Year(Datum) - 2000) / 4 = (Year(Datum) - 2000) \ 4 Then Prestup = Prestup - 1 'Ako je pocetna godina prestupna, vec je obracunato
        If (Razd - 2000) / 4 = (Razd - 2000) \ 4 Then Prestup = Prestup - 1 ' Ako je zadnja godina prestupna, vec je obracunata
        Dani = Dani + Prestup + (G - 1) * 365
    End If
    
End If
'Broj nedelja
Dol = Weekday(Datum)
D = (Dani + Dol) \ 7

End Function


