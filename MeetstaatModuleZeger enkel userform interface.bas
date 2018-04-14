Attribute VB_Name = "MeetstaatModule"
Option Explicit
Public Datum As Date
Public DatumDag As Integer
Public DatumMaand As Integer
Public DatumJaar As Integer
Public x As Integer
Public z As Integer
Public y As Integer
Public i As Integer


Public CurrentRow As Integer

Public Dagprijs As Integer

Public Locatie As String
Public Projectnummer As String
Public Beschrijving As String
Public Naam As String
Public Bedrijf As String
Public Projectnaam As String
Public Uitvoering As String

Public Omschrijving As String
Public Beginuur As Date
Public Einduur As Date

'Nu de kolommen soft maken
Public DayK As Integer
Public DateK As Integer
Public BeginuurK As Integer
Public EinduurK As Integer
Public LocatieK As Integer
Public UrenK As Integer
Public SoortPrijsK As Integer
Public BedrijfK As Integer
Public ContactpersoonK As Integer
Public ProjectnummerK As Integer
Public ProjectnaamK As Integer
Public UitvoeringK As Integer
Public DagprijsK As Integer
Public FacturatieK As Integer
Public OmschrijvingK As Integer
Public Prijs As Double

Sub LijnVerwijderen()
Dim x As Range
Dim y As Integer



On Error Resume Next
    Set x = Application.InputBox(Prompt:="Kies de lijn die u graag zou verwijderen.", Title:="Lijn verwijderen", Type:=8)
Err.Clear
On Error GoTo 0

If x Is Nothing Then
End
Else
y = x.Row
ActiveSheet.Cells(y, 1).EntireRow.Delete
End If


End Sub

Sub Aanmaken()

Call Inlezen

End Sub
Sub Inlezen()

'Kolommen even ingeven
DayK = 1
DateK = 2
BeginuurK = 3
EinduurK = 4
LocatieK = 5
UrenK = 6
SoortPrijsK = 7
BedrijfK = 8
ContactpersoonK = 9
ProjectnummerK = 10
ProjectnaamK = 11
UitvoeringK = 12
DagprijsK = 13
FacturatieK = 14
OmschrijvingK = 15

If Meetstaat.Cells(4, 1).Value = "" Then
CurrentRow = 4
Else
CurrentRow = Meetstaat.Range("A3").End(xlDown).Offset(1, 0).Row
End If


'Datum zou nu normaal al in orde moeten zijn
Beginuur = MeetstaatForm.CboBeginuur.Value & ":" & MeetstaatForm.CboBeginminuut.Value
Einduur = MeetstaatForm.CboEinduur.Value & ":" & MeetstaatForm.CboEindminuut.Value


Meetstaat.Cells(CurrentRow, DayK).Value = MeetstaatForm.TxtDag
Meetstaat.Cells(CurrentRow, DateK).Value = Datum
Meetstaat.Cells(CurrentRow, BeginuurK).Value = Beginuur
Meetstaat.Cells(CurrentRow, EinduurK).Value = Einduur
Meetstaat.Cells(CurrentRow, LocatieK).Value = MeetstaatForm.CboLocatie.Value
Meetstaat.Cells(CurrentRow, UrenK).Value = (Einduur - Beginuur) * 24 'probeer dit nog in vrije tijd te fixen
'Hier moet dan nog soortprijs komen
Meetstaat.Cells(CurrentRow, BedrijfK).Value = MeetstaatForm.CboBedrijf.Value
Meetstaat.Cells(CurrentRow, ContactpersoonK).Value = MeetstaatForm.TxtNaam.Value
Meetstaat.Cells(CurrentRow, ProjectnummerK).Value = MeetstaatForm.CboProjectnummer.Value
Meetstaat.Cells(CurrentRow, ProjectnaamK).Value = MeetstaatForm.TxtProjectnaam.Value
Meetstaat.Cells(CurrentRow, UitvoeringK).Value = MeetstaatForm.CboUitvoering.Value
Meetstaat.Cells(CurrentRow, DagprijsK).Value = MeetstaatForm.TxtDagprijs.Value
Meetstaat.Cells(CurrentRow, OmschrijvingK).Value = MeetstaatForm.TxtOmschrijving.Value

'Meetstaat.Cells(CurrentRow, FacturatieK).Value = Meetstaat.Cells(CurrentRow, DagprijsK).Value * Meetstaat.Cells(CurrentRow, UrenK).Value




'Sheets(1).Cells(369, 5).Value = Beginuur
'Sheets(1).Cells(369, 6).Value = Einduur
'Sheets(1).Cells(369, 7).Value = (Sheets(1).Cells(369, 6).Value - Sheets(1).Cells(369, 5).Value) * 24 'Om het in uren uit te drukken



'a = Einduur - Beginuur


End Sub

Sub test()


Sheets(1).Cells(6, 3).Value = "8:00"
Sheets(1).Cells(6, 4).Value = "13:00"
Sheets(1).Cells(6, 5).Value = (Sheets(1).Cells(369, 6).Value - Sheets(1).Cells(369, 5).Value) * 24 'Om het in uren uit te drukken


End Sub
