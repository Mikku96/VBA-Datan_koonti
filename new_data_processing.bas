Attribute VB_Name = "Module1"
''''''''''''''''''''''''''''''''''''''''
' VÄLILEHDEN OLEMASSAOLON TESTAUS (FUNKTIO)
' Tarkistetaan, että välilehteä olemassa
' Palauttaa True tai False
''''''''''''''''''''''''''''''''''''''''
Function WorksheetExists(shtName As String, Optional wb As Workbook) As Boolean
    Dim sht As Worksheet
    If wb Is Nothing Then Set wb = ThisWorkbook
    On Error Resume Next
    Set sht = wb.Sheets(shtName)
    On Error GoTo 0
    WorksheetExists = Not sht Is Nothing
End Function

Function AddressOfMax(rng As Range) As Range
    Set AddressOfMax = rng.Cells(WorksheetFunction.Match(WorksheetFunction.max(rng), rng, 0))

End Function
Sub mean_count_comment(mean As Range, errors As Range, number_of_obs As Integer, object As Integer, position As Integer, comments As String)
    Dim vali_mean As String, vali_max_virhe As String
    vali_mean = Application.WorksheetFunction.Round(mean.value, 2)
    vali_max_virhe = WorksheetFunction.Round(Application.Aggregate(4, 6, errors), 2)
    Sheets("Processed").Cells(1 + object, position * 4 - 2) = vali_mean
    Sheets("Processed").Cells(1 + object, position * 4 - 1) = vali_max_virhe
    Sheets("Processed").Cells(1 + object, position * 4) = number_of_obs
    Sheets("Processed").Cells(1 + object, position * 4 + 1) = comments
End Sub
Sub comment_only(object As Integer, position As Integer, value As String, comments As String)
    Sheets("Processed").Cells(1 + object, position * 4 - 2) = value
    Sheets("Processed").Cells(1 + object, position * 4 + -1) = value
    Sheets("Processed").Cells(1 + object, position * 4) = value
    Sheets("Processed").Cells(1 + object, position * 4 + 1) = comments
End Sub
Function counter() As Long

Dim counting As Long
Dim iRange As Range
Dim trg As Worksheet
Set trg = ThisWorkbook.Worksheets("Processed")
With trg.Range("A1:A100")

    'loop through each row from the used range
    For Each iRange In .Rows

        'check if the row contains a cell with a value
        If Application.CountA(iRange) > 0 Then

            'counts the number of rows non-empty Cells
            counting = counting + 1

        End If

    Next

End With
counter = counting - 1
End Function

Sub generate_scatterplot()
Dim ochartObj As ChartObject
Dim oChart As Chart

Set ochartObj = ActiveSheet.ChartObjects.Add(Top:=10, Left:=325, Width:=600, Height:=300)
Set oChart = ochartObj.Chart
Dim trg As Worksheet
Set trg = ThisWorkbook.Worksheets("Processed")
oChart.ChartType = xlXYScatter
oChart.ChartStyle = 245

Dim number_of_rows As Long
number_of_rows = counter

Dim objects As Range
Set objects = trg.Range("A" & 2 & ":A" & number_of_rows + 1)
Dim i As Integer
Dim group As Range
i = 2
oChart.SetSourceData Source:=trg.Range("A" & i & ":DQ" & i)
For Each cell In objects
    oChart.SeriesCollection(i - 1).XValues = trg.Range("DE2:DQ2")
    oChart.SeriesCollection(i - 1).Values = Union(trg.Range("CE" & i), trg.Range("CG" & i), trg.Range("CI" & i), trg.Range("CK" & i), trg.Range("CM" & i), trg.Range("CO" & i), trg.Range("CQ" & i), trg.Range("CS" & i), trg.Range("CU" & i), trg.Range("CW" & i), trg.Range("CY" & i), trg.Range("DA" & i), trg.Range("DC" & i))
    oChart.SeriesCollection(i - 1).ErrorBar _
Direction:=xlY, Include:=xlErrorBarIncludeBoth, _
 Type:=xlErrorBarTypeCustom, _
 Amount:=Union(trg.Range("CF" & i), trg.Range("CH" & i), trg.Range("CJ" & i), trg.Range("CL" & i), trg.Range("CN" & i), trg.Range("CP" & i), trg.Range("CR" & i), trg.Range("CT" & i), trg.Range("CV" & i), trg.Range("CX" & i), trg.Range("CZ" & i), trg.Range("DB" & i), trg.Range("DD" & i)), _
 MinusValues:=Union(trg.Range("CF" & i), trg.Range("CH" & i), trg.Range("CJ" & i), trg.Range("CL" & i), trg.Range("CN" & i), trg.Range("CP" & i), trg.Range("CR" & i), trg.Range("CT" & i), trg.Range("CV" & i), trg.Range("CX" & i), trg.Range("CZ" & i), trg.Range("DB" & i), trg.Range("DD" & i))
    oChart.SeriesCollection(i - 1).Name = cell.Text
    oChart.SeriesCollection(i - 1).MarkerSize = 10
    oChart.SeriesCollection(i - 1).MarkerStyle = 2
    i = i + 1
    If i = number_of_rows + 2 Then
    Else
        oChart.SeriesCollection.Add Source:=trg.Range("A" & i & ":DQ" & i)
    End If
    
Next cell
Dim LabelRange As Range
Set LabelRange = trg.Range("DE1:DQ1")


With oChart.SeriesCollection(1)
.ApplyDataLabels
        With .DataLabels
            .Format.TextFrame2.TextRange.InsertChartField msoChartFieldRange, LabelRange.Address(External:=True), 0
            .position = xlLabelPositionAbove
            .Font.Size = 13
            .Font.Bold = True
            .ShowCategoryName = False
            .ShowRange = True
            .ShowSeriesName = False
            .ShowValue = False
        End With
    End With
oChart.Axes(xlCategory).MinimumScale = 3.5
oChart.Axes(xlValue).MinimumScale = -18.5
oChart.Axes(xlValue).MaximumScale = -14.5


'oChart.SeriesCollection(2).XValues = trg.Range("DE2:DQ2")
'oChart.SeriesCollection(2).Values = trg.Range("B2:B21")

End Sub

Public Function Extract_string_mag(ByVal txt As String, ByVal char As String) As String
    If InStr(txt, char) > 0 Then
        Extract_string_mag = Split(txt, char)(0)
    Else
        Exctract_string_mag = "None"
    End If
    
End Function

Public Function Extract_string_error(ByVal txt As String, ByVal char As String) As String
    If InStr(txt, char) > 0 Then
        Extract_string_error = Split(txt, char)(1)
    Else
        Exctract_string_error = "None"
    End If
    
End Function

Public Function Extract_comment_error(ByVal txt As String, ByVal char As String) As String
        Extract_comment_error = Left(Split(txt, char)(1), 5)
    
End Function

Sub generate_files()

Dim Path As String
Dim FileNumber As Integer


Dim trg As Worksheet
Set trg = ThisWorkbook.Worksheets("Processed")

Dim number_of_rows As Long
number_of_rows = counter

Dim objects As Range
Dim saved_mag As Range
Dim comments As Range

Set objects = trg.Range("A" & 2 & ":A" & number_of_rows + 1)
Dim i As Integer
Dim j As Integer
Dim limiter As Integer
limiter = 1
i = 2
j = 0

Dim ArrayValues As Object

Set ArrayValues = CreateObject("System.Collections.ArrayList")

ArrayValues.Add "U"
ArrayValues.Add "B"
ArrayValues.Add "V"
ArrayValues.Add "R"
ArrayValues.Add "I"
ArrayValues.Add "J"
ArrayValues.Add "H"
ArrayValues.Add "Ks"
ArrayValues.Add "u'"
ArrayValues.Add "g'"
ArrayValues.Add "r'"
ArrayValues.Add "i'"
ArrayValues.Add "z'"


i = 2
j = 0
z = 1
limiter = 1

i = 2
j = 0
z = 1
limiter = 1
Dim saved_error As Range
For Each cell In objects
        If limiter = 1 Then
        Path = "C:\Users\mikku\Desktop\photo_output\" & cell.Text & ".txt"
        FileNumber = FreeFile
        Open Path For Output As FileNumber
            'Print #FileNumber, "FILTER  MAGNITUDE   ERROR"
            Set saved_mag = Union(trg.Range("B" & i), trg.Range("F" & i), trg.Range("J" & i), trg.Range("N" & i), trg.Range("R" & i), trg.Range("V" & i), trg.Range("Z" & i), trg.Range("AD" & i), trg.Range("AH" & i), trg.Range("AL" & i), trg.Range("AP" & i), trg.Range("AT" & i), trg.Range("AX" & i))
            Set saved_error = Union(trg.Range("C" & i), trg.Range("G" & i), trg.Range("K" & i), trg.Range("O" & i), trg.Range("S" & i), trg.Range("W" & i), trg.Range("AA" & i), trg.Range("AE" & i), trg.Range("AI" & i), trg.Range("AM" & i), trg.Range("AQ" & i), trg.Range("AU" & i), trg.Range("AY" & i))
                For z = 1 To saved_mag.Cells.Count
                        'MsgBox (UCase(trg.Cells(i, (3 * z) + 1).Text))
                        If Not trg.Cells(i, (z) * 4 - 1).value = " " Then
                            If Not trg.Cells(i, (z) * 4 - 1).value = "-" Then
                                Print #FileNumber, ArrayValues(j) & " " & trg.Cells(i, 4 * z - 2).Text & " " & trg.Cells(i, (4 * z) - 1).Text & " " & trg.Cells(i, 4 * z).Text & " " & "*" & trg.Cells(i, (4 * z) + 1).Text
                            End If
                        End If
                    j = j + 1
                Next
                j = 0
                'limiter = 2
        Close FileNumber
        i = i + 1
        End If
Next cell


End Sub
Sub change_error_columns()


Dim trg As Worksheet
Set trg = ThisWorkbook.Worksheets("Processed")

Dim number_of_rows As Long
number_of_rows = counter

Dim objects As Range
Dim saved_mag As Range
Dim comments As Range

Set objects = trg.Range("A" & 2 & ":A" & number_of_rows + 1)
Dim i As Integer
Dim j As Integer
Dim limiter As Integer
limiter = 1
i = 1


For i = 1 To number_of_rows
    For j = 1 To 13
        'MsgBox (trg.Cells(1 + i, j * 4 - 1))
        If Not trg.Cells(1 + i, j * 4 - 1).value = " " Then
            If Not trg.Cells(1 + i, j * 4 - 1).value = "-" Then
                trg.Cells(1 + i, j * 4 - 1) = Extract_comment_error(UCase(trg.Cells(1 + i, (j * 4) + 1)), "AT LEAST")
            End If
        End If
    Next j
Next i



End Sub


'-------------------------------------------------------------------
Sub CombineSheets() ' PÄÄOHJELMA

If WorksheetExists("Processed") Then
Else
Sheets.Add(After:=Sheets("RESULTS")).Name = "Processed"
End If

Sheets("Processed").Cells(1, 1) = "Target"

Sheets("Processed").Cells(1, 2) = "U mag."
Sheets("Processed").Cells(1, 3) = "U mag. err."
Sheets("Processed").Cells(1, 4) = "N"
Sheets("Processed").Cells(1, 5) = "Comments"

Sheets("Processed").Cells(1, 6) = "B mag."
Sheets("Processed").Cells(1, 7) = "B mag. err."
Sheets("Processed").Cells(1, 8) = "N"
Sheets("Processed").Cells(1, 9) = "Comments"

Sheets("Processed").Cells(1, 10) = "V mag."
Sheets("Processed").Cells(1, 11) = "V mag. err."
Sheets("Processed").Cells(1, 12) = "N"
Sheets("Processed").Cells(1, 13) = "Comments"

Sheets("Processed").Cells(1, 14) = "R mag."
Sheets("Processed").Cells(1, 15) = "R mag. err."
Sheets("Processed").Cells(1, 16) = "N"
Sheets("Processed").Cells(1, 17) = "Comments"

Sheets("Processed").Cells(1, 18) = "I mag."
Sheets("Processed").Cells(1, 19) = "I mag. err."
Sheets("Processed").Cells(1, 20) = "N"
Sheets("Processed").Cells(1, 21) = "Comments"

Sheets("Processed").Cells(1, 22) = "J mag."
Sheets("Processed").Cells(1, 23) = "J mag. err."
Sheets("Processed").Cells(1, 24) = "N"
Sheets("Processed").Cells(1, 25) = "Comments"

Sheets("Processed").Cells(1, 26) = "H mag."
Sheets("Processed").Cells(1, 27) = "H mag. err."
Sheets("Processed").Cells(1, 28) = "N"
Sheets("Processed").Cells(1, 29) = "Comments"

Sheets("Processed").Cells(1, 30) = "Ks mag."
Sheets("Processed").Cells(1, 31) = "Ks mag. err."
Sheets("Processed").Cells(1, 32) = "N"
Sheets("Processed").Cells(1, 33) = "Comments"

Sheets("Processed").Cells(1, 34) = "u mag."
Sheets("Processed").Cells(1, 35) = "u mag. err."
Sheets("Processed").Cells(1, 36) = "N"
Sheets("Processed").Cells(1, 37) = "Comments"

Sheets("Processed").Cells(1, 38) = "g mag."
Sheets("Processed").Cells(1, 39) = "g mag. err."
Sheets("Processed").Cells(1, 40) = "N"
Sheets("Processed").Cells(1, 41) = "Comments"

Sheets("Processed").Cells(1, 42) = "r mag."
Sheets("Processed").Cells(1, 43) = "r mag. err."
Sheets("Processed").Cells(1, 44) = "N"
Sheets("Processed").Cells(1, 45) = "Comments"

Sheets("Processed").Cells(1, 46) = "i mag."
Sheets("Processed").Cells(1, 47) = "i mag. err."
Sheets("Processed").Cells(1, 48) = "N"
Sheets("Processed").Cells(1, 49) = "Comments"

Sheets("Processed").Cells(1, 50) = "z mag."
Sheets("Processed").Cells(1, 51) = "z mag. err."
Sheets("Processed").Cells(1, 52) = "N"
Sheets("Processed").Cells(1, 53) = "Comments"

Dim x As Integer
x = 1
For Each Ws In Worksheets
     If Ws.Name = "RESULTS" Or Ws.Name = "TEMPLATE" Or Ws.Name = "Processed" Or Ws.Name = "Former" Then
     Else
     Sheets("Processed").Cells(x + 1, 1) = Ws.Name
     End If
     x = x + 1
Next Ws

Dim trg As Worksheet
Set trg = ThisWorkbook.Worksheets("Processed")

Dim i As Integer
Dim lower As Integer
Dim higher As Integer
Dim comment_high As Integer
Dim mean_value As Range
Dim error_range As Range
Dim commenting As String
Dim Count As Integer
Dim first_comment_test As String

x = 1
For Each Ws In Worksheets
    If Ws.Name = "TEMPLATE" Or Ws.Name = "RESULTS" Or Ws.Name = "Processed" Then
    Else
    For i = 1 To 13
        If i = 1 Then
            Count = Ws.Range("Z120") 'U filter
            Set mean_value = Ws.Range("Z121")
            Set error_range = Ws.Range("Y112:Y121")
            If InStr(1, UCase(Ws.Range("AA113")), "DIM") <> 0 Or InStr(1, UCase(Ws.Range("AA113")), "INVICIBLE") <> 0 Then
                'commenting = ws.Range("AA112") & "; " & ws.Range("AA113") & ";" &  ws.Range("AA117")
                commenting = Ws.Range("AA112") & "; " & Ws.Range("AA113") & ";" & Ws.Range("AA114") & ";" & Ws.Range("AA115") & ";" & Ws.Range("AA116") & ";" & Ws.Range("AA117")
            Else
                'commenting = ws.Range("AA112") & "; " & ws.Range("AA117")
                commenting = Ws.Range("AA112") & "; " & Ws.Range("AA113") & ";" & Ws.Range("AA114") & ";" & Ws.Range("AA115") & ";" & Ws.Range("AA116") & ";" & Ws.Range("AA117")
            End If
            lower = 112
            higher = 121
            comment_high = 117
            first_comment_test = Ws.Range("AA112")
        ElseIf i = 9 Then
            Count = Ws.Range("Z130") 'u filter
            Set mean_value = Ws.Range("Z131")
            Set error_range = Ws.Range("Y122:Y131")
            If InStr(1, UCase(Ws.Range("AA123")), "DIM") <> 0 Or InStr(1, UCase(Ws.Range("AA123")), "INVICIBLE") <> 0 Then
                'commenting = ws.Range("AA122") & "; " & ws.Range("AA123") & ";" & ws.Range("AA127")
                commenting = Ws.Range("AA122") & "; " & Ws.Range("AA123") & ";" & Ws.Range("AA124") & ";" & Ws.Range("AA125") & ";" & Ws.Range("AA126") & ";" & Ws.Range("AA127")
            Else
                'commenting = ws.Range("AA122") & "; " & ws.Range("AA127")
                commenting = Ws.Range("AA122") & "; " & Ws.Range("AA123") & ";" & Ws.Range("AA124") & ";" & Ws.Range("AA125") & ";" & Ws.Range("AA126") & ";" & Ws.Range("AA127")
            End If
            lower = 122
            higher = 131
            comment_high = 127
            first_comment_test = Ws.Range("AA122")
        Else
            If i < 9 Then
                Count = Ws.Range("Z" & 10 * i - 10) 'B filter to Ks filter
                Set mean_value = Ws.Range("Z" & 10 * i - 9)
                lower = 10 * i - 18
                higher = 10 * i - 9
                comment_high = lower + 5
                first_comment_test = Ws.Range("AA" & lower)
            Else
                Count = Ws.Range("Z" & 10 * i - 20) 'g filter to z filter
                Set mean_value = Ws.Range("Z" & 10 * i - 19)
                lower = 10 * i - 28
                higher = 10 * i - 19
                comment_high = lower + 5
                first_comment_test = Ws.Range("AA" & lower)
            End If
            Set error_range = Ws.Range("Y" & lower & ":Y" & higher)
            If InStr(1, UCase(Ws.Range("AA" & lower + 1)), "DIM") <> 0 Or InStr(1, UCase(Ws.Range("AA" & lower + 1)), "INVICIBLE") <> 0 Or InStr(1, UCase(Ws.Range("AA" & lower + 1)), "INVISIBLE") <> 0 Then
                'commenting = ws.Range("AA" & lower) & "; " & ws.Range("AA" & lower + 1) & ";" & ws.Range("AA" & comment_high)
                commenting = Ws.Range("AA" & lower) & "; " & Ws.Range("AA" & lower + 1) & ";" & Ws.Range("AA" & lower + 2) & ";" & Ws.Range("AA" & lower + 3) & ";" & Ws.Range("AA" & lower + 4) & ";" & Ws.Range("AA" & comment_high)
            Else
                'commenting = ws.Range("AA" & lower) & "; " & ws.Range("AA" & comment_high)
                commenting = Ws.Range("AA" & lower) & "; " & Ws.Range("AA" & lower + 1) & ";" & Ws.Range("AA" & lower + 2) & ";" & Ws.Range("AA" & lower + 3) & ";" & Ws.Range("AA" & lower + 4) & ";" & Ws.Range("AA" & comment_high)
            End If
        End If
        'Application.WorksheetFunction.max (ws.Range("Y" & lower, "Y" & higher))
        Dim maximum As Range
        Dim maximum_value As Long
        If InStr(1, first_comment_test, "measurement") <> 0 Then
            comment_only x, i, "-", commenting
        End If
        If IsNumeric(mean_value.value) = True Then
            mean_count_comment mean_value, error_range, Count, x, i, commenting
        End If
        If IsNumeric(mean_value.value) = False And InStr(1, first_comment_test, "measurement") = 0 Then
            comment_only x, i, " ", " "
        End If

    
    Next i
    End If
        
    x = x + 1

Next Ws
LastCol = Split(trg.Cells(1, Columns.Count).End(xlToLeft).Address, "$")(1)
trg.Range(("A2"), (LastCol & 2)).Columns.AutoFit



End Sub
