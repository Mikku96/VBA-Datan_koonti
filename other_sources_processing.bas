Attribute VB_Name = "Module2"

Sub Former_magnitudes() ' PÄÄOHJELMA

Dim target_name As String
Dim NTT_2023 As Range
Dim NTT_2023_times As Range
Dim PANSTARR As Range
Dim PANSTARR_time As String
Dim SkyMapper As Range
Dim SkyMapper_time As String
Dim SDSS As Range
Dim SDSS_time As String
Dim OTHER As Range
Dim OTHER_time As String
Dim OTHER_info As String
Dim JHK As Range
Dim JHK_time As Range

Dim Path As String
Dim FileNumber As Integer

Dim ArrayValues As Object
Dim array_position As Integer


If WorksheetExists("Historical") Then
Else
Sheets.Add(After:=Sheets("Processed")).Name = "Historical"
End If

Dim trg As Worksheet
Set trg = ThisWorkbook.Worksheets("Historical")

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

Dim ArrayTimes As Object
Set ArrayTimes = CreateObject("System.Collections.ArrayList")
ArrayTimes.Add "-"
ArrayTimes.Add "-"
ArrayTimes.Add "-"
ArrayTimes.Add "-"
ArrayTimes.Add "-"
ArrayTimes.Add "-"
ArrayTimes.Add "-"
ArrayTimes.Add "-"
ArrayTimes.Add "-"
ArrayTimes.Add "-"
ArrayTimes.Add "-"
ArrayTimes.Add "-"
ArrayTimes.Add "-"
ArrayTimes.Add "-"


array_position = 0
Dim i As Integer
i = 1
Dim j_filter As Integer
j_filter = 1

If Not trg.Cells(1, i) = "Target" Then
    
    trg.Cells(1, i) = "Target"
    For Each filter_name In ArrayValues
        trg.Cells(1, i + 1) = filter_name & "_NTT 2018"
        
        trg.Cells(1, i + 2) = filter_name & "_err"
        
        trg.Cells(1, i + 3) = "MJD"
        
        trg.Cells(1, i + 4) = filter_name & "_NTT 2017"
        
        trg.Cells(1, i + 5) = filter_name & "_err"
        
        trg.Cells(1, i + 6) = "MJD"
        
        If filter_name = "J" Or filter_name = "H" Or filter_name = "Ks" Then
            
            GoTo Continue
        Else
        
            GoTo Here
        End If
Here:
            
            trg.Cells(1, i + 7) = filter_name & "_SkyMapper"
            
            trg.Cells(1, i + 8) = filter_name & "_err"
            
            trg.Cells(1, i + 9) = "MJD"
            
            trg.Cells(1, i + 10) = filter_name & "_Pan-Starrs"
            
            trg.Cells(1, i + 11) = filter_name & "_err"
            
            trg.Cells(1, i + 12) = "MJD"
            
            trg.Cells(1, i + 13) = filter_name & "_SDSS"
            
            trg.Cells(1, i + 14) = filter_name & "_err"
            
            trg.Cells(1, i + 15) = "MJD"
        
            trg.Cells(1, i + 16) = filter_name & "_OTHER_(DES)"
            
            trg.Cells(1, i + 17) = filter_name & "_err"
            
            trg.Cells(1, i + 18) = "MJD"
            
            i = i + 19
            
            GoTo Filter_change
        
Continue:
        trg.Cells(1, i + 7) = filter_name & "_2MASS_or_VISTA"
            
        trg.Cells(1, i + 8) = filter_name & "_err"
            
        trg.Cells(1, i + 9) = "MJD"
            
        i = i + 10
Filter_change:
    Next filter_name

End If

i = 1

If IsEmpty(trg.Cells(i + 1, 1)) Then
    For Each Ws In Worksheets
         If Ws.Name = "RESULTS" Or Ws.Name = "TEMPLATE" Or Ws.Name = "Processed" Or Ws.Name = "Historical" Then
         Else
         Sheets("Historical").Cells(i + 1, 1) = Ws.Name
         End If
         i = i + 1
    Next Ws
End If



Dim number_of_targets As Integer
i = 1

'Dim sas As Worksheet
'Set sas = ThisWorkbook.Worksheets("ASASSN-14mc")
'
'Set NTT_2023 = sas.Range("AU16:AZ28")
'MsgBox (WorksheetFunction.Average(NTT_2023(2), NTT_2023(3), NTT_2023(4)))
j = 1
For Each Ws In Worksheets
     If Ws.Name = "RESULTS" Or Ws.Name = "TEMPLATE" Or Ws.Name = "Processed" Or Ws.Name = "Historical" Then
     Else
     target_name = Ws.Name

     Set NTT_2023 = Ws.Range("AU2:AV14")
     Set NTT_2023_times = Ws.Range("AU16:AZ28")
     Set SkyMapper = Ws.Range("AM2:AN14")
     SkyMapper_time = Ws.Range("AM16").Text
     Set PANSTARR = Ws.Range("AO2:AP14")
     PANSTARR_time = Ws.Range("AO16").Text
     Set SDSS = Ws.Range("AQ2:AR14")
     SDSS_time = Ws.Range("AQ16").Text
     Set OTHER = Ws.Range("AS2:AT14")
     OTHER_time = Ws.Range("AS16").Text
     Set JHK_time = Ws.Range("AW6:AW8")
     
  

     For i = 1 To 13
        Select Case NTT_2023_times(6 * i - 5).Text
            Case "U"
               ArrayTimes(1) = WorksheetFunction.Average(NTT_2023_times(6 * i - 4), NTT_2023_times(6 * i - 3), NTT_2023_times(6 * i - 3), NTT_2023_times(6 * i - 2), NTT_2023_times(6 * i - 1))
            Case "B"
               ArrayTimes(2) = WorksheetFunction.Average(NTT_2023_times(6 * i - 4), NTT_2023_times(6 * i - 3), NTT_2023_times(6 * i - 3), NTT_2023_times(6 * i - 2), NTT_2023_times(6 * i - 1))
            Case "V"
               ArrayTimes(3) = WorksheetFunction.Average(NTT_2023_times(6 * i - 4), NTT_2023_times(6 * i - 3), NTT_2023_times(6 * i - 3), NTT_2023_times(6 * i - 2), NTT_2023_times(6 * i - 1))
            Case "R"
               ArrayTimes(4) = WorksheetFunction.Average(NTT_2023_times(6 * i - 4), NTT_2023_times(6 * i - 3), NTT_2023_times(6 * i - 3), NTT_2023_times(6 * i - 2), NTT_2023_times(6 * i - 1))
            Case "I"
               ArrayTimes(5) = WorksheetFunction.Average(NTT_2023_times(6 * i - 4), NTT_2023_times(6 * i - 3), NTT_2023_times(6 * i - 3), NTT_2023_times(6 * i - 2), NTT_2023_times(6 * i - 1))
            Case "J"
               ArrayTimes(6) = WorksheetFunction.Average(NTT_2023_times(6 * i - 4), NTT_2023_times(6 * i - 3), NTT_2023_times(6 * i - 3), NTT_2023_times(6 * i - 2), NTT_2023_times(6 * i - 1))
            Case "H"
               ArrayTimes(7) = WorksheetFunction.Average(NTT_2023_times(6 * i - 4), NTT_2023_times(6 * i - 3), NTT_2023_times(6 * i - 3), NTT_2023_times(6 * i - 2), NTT_2023_times(6 * i - 1))
            Case "Ks"
               ArrayTimes(8) = WorksheetFunction.Average(NTT_2023_times(6 * i - 4), NTT_2023_times(6 * i - 3), NTT_2023_times(6 * i - 3), NTT_2023_times(6 * i - 2), NTT_2023_times(6 * i - 1))
            Case "u"
               ArrayTimes(9) = WorksheetFunction.Average(NTT_2023_times(6 * i - 4), NTT_2023_times(6 * i - 3), NTT_2023_times(6 * i - 3), NTT_2023_times(6 * i - 2), NTT_2023_times(6 * i - 1))
            Case "g"
               ArrayTimes(10) = WorksheetFunction.Average(NTT_2023_times(6 * i - 4), NTT_2023_times(6 * i - 3), NTT_2023_times(6 * i - 3), NTT_2023_times(6 * i - 2), NTT_2023_times(6 * i - 1))
            Case "r"
               ArrayTimes(11) = WorksheetFunction.Average(NTT_2023_times(6 * i - 4), NTT_2023_times(6 * i - 3), NTT_2023_times(6 * i - 3), NTT_2023_times(6 * i - 2), NTT_2023_times(6 * i - 1))
            Case "i"
               ArrayTimes(12) = WorksheetFunction.Average(NTT_2023_times(6 * i - 4), NTT_2023_times(6 * i - 3), NTT_2023_times(6 * i - 3), NTT_2023_times(6 * i - 2), NTT_2023_times(6 * i - 1))
            Case "z"
               ArrayTimes(13) = WorksheetFunction.Average(NTT_2023_times(6 * i - 4), NTT_2023_times(6 * i - 3), NTT_2023_times(6 * i - 3), NTT_2023_times(6 * i - 2), NTT_2023_times(6 * i - 1))
            End Select
        Next i

     For i = 1 To 13
    
     Select Case i
        Case 1 ' U
            If Not IsError(NTT_2023(23)) Then
            
                trg.Cells(j + 1, 2) = Application.WorksheetFunction.Round(NTT_2023(23), 2)
                trg.Cells(j + 1, 3) = NTT_2023(24).Text
                trg.Cells(j + 1, 4) = ArrayTimes(i)
                
            Else
                trg.Cells(j + 1, 2) = "-"
                trg.Cells(j + 1, 3) = "-"
                trg.Cells(j + 1, 4) = "-"
            End If
            
            If Not IsError(SkyMapper(23)) Then
                trg.Cells(j + 1, 8) = Application.WorksheetFunction.Round(SkyMapper(23), 2)
                If (Application.WorksheetFunction.Round(SkyMapper(24), 2) = 0) Then
                    trg.Cells(j + 1, 9) = Application.WorksheetFunction.RoundUp(SkyMapper(24), 2)
                Else
                    trg.Cells(j + 1, 9) = Application.WorksheetFunction.Round(SkyMapper(24), 2)
                End If
                trg.Cells(j + 1, 10) = SkyMapper_time
            Else
                trg.Cells(j + 1, 8) = "-"
                trg.Cells(j + 1, 9) = "-"
                trg.Cells(j + 1, 10) = "-"
            End If
            
            If Not IsError(PANSTARR(23)) Then
                trg.Cells(j + 1, 11) = Application.WorksheetFunction.Round(PANSTARR(23), 2)
                If Application.WorksheetFunction.Round(PANSTARR(24), 2) = 0 Then
                    trg.Cells(j + 1, 12) = Application.WorksheetFunction.RoundUp(PANSTARR(24), 2)
                Else
                    trg.Cells(j + 1, 12) = Application.WorksheetFunction.Round(PANSTARR(24), 2)
                End If
                trg.Cells(j + 1, 13) = PANSTARR_time
            Else
                trg.Cells(j + 1, 11) = "-"
                trg.Cells(j + 1, 12) = "-"
                trg.Cells(j + 1, 13) = "-"
                
            End If
            
            If Not IsError(SDSS(23)) Then
                trg.Cells(j + 1, 14) = Application.WorksheetFunction.Round(SDSS(23), 2)
                If Application.WorksheetFunction.Round(SDSS(24), 2) = 0 Then
                    trg.Cells(j + 1, 15) = Application.WorksheetFunction.RoundUp(SDSS(24), 2)
                Else
                    trg.Cells(j + 1, 15) = Application.WorksheetFunction.Round(SDSS(24), 2)
                End If
                trg.Cells(j + 1, 16) = SDSS_time
            Else
                trg.Cells(j + 1, 14) = "-"
                trg.Cells(j + 1, 15) = "-"
                trg.Cells(j + 1, 16) = "-"
            
            End If
            
            If Not IsError(OTHER(23)) Then
                trg.Cells(j + 1, 17) = Application.WorksheetFunction.Round(OTHER(23), 2)
                If Application.WorksheetFunction.Round(OTHER(24), 2) = 0 Then
                    trg.Cells(j + 1, 18) = Application.WorksheetFunction.RoundUp(OTHER(24), 2)
                Else
                    trg.Cells(j + 1, 18) = Application.WorksheetFunction.Round(OTHER(24), 2)
                End If
                trg.Cells(j + 1, 19) = OTHER_time
            Else
                trg.Cells(j + 1, 17) = "-"
                trg.Cells(j + 1, 18) = "-"
                trg.Cells(j + 1, 19) = "-"
            End If
            
            
            
        Case 2 To 5 ' B->I
            If Not IsError(NTT_2023(2 * i - 3)) Then
                trg.Cells(j + 1, (i - 1) * 19 + 2) = Application.WorksheetFunction.Round(NTT_2023(2 * i - 3), 2)
                trg.Cells(j + 1, (i - 1) * 19 + 3) = NTT_2023(2 * i - 2).Text
                trg.Cells(j + 1, (i - 1) * 19 + 4) = ArrayTimes(i)
            Else
                trg.Cells(j + 1, (i - 1) * 19 + 2) = "-"
                trg.Cells(j + 1, (i - 1) * 19 + 3) = "-"
                trg.Cells(j + 1, (i - 1) * 19 + 4) = "-"
            End If
            
            If Not IsError(SkyMapper(2 * i - 3)) Then
                trg.Cells(j + 1, (i - 1) * 19 + 8) = Application.WorksheetFunction.Round(SkyMapper(2 * i - 3), 2)
                If Application.WorksheetFunction.Round(SkyMapper(2 * i - 2), 2) = 0 Then
                    trg.Cells(j + 1, (i - 1) * 19 + 9) = Application.WorksheetFunction.RoundUp(SkyMapper(2 * i - 2), 2)
                Else
                    trg.Cells(j + 1, (i - 1) * 19 + 9) = Application.WorksheetFunction.Round(SkyMapper(2 * i - 2), 2)
                End If
                trg.Cells(j + 1, (i - 1) * 19 + 10) = SkyMapper_time
            Else
                trg.Cells(j + 1, (i - 1) * 19 + 8) = "-"
                trg.Cells(j + 1, (i - 1) * 19 + 9) = "-"
                trg.Cells(j + 1, (i - 1) * 19 + 10) = "-"
            End If
            
            If Not IsError(PANSTARR(2 * i - 3)) Then
                trg.Cells(j + 1, (i - 1) * 19 + 11) = Application.WorksheetFunction.Round(PANSTARR(2 * i - 3), 2)
                If Application.WorksheetFunction.Round(PANSTARR(2 * i - 2), 2) = 0 Then
                    trg.Cells(j + 1, (i - 1) * 19 + 12) = Application.WorksheetFunction.RoundUp(PANSTARR(2 * i - 2), 2)
                Else
                    trg.Cells(j + 1, (i - 1) * 19 + 12) = Application.WorksheetFunction.Round(PANSTARR(2 * i - 2), 2)
                End If
                trg.Cells(j + 1, (i - 1) * 19 + 13) = PANSTARR_time
            Else
                trg.Cells(j + 1, (i - 1) * 19 + 11) = "-"
                trg.Cells(j + 1, (i - 1) * 19 + 12) = "-"
                trg.Cells(j + 1, (i - 1) * 19 + 13) = "-"
            End If
            
            If Not IsError(SDSS(2 * i - 3)) Then
                trg.Cells(j + 1, (i - 1) * 19 + 14) = Application.WorksheetFunction.Round(SDSS(2 * i - 3), 2)
                If Application.WorksheetFunction.Round(SDSS(2 * i - 2), 2) = 0 Then
                    trg.Cells(j + 1, (i - 1) * 19 + 15) = Application.WorksheetFunction.RoundUp(SDSS(2 * i - 2), 2)
                Else
                     trg.Cells(j + 1, (i - 1) * 19 + 15) = Application.WorksheetFunction.Round(SDSS(2 * i - 2), 2)
                End If
                trg.Cells(j + 1, (i - 1) * 19 + 16) = SDSS_time
            Else
                trg.Cells(j + 1, (i - 1) * 19 + 14) = "-"
                trg.Cells(j + 1, (i - 1) * 19 + 15) = "-"
                trg.Cells(j + 1, (i - 1) * 19 + 16) = "-"
            End If
            
            If Not IsError(OTHER(2 * i - 3)) Then
                trg.Cells(j + 1, (i - 1) * 19 + 17) = Application.WorksheetFunction.Round(OTHER(2 * i - 3), 2)
                If Application.WorksheetFunction.Round(OTHER(2 * i - 2), 2) = 0 Then
                    trg.Cells(j + 1, (i - 1) * 19 + 18) = Application.WorksheetFunction.RoundUp(OTHER(2 * i - 2), 2)
                Else
                    trg.Cells(j + 1, (i - 1) * 19 + 18) = Application.WorksheetFunction.Round(OTHER(2 * i - 2), 2)
                End If
                trg.Cells(j + 1, (i - 1) * 19 + 19) = OTHER_time
            Else
                trg.Cells(j + 1, (i - 1) * 19 + 17) = "-"
                trg.Cells(j + 1, (i - 1) * 19 + 18) = "-"
                trg.Cells(j + 1, (i - 1) * 19 + 19) = "-"
            End If
          
            
        Case 6 To 8 ' J -> Ks
            If Not IsError(NTT_2023(2 * i - 3)) Then
            
                trg.Cells(j + 1, 37 + i * 10) = Application.WorksheetFunction.Round(NTT_2023(2 * i - 3), 2)
                trg.Cells(j + 1, 38 + i * 10) = NTT_2023(2 * i - 2).Text
                trg.Cells(j + 1, 39 + i * 10) = ArrayTimes(i)
            Else
                trg.Cells(j + 1, 37 + i * 10) = "-"
                trg.Cells(j + 1, 38 + i * 10) = "-"
                trg.Cells(j + 1, 39 + i * 10) = "-"
            End If
            
            If Not IsError(OTHER(2 * i - 3)) Then
            
                trg.Cells(j + 1, 43 + i * 10) = Application.WorksheetFunction.Round(OTHER(2 * i - 3), 2)
                If Application.WorksheetFunction.Round(OTHER(2 * i - 2), 2) = 0 Then
                    trg.Cells(j + 1, 44 + i * 10) = Application.WorksheetFunction.RoundUp(OTHER(2 * i - 2), 2)
                Else
                    trg.Cells(j + 1, 44 + i * 10) = Application.WorksheetFunction.Round(OTHER(2 * i - 2), 2)
                End If
                trg.Cells(j + 1, 45 + i * 10) = JHK_time(i - 5)
            Else
                trg.Cells(j + 1, 43 + i * 10) = "-"
                trg.Cells(j + 1, 44 + i * 10) = "-"
                trg.Cells(j + 1, 45 + i * 10) = "-"
            End If

        Case 9 ' u'
         If Not IsError(NTT_2023(25)) Then
            trg.Cells(j + 1, 127) = Application.WorksheetFunction.Round(NTT_2023(25), 2)
            trg.Cells(j + 1, 128) = NTT_2023(26).Text
            trg.Cells(j + 1, 129) = ArrayTimes(i)
        Else
            trg.Cells(j + 1, 127) = "-"
            trg.Cells(j + 1, 128) = "-"
            trg.Cells(j + 1, 129) = "-"
        End If
        
        If Not IsError(SkyMapper(25)) And Not SkyMapper(25) = "-" Then
            trg.Cells(j + 1, 133) = Application.WorksheetFunction.Round(SkyMapper(25), 2)
            If Application.WorksheetFunction.Round(SkyMapper(26), 2) = 0 Then
                trg.Cells(j + 1, 134) = Application.WorksheetFunction.RoundUp(SkyMapper(26), 2)
            Else
                trg.Cells(j + 1, 134) = Application.WorksheetFunction.Round(SkyMapper(26), 2)
            End If
            trg.Cells(j + 1, 135) = SkyMapper_time
        Else
            trg.Cells(j + 1, 133) = "-"
            trg.Cells(j + 1, 134) = "-"
            trg.Cells(j + 1, 135) = "-"
        End If
        
        If Not IsError(PANSTARR(25)) And Not PANSTARR(25) = "-" Then
            trg.Cells(j + 1, 136) = Application.WorksheetFunction.Round(PANSTARR(25), 2)
            If Application.WorksheetFunction.Round(PANSTARR(26), 2) = 0 Then
                trg.Cells(j + 1, 137) = Application.WorksheetFunction.RoundUp(PANSTARR(26), 2)
            Else
                 trg.Cells(j + 1, 137) = Application.WorksheetFunction.Round(PANSTARR(26), 2)
            End If
            trg.Cells(j + 1, 138) = PANSTARR_time
        Else
            trg.Cells(j + 1, 136) = "-"
            trg.Cells(j + 1, 137) = "-"
            trg.Cells(j + 1, 138) = "-"
        End If
        
        If Not IsError(SDSS(25)) And Not SDSS(25) = "-" Then
            trg.Cells(j + 1, 139) = Application.WorksheetFunction.Round(SDSS(25), 2)
            If Application.WorksheetFunction.Round(SDSS(26), 2) = 0 Then
                trg.Cells(j + 1, 140) = Application.WorksheetFunction.RoundUp(SDSS(26), 2)
            Else
                trg.Cells(j + 1, 140) = Application.WorksheetFunction.Round(SDSS(26), 2)
            End If
            trg.Cells(j + 1, 141) = SDSS_time
        Else
            trg.Cells(j + 1, 139) = "-"
            trg.Cells(j + 1, 140) = "-"
            trg.Cells(j + 1, 141) = "-"
        End If
        
        If Not IsError(OTHER(25)) And Not OTHER(25) = "-" Then
            trg.Cells(j + 1, 142) = Application.WorksheetFunction.Round(OTHER(25), 2)
            If Application.WorksheetFunction.Round(OTHER(26), 2) = 0 Then
                trg.Cells(j + 1, 143) = Application.WorksheetFunction.RoundUp(OTHER(26), 2)
            Else
                trg.Cells(j + 1, 143) = Application.WorksheetFunction.Round(OTHER(26), 2)
            End If
            trg.Cells(j + 1, 144) = OTHER_time
        Else
            trg.Cells(j + 1, 142) = "-"
            trg.Cells(j + 1, 143) = "-"
            trg.Cells(j + 1, 144) = "-"
        End If
        
 
        Case 10 To 13 ' g' -> r'
         If Not IsError(NTT_2023(2 * (i - 1) - 3)) Then
                trg.Cells(j + 1, -25 + 19 * (i - 1)) = Application.WorksheetFunction.Round(NTT_2023(2 * (i - 1) - 3), 2)
                trg.Cells(j + 1, -24 + 19 * (i - 1)) = NTT_2023(2 * (i - 1) - 2).Text
                trg.Cells(j + 1, -23 + 19 * (i - 1)) = ArrayTimes(i)
            
            Else
                trg.Cells(j + 1, -25 + 19 * (i - 1)) = "-"
                trg.Cells(j + 1, -24 + 19 * (i - 1)) = "-"
                trg.Cells(j + 1, -23 + 19 * (i - 1)) = "-"
         End If
         
         
          If Not IsError(SkyMapper(2 * (i - 1) - 3)) And Not SkyMapper(2 * (i - 1) - 3) = "-" Then
                trg.Cells(j + 1, -19 + 19 * (i - 1)) = Application.WorksheetFunction.Round(SkyMapper(2 * (i - 1) - 3), 2)
                If Application.WorksheetFunction.Round(SkyMapper(2 * (i - 1) - 2), 2) = 0 Then
                    trg.Cells(j + 1, -18 + 19 * (i - 1)) = Application.WorksheetFunction.RoundUp(SkyMapper(2 * (i - 1) - 2), 2)
                Else
                    trg.Cells(j + 1, -18 + 19 * (i - 1)) = Application.WorksheetFunction.Round(SkyMapper(2 * (i - 1) - 2), 2)
                End If
                trg.Cells(j + 1, -17 + 19 * (i - 1)) = SkyMapper_time
            
            Else
                trg.Cells(j + 1, -19 + 19 * (i - 1)) = "-"
                trg.Cells(j + 1, -18 + 19 * (i - 1)) = "-"
                trg.Cells(j + 1, -17 + 19 * (i - 1)) = "-"
         End If
         
          If Not IsError(PANSTARR(2 * (i - 1) - 3)) And Not PANSTARR(2 * (i - 1) - 3) = "-" Then
                trg.Cells(j + 1, -16 + 19 * (i - 1)) = Application.WorksheetFunction.Round(PANSTARR(2 * (i - 1) - 3), 2)
                If Application.WorksheetFunction.Round(PANSTARR(2 * (i - 1) - 2), 2) = 0 Then
                    trg.Cells(j + 1, -15 + 19 * (i - 1)) = Application.WorksheetFunction.RoundUp(PANSTARR(2 * (i - 1) - 2), 2)
                Else
                    trg.Cells(j + 1, -15 + 19 * (i - 1)) = Application.WorksheetFunction.Round(PANSTARR(2 * (i - 1) - 2), 2)
                End If
                trg.Cells(j + 1, -14 + 19 * (i - 1)) = PANSTARR_time
            
            Else
                trg.Cells(j + 1, -16 + 19 * (i - 1)) = "-"
                trg.Cells(j + 1, -15 + 19 * (i - 1)) = "-"
                trg.Cells(j + 1, -14 + 19 * (i - 1)) = "-"
         End If
         
          If Not IsError(SDSS(2 * (i - 1) - 3)) And Not SDSS(2 * (i - 1) - 3) = "-" Then
                trg.Cells(j + 1, -13 + 19 * (i - 1)) = Application.WorksheetFunction.Round(SDSS(2 * (i - 1) - 3), 2)
                If Application.WorksheetFunction.Round(SDSS(2 * (i - 1) - 2), 2) = 0 Then
                    trg.Cells(j + 1, -12 + 19 * (i - 1)) = Application.WorksheetFunction.RoundUp(SDSS(2 * (i - 1) - 2), 2)
                Else
                    trg.Cells(j + 1, -12 + 19 * (i - 1)) = Application.WorksheetFunction.Round(SDSS(2 * (i - 1) - 2), 2)
                End If
                trg.Cells(j + 1, -11 + 19 * (i - 1)) = SDSS_time
            
            Else
                trg.Cells(j + 1, -13 + 19 * (i - 1)) = "-"
                trg.Cells(j + 1, -12 + 19 * (i - 1)) = "-"
                trg.Cells(j + 1, -11 + 19 * (i - 1)) = "-"
         End If
         
          If Not IsError(OTHER(2 * (i - 1) - 3)) And Not OTHER(2 * (i - 1) - 3) = "-" Then
                trg.Cells(j + 1, -10 + 19 * (i - 1)) = Application.WorksheetFunction.Round(OTHER(2 * (i - 1) - 3), 2)
                If Application.WorksheetFunction.Round(OTHER(2 * (i - 1) - 2), 2) = 0 Then
                    trg.Cells(j + 1, -9 + 19 * (i - 1)) = Application.WorksheetFunction.RoundUp(OTHER(2 * (i - 1) - 2), 2)
                Else
                     trg.Cells(j + 1, -9 + 19 * (i - 1)) = Application.WorksheetFunction.Round(OTHER(2 * (i - 1) - 2), 2)
                End If
                trg.Cells(j + 1, -8 + 19 * (i - 1)) = OTHER_time
            
            Else
                trg.Cells(j + 1, -10 + 19 * (i - 1)) = "-"
                trg.Cells(j + 1, -9 + 19 * (i - 1)) = "-"
                trg.Cells(j + 1, -8 + 19 * (i - 1)) = "-"
         End If
         
         
         
         
         
         
     End Select
     Next i



    End If
    j = j + 1

For i = 1 To 13
    ArrayTimes(i) = "-"
Next i


Next Ws



End Sub



