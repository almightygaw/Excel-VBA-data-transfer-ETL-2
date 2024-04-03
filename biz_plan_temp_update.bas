Attribute VB_Name = "biz_plan_temp_update"
Sub biz_plan_temp_update()
  ' annual process, create template file to be used for new FY
'  Dim masterFilePath As String
'  masterFilePath = "Q:\FPO Business Development\Business Plans\Templates\FY25 Template - Draft\RecruitPackage (Sample Data) FY25 Manual Backup - MASTER FILE - 3.12.2024.xlsb"
'  Dim masterWb As Workbook
'  Set masterWb = Workbooks.Open(starteFilePath)
  
  Dim masterFeeSchedFilePath As String
  masterFeeSchedFilePath = "Q:\FPO Business Development\Fee Schedules\2024\Master Fee Schedule - 3.12.2024.xlsb"
  Dim masterFeeSched As Workbook
  Set masterFeeSched = Workbooks.Open(masterFeeSchedFilePath)
  
  Dim cptCategoryCrosswalk As Worksheet
  Set cptCategoryCrosswalk = masterFeeSched.Worksheets("CPT Category Crosswalk")
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  '' RVU FILE tab ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  Dim rvuFileWs As Worksheet
  Set rvuFileWs = masterFeeSched.Worksheets("RVU File")

  Dim sourceFilePath As String
  sourceFilePath = "Q:\FPO Business Development\Fee Schedules\2024\Backup\2024 National Physician Fee Schedule Relative Value.xlsx"
  Dim sourceWb As Workbook
  Set sourceWb = Workbooks.Open(sourceFilePath)
  Dim sourceWs As Worksheet
  Set sourceWs = sourceWb.Worksheets("PPRRVU24_V1214")

  Dim filePathString As String
  filePathString = rvuFileWs.Range("R1").Value2
  rvuFileWs.Range("R1").Value2 = Replace(filePathString, LTrim(Year(Date) - 1), LTrim(Str(Year(Date))))

  Dim doNotDeleteString As String
  doNotDeleteString = rvuFileWs.Range("AI1").Value2

  sourceWs.Copy Before:=rvuFileWs

  Application.DisplayAlerts = False
  rvuFileWs.Delete
  Application.DisplayAlerts = True

  With masterFeeSched.Worksheets("PPRRVU24_V1214").Range("R1")
    With .Range("R1")
      .Value2 = filePathString
      .HorizontalAlignment = xlLeft
      .Font.Bold = True
      .Font.Color = vbBlue
    End With

    With .Range("AI1")
      .Value2 = doNotDeleteString
      .HorizontalAlignment = xlLeft
      .Font.Bold = True
      .Font.Color = vbRed
      .Font.Size = 20
    End With
  End With

  masterFeeSched.Sheets("PPRRVU24_V1214").Name = "RVU File"
  Set rvuFileWs = masterFeeSched.Worksheets("RVU File")

  Set rvuFileWs = masterFeeSched.Worksheets("RVU File")

  Columns(1).Insert

  With rvuFileWs.Range("A10")
    .Value2 = "CPT+Modifier"
    .Font.Name = "Arial Narrow"
    .Font.Bold = True
    .Font.Size = 8
  End With

  For Each i In rvuFileWs.Range(Range("B11"), Range("B11").End(xlDown))
    If Not IsEmpty(i.Offset(0, 1)) Then
      i.Offset(0, -1).Formula = i.Value2 & "-" & i.Offset(0, 1).Value2
    Else
      i.Offset(0, -1).Formula = i.Value2
    End If
  Next i

  rvuFileWs.Columns("A").AutoFit

  With rvuFileWs.Range("AG10")
    .Value2 = "CPT Category Crosswalk"
    .Font.Name = "Arial Narrow"
    .Font.Size = 8
    .Font.Bold = True
  End With

  For Each i In rvuFileWs.Range(Range("A11"), Range("A11").End(xlDown))
    i.Offset(0, 32).Value2 = Application.WorksheetFunction.XLookup(i, cptCategoryCrosswalk.Range("A:A"), cptCategoryCrosswalk.Range("J:J"), "")
  Next i

  rvuFileWs.Select
  ActiveWindow.FreezePanes = False

  sourceWb.Close SaveChanges = False
  
  Dim m As Long
  m = 10
  
  For Each i In rvuFileWs.Range(Range("A11"), Range("A11").End(xlDown))
    m = m + 1
    If InStr(i.Value2, "#") > 0 Then
      rvuFileWs.Range("A" & m).EntireRow.Delete
    End If
  Next i
  
  For Each i In rvuFileWs.Range(Range("B11"), Range("B11").End(xlDown))
    i.Offset(0, 31).Formula = WorksheetFunction.XLookup(i.Value2, cptCategoryCrosswalk.Range("A:A"), cptCategoryCrosswalk.Range("J:J"), "")
  Next i
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  '' MEDICARE FEE SCHEDULE tab '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  Dim medicareFeeSched As Worksheet
  Set medicareFeeSched = masterFeeSched.Worksheets("Medicare Fee Schedule")

  Set sourceWb = Workbooks.Open("Q:\FPO Business Development\Fee Schedules\2024\Backup\Medicare\Medicare Fee Schedule data\2024 Medicare Fees.xlsx")
  Set sourceWs = sourceWb.Worksheets("2024 Medicare Fees")

  Dim x As Long
  x = 0

  For Each i In sourceWs.Range(Range("A1"), Range("A1").End(xlDown))
    x = x + 1
  Next i

  medicareFeeSched.Activate
  sourceWs.Range("A1:B" & x).Copy Destination:=medicareFeeSched.Range("A2:B" & x + 1)
  sourceWs.Range("C1:C" & x).Copy Destination:=medicareFeeSched.Range("D2:D" & x + 1)
  sourceWs.Range("L1:L" & x).Copy Destination:=medicareFeeSched.Range("E2:E" & x + 1)
  sourceWs.Range("D1:F" & x).Copy Destination:=medicareFeeSched.Range("F2:H" & x + 1)

  For Each i In medicareFeeSched.Range(Range("A3"), Range("A3").End(xlDown))
    If IsEmpty(i.Offset(0, 1)) Then
      i.Offset(0, 2).Formula = i.Value2
    Else
      i.Offset(0, 2).Formula = i.Value2 & "-" & i.Offset(0, 1).Value2
    End If
  Next i

  For Each i In medicareFeeSched.Range(Range("A3"), Range("A3").End(xlDown))
    i.Offset(0, 8).Value2 = "Medicare Fee Schedule"
    i.Offset(0, 9).Value2 = Application.WorksheetFunction.XLookup(i, cptCategoryCrosswalk.Range("A:A"), cptCategoryCrosswalk.Range("J:J"), "")
  Next i

  sourceWb.Close SaveChanges = False

  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  '' MEDICARE DRUG ASP DATA tab ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  Dim medicareDrugAsp As Worksheet
  Set medicareDrugAsp = masterFeeSched.Worksheets("Medicare Drug ASP Data")

  Set sourceWb = Workbooks.Open("Q:\FPO Business Development\Fee Schedules\2024\Backup\Medicare\January 2024 ASP Pricing File 122023.xlsx")
  Set sourceWs = sourceWb.Worksheets("Jan_24_ASP_byHCPCS")

  Dim medDrugArray As Variant
  medDrugArray = Array("A1:A8")

  For Each i In medDrugArray
    sourceWs.Range(i).UnMerge
    medicareDrugAsp.Range(i).UnMerge
    sourceWs.Range(i).Copy Destination:=medicareDrugAsp.Range(i)
  Next i


  For i = 1 To 8
    medicareDrugAsp.Range("A" & i & ":" & "J" & i).Merge Across = True
    sourceWs.Range("A" & i & ":" & "K" & i).Merge Across = True
  Next i


  x = 0

  For Each i In sourceWs.Range(Range("A9"), Range("A9").End(xlDown))
    x = x + 1
  Next i

  sourceWs.Range("A9:D" & x).Copy Destination:=medicareDrugAsp.Range("A9:D" & x)
  sourceWs.Range("F9:K" & x).Copy Destination:=medicareDrugAsp.Range("E9:J" & x)

  sourceWb.Close SaveChanges = False
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  '' MEDICARE CLIN DIAGNOSTIC LAB tab ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  Dim medicareClinDiag As Worksheet
  Set medicareClinDiag = masterFeeSched.Worksheets("Medicare Clin Diagnostic Lab")

  Set sourceWb = Workbooks.Open("Q:\FPO Business Development\Fee Schedules\2024\Backup\Medicare\2024 Medicare Clinical Diagnostics Code file.xlsx")
  Set sourceWs = sourceWb.Worksheets("CLAB2024Q1")

  sourceWs.Range("A:H").Copy Destination:=medicareClinDiag.Range("A:H")

  sourceWb.Close SaveChanges = False
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  '' MEDICAID FED SCHEDULE tab '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  Dim medicaidFeeSched As Worksheet
  Set medicaidFeeSched = masterFeeSched.Worksheets("Medicaid Fee Schedule")
  
'  medicaidFeeSched.Range(Range("A4"), Range("J4").End(xlDown)).Delete
  medicaidFeeSched.Rows(4 & ":" & medicaidFeeSched.Rows.Count).Delete

  ' Medicine
  Set sourceWb = Workbooks.Open("Q:\FPO Business Development\Fee Schedules\2024\Backup\Medicaid\NYS Medicaid Physician Medicine Services Fee Schedule.xlsx")
  Set sourceWs = sourceWb.Worksheets(1)
  
  Dim effectiveString As String
  effectiveString = sourceWs.Range("A2").Value & " (Medicaid Medicine)"
  
  x = 4
  
  Dim y As Long
  y = 4
  
  For Each i In sourceWs.Range(Range("A4"), Range("A4").End(xlDown))
    If Not InStr(i.Value2, "#") > 0 Then
      y = y + 1
    End If
  Next i
  
  sourceWs.Range("A" & x & ":I" & y).Copy Destination:=medicaidFeeSched.Range("A" & x & ":I" & y)
  medicaidFeeSched.Range("J" & x & ":J" & y).Value2 = "Medicine"
  
  Debug.Print "x = " & x & "     y = " & y
  
  Application.DisplayAlerts = False
  sourceWb.Close SaveChanges = False
  Application.DisplayAlerts = True


  ' Drugs
  Set sourceWb = Workbooks.Open("Q:\FPO Business Development\Fee Schedules\2024\Backup\Medicaid\NYS Medicaid Physician Drug and Drug Administration Services Fee Schedule.xlsx")
  Set sourceWs = sourceWb.Worksheets(1)

  effectiveString = effectiveString & "; " & sourceWs.Range("A2").Value2 & " (Medicaid Drug)"

  x = y + 1

  y = x
  
  Dim header As Long
  header = 3

  For Each i In sourceWs.Range(Range("A4"), Range("A4").End(xlDown))
    If Not InStr(i.Value2, "#") > 0 Then
      y = y + 1
    End If
  Next i

  Debug.Print "x = " & x & "     y = " & y

  sourceWs.Range("A4:F" & (y - x) + header).Copy Destination:=medicaidFeeSched.Range("A" & x & ":F" & y)
  sourceWs.Range("G4:G" & (y - x) + header).Copy Destination:=medicaidFeeSched.Range("I" & x & ":I" & y)
  medicaidFeeSched.Range("J" & x & ":J" & y - 1).Value2 = "Drugs"

  Application.DisplayAlerts = False
  sourceWb.Close SaveChanges = False
  Application.DisplayAlerts = True


  ' Radiology
  Set sourceWb = Workbooks.Open("Q:\FPO Business Development\Fee Schedules\2024\Backup\Medicaid\NYS Medicaid Physician Radiology Services Fee Schedule.xlsx")
  Set sourceWs = sourceWb.Worksheets(1)
  
  effectiveString = effectiveString & "; " & sourceWs.Range("A2").Value2 & " (Medicaid Radiology)"

  sourceWs.Activate
  ActiveSheet.Cells.UnMerge
  
  x = y

  y = x
  
  header = 4
  
  For Each i In sourceWs.Range(Range("A5"), Range("A5").End(xlDown))
    If Not InStr(i.Value2, "#") > 0 Then
      y = y + 1
    End If
  Next i

  Debug.Print "x = " & x & "     y = " & y

  sourceWs.Range("A5:I" & (y - x) + header).Copy Destination:=medicaidFeeSched.Range("A" & x & ":I" & y)
  medicaidFeeSched.Range("J" & x & ":J" & y - 1).Value2 = "Radiology"

  Application.DisplayAlerts = False
  sourceWb.Close SaveChanges = False
  Application.DisplayAlerts = True


  ' Surgery
  Set sourceWb = Workbooks.Open("Q:\FPO Business Development\Fee Schedules\2024\Backup\Medicaid\NYS Medicaid Physician Surgery Services Fee Schedule.xlsx")
  Set sourceWs = sourceWb.Sheets(1)
  
  effectiveString = effectiveString & "; " & sourceWs.Range("A2").Value2 & " (Medicaid Surgery)"

  x = y

  y = x
  
  header = 4

  For Each i In sourceWs.Range(Range("A4"), Range("A4").End(xlDown))
    If Not InStr(i.Value2, "#") > 0 Then
      y = y + 1
    End If
  Next i

  Debug.Print "x = " & x & "     y = " & y
  
  sourceWs.Range("A4:I" & (y - x) + header).Copy Destination:=medicaidFeeSched.Range("A" & x & ":I" & y)
  medicaidFeeSched.Range("J" & x & ":J" & y - 1).Value2 = "Surgery"

  Application.DisplayAlerts = False
  sourceWb.Close SaveChanges = False
  Application.DisplayAlerts = True
  
  ' header, align, formatting
  medicaidFeeSched.Range("A2").Value2 = effectiveString

'  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  ' FEE SCHEDULE tab ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  Dim feeSchedTab As Worksheet
'  Set feeSchedTab = masterFeeSched.Worksheets("Fee Schedule")
  
  
  
  
  
  

End Sub
