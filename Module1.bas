Attribute VB_Name = "Module1"
'Attribute VB_Name = "sheetCreateByColumnFilter"
Option Explicit

Sub sheetCreateByColumnFilter()
    Dim sheetSource As Worksheet, sheetTarget As Worksheet, sheetFiltered As Worksheet
    Dim rangeDataSource As Range, rangeFilter As Range, rangeTarget As Range, rng As Range
    Dim counterRows As Long
    Dim counterColumns As Integer, columnFilter As Integer, rowFilter As Long
    Dim sheetName As String, message As String
    
    'Set source sheet or sheet to be copied.
    Set sheetSource = ActiveSheet
    
    'Select the Column for filtering.
top:
        On Error Resume Next
        '---------- Spanish version ----------
        Set rangeTarget = Application.InputBox("Selecciona la columna a filtrar", "Crear hojas filtradas por valores únicos", , , , , , 8)
        '---------- English version ----------
        'Set rangeTarget = Application.InputBox("Select Field Name To Filter", "Range Input", , , , , , 8)
        On Error GoTo 0
        If rangeTarget Is Nothing Then
            Exit Sub
        ElseIf rangeTarget.Columns.Count > 1 Then
            GoTo top
        End If
        
        columnFilter = rangeTarget.Column
        rowFilter = rangeTarget.Row
        
        With Application
            .ScreenUpdating = False
            .DisplayAlerts = False
        End With
        
        On Error GoTo progend
        
        'add filter sheet
        Set sheetFiltered = Sheets.Add
        
        With sheetSource
            .Activate
            .Unprotect Password:=""  'add password if needed
            
            counterRows = .Cells(.Rows.Count, columnFilter).End(xlUp).Row
            counterColumns = .Cells(rowFilter, .Columns.Count).End(xlToLeft).Column
            
            If columnFilter > counterColumns Then
                Err.Raise 65000, "", "FilterCol Setting Is Outside Data Range.", "", 0
            End If
            
            Set rangeDataSource = .Range(.Cells(rowFilter, 1), .Cells(counterRows, counterColumns))
            
            'extract Unique values from FilterCol
            .Range(.Cells(rowFilter, columnFilter), .Cells(counterRows, columnFilter)).AdvancedFilter _
                Action:=xlFilterCopy, CopyToRange:=sheetFiltered.Range("A1"), Unique:=True
            
            counterRows = sheetFiltered.Cells(sheetFiltered.Rows.Count, "A").End(xlUp).Row
            
            'set Criteria
            sheetFiltered.Range("B1").Value = sheetFiltered.Range("A1").Value
            
            For Each rangeFilter In sheetFiltered.Range("A2:A" & counterRows)
            
                'check for blank cell in range
                If Len(rangeFilter.Value) > 0 Then
                
                'add the FilterRange to criteria
                sheetFiltered.Range("B2").Value = rangeFilter.Value
                'ensure tab name limit not exceeded
                sheetName = Trim(Left(rangeFilter.Value, 31))
                
                'check if sheet exists
                On Error Resume Next
                Set sheetTarget = Worksheets(sheetName)
                If sheetTarget Is Nothing Then
                    'if not, add new sheet
                    Set sheetTarget = Sheets.Add(After:=Worksheets(Worksheets.Count))
                    sheetTarget.Name = sheetName
                Else
                    'clear existing data
                    sheetTarget.UsedRange.ClearContents
                End If
                On Error GoTo progend
                'apply filter
                rangeDataSource.AdvancedFilter Action:=xlFilterCopy, CriteriaRange:=sheetFiltered.Range("B1:B2"), _
                    CopyToRange:=sheetTarget.Range("A1"), Unique:=False
                
                End If
                sheetTarget.UsedRange.Columns.AutoFit
                Set sheetTarget = Nothing
            Next
        
        .Select
    End With
    
progend:
        sheetFiltered.Delete
        With Application
            .ScreenUpdating = True: .DisplayAlerts = True
        End With
        
        If Err > 0 Then MsgBox (Error(Err)), 16, "Error"
End Sub


