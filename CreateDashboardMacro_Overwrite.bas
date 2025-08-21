' Consolidated Timesheet Dashboard Macro (Handles Existing Dashboard Sheet)
' This macro creates Pivot Tables, Charts, and Slicers for the consolidated data
Sub CreateDashboard()
    Dim ws As Worksheet
    Dim pvtCache As PivotCache
    Dim pvt As PivotTable
    Dim dataRange As Range
    Dim lastRow As Long
    Dim lastCol As Long
    Dim chartObj As Shape
    Dim sh As Worksheet
    
    ' Delete existing Dashboard sheet if it exists
    Application.DisplayAlerts = False
    For Each sh In ThisWorkbook.Sheets
        If sh.Name = "Dashboard" Then
            sh.Delete
            Exit For
        End If
    Next sh
    Application.DisplayAlerts = True
    
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = "Dashboard"
    
    ' Find data range
    With ThisWorkbook.Sheets(1)
        lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        lastCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        Set dataRange = .Range(.Cells(1, 1), .Cells(lastRow, lastCol))
    End With
    
    ' Create Pivot Cache
    Set pvtCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=dataRange)
    
    ' 1. Total and Available Funding
    Set pvt = ws.PivotTables.Add(PivotCache:=pvtCache, TableDestination:=ws.Range("A3"), TableName:="FundingPivot")
    With pvt
        .PivotFields("Month").Orientation = xlRowField
        .PivotFields("Allocated Funding").Orientation = xlDataField
        .PivotFields("Actual Spend").Orientation = xlDataField
    End With
    Set chartObj = ws.Shapes.AddChart2(251, xlColumnClustered, 300, 10, 300, 200)
    chartObj.Chart.SetSourceData pvt.TableRange1
    
    ' 2. Breakdown by Type
    Set pvt = ws.PivotTables.Add(PivotCache:=pvtCache, TableDestination:=ws.Range("A20"), TableName:="TypePivot")
    With pvt
        .PivotFields("Type").Orientation = xlRowField
        .AddDataField .PivotFields("Resource Name"), "Count of Resource Name", xlCount
    End With
    Set chartObj = ws.Shapes.AddChart2(251, xlPie, 300, 220, 300, 200)
    chartObj.Chart.SetSourceData pvt.TableRange1
    
    ' 3. Headcount by Status
    Set pvt = ws.PivotTables.Add(PivotCache:=pvtCache, TableDestination:=ws.Range("A40"), TableName:="StatusPivot")
    With pvt
        .PivotFields("Status").Orientation = xlRowField
        .AddDataField .PivotFields("Resource Name"), "Count of Resource Name", xlCount
    End With
    Set chartObj = ws.Shapes.AddChart2(251, xlBarClustered, 300, 430, 300, 200)
    chartObj.Chart.SetSourceData pvt.TableRange1
    
    ' 4. Projections vs. Actuals per Person
    Set pvt = ws.PivotTables.Add(PivotCache:=pvtCache, TableDestination:=ws.Range("A60"), TableName:="ProjActualPivot")
    With pvt
        .PivotFields("Resource Name").Orientation = xlRowField
        .PivotFields("Month").Orientation = xlColumnField
        .AddDataField .PivotFields("Projected Hours"), "Sum of Projected Hours", xlSum
        .AddDataField .PivotFields("Hours Worked"), "Sum of Hours Worked", xlSum
    End With
    Set chartObj = ws.Shapes.AddChart2(251, xlColumnClustered, 300, 640, 300, 200)
    chartObj.Chart.SetSourceData pvt.TableRange1
    
    MsgBox "Dashboard created!"
End Sub
