' Enhanced Dashboard Macro: Adds Slicers, Chart Titles, and Available Funding Calculation
Sub EnhanceDashboard()
    Dim ws As Worksheet
    Dim pvt As PivotTable
    Dim chartObj As ChartObject
    Dim sc As SlicerCache
    Dim i As Integer
    
    Set ws = ThisWorkbook.Sheets("Dashboard")
    
    ' 1. Add Slicers to FundingPivot
    Set pvt = ws.PivotTables("FundingPivot")
    Set sc = ThisWorkbook.SlicerCaches.Add2(pvt, "Month", "MonthSlicer")
    sc.Slicers.Add ws, , "Month", "Month", 650, 10, 100, 200
    Set sc = ThisWorkbook.SlicerCaches.Add2(pvt, "Resource Name", "ResourceNameSlicer")
    sc.Slicers.Add ws, , "Resource Name", "Resource Name", 650, 220, 100, 200
    Set sc = ThisWorkbook.SlicerCaches.Add2(pvt, "Status", "StatusSlicer")
    sc.Slicers.Add ws, , "Status", "Status", 650, 430, 100, 200
    
    ' 2. Add Chart Titles and Format
    For i = 1 To ws.ChartObjects.Count
        Set chartObj = ws.ChartObjects(i)
        Select Case i
            Case 1
                chartObj.Chart.HasTitle = True
                chartObj.Chart.ChartTitle.Text = "Total and Available Funding"
            Case 2
                chartObj.Chart.HasTitle = True
                chartObj.Chart.ChartTitle.Text = "Breakdown by Type"
            Case 3
                chartObj.Chart.HasTitle = True
                chartObj.Chart.ChartTitle.Text = "Headcount by Status"
            Case 4
                chartObj.Chart.HasTitle = True
                chartObj.Chart.ChartTitle.Text = "Projections vs. Actuals per Person"
        End Select
    Next i
    
    ' 3. Add Available Funding as a Calculated Field to FundingPivot
    On Error Resume Next
    pvt.CalculatedFields.Add "Available Funding", "='Allocated Funding'-'Actual Spend'"
    pvt.AddDataField pvt.PivotFields("Available Funding"), "Available Funding", xlSum
    On Error GoTo 0
    
    MsgBox "Enhancements complete!"
End Sub
