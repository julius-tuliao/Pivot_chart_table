Attribute VB_Name = "create_charts"
Option Explicit

Sub create_chart_bdo_Click()
Call chart_add_bdo
End Sub

Sub chart_add_bdo()
Dim lastrow As Long
'We turn off ScreenUpdating to help our macro run faster
Application.ScreenUpdating = False
'delete chart
Dim chtObj As ChartObject
For Each chtObj In ActiveSheet.ChartObjects
chtObj.Delete
Next

'get data
With ActiveSheet
lastrow = .Range("N" & .Rows.Count).End(xlUp).Row + 5
End With


'we Declare a variable for the PivotTable

Dim myPT As PivotTable
'we set the myPT variable for our first PivotTable using index no. 1
Set myPT = Worksheets("bdo").PivotTables(1)
'Select PivotTable.
myPT.PivotSelect ("")
'Add the chart.
Charts.Add
'Place it on the PivotTable’s worksheet.
ActiveChart.Location Where:=xlLocationAsObject, _
name:=myPT.Parent.name
'Position the PivotChart so its top left corner
'occupies cell H23, a few rows below the PivotTable.
ActiveChart.Parent.Left = Range("n" & lastrow).Left
ActiveChart.Parent.Top = Range("n" & lastrow).Top

 'resize chart
 ActiveChart.Parent.Width = Range("N" & lastrow & ":AF" & lastrow + 28).Width
 ActiveChart.Parent.Height = Range("N" & lastrow & ":AF" & lastrow + 28).Height
    

'add label
    ActiveChart.PlotArea.Select
    ActiveChart.SetElement (msoElementDataLabelCenter)

'Deselect the PivotChart.
Range("A1").Select
'Turn on ScreenUpdating.
Application.ScreenUpdating = True
End Sub

Sub create_chart_psb_Click()
Dim lastrow As Long
'We turn off ScreenUpdating to help our macro run faster
Application.ScreenUpdating = False
'delete chart
Dim chtObj As ChartObject
For Each chtObj In ActiveSheet.ChartObjects
chtObj.Delete
Next
'get data
With ActiveSheet
lastrow = .Range("N" & .Rows.Count).End(xlUp).Row + 5
End With

'we Declare a variable for the PivotTable

Dim myPT As PivotTable
'we set the myPT variable for our first PivotTable using index no. 1
Set myPT = Worksheets("psb").PivotTables(1)
'Select PivotTable.
myPT.PivotSelect ("")
'Add the chart.
Charts.Add
'Place it on the PivotTable’s worksheet.
ActiveChart.Location Where:=xlLocationAsObject, _
name:=myPT.Parent.name
'Position the PivotChart so its top left corner
'occupies cell H23, a few rows below the PivotTable.
ActiveChart.Parent.Left = Range("n" & lastrow).Left
ActiveChart.Parent.Top = Range("n" & lastrow).Top

 'resize chart
 ActiveChart.Parent.Width = Range("N" & lastrow & ":AF" & lastrow + 28).Width
 ActiveChart.Parent.Height = Range("N" & lastrow & ":AF" & lastrow + 28).Height
'add label
    ActiveChart.PlotArea.Select
    ActiveChart.SetElement (msoElementDataLabelCenter)

'Deselect the PivotChart.
Range("A1").Select
'Turn on ScreenUpdating.
Application.ScreenUpdating = True
End Sub

Sub create_chart_lks_Click()
Dim lastrow As Long
'We turn off ScreenUpdating to help our macro run faster
Application.ScreenUpdating = False
'get data
With ActiveSheet
lastrow = .Range("N" & .Rows.Count).End(xlUp).Row + 5
End With
'delete chart
Dim chtObj As ChartObject
For Each chtObj In ActiveSheet.ChartObjects
chtObj.Delete
Next
'we Declare a variable for the PivotTable

Dim myPT As PivotTable
'we set the myPT variable for our first PivotTable using index no. 1
Set myPT = Worksheets("lks").PivotTables(1)
'Select PivotTable.
myPT.PivotSelect ("")
'Add the chart.
Charts.Add
'Place it on the PivotTable’s worksheet.
ActiveChart.Location Where:=xlLocationAsObject, _
name:=myPT.Parent.name
'Position the PivotChart so its top left corner
'occupies cell H23, a few rows below the PivotTable.
ActiveChart.Parent.Left = Range("n" & lastrow).Left
ActiveChart.Parent.Top = Range("n" & lastrow).Top

 'resize chart
 ActiveChart.Parent.Width = Range("N" & lastrow & ":AF" & lastrow + 28).Width
 ActiveChart.Parent.Height = Range("N" & lastrow & ":AF" & lastrow + 28).Height
'add label
    ActiveChart.PlotArea.Select
    ActiveChart.SetElement (msoElementDataLabelCenter)

'Deselect the PivotChart.
Range("A1").Select
'Turn on ScreenUpdating.
Application.ScreenUpdating = True
End Sub
Sub create_chart_pif_Click()
Dim lastrow As Long
'We turn off ScreenUpdating to help our macro run faster
Application.ScreenUpdating = False
'get data
With ActiveSheet
lastrow = .Range("N" & .Rows.Count).End(xlUp).Row + 5
End With
'delete chart
Dim chtObj As ChartObject
For Each chtObj In ActiveSheet.ChartObjects
chtObj.Delete
Next
'we Declare a variable for the PivotTable

Dim myPT As PivotTable
'we set the myPT variable for our first PivotTable using index no. 1
Set myPT = Worksheets("pif").PivotTables(1)
'Select PivotTable.
myPT.PivotSelect ("")
'Add the chart.
Charts.Add
'Place it on the PivotTable’s worksheet.
ActiveChart.Location Where:=xlLocationAsObject, _
name:=myPT.Parent.name
'Position the PivotChart so its top left corner
'occupies cell H23, a few rows below the PivotTable.
ActiveChart.Parent.Left = Range("n" & lastrow).Left
ActiveChart.Parent.Top = Range("n" & lastrow).Top

 'resize chart
 ActiveChart.Parent.Width = Range("N" & lastrow & ":AF" & lastrow + 28).Width
 ActiveChart.Parent.Height = Range("N" & lastrow & ":AF" & lastrow + 28).Height

'add label
    ActiveChart.PlotArea.Select
    ActiveChart.SetElement (msoElementDataLabelCenter)
'Deselect the PivotChart.
Range("A1").Select
'Turn on ScreenUpdating.
Application.ScreenUpdating = True
End Sub
Sub create_chart_mcc_Click()
Dim lastrow As Long
'We turn off ScreenUpdating to help our macro run faster
Application.ScreenUpdating = False
'get data
With ActiveSheet
lastrow = .Range("N" & .Rows.Count).End(xlUp).Row + 5
End With
'delete chart
Dim chtObj As ChartObject
For Each chtObj In ActiveSheet.ChartObjects
chtObj.Delete
Next
'we Declare a variable for the PivotTable

Dim myPT As PivotTable
'we set the myPT variable for our first PivotTable using index no. 1
Set myPT = Worksheets("mcc").PivotTables(1)
'Select PivotTable.
myPT.PivotSelect ("")
'Add the chart.
Charts.Add
'Place it on the PivotTable’s worksheet.
ActiveChart.Location Where:=xlLocationAsObject, _
name:=myPT.Parent.name
'Position the PivotChart so its top left corner
'occupies cell H23, a few rows below the PivotTable.
ActiveChart.Parent.Left = Range("n" & lastrow).Left
ActiveChart.Parent.Top = Range("n" & lastrow).Top

 'resize chart
 ActiveChart.Parent.Width = Range("N" & lastrow & ":AF" & lastrow + 28).Width
 ActiveChart.Parent.Height = Range("N" & lastrow & ":AF" & lastrow + 28).Height

'add label
    ActiveChart.PlotArea.Select
    ActiveChart.SetElement (msoElementDataLabelCenter)

'Deselect the PivotChart.
Range("A1").Select
'Turn on ScreenUpdating.
Application.ScreenUpdating = True
End Sub
Sub create_chart_hsm_Click()
Dim lastrow As Long
'We turn off ScreenUpdating to help our macro run faster
Application.ScreenUpdating = False
'get data
With ActiveSheet
lastrow = .Range("N" & .Rows.Count).End(xlUp).Row + 5
End With
'delete chart
Dim chtObj As ChartObject
For Each chtObj In ActiveSheet.ChartObjects
chtObj.Delete
Next
'we Declare a variable for the PivotTable

Dim myPT As PivotTable
'we set the myPT variable for our first PivotTable using index no. 1
Set myPT = Worksheets("hsm").PivotTables(1)
'Select PivotTable.
myPT.PivotSelect ("")
'Add the chart.
Charts.Add
'Place it on the PivotTable’s worksheet.
ActiveChart.Location Where:=xlLocationAsObject, _
name:=myPT.Parent.name
'Position the PivotChart so its top left corner
'occupies cell H23, a few rows below the PivotTable.
ActiveChart.Parent.Left = Range("n" & lastrow).Left
ActiveChart.Parent.Top = Range("n" & lastrow).Top

 'resize chart
 ActiveChart.Parent.Width = Range("N" & lastrow & ":AF" & lastrow + 28).Width
 ActiveChart.Parent.Height = Range("N" & lastrow & ":AF" & lastrow + 28).Height

'add label
    ActiveChart.PlotArea.Select
    ActiveChart.SetElement (msoElementDataLabelCenter)
'Deselect the PivotChart.
Range("A1").Select
'Turn on ScreenUpdating.
Application.ScreenUpdating = True
End Sub
Sub create_chart_bpi_Click()
Dim lastrow As Long
'We turn off ScreenUpdating to help our macro run faster
Application.ScreenUpdating = False
'get data
With ActiveSheet
lastrow = .Range("N" & .Rows.Count).End(xlUp).Row + 5
End With
'delete chart
Dim chtObj As ChartObject
For Each chtObj In ActiveSheet.ChartObjects
chtObj.Delete
Next
'we Declare a variable for the PivotTable

Dim myPT As PivotTable
'we set the myPT variable for our first PivotTable using index no. 1
Set myPT = Worksheets("bpi").PivotTables(1)
'Select PivotTable.
myPT.PivotSelect ("")
'Add the chart.
Charts.Add
'Place it on the PivotTable’s worksheet.
ActiveChart.Location Where:=xlLocationAsObject, _
name:=myPT.Parent.name
'Position the PivotChart so its top left corner
'occupies cell H23, a few rows below the PivotTable.
ActiveChart.Parent.Left = Range("n" & lastrow).Left
ActiveChart.Parent.Top = Range("n" & lastrow).Top

 'resize chart
 ActiveChart.Parent.Width = Range("N" & lastrow & ":AF" & lastrow + 28).Width
 ActiveChart.Parent.Height = Range("N" & lastrow & ":AF" & lastrow + 28).Height

'add label
    ActiveChart.PlotArea.Select
    ActiveChart.SetElement (msoElementDataLabelCenter)

'Deselect the PivotChart.
Range("A1").Select
'Turn on ScreenUpdating.
Application.ScreenUpdating = True
End Sub
Sub create_chart_fcv_Click()
Dim lastrow As Long
'We turn off ScreenUpdating to help our macro run faster
Application.ScreenUpdating = False
'get data
With ActiveSheet
lastrow = .Range("N" & .Rows.Count).End(xlUp).Row + 5
End With
'delete chart
Dim chtObj As ChartObject
For Each chtObj In ActiveSheet.ChartObjects
chtObj.Delete
Next
'we Declare a variable for the PivotTable

Dim myPT As PivotTable
'we set the myPT variable for our first PivotTable using index no. 1
Set myPT = Worksheets("fcv").PivotTables(1)
'Select PivotTable.
myPT.PivotSelect ("")
'Add the chart.
Charts.Add
'Place it on the PivotTable’s worksheet.
ActiveChart.Location Where:=xlLocationAsObject, _
name:=myPT.Parent.name
'Position the PivotChart so its top left corner
'occupies cell H23, a few rows below the PivotTable.
ActiveChart.Parent.Left = Range("n" & lastrow).Left
ActiveChart.Parent.Top = Range("n" & lastrow).Top

 'resize chart
 ActiveChart.Parent.Width = Range("N" & lastrow & ":AF" & lastrow + 28).Width
 ActiveChart.Parent.Height = Range("N" & lastrow & ":AF" & lastrow + 28).Height

'add label
    ActiveChart.PlotArea.Select
    ActiveChart.SetElement (msoElementDataLabelCenter)
'Deselect the PivotChart.
Range("A1").Select
'Turn on ScreenUpdating.
Application.ScreenUpdating = True
End Sub
Sub create_chart_alldata_Click()
Dim lastrow As Long
'We turn off ScreenUpdating to help our macro run faster
Application.ScreenUpdating = False

'get data
With Worksheets("All Data")
lastrow = .Range("P" & .Rows.Count).End(xlUp).Row + 5
End With

'delete chart
Dim chtObj As ChartObject
For Each chtObj In ActiveSheet.ChartObjects
chtObj.Delete
Next
'we Declare a variable for the PivotTable

Dim myPT As PivotTable
'we set the myPT variable for our first PivotTable using index no. 1
Set myPT = Worksheets("All Data").PivotTables(1)
'Select PivotTable.
myPT.PivotSelect ("")
'Add the chart.
Charts.Add
'Place it on the PivotTable’s worksheet.
ActiveChart.Location Where:=xlLocationAsObject, _
name:=myPT.Parent.name
'Position the PivotChart so its top left corner
'occupies cell H23, a few rows below the PivotTable.
ActiveChart.Parent.Left = Range("p" & lastrow).Left
ActiveChart.Parent.Top = Range("p" & lastrow).Top



'add label
    ActiveChart.PlotArea.Select
    ActiveChart.SetElement (msoElementDataLabelCenter)
    
 'resize chart
 ActiveChart.Parent.Width = Range("p" & lastrow & ":AH" & lastrow + 28).Width
 ActiveChart.Parent.Height = Range("p" & lastrow & ":AH" & lastrow + 28).Height
    
    
'Deselect the PivotChart.
Range("A1").Select
'Turn on ScreenUpdating.
Application.ScreenUpdating = True
End Sub
