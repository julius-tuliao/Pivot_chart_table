Attribute VB_Name = "pivot_modification"
Option Explicit
Sub Cache_fcv()
Worksheets("fcv").Select
Dim lastrow As Long
With ActiveSheet
lastrow = .Range("A" & .Rows.Count).End(xlUp).Row
End With
If lastrow > 1 Then
ThisWorkbook.PivotCaches.Add _
(SourceType:=xlDatabase, _
SourceData:=Worksheets("fcv").Range("A1").CurrentRegion).CreatePivotTable _
TableDestination:="R7C" & Range("A1").CurrentRegion.Columns.Count + 7

'Manipulate pivottables
With Worksheets("fcv").PivotTables(1)
'First row field
With .PivotFields("Full_Name")
.Orientation = xlRowField
.Position = 1
End With

'Legend Filter field
  With .PivotFields("time")
        .Orientation = xlColumnField
        .Position = 1
    End With
    


'Report Filter field
With .PivotFields("agroup")
.Orientation = xlPageField
.Position = 1
End With



'Report Filter field
With .PivotFields("Week")
.Orientation = xlPageField
.Position = 1
End With


'Report Filter field
With .PivotFields("Day")
.Orientation = xlPageField
.Position = 1
End With

'Report Filter field
With .PivotFields("Month")
.Orientation = xlPageField
.Position = 1
End With

'Order Amount or numerical data  in the Values field
.AddDataField Worksheets("fcv").PivotTables(1).PivotFields("Total Count"), _
"Sum of Amount", xlSum
End With


With Worksheets("fcv").PivotTables(1)
.ShowTableStyleRowStripes = True
 .ShowTableStyleColumnStripes = True
    
.TableStyle2 = "PivotStyleMedium4"
End With
 
Worksheets("fcv").Columns("N:AF").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With

Else
End If
End Sub
Sub Cache_bpi()
Worksheets("bpi").Select
Dim lastrow As Long
With ActiveSheet
lastrow = .Range("A" & .Rows.Count).End(xlUp).Row
End With
If lastrow > 1 Then
ThisWorkbook.PivotCaches.Add _
(SourceType:=xlDatabase, _
SourceData:=Worksheets("bpi").Range("A1").CurrentRegion).CreatePivotTable _
TableDestination:="R7C" & Range("A1").CurrentRegion.Columns.Count + 7

'Manipulate pivottables
With Worksheets("bpi").PivotTables(1)
'First row field
With .PivotFields("Full_Name")
.Orientation = xlRowField
.Position = 1
End With

'Legend Filter field
  With .PivotFields("time")
        .Orientation = xlColumnField
        .Position = 1
    End With
    

'Report Filter field
With .PivotFields("agroup")
.Orientation = xlPageField
.Position = 1
End With



'Report Filter field
With .PivotFields("Week")
.Orientation = xlPageField
.Position = 1
End With


'Report Filter field
With .PivotFields("Day")
.Orientation = xlPageField
.Position = 1
End With

'Report Filter field
With .PivotFields("Month")
.Orientation = xlPageField
.Position = 1
End With

'Order Amount or numerical data  in the Values field
.AddDataField Worksheets("bpi").PivotTables(1).PivotFields("Total Count"), _
"Sum of Amount", xlSum
End With


With Worksheets("bpi").PivotTables(1)
.ShowTableStyleRowStripes = True
 .ShowTableStyleColumnStripes = True
    
.TableStyle2 = "PivotStyleMedium4"
End With
 
Worksheets("bpi").Columns("N:AF").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With

Else
End If
End Sub
Sub Cache_hsm()
Worksheets("hsm").Select
Dim lastrow As Long
With ActiveSheet
lastrow = .Range("A" & .Rows.Count).End(xlUp).Row
End With
If lastrow > 1 Then
ThisWorkbook.PivotCaches.Add _
(SourceType:=xlDatabase, _
SourceData:=Worksheets("hsm").Range("A1").CurrentRegion).CreatePivotTable _
TableDestination:="R7C" & Range("A1").CurrentRegion.Columns.Count + 7

'Manipulate pivottables
With Worksheets("hsm").PivotTables(1)
'First row field
With .PivotFields("Full_Name")
.Orientation = xlRowField
.Position = 1
End With

'Report Filter field
With .PivotFields("agroup")
.Orientation = xlPageField
.Position = 1
End With

'Legend Filter field
  With .PivotFields("time")
        .Orientation = xlColumnField
        .Position = 1
    End With
    



'Report Filter field
With .PivotFields("Week")
.Orientation = xlPageField
.Position = 1
End With


'Report Filter field
With .PivotFields("Day")
.Orientation = xlPageField
.Position = 1
End With

'Report Filter field
With .PivotFields("Month")
.Orientation = xlPageField
.Position = 1
End With

'Order Amount or numerical data  in the Values field
.AddDataField Worksheets("hsm").PivotTables(1).PivotFields("Total Count"), _
"Sum of Amount", xlSum
End With


With Worksheets("hsm").PivotTables(1)
.ShowTableStyleRowStripes = True
 .ShowTableStyleColumnStripes = True
    
.TableStyle2 = "PivotStyleMedium4"
End With
 
Worksheets("hsm").Columns("N:AF").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With


Else
End If
End Sub
Sub Cache_mcc()
Worksheets("mcc").Select
Dim lastrow As Long
With ActiveSheet
lastrow = .Range("A" & .Rows.Count).End(xlUp).Row
End With
If lastrow > 1 Then
ThisWorkbook.PivotCaches.Add _
(SourceType:=xlDatabase, _
SourceData:=Worksheets("mcc").Range("A1").CurrentRegion).CreatePivotTable _
TableDestination:="R7C" & Range("A1").CurrentRegion.Columns.Count + 7

'Manipulate pivottables
With Worksheets("mcc").PivotTables(1)
'First row field
With .PivotFields("Full_Name")
.Orientation = xlRowField
.Position = 1
End With

'Report Filter field
With .PivotFields("agroup")
.Orientation = xlPageField
.Position = 1
End With

'Legend Filter field
  With .PivotFields("time")
        .Orientation = xlColumnField
        .Position = 1
    End With
    


'Report Filter field
With .PivotFields("Week")
.Orientation = xlPageField
.Position = 1
End With


'Report Filter field
With .PivotFields("Day")
.Orientation = xlPageField
.Position = 1
End With

'Report Filter field
With .PivotFields("Month")
.Orientation = xlPageField
.Position = 1
End With

'Order Amount or numerical data  in the Values field
.AddDataField Worksheets("mcc").PivotTables(1).PivotFields("Total Count"), _
"Sum of Amount", xlSum
End With

With Worksheets("mcc").PivotTables(1)
.ShowTableStyleRowStripes = True
 .ShowTableStyleColumnStripes = True
    
.TableStyle2 = "PivotStyleMedium4"
End With
 
Worksheets("mcc").Columns("N:AF").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With

Else
End If
End Sub
Sub Cache_pif()
Worksheets("pif").Select
Dim lastrow As Long
With ActiveSheet
lastrow = .Range("A" & .Rows.Count).End(xlUp).Row
End With
If lastrow > 1 Then
ThisWorkbook.PivotCaches.Add _
(SourceType:=xlDatabase, _
SourceData:=Worksheets("pif").Range("A1").CurrentRegion).CreatePivotTable _
TableDestination:="R7C" & Range("A1").CurrentRegion.Columns.Count + 7

'Manipulate pivottables
With Worksheets("pif").PivotTables(1)
'First row field
With .PivotFields("Full_Name")
.Orientation = xlRowField
.Position = 1
End With

'Report Filter field
With .PivotFields("agroup")
.Orientation = xlPageField
.Position = 1
End With


'Legend Filter field
  With .PivotFields("time")
        .Orientation = xlColumnField
        .Position = 1
    End With
    


'Report Filter field
With .PivotFields("Week")
.Orientation = xlPageField
.Position = 1
End With


'Report Filter field
With .PivotFields("Day")
.Orientation = xlPageField
.Position = 1
End With

'Report Filter field
With .PivotFields("Month")
.Orientation = xlPageField
.Position = 1
End With

'Order Amount or numerical data  in the Values field
.AddDataField Worksheets("pif").PivotTables(1).PivotFields("Total Count"), _
"Sum of Amount", xlSum
End With


With Worksheets("pif").PivotTables(1)
.ShowTableStyleRowStripes = True
 .ShowTableStyleColumnStripes = True
    
.TableStyle2 = "PivotStyleMedium4"
End With
 
Worksheets("pif").Columns("N:AF").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With

Else
End If
End Sub

Sub Cache_lks()
Dim lastrow As Long
Worksheets("lks").Select
With ActiveSheet
lastrow = .Range("A" & .Rows.Count).End(xlUp).Row
End With
If lastrow > 1 Then
ThisWorkbook.PivotCaches.Add _
(SourceType:=xlDatabase, _
SourceData:=Worksheets("lks").Range("A1").CurrentRegion).CreatePivotTable _
TableDestination:="R7C" & Range("A1").CurrentRegion.Columns.Count + 7

'Manipulate pivottables
With Worksheets("lks").PivotTables(1)
'First row field
With .PivotFields("Full_Name")
.Orientation = xlRowField
.Position = 1
End With

'Report Filter field
With .PivotFields("agroup")
.Orientation = xlPageField
.Position = 1
End With


'Legend Filter field
  With .PivotFields("time")
        .Orientation = xlColumnField
        .Position = 1
    End With
    


'Report Filter field
With .PivotFields("Week")
.Orientation = xlPageField
.Position = 1
End With


'Report Filter field
With .PivotFields("Day")
.Orientation = xlPageField
.Position = 1
End With

'Report Filter field
With .PivotFields("Month")
.Orientation = xlPageField
.Position = 1
End With

'Order Amount or numerical data  in the Values field
.AddDataField Worksheets("lks").PivotTables(1).PivotFields("Total Count"), _
"Sum of Amount", xlSum
End With


With Worksheets("lks").PivotTables(1)
.ShowTableStyleRowStripes = True
 .ShowTableStyleColumnStripes = True
    
.TableStyle2 = "PivotStyleMedium4"
End With
 
Worksheets("lks").Columns("N:AF").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With


Else
End If
End Sub

Sub Cache_psb()
Worksheets("psb").Select
Dim lastrow As Long
With ActiveSheet
lastrow = .Range("A" & .Rows.Count).End(xlUp).Row
End With
If lastrow > 1 Then
ThisWorkbook.PivotCaches.Add _
(SourceType:=xlDatabase, _
SourceData:=Worksheets("psb").Range("A1").CurrentRegion).CreatePivotTable _
TableDestination:="R7C" & Range("A1").CurrentRegion.Columns.Count + 7

'Manipulate pivottables
With Worksheets("psb").PivotTables(1)
'First row field
With .PivotFields("Full_Name")
.Orientation = xlRowField
.Position = 1
End With

'Report Filter field
With .PivotFields("agroup")
.Orientation = xlPageField
.Position = 1
End With


'Legend Filter field
  With .PivotFields("time")
        .Orientation = xlColumnField
        .Position = 1
    End With
    


'Report Filter field
With .PivotFields("Week")
.Orientation = xlPageField
.Position = 1
End With


'Report Filter field
With .PivotFields("Day")
.Orientation = xlPageField
.Position = 1
End With

'Report Filter field
With .PivotFields("Month")
.Orientation = xlPageField
.Position = 1
End With

'Order Amount or numerical data  in the Values field
.AddDataField Worksheets("psb").PivotTables(1).PivotFields("Total Count"), _
"Sum of Amount", xlSum
End With


With Worksheets("psb").PivotTables(1)
.ShowTableStyleRowStripes = True
 .ShowTableStyleColumnStripes = True
    
.TableStyle2 = "PivotStyleMedium4"
End With
 
Worksheets("psb").Columns("N:AF").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With


Else
End If
End Sub

Sub Cache_ewb()
Worksheets("ewb").Select
Dim lastrow As Long
With ActiveSheet
lastrow = .Range("A" & .Rows.Count).End(xlUp).Row
End With
If lastrow > 1 Then
ThisWorkbook.PivotCaches.Add _
(SourceType:=xlDatabase, _
SourceData:=Worksheets("ewb").Range("A1").CurrentRegion).CreatePivotTable _
TableDestination:="R7C" & Range("A1").CurrentRegion.Columns.Count + 7

'Manipulate pivottables
With Worksheets("ewb").PivotTables(1)
'First row field
With .PivotFields("Full_Name")
.Orientation = xlRowField
.Position = 1
End With

'Report Filter field
With .PivotFields("agroup")
.Orientation = xlPageField
.Position = 1
End With


'Legend Filter field
  With .PivotFields("time")
        .Orientation = xlColumnField
        .Position = 1
    End With
    


'Report Filter field
With .PivotFields("Week")
.Orientation = xlPageField
.Position = 1
End With


'Report Filter field
With .PivotFields("Day")
.Orientation = xlPageField
.Position = 1
End With

'Report Filter field
With .PivotFields("Month")
.Orientation = xlPageField
.Position = 1
End With

'Order Amount or numerical data  in the Values field
.AddDataField Worksheets("ewb").PivotTables(1).PivotFields("Total Count"), _
"Sum of Amount", xlSum
End With


With Worksheets("ewb").PivotTables(1)
.ShowTableStyleRowStripes = True
 .ShowTableStyleColumnStripes = True
    
.TableStyle2 = "PivotStyleMedium4"
End With
 
Worksheets("ewb").Columns("N:AF").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With


Else
End If
End Sub

Sub Cache_bdo()
Worksheets("bdo").Select
Dim lastrow As Long
With ActiveSheet
lastrow = .Range("A" & .Rows.Count).End(xlUp).Row
End With
If lastrow > 1 Then
ThisWorkbook.PivotCaches.Add _
(SourceType:=xlDatabase, _
SourceData:=Worksheets("bdo").Range("A1").CurrentRegion).CreatePivotTable _
TableDestination:="R7C" & Range("A1").CurrentRegion.Columns.Count + 7

'Manipulate pivottables
With Worksheets("bdo").PivotTables(1)
'First row field
With .PivotFields("Full_Name")
.Orientation = xlRowField
.Position = 1
End With

'Report Filter field
With .PivotFields("agroup")
.Orientation = xlPageField
.Position = 1
End With


'Legend Filter field
  With .PivotFields("time")
        .Orientation = xlColumnField
        .Position = 1
    End With
    


'Report Filter field
With .PivotFields("Week")
.Orientation = xlPageField
.Position = 1
End With


'Report Filter field
With .PivotFields("Day")
.Orientation = xlPageField
.Position = 1
End With

'Report Filter field
With .PivotFields("Month")
.Orientation = xlPageField
.Position = 1
End With

'Order Amount or numerical data  in the Values field
.AddDataField Worksheets("bdo").PivotTables(1).PivotFields("Total Count"), _
"Sum of Amount", xlSum
End With


With Worksheets("bdo").PivotTables(1)
.ShowTableStyleRowStripes = True
 .ShowTableStyleColumnStripes = True
    
.TableStyle2 = "PivotStyleMedium4"
End With
 
Worksheets("bdo").Columns("N:AF").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With


Else
End If
End Sub


Sub Display_single_bank_Click()


End Sub

Sub filter_pivot()
Dim myPivotField As PivotField
Dim myPivotItem As PivotItem
Dim pivot As String
Set myPivotField = _
Worksheets("ewb").PivotTables(1).PivotFields(Index:="Full_Name")
For Each myPivotItem In myPivotField.PivotItems
myPivotItem.Visible = True
Next myPivotItem
pivot = InputBox("Agent Name. Input name indicated on the Full_Name Column")

Set myPivotField = _
Worksheets("ewb").PivotTables(1).PivotFields(Index:="Full_Name")
For Each myPivotItem In myPivotField.PivotItems
If myPivotItem.name = pivot Then
myPivotItem.Visible = True
Else
myPivotItem.Visible = False
End If
Next myPivotItem
End Sub


Sub addchart_Click()
Call chart_add_ewb
End Sub

Sub chart_add_ewb()
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
Set myPT = Worksheets("ewb").PivotTables(1)
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

Sub delete_all_charts()
Dim chtObj As ChartObject
For Each chtObj In Worksheets("All Data").ChartObjects
chtObj.Delete
Next

End Sub


Sub DeleteAllPivotTablesInWorkbook()
'Updateby20140618
Dim xWs As Worksheet
Dim xPT As PivotTable
For Each xWs In Application.ActiveWorkbook.Worksheets
    For Each xPT In xWs.PivotTables
        xWs.Range(xPT.TableRange2.Address).Delete Shift:=xlUp
    Next
Next
End Sub


Sub Cache_All_Data()
Worksheets("All Data").Select

Call getaverage_alldataonly

Dim lastrow As Long
With ActiveSheet
lastrow = .Range("A" & .Rows.Count).End(xlUp).Row
End With
If lastrow > 1 Then
ThisWorkbook.PivotCaches.Add _
(SourceType:=xlDatabase, _
SourceData:=Worksheets("All Data").Range("A1").CurrentRegion).CreatePivotTable _
TableDestination:="R8C" & Range("A1").CurrentRegion.Columns.Count + 8




'Manipulate pivottables
With Worksheets("All Data").PivotTables(1)
'First row field
With .PivotFields("agroup")
.Orientation = xlRowField
.Position = 1
End With


'Legend Filter field
  With .PivotFields("time")
        .Orientation = xlColumnField
        .Position = 1
    End With
    

'Report Filter field
With .PivotFields("Full_Name")
.Orientation = xlPageField
.Position = 1
End With



'Report Filter field
With .PivotFields("Week")
.Orientation = xlPageField
.Position = 1
End With


'Report Filter field
With .PivotFields("Day")
.Orientation = xlPageField
.Position = 1
End With

'Report Filter field
With .PivotFields("Month")
.Orientation = xlPageField
.Position = 1
End With

'Order Amount or numerical data  in the Values field
.AddDataField Worksheets("All Data").PivotTables(1).PivotFields("Total Average"), _
"Sum of Amount", xlSum

'fix border centeer

End With

With Worksheets("All Data").PivotTables(1)
.ShowTableStyleRowStripes = True
 .ShowTableStyleColumnStripes = True
    
.TableStyle2 = "PivotStyleMedium4"
End With
 
Worksheets("All Data").Columns("P:AH").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With


      Worksheets("All Data").Range("R3").Value = "FROM:"
   Worksheets("All Data").Range("R4").Value = "TO:"

    Worksheets("All Data").Range("S3:S4").Font.Bold = True
    Selection.Font.Size = 14
  
  
   Worksheets("All Data").Columns("Q:AH").NumberFormat = "0"
    
       Worksheets("All Data").Range("S3").Value = Worksheets("Main").Range("B10").Value
   Worksheets("All Data").Range("S4").Value = Worksheets("Main").Range("E10").Value

    Worksheets("All Data").Range("S3:S4").NumberFormat = "m/d/yyyy"
Else
End If
End Sub
