Attribute VB_Name = "PDF_CREATOR"
Option Explicit


Sub Create_PDF_ALL_Click()
Dim lastrow As Long
 Application.ScreenUpdating = False
 Dim path As String
 'get data
With Worksheets("All Data")
lastrow = .Range("P" & .Rows.Count).End(xlUp).Row + 5
End With
MsgBox (lastrow)
 path = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\"
   ActiveSheet.Range("P1:AH" & lastrow + 28).ExportAsFixedFormat Type:=xlTypePDF, _
            Filename:=path & "All DATA " & Format(Now(), "mmddyyyy") & ".PDF", _
            OpenAfterPublish:=False
            Application.ScreenUpdating = True
End Sub

Sub PDF_CREATOR_BDO_Click()
Dim lastrow As Long
 Application.ScreenUpdating = False
 Dim path As String
 'get data
With ActiveSheet
lastrow = .Range("N" & .Rows.Count).End(xlUp).Row + 5
End With

 path = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\"
   ActiveSheet.Range("N1:AF" & lastrow + 28).ExportAsFixedFormat Type:=xlTypePDF, _
            Filename:=path & "BDO " & Format(Now(), "mmddyyyy") & ".PDF", _
            OpenAfterPublish:=False
            Application.ScreenUpdating = True
End Sub
Sub create_PDF_psb_Click()
Dim lastrow As Long
 Application.ScreenUpdating = False
 Dim path As String
 'get data
With ActiveSheet
lastrow = .Range("N" & .Rows.Count).End(xlUp).Row + 5
End With

 path = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\"
   ActiveSheet.Range("N1:AF" & lastrow + 28).ExportAsFixedFormat Type:=xlTypePDF, _
            Filename:=path & "PSB " & Format(Now(), "mmddyyyy") & ".PDF", _
            OpenAfterPublish:=False
            Application.ScreenUpdating = True
End Sub
Sub pdf_lks_Click()
Dim lastrow As Long
 Application.ScreenUpdating = False
 Dim path As String
 'get data
With ActiveSheet
lastrow = .Range("N" & .Rows.Count).End(xlUp).Row + 5
End With

 path = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\"
   ActiveSheet.Range("N1:AF" & lastrow + 28).ExportAsFixedFormat Type:=xlTypePDF, _
            Filename:=path & "LKS " & Format(Now(), "mmddyyyy") & ".PDF", _
            OpenAfterPublish:=False
            Application.ScreenUpdating = True
End Sub
Sub PDF_CREATOR_PIF_Click()
Dim lastrow As Long
 Application.ScreenUpdating = False
 Dim path As String
 'get data
With ActiveSheet
lastrow = .Range("N" & .Rows.Count).End(xlUp).Row + 5
End With

 path = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\"
   ActiveSheet.Range("N1:AF" & lastrow + 28).ExportAsFixedFormat Type:=xlTypePDF, _
            Filename:=path & "PIF " & Format(Now(), "mmddyyyy") & ".PDF", _
            OpenAfterPublish:=False
            Application.ScreenUpdating = True
End Sub
Sub PDF_CREATOR_MCC_Click()
Dim lastrow As Long
 Application.ScreenUpdating = False
 Dim path As String
 'get data
With ActiveSheet
lastrow = .Range("N" & .Rows.Count).End(xlUp).Row + 5
End With

 path = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\"
   ActiveSheet.Range("N1:AF" & lastrow + 28).ExportAsFixedFormat Type:=xlTypePDF, _
            Filename:=path & "MCC " & Format(Now(), "mmddyyyy") & ".PDF", _
            OpenAfterPublish:=False
            Application.ScreenUpdating = True
End Sub
Sub PDF_HSM_Click()
Dim lastrow As Long
 Application.ScreenUpdating = False
 Dim path As String
 'get data
With ActiveSheet
lastrow = .Range("N" & .Rows.Count).End(xlUp).Row + 5
End With

 path = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\"
   ActiveSheet.Range("N1:AF" & lastrow + 28).ExportAsFixedFormat Type:=xlTypePDF, _
            Filename:=path & "HSBC " & Format(Now(), "mmddyyyy") & ".PDF", _
            OpenAfterPublish:=False
            Application.ScreenUpdating = True
End Sub
Sub Create_Pdf_ewb_Click()
Dim lastrow As Long
 Application.ScreenUpdating = False
 Dim path As String
 'get data
With ActiveSheet
lastrow = .Range("N" & .Rows.Count).End(xlUp).Row + 5
End With

 path = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\"
   ActiveSheet.Range("N1:AF" & lastrow + 28).ExportAsFixedFormat Type:=xlTypePDF, _
            Filename:=path & "EWB " & Format(Now(), "mmddyyyy") & ".PDF", _
            OpenAfterPublish:=False
            Application.ScreenUpdating = True
End Sub
Sub PDF_BPI_CREAT_Click()
Dim lastrow As Long
 Application.ScreenUpdating = False
 Dim path As String
 'get data
With ActiveSheet
lastrow = .Range("N" & .Rows.Count).End(xlUp).Row + 5
End With

 path = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\"
   ActiveSheet.Range("N1:AF" & lastrow + 28).ExportAsFixedFormat Type:=xlTypePDF, _
            Filename:=path & "BPI " & Format(Now(), "mmddyyyy") & ".PDF", _
            OpenAfterPublish:=False
            Application.ScreenUpdating = True
End Sub
Sub Create_PDF_FCV_Click()
Dim lastrow As Long
 Application.ScreenUpdating = False
 Dim path As String
 'get data
With ActiveSheet
lastrow = .Range("N" & .Rows.Count).End(xlUp).Row + 5
End With

 path = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\"
   ActiveSheet.Range("N1:AF" & lastrow + 28).ExportAsFixedFormat Type:=xlTypePDF, _
            Filename:=path & "FCV " & Format(Now(), "mmddyyyy") & ".PDF", _
            OpenAfterPublish:=False
            Application.ScreenUpdating = True
End Sub
