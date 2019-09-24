Attribute VB_Name = "Module1"
' Steps:
' 1. Loop through every worksheet and select the state contents.
' 2. Copy the state contents and paste it into the Combined_Data tab

Sub WellsFargo_HWTask3()
    
    ' Add a sheet named "Combined Data"
    Sheets.Add.Name = "Combined_Data"
    
    ' Move created sheet to be first sheet
    Sheets("Combined_Data").Move Before:=Sheets(1)
    
    ' Specify the location of the combined sheet
    Set combined_sheet = Worksheets("Combined_Data")

    Dim ws As Worksheet

    ' Loop through all sheets
    For Each ws In Worksheets
        If ws.Name <> "Combined_Data" Then
        
        ' Find the last row of the combined sheet after each paste
        ' Add 1 to get first empty row
        LastRow = combined_sheet.Cells(Rows.Count, "A").End(xlUp).Row + 1


        ' Find the last row of each worksheet
        ' Subtract one to return the number of rows without header
        ' HINT: Use similar logic as above, but this time subtract a row
        ' ****************
        ' [YOUR CODE HERE]
        LastRow2 = ws.Cells(Rows.Count, "A").End(xlUp).Row - 1
        
        ' ****************
        
        ' Copy the contents of each state sheet into the combined sheet
        ' ****************
        ' [YOUR CODE HERE]
        ws.Range("A2:G" & LastRow2 + 1).Copy Destination:=Sheets("Combined_Data").Range("A" & LastRow)

        '****************
        End If
    Next ws

' Copy the headers from sheet 1
    combined_sheet.Range("A1:G1").Value = Sheets(2).Range("A1:G1").Value
    
    ' Autofit to display data
    combined_sheet.Columns("A:G").AutoFit
End Sub


