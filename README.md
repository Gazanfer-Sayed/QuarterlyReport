# Quarterly Report with Excel VBA

Description: This *EXCEL VBA* based project consists of advance macros that are used for automating tasks to decrease manual work. Additionally, I developed a userform which helps users generate reports and more with just a click.

The process followed and macros developed in this project are as follows:

1. InsertHeaders Macro:
   * Inserts a new row at the top of the sheet.
   * Adds list headers for “Division,” “Category,” and monthly columns (“Jan,” “Feb,” “Mar,” and “Total”).
     

 ```
   Sub InsertHeaders()

      Rows("1:1").Select
      Selection.Insert Shift:=xlDown
      Range("A1").Select
      ActiveCell.FormulaR1C1 = "Division"
      Range("B1").Select
      ActiveCell.FormulaR1C1 = "Category"
      Range("C1").Select
      ActiveCell.FormulaR1C1 = "Jan"
      Range("D1").Select
      ActiveCell.FormulaR1C1 = "Feb"
      Range("E1").Select
      ActiveCell.FormulaR1C1 = "Mar"
      Range("F1").Select
      ActiveCell.FormulaR1C1 = "Total"
      Range("A2").Select
   
    End Sub
 ```

2. FormatHeaders Macro:
   * Formats the list headers:
      - Makes them bold.
      - Applies a solid fill color (accent theme).
      - Sets font color to dark theme.
      - Adds medium-weight bottom border.
      - Adjusts font size.
    * Formats data in columns C to F as currency.

   ```
    Sub FormatHeaders()

        Range("A1:F1").Select
        Selection.Font.Bold = True
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent1
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        With Selection.Font
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = 0
        End With
        Selection.Borders(xlDiagonalDown).LineStyle = xlNone
        Selection.Borders(xlDiagonalUp).LineStyle = xlNone
        Selection.Borders(xlEdgeLeft).LineStyle = xlNone
        Selection.Borders(xlEdgeTop).LineStyle = xlNone
        With Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        Selection.Borders(xlEdgeRight).LineStyle = xlNone
        Selection.Borders(xlInsideVertical).LineStyle = xlNone
        Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
        Selection.Font.Size = 12
        Range("C2").Select
        Range(Selection, Selection.End(xlDown)).Select
        Range(Selection, Selection.End(xlToRight)).Select
        Selection.Style = "Currency"
        Columns("B:F").Select
        Columns("B:F").EntireColumn.AutoFit
        Range("A2").Select
    End Sub
   ```
  
3. AutomateTotalSUM Macro:
  * Calculates the total sum for the “Total” column dynamically based on data range.
    
```
  Public Sub AutomateTotalSUM()
      Dim lastCell As String
      Dim ws As Worksheet
  
          
          Range("F2").Select
          
          Selection.End(xlDown).Select
          
          lastCell = ActiveCell.Address(False, False)
          
          ActiveCell.Offset(1, 0).Select
          
          ActiveCell.Value = "=sum(F2:" & lastCell & ")"
      
  End Sub
```

4. YearlyReport Macro:

  * Iterates through worksheets (except “YEARLY REPORT” & NULL worksheets) and performs formatting and calcualtions.
  * Copies data from each sheet.
  * Pastes it into the “YEARLY REPORT” sheet.
  * Ensures proper formatting and total calculation to the Yearly Report sheet

  ```
    Sub YearlyReport()

      Dim ws As Worksheet
      Dim firstTime As Boolean
      
      firstTime = True
      
      For Each ws In Worksheets
          Worksheets(ws.Name).Select
          
          If ws.Name <> "YEARLY REPORT" And ActiveSheet.Range("A1") <> "" Then
              InsertHeaders
              FormatHeaders
              AutomateTotalSUM
              
              ' select current data
              Range("A2").Select
              Range(Selection, Selection.End(xlDown)).Select
              Range(Selection, Selection.End(xlToRight)).Select
              
              ' copy data
              Selection.Copy
              
              ' select yearly report
              Worksheets("YEARLY REPORT").Select
              
              ' paste data
              Range("A30000").Select
              Selection.End(xlUp).Select
              
              If firstTime <> True Then
                  ActiveCell.Offset(1, 0).Select
              Else
                  firstTime = False
              End If
              
              ActiveSheet.Paste
          End If
      ' move to the next sheet in the loop
      Next ws
      
      Worksheets("Yearly Report").Select
      InsertHeaders
      FormatHeaders
      AutomateTotalSUM

    End Sub
  ```

5. User Form:

  * Opens automatically when the workbook is opened.
  * Allows users to add new worksheets with custom names.
  * Provides a button to run the yearly report generation process.
  * Includes a combo box to select specific sheets.

```
Private Sub btnAddWorksheet_Click()

'declaring variable required for the error handler code
    Dim tryAgain As Integer
    On Error GoTo errorHandler

' adding worksheet in the 1st place
    Worksheets.Add before:=Worksheets(1)
    
'asking user a new name for the worksheet
    ActiveSheet.Name = InputBox("Please enter a new name for the worksheet")
    
' code for handling errors
errorHandler:
    tryAgain = MsgBox("Invalid input. Try again?", vbYesNo)
    
    If tryAgain = 6 Then
        Application.DisplayAlerts = False
        ActiveSheet.Delete
        btnAddWorksheet_Click
    Else
        Application.DisplayAlerts = False
        ActiveSheet.Delete
    End If
    
End Sub

---
Private Sub btnRunReport_Click()
    
    YearlyReport
    
End Sub

---
Private Sub cblWhichSheet_Change()

    Worksheets(Me.cblWhichSheet.Value).Select

End Sub

---
Private Sub UserForm_Initialize()

    Dim i As Integer
    i = 1

'Logic for adding all the sheet names in the combo box
    Do While i <= Worksheets.Count
    
        Me.cblWhichSheet.AddItem (Worksheets(i).Name)
        
        i = i + 1
        
    Loop

End Sub

----
Private Sub Workbook_Open()
    
    FRMReport.Show
    
End Sub

```

