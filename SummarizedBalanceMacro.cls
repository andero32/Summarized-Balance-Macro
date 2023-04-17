VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub SummarizedGL()
    Dim lastRow As Long
    Dim acctNumber As String
    Dim currentTotal As Double
    Dim summaryRow As Long
    Dim firstRow As Long
    
    'Get the last row in Column A
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    
    'Start summary on row 2 in Column C
    summaryRow = 2
    
    'Loop through all rows in Column A
    For i = 2 To lastRow
        'Check if the account number is the same as the previous row
        If Cells(i, "A").Value <> acctNumber Then
            'If it's a new account number, write the previous total to the summary table
            If i > 2 Then
                'Write the current total to the first row of the current account number
                Range("C" & firstRow).Value = currentTotal
                
                'Reset first row counter
                firstRow = i
                
                'Increment the summary row counter
                summaryRow = summaryRow + 1
            Else
                'Set the first row counter
                firstRow = i
            End If
            
            'Update the current account number and reset the current total
            acctNumber = Cells(i, "A").Value
            currentTotal = 0
        End If
        
        'Add the current balance to the total for this account number
        currentTotal = currentTotal + Cells(i, "B").Value
    Next i
    
    'Write the last total to the summary table
    Range("C" & firstRow).Value = currentTotal
End Sub


