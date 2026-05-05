'Only works when ran from the sheet containing the pivot(ActiveSheet)!
'This only works if Work1 and Work2 are both sorted by Invoice Numbers in the same direction!

Sub ConcatenateInvoiceTexts()

    Dim count As Integer
    Dim i As Integer
    Dim invoiceText As String
    Dim invoiceCount As Integer
    Dim invoiceLine As Integer
    
    Set ws = ActiveSheet
    count = ws.Range("A6", ws.Range("A6").End(xlDown)).Rows.count
    
    invoiceLine = 0
    For i = 0 To count - 2
        invoiceCount = ws.Cells(i + 6, 5).Value
        
        'Concatenate strings based on the number of occurrences on Work1
        For j = 1 To invoiceCount
            invoiceText = invoiceText & " " & Worksheets("Work1").Cells(invoiceLine + 1 + j, 11).Value
        Next j
        
        invoiceLine = invoiceLine + invoiceCount

        'Format the concatenated string
        invoiceText = Replace(invoiceText, "(", "")
        invoiceText = Replace(invoiceText, ")", "")
        invoiceText = Trim(invoiceText)
        
        'Write the formatted string to the cell
        ws.Cells(i + 6, 6).Value = invoiceText

        'Empty the variable for reuse in the next iteration
        invoiceText = ""
    Next i
    
End Sub

'Only works when ran from the sheet containing the pivot(ActiveSheet)!
'Only needs Work1 to be sorted by Invoice Numbers!

Sub ConcatenateInvoiceTexts_Advanced()

    Dim count As Integer
    Dim i As Integer
    Dim invoiceText As String
    Dim invoiceCount As Integer
    Dim currentInvoice As Variant
    Dim result As Range
    Dim cellValue As Variant
    
    Set ws = ActiveSheet
    count = ws.Range("A6", ws.Range("A6").End(xlDown)).Rows.count 'Number of items in the Pivot Table
    
    For i = 0 To count - 2
   
        currentInvoice = ws.Cells(i + 6, 1).Value
        
        Set result = Worksheets("Work1").Range("B:B").Find( _
                        What:=currentInvoice, LookIn:=xlValues, LookAt:=xlWhole)
                        
        If result Is Nothing Then
            'No match found > leave cell empty
            ws.Cells(i + 6, 6).Value = ""
            GoTo next_i
        End If
       
        invoiceCount = ws.Cells(i + 6, 5).Value
        
        'Concatenate strings based on the number of occurrences on Work1
        For j = 0 To invoiceCount - 1
            cellValue = Worksheets("Work1").Cells(result.Row + j, 11).Value
            invoiceText = invoiceText & " " & cellValue
        Next j
         
        'Format the concatenated string
        invoiceText = Replace(invoiceText, "(", "")
        invoiceText = Replace(invoiceText, ")", "")
        invoiceText = Trim(invoiceText)
        
        'Write the formatted string to the cell
        ws.Cells(i + 6, 7).Value = invoiceText

        'Empty the variable for reuse in the next iteration
        invoiceText = ""
next_i:
    Next i
    
End Sub


