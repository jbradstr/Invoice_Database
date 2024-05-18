# Invoice_Database
I created an interactive invoice tailored to the client's specifications.  They wanted to be able to choose a state in the state field and then only have the addresses show up that were tied to that state in the address field.  I used depending drop down lists to accomplish this.  An addition they also wanted an automatic invoice number to be generated from the state abbreviation, the numbers of the street address followed by the first letters of the stree name, then the first letter of the state name, 4 static zeros at the end and finally a number that would increment based on if the same invoice number was ever found in another worksheet where all of the invoices would be submitted.  I created a custom function for this invoice number which you can find below.  I used ChatGPT to assist me in the base creation of the function and from there I tailored it to fit the client's specific situation. The client also wanted an easy way to get the layout in a pdf so that they could print it.  I created a button in which I tied a macro below to it so that it would print out the invoice number as the name and tie -EST to the end of it (see below macro code). The last button I created helps the user submit the invoice to another worksheet where it documents all of the fields of the worksheet into rows.  The description field dictates how many rows will be created as all of the other information collected will be the same for each row except the descriptions.


Invoice Number Function: (In Progress)

```vbscript
Function GenerateBaseInvoiceNumber(address As String, city As String, state As String) As String

    Dim ai As Worksheet: Set ai = ThisWorkbook.Worksheets("Actual_Invoice")
    Dim allInvoices As Worksheet: Set allInvoices = ThisWorkbook.Worksheets("All_Invoices")
    
    Dim baseInvoiceNumber As String
    Dim addressParts() As String
    Dim i As Integer
    Dim word As String
    Dim counter As Integer
    Dim finalInvoiceNumber As String
    
    ' Extract the number from the street address
    addressParts = Split(ai.Cells(9, 2), " ")
    
    For i = LBound(addressParts) To UBound(addressParts)
        If IsNumeric(addressParts(i)) Then
            baseInvoiceNumber = addressParts(i)
            Exit For
        End If
    Next i
    
    ' Add first letters of non-numeric words in street address
    For i = LBound(addressParts) To UBound(addressParts)
        If Not IsNumeric(addressParts(i)) Then
            word = addressParts(i)
            If Len(word) > 1 Then
                baseInvoiceNumber = baseInvoiceNumber & Left(word, 1)
            End If
        End If
    Next i
    
    ' Add first letters of city
    If Len(ai.Cells(10, 2)) > 1 Then
        baseInvoiceNumber = baseInvoiceNumber & Left(ai.Cells(10, 2), 1)
    End If
    
    ' Add state abbreviation
    If Len(ai.Cells(11, 2)) >= 2 Then
        baseInvoiceNumber = Left(ai.Cells(11, 2), 2) & baseInvoiceNumber
    End If
    
    ' Add the fixed string of four zeros
    baseInvoiceNumber = baseInvoiceNumber & "0000"
    
    ' Initialize counter starting from 1
    counter = 1
    finalInvoiceNumber = baseInvoiceNumber & counter
    
    ' Check for unique invoice number
    Do While Not IsUniqueInvoiceNumber(finalInvoiceNumber, allInvoices)
        counter = counter + 1
        finalInvoiceNumber = baseInvoiceNumber & counter
    Loop
    
    GenerateBaseInvoiceNumber = finalInvoiceNumber
End Function

```

Print invoice as a PDF to a folder on the desktop:

```vbscript


Public Sub print_invoice_as_pdf()

    Dim ws As Worksheet
    Dim invoiceNumber As String
    Dim filePath As String
    Dim invoiceCell As Range
    Dim desktopPath As String
    Dim folderName As String

    ' Set the worksheet
    Dim ai As Worksheet: Set ai = ThisWorkbook.Worksheets("Actual_Invoice")

    ' setting cell where invoice # can be found
    Set invoiceCell = ai.Range("G4")

    ' Get the invoice number from the specified cell
    invoiceNumber = invoiceCell.Value

    ' Check if the invoice number is not empty
    If invoiceNumber = "" Then
        MsgBox "Invoice number cell is empty.", vbExclamation
        Exit Sub
    End If

    ' Get the path to the desktop
    desktopPath = Environ("USERPROFILE") & "\Desktop"

    ' Specify the folder name on the desktop
    folderName = "Invoice_test" ' Change this to your specific folder name

    ' Define the file path and name
    filePath = desktopPath & "\" & folderName & "\" & invoiceNumber & "-EST.pdf"

    ' Check if the folder exists, create it if it doesn't
    If Dir(desktopPath & "\" & folderName, vbDirectory) = "" Then
        MkDir desktopPath & "\" & folderName
    End If

    ' Export the first printable area to PDF
    ai.ExportAsFixedFormat Type:=xlTypePDF, Filename:=filePath, Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False

    ' Notify the user
    MsgBox "Invoice has been printed to PDF: " & filePath, vbInformation

End Sub
```

Export invoice data to worksheet (In Progress)

```vbscript
Sub ExportInvoiceData()
    Dim wsInput As Worksheet
    Dim wsOutput As Worksheet
    Dim outputRow As Long
    Dim descriptionCount As Long
    Dim i As Long
    Dim cell As Range
    Dim descriptionRange As Range
    
    ' Set references to the input and output worksheets
    Set wsInput = ThisWorkbook.Sheets("Actual_Invoice")
    On Error Resume Next
    Set wsOutput = ThisWorkbook.Sheets("All_Invoices")
    On Error GoTo 0
    If wsOutput Is Nothing Then
        Set wsOutput = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsOutput.Name = "All_Invoices"
    Else
        wsOutput.Cells.Clear 'might have to comment this out so it doesn't get rid of the data
    End If

    ' Set headers for the output worksheet
    With wsOutput
        .Cells(1, 1).Value = "WO-Invoice #"
        .Cells(1, 2).Value = "Invoice Date"
        .Cells(1, 3).Value = "Company Name:"
        .Cells(1, 4).Value = "POC:"
        .Cells(1, 5).Value = "Service Address:"
        .Cells(1, 6).Value = "City:"
        .Cells(1, 7).Value = "State:"
        .Cells(1, 8).Value = "Zip:"
        .Cells(1, 9).Value = "POC Phone:"
        .Cells(1, 10).Value = "Description"
        .Cells(1, 11).Value = "Qty"
        .Cells(1, 12).Value = "Hours"
        .Cells(1, 13).Value = "Rate"
        .Cells(1, 14).Value = "Amount"
        .Cells(1, 15).Value = "Subtotal"
        .Cells(1, 16).Value = "Tax Rate"
        .Cells(1, 17).Value = "Sales Tax"
        .Cells(1, 18).Value = "Other"
        .Cells(1, 19).Value = "Total"
    End With

    ' Define the range for descriptions
    Set descriptionRange = wsInput.Range("A16:A31")
    
    ' Count the number of non-empty description cells
    descriptionCount = 0
    For Each cell In descriptionRange
        If cell.Value <> "" Then
            descriptionCount = descriptionCount + 1
        End If
    Next cell


    ' Initialize the output row
    outputRow = 2


    ' Loop through each description
    For i = 1 To descriptionCount
        ' Copy data to the output worksheet
        wsOutput.Cells(outputRow, 1).Value = wsInput.Range("G4").Value ' WO-Invoice #
        wsOutput.Cells(outputRow, 2).Value = wsInput.Range("B2").Value ' Invoice Date
        wsOutput.Cells(outputRow, 3).Value = wsInput.Range("B3").Value ' Company Name
        wsOutput.Cells(outputRow, 4).Value = wsInput.Range("B4").Value ' POC
        wsOutput.Cells(outputRow, 5).Value = wsInput.Range("B5").Value ' Service Address
        wsOutput.Cells(outputRow, 6).Value = wsInput.Range("B6").Value ' City
        wsOutput.Cells(outputRow, 7).Value = wsInput.Range("B7").Value ' State
        wsOutput.Cells(outputRow, 8).Value = wsInput.Range("B8").Value ' Zip
        wsOutput.Cells(outputRow, 9).Value = wsInput.Range("B9").Value ' POC Phone
        
        ' Assume descriptions are in B21:B(n), quantities in C21:C(n), etc.
        wsOutput.Cells(outputRow, 10).Value = wsInput.Cells(21 + i - 1, 2).Value ' Description
        wsOutput.Cells(outputRow, 11).Value = wsInput.Cells(21 + i - 1, 3).Value ' Qty
        wsOutput.Cells(outputRow, 12).Value = wsInput.Cells(21 + i - 1, 4).Value ' Hours
        wsOutput.Cells(outputRow, 13).Value = wsInput.Cells(21 + i - 1, 5).Value ' Rate
        wsOutput.Cells(outputRow, 14).Value = wsInput.Cells(21 + i - 1, 6).Value ' Amount
        
        ' Subtotal, Tax Rate, Sales Tax, Other, Total from fixed cells
        wsOutput.Cells(outputRow, 15).Value = wsInput.Range("B10").Value ' Subtotal
        wsOutput.Cells(outputRow, 16).Value = wsInput.Range("B11").Value ' Tax Rate
        wsOutput.Cells(outputRow, 17).Value = wsInput.Range("B12").Value ' Sales Tax
        wsOutput.Cells(outputRow, 18).Value = wsInput.Range("B13").Value ' Other
        wsOutput.Cells(outputRow, 19).Value = wsInput.Range("B14").Value ' Total
        
        ' Move to the next output row
        outputRow = outputRow + 1
    Next i
    
    MsgBox "Invoice data has been exported successfully!", vbInformation
End Sub
```
