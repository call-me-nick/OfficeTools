# OfficeTools

Visual Basic Macro Placeholder:
~~~
Sub Links()

    Dim userInput As String
    Dim inputParts() As String
    Dim addressParts() As String
    
    Dim hyperlinkAddress As String
    Dim hyperlinkSubDestination As String
    Dim hyperlinkScreenTip As String
    
    
    userInput = InputBox("Enter the hyperlink address, and destination link (if any), separated with ';' :      Example: <my_file.pdf>;<my_destination_123>")
    
    inputParts = Split(userInput, ";")
    
    ' Check for number of valid tokens
    If UBound(inputParts) >= 0 Then
        hyperlinkAddress = Trim(inputParts(0))
        
        ' Create ScreenTip with just end destination file name
        ' addressParts = Split(hyperlinkAddress, "\")
        ' hyperlinkScreenTip = addressParts(UBound(addressParts))
        hyperlinkScreenTip = hyperlinkAddress
        
        ' Assign destination or bookmark reference
        If UBound(inputParts) >= 1 Then
            hyperlinkSubDestination = Trim(inputParts(1))
        Else
            hyperlinkSubDestination = ""
        End If

        ' Check for Selection in ActiveDocument
        If Not Selection Is Nothing Then
            ActiveDocument.Hyperlinks.Add Anchor:=Selection.Range, Address:=hyperlinkAddress, SubAddress:=hyperlinkSubDestination, ScreenTip:=hyperlinkScreenTip
        Else
            MsgBox "Please select a range of text in the active document.", vbExclamation
        End If
    
    Else
        MsgBox "Please provide all required information (hyperlink address and optional page number).", vbExclamation
    End If
    
    
End Sub
~~~
