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


How to use Macro:


Shortcut: Shift+Alt+M
- Runs Links Macro
- Need to Highlight text manually or use Find function to highlight section


You can create a hyperlink using one of the below formats:

< File >
hyperlinks-in-briefs-bookmarks-cross-references.pdf

< File >;< Destination | Bookmark >
Reply-hyperlink to How to-macro.docx;TestDestination


NOTE:
- A File can be the document you are currently focused on or another document as long as you specify the path to it. 
- By default if you are in the same folder you include the file name.


If you need to reference a file then you need to specify a path relative to your location:

1. Right Click File
2. Copy as path

Example:

"C:\Users\TRettinghouse\Desktop\Resources\Practice to Hyperlink Brief\Sample folder for testing\2023-12-22 DRAFT FR4 Petitioners Reply -w hyperlink box checked for TOC and first case linked.docx"

3. Trim relative to file containing the hyperlink

(Relative to Reply-hyperlink to How to-macro.docx)

Sample folder for testing\2023-12-22 DRAFT FR4 Petitioners Reply -first case linked to 2 pages test.pdf

