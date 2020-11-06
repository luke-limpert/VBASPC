Attribute VB_Name = "Module4"
Sub Abnormality_Report()

' Add a reference to the Word-library via VBE > Tools > References > Microsoft Word xx.x Object Library.
' Create a folder named C:\Temp or edit the filnames in the code.
'
    Dim wrdApp As Word.Application
    Dim wrdDoc As Word.Document
    Dim bWeStartedWord As Boolean
    
    Dim Title As String
    
    Dim i As Integer
    
    On Error Resume Next
    Set wrdApp = GetObject(, "Word.Application")
    On Error GoTo 0
    If wrdApp Is Nothing Then
        Set wrdApp = CreateObject("Word.Application")
        bWeStartedWord = True
    End If
    wrdApp.Visible = True 'optional!
    
    Set wrdDoc = wrdApp.Documents.Add ' create a new document
    ' or open an existing document:
    'Set wrdDoc = wrdApp.Documents.Open("C:\Foldername\Filename.docx")
    
    ' example word operations:
    With wrdDoc
    
    .Content.InsertAfter "Weekly Abnormality Tracking report-out"
    .Content.InsertParagraphAfter
    .Content.InsertAfter "Abnormalities by part"
    .Content.InsertParagraphAfter
        
    'copy data from A1:A10 into the word doc:
        
        For i = 0 To 49
        
            Select Case i
    
    Case 0
Title = "Nissan L42P Assist: Case Depth"
'Using Find Function
        LastRow = Sheet15.Range("A" & Rows.Count).End(xlUp).Row
        
'Create Values

        Spread1 = Sheet15.Cells(LastRow, "J").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet15.Cells(LastRow, "J").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread2 = Sheet15.Cells(LastRow, "J").Value

        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet15.Cells(LastRow, "J").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread3 = Sheet15.Cells(LastRow, "J").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet15.Cells(LastRow, "J").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread4 = Sheet15.Cells(LastRow, "J").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet15.Cells(LastRow, "J").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread5 = Sheet15.Cells(LastRow, "J").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet15.Cells(LastRow, "J").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread6 = Sheet15.Cells(LastRow, "J").Value

    Case 1
Title = "Nissan L42P Sensor: Case Depth"
'Using Find Function
        LastRow = Sheet16.Range("A" & Rows.Count).End(xlUp).Row
        
'Create Values

        Spread1 = Sheet16.Cells(LastRow, "J").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet16.Cells(LastRow, "J").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread2 = Sheet16.Cells(LastRow, "J").Value

        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet16.Cells(LastRow, "J").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread3 = Sheet16.Cells(LastRow, "J").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet16.Cells(LastRow, "J").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread4 = Sheet16.Cells(LastRow, "J").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet16.Cells(LastRow, "J").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread5 = Sheet16.Cells(LastRow, "J").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet16.Cells(LastRow, "J").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread6 = Sheet16.Cells(LastRow, "J").Value

    Case 2
Title = "Nissan L42P Assist: Hardness"
'Using Find Function
        LastRow = Sheet15.Range("A" & Rows.Count).End(xlUp).Row
        
'Create Values

        Spread1 = Sheet15.Cells(LastRow, "I").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet15.Cells(LastRow, "I").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread2 = Sheet15.Cells(LastRow, "I").Value

        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet15.Cells(LastRow, "I").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread3 = Sheet15.Cells(LastRow, "I").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet15.Cells(LastRow, "I").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread4 = Sheet15.Cells(LastRow, "I").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet15.Cells(LastRow, "I").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread5 = Sheet15.Cells(LastRow, "I").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet15.Cells(LastRow, "I").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread6 = Sheet15.Cells(LastRow, "I").Value

    Case 3
Title = "Nissan L42P Sensor: Hardness"
'Using Find Function
        LastRow = Sheet16.Range("A" & Rows.Count).End(xlUp).Row
        
'Create Values

        Spread1 = Sheet16.Cells(LastRow, "I").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet16.Cells(LastRow, "I").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread2 = Sheet16.Cells(LastRow, "I").Value

        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet16.Cells(LastRow, "I").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread3 = Sheet16.Cells(LastRow, "I").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet16.Cells(LastRow, "I").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread4 = Sheet16.Cells(LastRow, "I").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet16.Cells(LastRow, "I").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread5 = Sheet16.Cells(LastRow, "I").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet16.Cells(LastRow, "I").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread6 = Sheet16.Cells(LastRow, "I").Value

    Case 4
Title = "BMW LHD CGR Assist Case Depth"
'Using Find Function
    
    LastRow = Sheet1.Range("A" & Rows.Count).End(xlUp).Row
    
'Create Values

        Spread1 = Sheet1.Cells(LastRow, "J").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet1.Cells(LastRow, "J").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread2 = Sheet1.Cells(LastRow, "J").Value

        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet1.Cells(LastRow, "J").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread3 = Sheet1.Cells(LastRow, "J").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet1.Cells(LastRow, "J").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread4 = Sheet1.Cells(LastRow, "J").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet1.Cells(LastRow, "J").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread5 = Sheet1.Cells(LastRow, "J").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet1.Cells(LastRow, "J").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread6 = Sheet1.Cells(LastRow, "J").Value

    Case 5
Title = "BMW LHD CGR Sensor Case Depth"
'Using Find Function
    
    LastRow = Sheet2.Range("A" & Rows.Count).End(xlUp).Row
    
'Create Values

        Spread1 = Sheet2.Cells(LastRow, "J").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet2.Cells(LastRow, "J").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread2 = Sheet2.Cells(LastRow, "J").Value

        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet2.Cells(LastRow, "J").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread3 = Sheet2.Cells(LastRow, "J").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet2.Cells(LastRow, "J").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread4 = Sheet2.Cells(LastRow, "J").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet2.Cells(LastRow, "J").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread5 = Sheet2.Cells(LastRow, "J").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet2.Cells(LastRow, "J").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread6 = Sheet2.Cells(LastRow, "J").Value
        
    Case 6
Title = "BMW LHD CGR Assist Hardness"
'Using Find Function
    
    LastRow = Sheet1.Range("A" & Rows.Count).End(xlUp).Row
    
'Create Values

        Spread1 = Sheet1.Cells(LastRow, "I").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet1.Cells(LastRow, "I").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread2 = Sheet1.Cells(LastRow, "I").Value

        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet1.Cells(LastRow, "I").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread3 = Sheet1.Cells(LastRow, "I").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet1.Cells(LastRow, "I").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread4 = Sheet1.Cells(LastRow, "I").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet1.Cells(LastRow, "I").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread5 = Sheet1.Cells(LastRow, "I").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet1.Cells(LastRow, "I").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread6 = Sheet1.Cells(LastRow, "I").Value

    Case 7
Title = "BMW LHD CGR Sensor Hardness"
'Using Find Function
    
    LastRow = Sheet2.Range("A" & Rows.Count).End(xlUp).Row
    
'Create Values

        Spread1 = Sheet2.Cells(LastRow, "I").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet2.Cells(LastRow, "I").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread2 = Sheet2.Cells(LastRow, "I").Value

        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet2.Cells(LastRow, "I").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread3 = Sheet2.Cells(LastRow, "I").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet2.Cells(LastRow, "I").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread4 = Sheet2.Cells(LastRow, "I").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet2.Cells(LastRow, "I").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread5 = Sheet2.Cells(LastRow, "I").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet2.Cells(LastRow, "I").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread6 = Sheet2.Cells(LastRow, "I").Value

    Case 8
Title = "BMW RHD CGR Assist: Case Depth"
'Using Find Function
    
    LastRow = Sheet3.Range("A" & Rows.Count).End(xlUp).Row
    
'Create Values

        Spread1 = Sheet3.Cells(LastRow, "J").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet3.Cells(LastRow, "J").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread2 = Sheet3.Cells(LastRow, "J").Value

        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet3.Cells(LastRow, "J").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread3 = Sheet3.Cells(LastRow, "J").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet3.Cells(LastRow, "J").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread4 = Sheet3.Cells(LastRow, "J").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet3.Cells(LastRow, "J").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread5 = Sheet3.Cells(LastRow, "J").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet3.Cells(LastRow, "J").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread6 = Sheet3.Cells(LastRow, "J").Value

    Case 9
Title = "BMW RHD CGR Sensor: Case Depth"
'Using Find Function
    
    LastRow = Sheet4.Range("A" & Rows.Count).End(xlUp).Row
    
'Create Values

        Spread1 = Sheet4.Cells(LastRow, "J").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet4.Cells(LastRow, "J").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread2 = Sheet4.Cells(LastRow, "J").Value

        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet4.Cells(LastRow, "J").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread3 = Sheet4.Cells(LastRow, "J").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet4.Cells(LastRow, "J").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread4 = Sheet4.Cells(LastRow, "J").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet4.Cells(LastRow, "J").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread5 = Sheet4.Cells(LastRow, "J").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet4.Cells(LastRow, "J").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread6 = Sheet4.Cells(LastRow, "J").Value

    Case 10
Title = "BMW RHD CGR Assist: Hardness"
'Using Find Function

    LastRow = Sheet3.Range("A" & Rows.Count).End(xlUp).Row
    
'Create Values

        Spread1 = Sheet3.Cells(LastRow, "I").Value
        
        LastRow = LastRow - 1
    
        Do Until IsNumeric(Sheet3.Cells(LastRow, "I").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread2 = Sheet3.Cells(LastRow, "I").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet3.Cells(LastRow, "I").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread3 = Sheet3.Cells(LastRow, "I").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet3.Cells(LastRow, "I").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread4 = Sheet3.Cells(LastRow, "I").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet3.Cells(LastRow, "I").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread5 = Sheet3.Cells(LastRow, "I").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet3.Cells(LastRow, "I").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread6 = Sheet3.Cells(LastRow, "I").Value
        
    Case 11
Title = "BMW RHD CGR Sensor: Hardness"
'Using Find Function
    
    LastRow = Sheet4.Range("A" & Rows.Count).End(xlUp).Row
    
'Create Values

        Spread1 = Sheet4.Cells(LastRow, "I").Value

        LastRow = LastRow - 1
    
        Do Until IsNumeric(Sheet4.Cells(LastRow, "I").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread2 = Sheet4.Cells(LastRow, "I").Value

        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet4.Cells(LastRow, "I").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread3 = Sheet4.Cells(LastRow, "I").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet4.Cells(LastRow, "I").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread4 = Sheet4.Cells(LastRow, "I").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet4.Cells(LastRow, "I").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread5 = Sheet4.Cells(LastRow, "I").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet4.Cells(LastRow, "I").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread6 = Sheet4.Cells(LastRow, "I").Value

    Case 12
Title = "BMW LHD VGR Assist: Case Depth"
'Using Find Function
    
    LastRow = Sheet5.Range("A" & Rows.Count).End(xlUp).Row
    
'Create Values

        Spread1 = Sheet5.Cells(LastRow, "J").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet5.Cells(LastRow, "J").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread2 = Sheet5.Cells(LastRow, "J").Value

        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet5.Cells(LastRow, "J").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread3 = Sheet5.Cells(LastRow, "J").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet5.Cells(LastRow, "J").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread4 = Sheet5.Cells(LastRow, "J").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet5.Cells(LastRow, "J").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread5 = Sheet5.Cells(LastRow, "J").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet5.Cells(LastRow, "J").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread6 = Sheet5.Cells(LastRow, "J").Value
        
    Case 13
Title = "BMW LHD VGR Assist: Hardness"
'Using Find Function
    
    LastRow = Sheet5.Range("A" & Rows.Count).End(xlUp).Row
    
'Create Values

        Spread1 = Sheet5.Cells(LastRow, "I").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet5.Cells(LastRow, "I").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread2 = Sheet5.Cells(LastRow, "I").Value

        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet5.Cells(LastRow, "I").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread3 = Sheet5.Cells(LastRow, "I").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet5.Cells(LastRow, "I").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread4 = Sheet5.Cells(LastRow, "I").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet5.Cells(LastRow, "I").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread5 = Sheet5.Cells(LastRow, "I").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet5.Cells(LastRow, "I").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread6 = Sheet5.Cells(LastRow, "I").Value

    Case 14
Title = "BMW RHD VGR Assist: Case Depth"
'Using Find Function
    
    LastRow = Sheet6.Range("A" & Rows.Count).End(xlUp).Row

'Create Values

        Spread1 = Sheet6.Cells(LastRow, "J").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet6.Cells(LastRow, "J").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread2 = Sheet6.Cells(LastRow, "J").Value

        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet6.Cells(LastRow, "J").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread3 = Sheet6.Cells(LastRow, "J").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet6.Cells(LastRow, "J").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread4 = Sheet6.Cells(LastRow, "J").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet6.Cells(LastRow, "J").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread5 = Sheet6.Cells(LastRow, "J").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet6.Cells(LastRow, "J").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread6 = Sheet6.Cells(LastRow, "J").Value
        
    Case 15
Title = "BMW RHD VGR Assist: Hardness"
'Using Find Function
    
    LastRow = Sheet6.Range("A" & Rows.Count).End(xlUp).Row
    
'Create Values

        Spread1 = Sheet6.Cells(LastRow, "I").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet6.Cells(LastRow, "I").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread2 = Sheet6.Cells(LastRow, "I").Value

        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet6.Cells(LastRow, "I").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread3 = Sheet6.Cells(LastRow, "I").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet6.Cells(LastRow, "I").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread4 = Sheet6.Cells(LastRow, "I").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet6.Cells(LastRow, "I").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread5 = Sheet6.Cells(LastRow, "I").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet6.Cells(LastRow, "I").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread6 = Sheet6.Cells(LastRow, "I").Value

    Case 16
Title = "Honda THR Assist: Case Depth"
'Using Find Function
    
    LastRow = Sheet7.Range("A" & Rows.Count).End(xlUp).Row
    
'Create Values

        Spread1 = Sheet7.Cells(LastRow, "J").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet7.Cells(LastRow, "J").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread2 = Sheet7.Cells(LastRow, "J").Value

        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet7.Cells(LastRow, "J").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread3 = Sheet7.Cells(LastRow, "J").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet7.Cells(LastRow, "J").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread4 = Sheet7.Cells(LastRow, "J").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet7.Cells(LastRow, "J").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread5 = Sheet7.Cells(LastRow, "J").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet7.Cells(LastRow, "J").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread6 = Sheet7.Cells(LastRow, "J").Value

    Case 17
Title = "Honda THR Sensor: Case Depth"
'Using Find Function
    
    LastRow = Sheet8.Range("A" & Rows.Count).End(xlUp).Row

'Create Values

        Spread1 = Sheet8.Cells(LastRow, "J").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet8.Cells(LastRow, "J").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread2 = Sheet8.Cells(LastRow, "J").Value

        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet8.Cells(LastRow, "J").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread3 = Sheet8.Cells(LastRow, "J").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet8.Cells(LastRow, "J").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread4 = Sheet8.Cells(LastRow, "J").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet8.Cells(LastRow, "J").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread5 = Sheet8.Cells(LastRow, "J").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet8.Cells(LastRow, "J").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread6 = Sheet8.Cells(LastRow, "J").Value

    Case 18
Title = "Honda THR Assist: Hardness"
'Using Find Function
    
    LastRow = Sheet7.Range("A" & Rows.Count).End(xlUp).Row
    
'Create Values

        Spread1 = Sheet7.Cells(LastRow, "I").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet7.Cells(LastRow, "I").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread2 = Sheet7.Cells(LastRow, "I").Value

        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet7.Cells(LastRow, "I").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread3 = Sheet7.Cells(LastRow, "I").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet7.Cells(LastRow, "I").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread4 = Sheet7.Cells(LastRow, "I").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet7.Cells(LastRow, "I").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread5 = Sheet7.Cells(LastRow, "I").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet7.Cells(LastRow, "I").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread6 = Sheet7.Cells(LastRow, "I").Value
    
    Case 19
Title = "Honda THR Sensor: Hardness"
'Using Find Function
    
    LastRow = Sheet8.Range("A" & Rows.Count).End(xlUp).Row
    
'Create Values

        Spread1 = Sheet8.Cells(LastRow, "I").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet8.Cells(LastRow, "I").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread2 = Sheet8.Cells(LastRow, "I").Value

        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet8.Cells(LastRow, "I").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread3 = Sheet8.Cells(LastRow, "I").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet8.Cells(LastRow, "I").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread4 = Sheet8.Cells(LastRow, "I").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet8.Cells(LastRow, "I").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread5 = Sheet8.Cells(LastRow, "I").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet8.Cells(LastRow, "I").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread6 = Sheet8.Cells(LastRow, "I").Value

    Case 20
Title = "Honda TJB Assist: Case Depth"
'Using Find Function

    LastRow = Sheet17.Range("A" & Rows.Count).End(xlUp).Row
    
'Create Values

        Spread1 = Sheet17.Cells(LastRow, "J").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet17.Cells(LastRow, "J").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread2 = Sheet17.Cells(LastRow, "J").Value

        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet17.Cells(LastRow, "J").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread3 = Sheet17.Cells(LastRow, "J").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet17.Cells(LastRow, "J").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread4 = Sheet17.Cells(LastRow, "J").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet17.Cells(LastRow, "J").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread5 = Sheet17.Cells(LastRow, "J").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet17.Cells(LastRow, "J").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread6 = Sheet17.Cells(LastRow, "J").Value

    Case 21
Title = "Honda TJB Assist: Hardness"
'Using Find Function

    LastRow = Sheet17.Range("A" & Rows.Count).End(xlUp).Row
    
'Create Values
        
        Spread1 = Sheet17.Cells(LastRow, "I").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet17.Cells(LastRow, "I").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread2 = Sheet17.Cells(LastRow, "I").Value

        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet17.Cells(LastRow, "I").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread3 = Sheet17.Cells(LastRow, "I").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet17.Cells(LastRow, "I").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread4 = Sheet17.Cells(LastRow, "I").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet17.Cells(LastRow, "I").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread5 = Sheet17.Cells(LastRow, "I").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet17.Cells(LastRow, "I").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread6 = Sheet17.Cells(LastRow, "I").Value
        
    Case 22
Title = "09PL Ball Screw: Hardness A"
'Using Find Function

    LastRow = Sheet12.Range("A" & Rows.Count).End(xlUp).Row
    
'Create Values
        
        Spread1 = Sheet12.Cells(LastRow, "S").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet12.Cells(LastRow, "S").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread2 = Sheet12.Cells(LastRow, "S").Value

        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet12.Cells(LastRow, "S").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread3 = Sheet12.Cells(LastRow, "S").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet12.Cells(LastRow, "S").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread4 = Sheet12.Cells(LastRow, "S").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet12.Cells(LastRow, "S").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread5 = Sheet12.Cells(LastRow, "S").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet12.Cells(LastRow, "S").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread6 = Sheet12.Cells(LastRow, "S").Value
        
    Case 23
Title = "09PL Ball Screw: Hardness B"
'Using Find Function

    LastRow = Sheet12.Range("A" & Rows.Count).End(xlUp).Row
    
'Create Values
        
        Spread1 = Sheet12.Cells(LastRow, "V").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet12.Cells(LastRow, "V").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread2 = Sheet12.Cells(LastRow, "V").Value

        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet12.Cells(LastRow, "V").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread3 = Sheet12.Cells(LastRow, "V").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet12.Cells(LastRow, "V").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread4 = Sheet12.Cells(LastRow, "V").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet12.Cells(LastRow, "V").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread5 = Sheet12.Cells(LastRow, "V").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet12.Cells(LastRow, "V").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread6 = Sheet12.Cells(LastRow, "V").Value
        
    Case 24
Title = "09PL Ball Screw: Case Depth"
'Using Find Function

    LastRow = Sheet12.Range("A" & Rows.Count).End(xlUp).Row
    
'Create Values
        
        Spread1 = Sheet12.Cells(LastRow, "T").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet12.Cells(LastRow, "T").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread2 = Sheet12.Cells(LastRow, "T").Value

        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet12.Cells(LastRow, "T").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread3 = Sheet12.Cells(LastRow, "T").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet12.Cells(LastRow, "T").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread4 = Sheet12.Cells(LastRow, "T").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet12.Cells(LastRow, "T").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread5 = Sheet12.Cells(LastRow, "T").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet12.Cells(LastRow, "T").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread6 = Sheet12.Cells(LastRow, "T").Value
        
    Case 25
Title = "09PL Rack Side: Tooth Hardness"
'Using Find Function
    
    LastRow = Sheet13.Range("A" & Rows.Count).End(xlUp).Row
    
'Create Values

        Spread1 = Sheet13.Cells(LastRow, "I").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet13.Cells(LastRow, "I").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread2 = Sheet13.Cells(LastRow, "I").Value

        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet13.Cells(LastRow, "I").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread3 = Sheet13.Cells(LastRow, "I").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet13.Cells(LastRow, "I").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread4 = Sheet13.Cells(LastRow, "I").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet13.Cells(LastRow, "I").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread5 = Sheet13.Cells(LastRow, "I").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet13.Cells(LastRow, "I").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread6 = Sheet13.Cells(LastRow, "I").Value

    Case 26
Title = "09PL Rack Side: Case Depth"
'Using Find Function

    LastRow = Sheet13.Range("A" & Rows.Count).End(xlUp).Row
    
'Create Values

        Spread1 = Sheet13.Cells(LastRow, "J").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet13.Cells(LastRow, "J").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread2 = Sheet13.Cells(LastRow, "J").Value

        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet13.Cells(LastRow, "J").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread3 = Sheet13.Cells(LastRow, "J").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet13.Cells(LastRow, "J").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread4 = Sheet13.Cells(LastRow, "J").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet13.Cells(LastRow, "J").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread5 = Sheet13.Cells(LastRow, "J").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet13.Cells(LastRow, "J").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread6 = Sheet13.Cells(LastRow, "J").Value
        
    Case 27
Title = "BMW G2X Assist: Tooth Hardness"
'Using Find Function
    
    LastRow = Sheet18.Range("A" & Rows.Count).End(xlUp).Row
    
'Create Values

        Spread1 = Sheet18.Cells(LastRow, "I").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet18.Cells(LastRow, "I").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread2 = Sheet18.Cells(LastRow, "I").Value

        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet18.Cells(LastRow, "I").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread3 = Sheet18.Cells(LastRow, "I").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet18.Cells(LastRow, "I").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread4 = Sheet18.Cells(LastRow, "I").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet18.Cells(LastRow, "I").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread5 = Sheet18.Cells(LastRow, "I").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet18.Cells(LastRow, "I").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread6 = Sheet18.Cells(LastRow, "I").Value

    Case 28
Title = "BMW G2X Assist: Case Depth"
'Using Find Function

    LastRow = Sheet18.Range("A" & Rows.Count).End(xlUp).Row
    
'Create Values

        Spread1 = Sheet18.Cells(LastRow, "J").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet18.Cells(LastRow, "J").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread2 = Sheet18.Cells(LastRow, "J").Value

        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet18.Cells(LastRow, "J").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread3 = Sheet18.Cells(LastRow, "J").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet18.Cells(LastRow, "J").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread4 = Sheet18.Cells(LastRow, "J").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet18.Cells(LastRow, "J").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread5 = Sheet18.Cells(LastRow, "J").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet18.Cells(LastRow, "J").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread6 = Sheet18.Cells(LastRow, "J").Value
        
    Case 29
Title = "BMW G2X Sensor: Tooth Hardness"
'Using Find Function
    
    LastRow = Sheet19.Range("A" & Rows.Count).End(xlUp).Row
    
'Create Values

        Spread1 = Sheet19.Cells(LastRow, "I").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet19.Cells(LastRow, "I").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread2 = Sheet19.Cells(LastRow, "I").Value

        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet19.Cells(LastRow, "I").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread3 = Sheet19.Cells(LastRow, "I").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet19.Cells(LastRow, "I").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread4 = Sheet19.Cells(LastRow, "I").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet19.Cells(LastRow, "I").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread5 = Sheet19.Cells(LastRow, "I").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet19.Cells(LastRow, "I").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread6 = Sheet19.Cells(LastRow, "I").Value

    Case 30
Title = "BMW G2X Sensor: Case Depth"
'Using Find Function

    LastRow = Sheet19.Range("A" & Rows.Count).End(xlUp).Row
    
'Create Values

        Spread1 = Sheet19.Cells(LastRow, "J").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet19.Cells(LastRow, "J").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread2 = Sheet19.Cells(LastRow, "J").Value

        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet19.Cells(LastRow, "J").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread3 = Sheet19.Cells(LastRow, "J").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet19.Cells(LastRow, "J").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread4 = Sheet19.Cells(LastRow, "J").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet19.Cells(LastRow, "J").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread5 = Sheet19.Cells(LastRow, "J").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet19.Cells(LastRow, "J").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread6 = Sheet19.Cells(LastRow, "J").Value
        
    Case 31
Title = "15PL Rack Side: Tooth Hardness"
'Using Find Function
    
    LastRow = Sheet20.Range("A" & Rows.Count).End(xlUp).Row
    
'Create Values

        Spread1 = Sheet20.Cells(LastRow, "I").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet20.Cells(LastRow, "I").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread2 = Sheet20.Cells(LastRow, "I").Value

        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet20.Cells(LastRow, "I").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread3 = Sheet20.Cells(LastRow, "I").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet20.Cells(LastRow, "I").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread4 = Sheet20.Cells(LastRow, "I").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet20.Cells(LastRow, "I").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread5 = Sheet20.Cells(LastRow, "I").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet20.Cells(LastRow, "I").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread6 = Sheet20.Cells(LastRow, "I").Value

    Case 32
Title = "15PL Rack Side: Case Depth"
'Using Find Function

    LastRow = Sheet20.Range("A" & Rows.Count).End(xlUp).Row
    
'Create Values

        Spread1 = Sheet20.Cells(LastRow, "J").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet20.Cells(LastRow, "J").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread2 = Sheet20.Cells(LastRow, "J").Value

        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet20.Cells(LastRow, "J").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread3 = Sheet20.Cells(LastRow, "J").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet20.Cells(LastRow, "J").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread4 = Sheet20.Cells(LastRow, "J").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet20.Cells(LastRow, "J").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread5 = Sheet20.Cells(LastRow, "J").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet20.Cells(LastRow, "J").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread6 = Sheet20.Cells(LastRow, "J").Value
        
    Case 33
Title = "15PL Ball Screw: Hardness A"
'Using Find Function

    LastRow = Sheet21.Range("A" & Rows.Count).End(xlUp).Row
    
'Create Values
        
        Spread1 = Sheet21.Cells(LastRow, "S").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet21.Cells(LastRow, "S").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread2 = Sheet21.Cells(LastRow, "S").Value

        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet21.Cells(LastRow, "S").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread3 = Sheet21.Cells(LastRow, "S").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet21.Cells(LastRow, "S").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread4 = Sheet21.Cells(LastRow, "S").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet21.Cells(LastRow, "S").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread5 = Sheet21.Cells(LastRow, "S").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet21.Cells(LastRow, "S").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread6 = Sheet21.Cells(LastRow, "S").Value
        
    Case 23
Title = "015PL Ball Screw: Hardness B"
'Using Find Function

    LastRow = Sheet21.Range("A" & Rows.Count).End(xlUp).Row
    
'Create Values
        
        Spread1 = Sheet21.Cells(LastRow, "V").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet21.Cells(LastRow, "V").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread2 = Sheet21.Cells(LastRow, "V").Value

        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet21.Cells(LastRow, "V").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread3 = Sheet21.Cells(LastRow, "V").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet21.Cells(LastRow, "V").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread4 = Sheet21.Cells(LastRow, "V").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet21.Cells(LastRow, "V").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread5 = Sheet21.Cells(LastRow, "V").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet21.Cells(LastRow, "V").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread6 = Sheet21.Cells(LastRow, "V").Value
        
    Case 35
Title = "09PL Ball Screw: Case Depth"
'Using Find Function

    LastRow = Sheet21.Range("A" & Rows.Count).End(xlUp).Row
    
'Create Values
        
        Spread1 = Sheet21.Cells(LastRow, "T").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet21.Cells(LastRow, "T").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread2 = Sheet21.Cells(LastRow, "T").Value

        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet21.Cells(LastRow, "T").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread3 = Sheet21.Cells(LastRow, "T").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet21.Cells(LastRow, "T").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread4 = Sheet21.Cells(LastRow, "T").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet21.Cells(LastRow, "T").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread5 = Sheet21.Cells(LastRow, "T").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet21.Cells(LastRow, "T").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread6 = Sheet21.Cells(LastRow, "T").Value

    Case 36
Title = "Nissan L42P Assist: Root Hardness"
'Using Find Function

    LastRow = Sheet15.Range("A" & Rows.Count).End(xlUp).Row
    
'Create Values
        
        Spread1 = Sheet15.Cells(LastRow, "M").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet15.Cells(LastRow, "M").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread2 = Sheet15.Cells(LastRow, "M").Value

        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet15.Cells(LastRow, "M").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread3 = Sheet15.Cells(LastRow, "M").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet15.Cells(LastRow, "M").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread4 = Sheet15.Cells(LastRow, "M").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet15.Cells(LastRow, "M").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread5 = Sheet15.Cells(LastRow, "M").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet15.Cells(LastRow, "M").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread6 = Sheet15.Cells(LastRow, "M").Value

    Case 37
Title = "Nissan L42P Sensor: Root Hardness"
'Using Find Function

    LastRow = Sheet16.Range("A" & Rows.Count).End(xlUp).Row
    
'Create Values
        
        Spread1 = Sheet16.Cells(LastRow, "M").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet16.Cells(LastRow, "M").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread2 = Sheet16.Cells(LastRow, "M").Value

        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet16.Cells(LastRow, "M").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread3 = Sheet16.Cells(LastRow, "M").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet16.Cells(LastRow, "M").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread4 = Sheet16.Cells(LastRow, "M").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet16.Cells(LastRow, "M").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread5 = Sheet16.Cells(LastRow, "M").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet16.Cells(LastRow, "M").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread6 = Sheet16.Cells(LastRow, "M").Value
        
    Case 38
Title = "BMW LHD CGR Assist: Root Hardness"
'Using Find Function

    LastRow = Sheet1.Range("A" & Rows.Count).End(xlUp).Row
    
'Create Values
        
        Spread1 = Sheet1.Cells(LastRow, "M").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet1.Cells(LastRow, "M").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread2 = Sheet1.Cells(LastRow, "M").Value

        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet1.Cells(LastRow, "M").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread3 = Sheet1.Cells(LastRow, "M").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet1.Cells(LastRow, "M").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread4 = Sheet1.Cells(LastRow, "M").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet1.Cells(LastRow, "M").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread5 = Sheet1.Cells(LastRow, "M").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet1.Cells(LastRow, "M").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread6 = Sheet1.Cells(LastRow, "M").Value
        
    Case 39
Title = "BMW LHD CGR Sensor: Root Hardness"
'Using Find Function

    LastRow = Sheet2.Range("A" & Rows.Count).End(xlUp).Row
    
'Create Values
        
        Spread1 = Sheet2.Cells(LastRow, "M").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet2.Cells(LastRow, "M").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread2 = Sheet2.Cells(LastRow, "M").Value

        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet2.Cells(LastRow, "M").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread3 = Sheet2.Cells(LastRow, "M").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet2.Cells(LastRow, "M").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread4 = Sheet2.Cells(LastRow, "M").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet2.Cells(LastRow, "M").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread5 = Sheet2.Cells(LastRow, "M").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet2.Cells(LastRow, "M").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread6 = Sheet2.Cells(LastRow, "M").Value
        
    Case 40
Title = "BMW RHD CGR Assist: Root Hardness"
'Using Find Function

    LastRow = Sheet3.Range("A" & Rows.Count).End(xlUp).Row
    
'Create Values
        
        Spread1 = Sheet3.Cells(LastRow, "M").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet3.Cells(LastRow, "M").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread2 = Sheet3.Cells(LastRow, "M").Value

        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet3.Cells(LastRow, "M").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread3 = Sheet3.Cells(LastRow, "M").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet3.Cells(LastRow, "M").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread4 = Sheet3.Cells(LastRow, "M").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet3.Cells(LastRow, "M").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread5 = Sheet3.Cells(LastRow, "M").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet3.Cells(LastRow, "M").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread6 = Sheet3.Cells(LastRow, "M").Value
        
    Case 41
Title = "BMW RHD CGR Sensor: Root Hardness"
'Using Find Function

    LastRow = Sheet4.Range("A" & Rows.Count).End(xlUp).Row
    
'Create Values
        
        Spread1 = Sheet4.Cells(LastRow, "M").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet4.Cells(LastRow, "M").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread2 = Sheet4.Cells(LastRow, "M").Value

        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet4.Cells(LastRow, "M").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread3 = Sheet4.Cells(LastRow, "M").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet4.Cells(LastRow, "M").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread4 = Sheet4.Cells(LastRow, "M").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet4.Cells(LastRow, "M").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread5 = Sheet4.Cells(LastRow, "M").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet4.Cells(LastRow, "M").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread6 = Sheet4.Cells(LastRow, "M").Value
        
    Case 42
Title = "BMW LHD VGR Assist: Root Hardness"
'Using Find Function

    LastRow = Sheet5.Range("A" & Rows.Count).End(xlUp).Row
    
'Create Values
        
        Spread1 = Sheet5.Cells(LastRow, "M").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet5.Cells(LastRow, "M").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread2 = Sheet5.Cells(LastRow, "M").Value

        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet5.Cells(LastRow, "M").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread3 = Sheet5.Cells(LastRow, "M").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet5.Cells(LastRow, "M").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread4 = Sheet5.Cells(LastRow, "M").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet5.Cells(LastRow, "M").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread5 = Sheet5.Cells(LastRow, "M").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet5.Cells(LastRow, "M").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread6 = Sheet5.Cells(LastRow, "M").Value
        
    Case 43
Title = "BMW RHD VGR Assist: Root Hardness"
'Using Find Function

    LastRow = Sheet6.Range("A" & Rows.Count).End(xlUp).Row
    
'Create Values
        
        Spread1 = Sheet6.Cells(LastRow, "M").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet6.Cells(LastRow, "M").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread2 = Sheet6.Cells(LastRow, "M").Value

        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet6.Cells(LastRow, "M").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread3 = Sheet6.Cells(LastRow, "M").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet6.Cells(LastRow, "M").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread4 = Sheet6.Cells(LastRow, "M").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet6.Cells(LastRow, "M").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread5 = Sheet6.Cells(LastRow, "M").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet6.Cells(LastRow, "M").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread6 = Sheet6.Cells(LastRow, "M").Value
        
    Case 44
Title = "Honda THR Assist: Root Hardness"
'Using Find Function

    LastRow = Sheet7.Range("A" & Rows.Count).End(xlUp).Row
    
'Create Values
        
        Spread1 = Sheet7.Cells(LastRow, "M").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet7.Cells(LastRow, "M").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread2 = Sheet7.Cells(LastRow, "M").Value

        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet7.Cells(LastRow, "M").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread3 = Sheet7.Cells(LastRow, "M").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet7.Cells(LastRow, "M").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread4 = Sheet7.Cells(LastRow, "M").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet7.Cells(LastRow, "M").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread5 = Sheet7.Cells(LastRow, "M").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet7.Cells(LastRow, "M").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread6 = Sheet7.Cells(LastRow, "M").Value
        
    Case 45
Title = "Honda THR Sensor: Root Hardness"
'Using Find Function

    LastRow = Sheet8.Range("A" & Rows.Count).End(xlUp).Row
    
'Create Values
        
        Spread1 = Sheet8.Cells(LastRow, "M").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet8.Cells(LastRow, "M").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread2 = Sheet8.Cells(LastRow, "M").Value

        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet8.Cells(LastRow, "M").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread3 = Sheet8.Cells(LastRow, "M").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet8.Cells(LastRow, "M").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread4 = Sheet8.Cells(LastRow, "M").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet8.Cells(LastRow, "M").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread5 = Sheet8.Cells(LastRow, "M").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet8.Cells(LastRow, "M").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread6 = Sheet8.Cells(LastRow, "M").Value
        
    Case 46
Title = "Honda TJB Assist: Root Hardness"
'Using Find Function

    LastRow = Sheet17.Range("A" & Rows.Count).End(xlUp).Row
    
'Create Values
        
        Spread1 = Sheet17.Cells(LastRow, "M").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet17.Cells(LastRow, "M").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread2 = Sheet17.Cells(LastRow, "M").Value

        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet17.Cells(LastRow, "M").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread3 = Sheet17.Cells(LastRow, "M").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet17.Cells(LastRow, "M").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread4 = Sheet17.Cells(LastRow, "M").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet17.Cells(LastRow, "M").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread5 = Sheet17.Cells(LastRow, "M").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet17.Cells(LastRow, "M").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread6 = Sheet17.Cells(LastRow, "M").Value
        
    Case 47
Title = "09PL Rack Side: Root Hardness"
'Using Find Function

    LastRow = Sheet13.Range("A" & Rows.Count).End(xlUp).Row
    
'Create Values
        
        Spread1 = Sheet13.Cells(LastRow, "M").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet13.Cells(LastRow, "M").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread2 = Sheet13.Cells(LastRow, "M").Value

        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet13.Cells(LastRow, "M").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread3 = Sheet13.Cells(LastRow, "M").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet13.Cells(LastRow, "M").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread4 = Sheet13.Cells(LastRow, "M").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet13.Cells(LastRow, "M").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread5 = Sheet13.Cells(LastRow, "M").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet13.Cells(LastRow, "M").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread6 = Sheet13.Cells(LastRow, "M").Value
        
    Case 48
Title = "BMW G2X Assist: Root Hardness"
'Using Find Function

    LastRow = Sheet18.Range("A" & Rows.Count).End(xlUp).Row
    
'Create Values
        
        Spread1 = Sheet18.Cells(LastRow, "M").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet18.Cells(LastRow, "M").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread2 = Sheet18.Cells(LastRow, "M").Value

        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet18.Cells(LastRow, "M").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread3 = Sheet18.Cells(LastRow, "M").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet18.Cells(LastRow, "M").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread4 = Sheet18.Cells(LastRow, "M").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet18.Cells(LastRow, "M").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread5 = Sheet18.Cells(LastRow, "M").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet18.Cells(LastRow, "M").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread6 = Sheet18.Cells(LastRow, "M").Value
        
    Case 49
Title = "BMW G2X Sensor: Root Hardness"
'Using Find Function

    LastRow = Sheet19.Range("A" & Rows.Count).End(xlUp).Row
    
'Create Values
        
        Spread1 = Sheet19.Cells(LastRow, "M").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet19.Cells(LastRow, "M").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread2 = Sheet19.Cells(LastRow, "M").Value

        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet19.Cells(LastRow, "M").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread3 = Sheet19.Cells(LastRow, "M").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet19.Cells(LastRow, "M").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread4 = Sheet19.Cells(LastRow, "M").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet19.Cells(LastRow, "M").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread5 = Sheet19.Cells(LastRow, "M").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet19.Cells(LastRow, "M").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread6 = Sheet19.Cells(LastRow, "M").Value
        
    Case 50
Title = "015PL Rack Side: Root Hardness"
'Using Find Function

    LastRow = Sheet20.Range("A" & Rows.Count).End(xlUp).Row
    
'Create Values
        
        Spread1 = Sheet20.Cells(LastRow, "M").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet20.Cells(LastRow, "M").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread2 = Sheet20.Cells(LastRow, "M").Value

        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet20.Cells(LastRow, "M").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread3 = Sheet20.Cells(LastRow, "M").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet20.Cells(LastRow, "M").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread4 = Sheet20.Cells(LastRow, "M").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet20.Cells(LastRow, "M").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread5 = Sheet20.Cells(LastRow, "M").Value
        
        LastRow = LastRow - 1
        
        Do Until IsNumeric(Sheet20.Cells(LastRow, "M").Value) = True
        
        LastRow = LastRow - 1
        
        Loop
        
        Spread6 = Sheet20.Cells(LastRow, "M").Value

End Select

If Spread1 <> "" Then
    If Abs(Spread1 - Spread2) > 30 Or Abs(Spread2 - Spread3) > 30 Or Abs(Spread3 - Spread4) > 30 Or Abs(Spread4 - Spread5) > 30 Or Abs(Spread5 - Spread6) > 30 Then
        .Content.InsertAfter (Title)
        .Content.InsertAfter (" ! ! ! ! ! ! ! ! ! ! ")
        .Content.InsertAfter ("Abnormality: Hardness Variance")
        .Content.InsertParagraphAfter
    End If
End If

If Spread1 <> "" Then
    If Spread1 < 7 Or Spread2 < 7 Or Spread3 < 7 Or Spread4 < 7 Or Spread5 < 7 Or Spread6 < 7 Then
        If Abs(Spread1 - Spread2) > 0.2 Or Abs(Spread2 - Spread3) > 0.2 Or Abs(Spread3 - Spread4) > 0.2 Or Abs(Spread4 - Spread5) > 0.2 Or Abs(Spread5 - Spread6) > 0.2 Then
            .Content.InsertAfter (Title)
            .Content.InsertAfter (" ! ! ! ! ! ! ! ! ! ! ")
            .Content.InsertAfter ("Abnormality: Case Depth Variance")
            .Content.InsertParagraphAfter
        End If
    End If
End If

'XBar = (1 + 2 + 3 + 4 + 5 + 6) / 6
'YBar = (Spread1 + Spread2 + Spread3 + Spread4 + Spread5 + Spread6) / 6
'i = ((1 - XBar) * (Spread6 - YBar)) + ((2 - XBar) * (Spread5 - YBar)) + ((3 - XBar) * (Spread4 - YBar)) + ((4 - XBar) * (Spread3 - YBar)) + ((5 - XBar) * (Spread2 - YBar)) + ((6 - XBar) * (Spread1 - YBar))
'j = (((1 - XBar) * (1 - XBar)) + ((2 - XBar) * (2 - XBar)) + ((3 - XBar) * (3 - XBar)) + ((4 - XBar) * (4 - XBar)) + ((5 - XBar) * (5 - XBar)) + ((6 - XBar) * (6 - XBar)))
'm = i / j

'If m > 0 Then
    'MsgBox ("Upward Trend")
'End If

'If m < 0 Then
    'MsgBox ("Downward Trend")
'End If

        Next i

    End With
    
    Set wrdDoc = Nothing
    Set wrdApp = Nothing

End Sub
