VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "UserForm2"
   ClientHeight    =   8028
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   12504
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CheckBox1_Click()
'Hardness Upper Limit

If CheckBox1.Value = False Then

    UserForm2.Image2.BackStyle = fmBackStyleTransparent

End If

If CheckBox1.Value = True Then

    UserForm2.Image2.BackStyle = fmBackStyleOpaque

End If

End Sub

Private Sub CheckBox2_Click()
'Hardness Lower Limit

If CheckBox2.Value = False Then

    UserForm2.Image3.BackStyle = fmBackStyleTransparent

End If

If CheckBox2.Value = True Then

    UserForm2.Image3.BackStyle = fmBackStyleOpaque

End If

End Sub

Private Sub CheckBox3_Click()
'Case Depth Upper Limit

If CheckBox3.Value = False Then

    UserForm2.Image4.BackStyle = fmBackStyleTransparent

End If

If CheckBox3.Value = True Then

    UserForm2.Image4.BackStyle = fmBackStyleOpaque

End If

End Sub

Private Sub CheckBox4_Click()
'Case Depth Upper Limit

If CheckBox4.Value = False Then

    UserForm2.Image5.BackStyle = fmBackStyleTransparent

End If

If CheckBox4.Value = True Then

    UserForm2.Image5.BackStyle = fmBackStyleOpaque

End If

End Sub

Private Sub CommandButton2_Click()

If ComboBox1.Text = "Select a chart" Then

    MsgBox ("Select a chart from the dropdown list")
    Exit Sub
    
End If

CheckBox1.Value = False
CheckBox2.Value = False
CheckBox3.Value = False
CheckBox4.Value = False

Dim LastRow As Long
Dim MyChart As Chart
Dim ChartData As Range
Dim ChartIndex As Integer
Dim ChartName As String
Dim imageName As String
Dim FirstRow As Long
Dim ChartZero As Double

ChartZero = 0#

ChartIndex = ComboBox1.ListIndex

Select Case ChartIndex

    'Nissan L42P Assist Case Depth
    Case 0
        
    'Using Find Function
    
    LastRow = Sheet15.Range("A" & Rows.Count).End(xlUp).Row
    FirstRow = LastRow - 25
    
    If FirstRow < 10 Then
        FirstRow = 10
    End If
        
        Set ChartData = Sheet15.Range("J" & FirstRow & ":J" & LastRow)
        ChartName = Sheet15.Range("J5")
    
    Application.ScreenUpdating = False
    
    Set MyChart = Sheet15.Shapes.AddChart(xlLineMarkers).Chart
    
    Do While MyChart.SeriesCollection.Count > 0
    
    MyChart.FullSeriesCollection(1).Delete
    
    Loop
    
    MyChart.SeriesCollection.NewSeries
    MyChart.SeriesCollection(1).Name = ChartName
    MyChart.SeriesCollection(1).Values = ChartData
    
    'Scale
    MyChart.Axes(xlValue).MinimumScale = 4
    MyChart.Axes(xlValue).MaximumScale = 6.2
    
    'Zoom
    Sheet15.Activate
    ActiveWindow.Zoom = 125
    Sheet14.Activate
    
    'Create Image
    imageName = Application.DefaultFilePath & Application.PathSeparator & "TempChart.gif"
    MyChart.Export Filename:=imageName, FilterName:="GIF"
    Sheet15.ChartObjects(1).Delete
    Application.ScreenUpdating = True
    UserForm2.Image1.Picture = LoadPicture(imageName)
        
    'Nissan L42P Sensor Case Depth
    Case 1
    
    'Using Find Function
    
    LastRow = Sheet16.Range("A" & Rows.Count).End(xlUp).Row
    FirstRow = LastRow - 25
    
    If FirstRow < 10 Then
        FirstRow = 10
    End If
        
        Set ChartData = Sheet16.Range("J" & FirstRow & ":J" & LastRow)
        ChartName = Sheet16.Range("J5")
        
    Application.ScreenUpdating = False
    
    ChartData.Replace What:="", Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    ChartData.Replace What:=ChartZero, Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    
    Set MyChart = Sheet16.Shapes.AddChart(xlLineMarkers).Chart

    Do While MyChart.SeriesCollection.Count > 0
    
    MyChart.FullSeriesCollection(1).Delete
    
    Loop
            
    MyChart.SeriesCollection.NewSeries
    MyChart.SeriesCollection(1).Name = ChartName
    MyChart.SeriesCollection(1).Values = ChartData
    
    'Scale
    MyChart.Axes(xlValue).MinimumScale = 4
    MyChart.Axes(xlValue).MaximumScale = 6.2
    
    'Zoom
    Sheet16.Activate
    ActiveWindow.Zoom = 125
    Sheet14.Activate
    
    'Create Image
    imageName = Application.DefaultFilePath & Application.PathSeparator & "TempChart.gif"
    MyChart.Export Filename:=imageName, FilterName:="GIF"
    Sheet16.ChartObjects(1).Delete
    Application.ScreenUpdating = True
    UserForm2.Image1.Picture = LoadPicture(imageName)
                   
    'Nissan L42P Assist Hardness
    Case 2
    
    'Using Find Function
    
    LastRow = Sheet15.Range("A" & Rows.Count).End(xlUp).Row
    FirstRow = LastRow - 25
    
    If FirstRow < 10 Then
        FirstRow = 10
    End If
        
        Set ChartData = Sheet15.Range("I" & FirstRow & ":I" & LastRow)
        ChartName = Sheet15.Range("H5")
        
    Application.ScreenUpdating = False
    
    ChartData.Replace What:="", Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    ChartData.Replace What:=ChartZero, Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    
    Set MyChart = Sheet15.Shapes.AddChart(xlLineMarkers).Chart
    
    Do While MyChart.SeriesCollection.Count > 0
    
    MyChart.FullSeriesCollection(1).Delete
    
    Loop

    MyChart.SeriesCollection.NewSeries
    MyChart.SeriesCollection(1).Name = ChartName
    MyChart.SeriesCollection(1).Values = ChartData
    
    'Scale
    MyChart.Axes(xlValue).MinimumScale = 580
    MyChart.Axes(xlValue).MaximumScale = 735
    
    'Zoom
    Sheet15.Activate
    ActiveWindow.Zoom = 125
    Sheet14.Activate
    
    'Create Image
    imageName = Application.DefaultFilePath & Application.PathSeparator & "TempChart.gif"
    MyChart.Export Filename:=imageName, FilterName:="GIF"
    Sheet15.ChartObjects(1).Delete
    Application.ScreenUpdating = True
    UserForm2.Image1.Picture = LoadPicture(imageName)
        
    'Nissan L42P Sensor Hardness
    Case 3
    
    'Using Find Function
    
    LastRow = Sheet16.Range("A" & Rows.Count).End(xlUp).Row
    FirstRow = LastRow - 25
    
    If FirstRow < 10 Then
        FirstRow = 10
    End If
        
        Set ChartData = Sheet16.Range("I" & FirstRow & ":I" & LastRow)
        ChartName = Sheet16.Range("H5")
        
    Application.ScreenUpdating = False
    
    ChartData.Replace What:="", Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    ChartData.Replace What:=ChartZero, Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    
    Set MyChart = Sheet16.Shapes.AddChart(xlLineMarkers).Chart
    
    Do While MyChart.SeriesCollection.Count > 0
    
    MyChart.FullSeriesCollection(1).Delete
    
    Loop
    
    MyChart.SeriesCollection.NewSeries
    MyChart.SeriesCollection(1).Name = ChartName
    MyChart.SeriesCollection(1).Values = ChartData
    
    'Scale
    MyChart.Axes(xlValue).MinimumScale = 580
    MyChart.Axes(xlValue).MaximumScale = 735
    
    'Zoom
    Sheet16.Activate
    ActiveWindow.Zoom = 125
    Sheet14.Activate
    
    'Create Image
    
    imageName = Application.DefaultFilePath & Application.PathSeparator & "TempChart.gif"
    MyChart.Export Filename:=imageName, FilterName:="GIF"
    Sheet16.ChartObjects(1).Delete
    Application.ScreenUpdating = True
    UserForm2.Image1.Picture = LoadPicture(imageName)
              
    'BMW LHD CGR Assist Case Depth
    Case 4
    
    'Using Find Function
    
    LastRow = Sheet1.Range("A" & Rows.Count).End(xlUp).Row
    FirstRow = LastRow - 25
    
    If FirstRow < 10 Then
        FirstRow = 10
    End If
        
        Set ChartData = Sheet1.Range("J" & FirstRow & ":J" & LastRow)
        ChartName = Sheet1.Range("J5")
        
    Application.ScreenUpdating = False
    
    ChartData.Replace What:="", Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    ChartData.Replace What:=ChartZero, Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    
    Set MyChart = Sheet1.Shapes.AddChart(xlLineMarkers).Chart
    
    Do While MyChart.SeriesCollection.Count > 0
    
    MyChart.FullSeriesCollection(1).Delete
    
    Loop
    
    MyChart.SeriesCollection.NewSeries
    MyChart.SeriesCollection(1).Name = ChartName
    MyChart.SeriesCollection(1).Values = ChartData
    
    'Scale
    MyChart.Axes(xlValue).MinimumScale = 4
    MyChart.Axes(xlValue).MaximumScale = 6.2
    
    'Zoom
    Sheet1.Activate
    ActiveWindow.Zoom = 125
    Sheet14.Activate
    
    'Create Image
    imageName = Application.DefaultFilePath & Application.PathSeparator & "TempChart.gif"
    MyChart.Export Filename:=imageName, FilterName:="GIF"
    Sheet1.ChartObjects(1).Delete
    Application.ScreenUpdating = True
    UserForm2.Image1.Picture = LoadPicture(imageName)
        
    'BMW LHD CGR Sensor Case Depth
    Case 5
    
    'Using Find Function
    
    LastRow = Sheet2.Range("A" & Rows.Count).End(xlUp).Row
    FirstRow = LastRow - 25
    
    If FirstRow < 10 Then
        FirstRow = 10
    End If
        
        Set ChartData = Sheet2.Range("J" & FirstRow & ":J" & LastRow)
        ChartName = Sheet2.Range("J5")
        
    Application.ScreenUpdating = False
    
    ChartData.Replace What:="", Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    ChartData.Replace What:=ChartZero, Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    
    Set MyChart = Sheet2.Shapes.AddChart(xlLineMarkers).Chart
    
    Do While MyChart.SeriesCollection.Count > 0
    
    MyChart.FullSeriesCollection(1).Delete
    
    Loop
    
    MyChart.SeriesCollection.NewSeries
    MyChart.SeriesCollection(1).Name = ChartName
    MyChart.SeriesCollection(1).Values = ChartData
        
    'Scale
    MyChart.Axes(xlValue).MinimumScale = 4
    MyChart.Axes(xlValue).MaximumScale = 6.2
    
    'Zoom
    Sheet2.Activate
    ActiveWindow.Zoom = 125
    Sheet14.Activate
    
    'Create Image
    imageName = Application.DefaultFilePath & Application.PathSeparator & "TempChart.gif"
    MyChart.Export Filename:=imageName, FilterName:="GIF"
    Sheet2.ChartObjects(1).Delete
    Application.ScreenUpdating = True
    UserForm2.Image1.Picture = LoadPicture(imageName)
        
    'BMW LHD CGR Assist Hardness
    Case 6
    
    'Using Find Function
    
    LastRow = Sheet1.Range("A" & Rows.Count).End(xlUp).Row
    FirstRow = LastRow - 25
    
    If FirstRow < 10 Then
        FirstRow = 10
    End If
        
        Set ChartData = Sheet1.Range("I" & FirstRow & ":I" & LastRow)
        ChartName = Sheet1.Range("H5")
        
    Application.ScreenUpdating = False
    
    ChartData.Replace What:="", Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    ChartData.Replace What:=ChartZero, Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    
    Set MyChart = Sheet1.Shapes.AddChart(xlLineMarkers).Chart
    
    Do While MyChart.SeriesCollection.Count > 0
    
    MyChart.FullSeriesCollection(1).Delete
    
    Loop
    
    MyChart.SeriesCollection.NewSeries
    MyChart.SeriesCollection(1).Name = ChartName
    MyChart.SeriesCollection(1).Values = ChartData
        
    'Scale
    MyChart.Axes(xlValue).MinimumScale = 580
    MyChart.Axes(xlValue).MaximumScale = 735
    
    'Zoom
    Sheet1.Activate
    ActiveWindow.Zoom = 125
    Sheet14.Activate
    
    'Create Image
    imageName = Application.DefaultFilePath & Application.PathSeparator & "TempChart.gif"
    MyChart.Export Filename:=imageName, FilterName:="GIF"
    Sheet1.ChartObjects(1).Delete
    Application.ScreenUpdating = True
    UserForm2.Image1.Picture = LoadPicture(imageName)
    
    'BMW LHD CGR Sensor Hardness
    Case 7
    
    'Using Find Function
    
    LastRow = Sheet2.Range("A" & Rows.Count).End(xlUp).Row
    FirstRow = LastRow - 25
    
    If FirstRow < 10 Then
        FirstRow = 10
    End If
        
        Set ChartData = Sheet2.Range("I" & FirstRow & ":I" & LastRow)
        ChartName = Sheet2.Range("H5")
        
    Application.ScreenUpdating = False
    
    ChartData.Replace What:="", Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    ChartData.Replace What:=ChartZero, Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    
    Set MyChart = Sheet2.Shapes.AddChart(xlLineMarkers).Chart
    
    Do While MyChart.SeriesCollection.Count > 0
    
    MyChart.FullSeriesCollection(1).Delete
    
    Loop
    
    MyChart.SeriesCollection.NewSeries
    MyChart.SeriesCollection(1).Name = ChartName
    MyChart.SeriesCollection(1).Values = ChartData
    
    'Scale
    MyChart.Axes(xlValue).MinimumScale = 580
    MyChart.Axes(xlValue).MaximumScale = 735
    
    'Zoom
    Sheet2.Activate
    ActiveWindow.Zoom = 125
    Sheet14.Activate
    
    'Create Image
    imageName = Application.DefaultFilePath & Application.PathSeparator & "TempChart.gif"
    MyChart.Export Filename:=imageName, FilterName:="GIF"
    Sheet2.ChartObjects(1).Delete
    Application.ScreenUpdating = True
    UserForm2.Image1.Picture = LoadPicture(imageName)
                   
    'BMW RHD CGR Assist: Case Depth
    Case 8
    
    'Using Find Function
    
    LastRow = Sheet3.Range("A" & Rows.Count).End(xlUp).Row
    FirstRow = LastRow - 25
    
    If FirstRow < 10 Then
        FirstRow = 10
    End If
        
        Set ChartData = Sheet3.Range("J" & FirstRow & ":J" & LastRow)
        ChartName = Sheet3.Range("J5")
        
    Application.ScreenUpdating = False
    
    ChartData.Replace What:="", Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    ChartData.Replace What:=ChartZero, Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    
    Set MyChart = Sheet3.Shapes.AddChart(xlLineMarkers).Chart
    
    Do While MyChart.SeriesCollection.Count > 0
    
    MyChart.FullSeriesCollection(1).Delete
    
    Loop
    
    MyChart.SeriesCollection.NewSeries
    MyChart.SeriesCollection(1).Name = ChartName
    MyChart.SeriesCollection(1).Values = ChartData
        
    'Scale
    MyChart.Axes(xlValue).MinimumScale = 4
    MyChart.Axes(xlValue).MaximumScale = 6.2
    
    'Zoom
    Sheet3.Activate
    ActiveWindow.Zoom = 125
    Sheet14.Activate
    
    'Create Image
    imageName = Application.DefaultFilePath & Application.PathSeparator & "TempChart.gif"
    MyChart.Export Filename:=imageName, FilterName:="GIF"
    Sheet3.ChartObjects(1).Delete
    Application.ScreenUpdating = True
    UserForm2.Image1.Picture = LoadPicture(imageName)
                   
    'BMW RHD CGR Sensor: Case Depth
    Case 9
    
    'Using Find Function
    
    LastRow = Sheet4.Range("A" & Rows.Count).End(xlUp).Row
    FirstRow = LastRow - 25
    
    If FirstRow < 10 Then
        FirstRow = 10
    End If
        
        Set ChartData = Sheet4.Range("J" & FirstRow & ":J" & LastRow)
        ChartName = Sheet4.Range("J5")
        
    Application.ScreenUpdating = False
    
    ChartData.Replace What:="", Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    ChartData.Replace What:=ChartZero, Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    
    Set MyChart = Sheet4.Shapes.AddChart(xlLineMarkers).Chart
    
    Do While MyChart.SeriesCollection.Count > 0
    
    MyChart.FullSeriesCollection(1).Delete
    
    Loop
    
    MyChart.SeriesCollection.NewSeries
    MyChart.SeriesCollection(1).Name = ChartName
    MyChart.SeriesCollection(1).Values = ChartData
    
    'Scale
    MyChart.Axes(xlValue).MinimumScale = 4
    MyChart.Axes(xlValue).MaximumScale = 6.2
    
    'Zoom
    Sheet4.Activate
    ActiveWindow.Zoom = 125
    Sheet14.Activate
    
    'Create Image
    imageName = Application.DefaultFilePath & Application.PathSeparator & "TempChart.gif"
    MyChart.Export Filename:=imageName, FilterName:="GIF"
    Sheet4.ChartObjects(1).Delete
    Application.ScreenUpdating = True
    UserForm2.Image1.Picture = LoadPicture(imageName)
                   
    'BMW RHD CGR Assist: Hardness
    Case 10
    
    'Using Find Function
    
    LastRow = Sheet3.Range("A" & Rows.Count).End(xlUp).Row
    FirstRow = LastRow - 25
    
    If FirstRow < 10 Then
        FirstRow = 10
    End If
            
        Set ChartData = Sheet3.Range("I" & FirstRow & ":I" & LastRow)
        ChartName = Sheet3.Range("H5")
        
    Application.ScreenUpdating = False
    
    ChartData.Replace What:="", Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    ChartData.Replace What:=ChartZero, Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    
    Set MyChart = Sheet3.Shapes.AddChart(xlLineMarkers).Chart
    
    Do While MyChart.SeriesCollection.Count > 0
    
    MyChart.FullSeriesCollection(1).Delete
    
    Loop
    
    MyChart.SeriesCollection.NewSeries
    MyChart.SeriesCollection(1).Name = ChartName
    MyChart.SeriesCollection(1).Values = ChartData
    
    'Scale
    MyChart.Axes(xlValue).MinimumScale = 580
    MyChart.Axes(xlValue).MaximumScale = 735
    
    'Zoom
    Sheet3.Activate
    ActiveWindow.Zoom = 125
    Sheet14.Activate
    
    'Create Image
    imageName = Application.DefaultFilePath & Application.PathSeparator & "TempChart.gif"
    MyChart.Export Filename:=imageName, FilterName:="GIF"
    Sheet3.ChartObjects(1).Delete
    Application.ScreenUpdating = True
    UserForm2.Image1.Picture = LoadPicture(imageName)
                   
    'BMW RHD CGR Sensor: Hardness
    Case 11
    
    'Using Find Function
    
    LastRow = Sheet4.Range("A" & Rows.Count).End(xlUp).Row
    FirstRow = LastRow - 25
    
    If FirstRow < 10 Then
        FirstRow = 10
    End If
        
        Set ChartData = Sheet4.Range("I" & FirstRow & ":I" & LastRow)
        ChartName = Sheet4.Range("H5")
        
    Application.ScreenUpdating = False
    
    ChartData.Replace What:="", Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    ChartData.Replace What:=ChartZero, Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    
    Set MyChart = Sheet4.Shapes.AddChart(xlLineMarkers).Chart
    
    Do While MyChart.SeriesCollection.Count > 0
    
    MyChart.FullSeriesCollection(1).Delete
    
    Loop
    
    MyChart.SeriesCollection.NewSeries
    MyChart.SeriesCollection(1).Name = ChartName
    MyChart.SeriesCollection(1).Values = ChartData
    
    'Scale
    MyChart.Axes(xlValue).MinimumScale = 580
    MyChart.Axes(xlValue).MaximumScale = 735
    
    'Zoom
    Sheet4.Activate
    ActiveWindow.Zoom = 125
    Sheet14.Activate
    
    'Create Image
    imageName = Application.DefaultFilePath & Application.PathSeparator & "TempChart.gif"
    MyChart.Export Filename:=imageName, FilterName:="GIF"
    Sheet4.ChartObjects(1).Delete
    Application.ScreenUpdating = True
    UserForm2.Image1.Picture = LoadPicture(imageName)
                   
    'BMW LHD VGR Assist: Case Depth
    Case 12
    
    'Using Find Function
    
    LastRow = Sheet5.Range("A" & Rows.Count).End(xlUp).Row
    FirstRow = LastRow - 25
    
    If FirstRow < 10 Then
        FirstRow = 10
    End If
        
        Set ChartData = Sheet5.Range("J" & FirstRow & ":J" & LastRow)
        ChartName = Sheet5.Range("J5")
        
    Application.ScreenUpdating = False
    
    ChartData.Replace What:="", Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    ChartData.Replace What:=ChartZero, Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    
    Set MyChart = Sheet5.Shapes.AddChart(xlLineMarkers).Chart
     
    Do While MyChart.SeriesCollection.Count > 0
    
    MyChart.FullSeriesCollection(1).Delete
    
    Loop
    
    MyChart.SeriesCollection.NewSeries
    MyChart.SeriesCollection(1).Name = ChartName
    MyChart.SeriesCollection(1).Values = ChartData

    'Scale
    MyChart.Axes(xlValue).MinimumScale = 4
    MyChart.Axes(xlValue).MaximumScale = 6.2
    
    'Zoom
    Sheet5.Activate
    ActiveWindow.Zoom = 125
    Sheet14.Activate
    
    'Create Image
    imageName = Application.DefaultFilePath & Application.PathSeparator & "TempChart.gif"
    MyChart.Export Filename:=imageName, FilterName:="GIF"
    Sheet5.ChartObjects(1).Delete
    Application.ScreenUpdating = True
    UserForm2.Image1.Picture = LoadPicture(imageName)
                   
    'BMW LHD VGR Assist: Hardness
    Case 13
    
    'Using Find Function
    
    LastRow = Sheet5.Range("A" & Rows.Count).End(xlUp).Row
    FirstRow = LastRow - 25
    
    If FirstRow < 10 Then
        FirstRow = 10
    End If
        
        Set ChartData = Sheet5.Range("I" & FirstRow & ":I" & LastRow)
        ChartName = Sheet5.Range("H5")
        
    Application.ScreenUpdating = False
    
    ChartData.Replace What:="", Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    ChartData.Replace What:=ChartZero, Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    
    Set MyChart = Sheet5.Shapes.AddChart(xlLineMarkers).Chart
    
    Do While MyChart.SeriesCollection.Count > 0
    
    MyChart.FullSeriesCollection(1).Delete
    
    Loop
    
    MyChart.SeriesCollection.NewSeries
    MyChart.SeriesCollection(1).Name = ChartName
    MyChart.SeriesCollection(1).Values = ChartData
    
    'Scale
    MyChart.Axes(xlValue).MinimumScale = 530
    MyChart.Axes(xlValue).MaximumScale = 665
    
    'Zoom
    Sheet5.Activate
    ActiveWindow.Zoom = 125
    Sheet14.Activate
    
    'Create Image
    imageName = Application.DefaultFilePath & Application.PathSeparator & "TempChart.gif"
    MyChart.Export Filename:=imageName, FilterName:="GIF"
    Sheet5.ChartObjects(1).Delete
    Application.ScreenUpdating = True
    UserForm2.Image1.Picture = LoadPicture(imageName)
                   
    'BMW RHD VGR Assist: Case Depth
    Case 14
    
    'Using Find Function
    
    LastRow = Sheet6.Range("A" & Rows.Count).End(xlUp).Row
    FirstRow = LastRow - 25
    
    If FirstRow < 10 Then
        FirstRow = 10
    End If
        
        Set ChartData = Sheet6.Range("J" & FirstRow & ":J" & LastRow)
        ChartName = Sheet6.Range("J5")
        
    Application.ScreenUpdating = False
    
    ChartData.Replace What:="", Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    ChartData.Replace What:=ChartZero, Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    
    Set MyChart = Sheet6.Shapes.AddChart(xlLineMarkers).Chart
    
    Do While MyChart.SeriesCollection.Count > 0
    
    MyChart.FullSeriesCollection(1).Delete
    
    Loop
    
    MyChart.SeriesCollection.NewSeries
    MyChart.SeriesCollection(1).Name = ChartName
    MyChart.SeriesCollection(1).Values = ChartData
    
    'Scale
    MyChart.Axes(xlValue).MinimumScale = 4
    MyChart.Axes(xlValue).MaximumScale = 6.2
    
    'Zoom
    Sheet6.Activate
    ActiveWindow.Zoom = 125
    Sheet14.Activate
    
    'Create Image
    imageName = Application.DefaultFilePath & Application.PathSeparator & "TempChart.gif"
    MyChart.Export Filename:=imageName, FilterName:="GIF"
    Sheet6.ChartObjects(1).Delete
    Application.ScreenUpdating = True
    UserForm2.Image1.Picture = LoadPicture(imageName)
                   
    'BMW RHD VGR Assist: Hardness
    Case 15
    
    'Using Find Function
    
    LastRow = Sheet6.Range("A" & Rows.Count).End(xlUp).Row
    FirstRow = LastRow - 25
    
    If FirstRow < 10 Then
        FirstRow = 10
    End If
            
        Set ChartData = Sheet6.Range("I" & FirstRow & ":I" & LastRow)
        ChartName = Sheet6.Range("H5")
        
    Application.ScreenUpdating = False
    
    ChartData.Replace What:="", Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    ChartData.Replace What:=ChartZero, Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    
    Set MyChart = Sheet6.Shapes.AddChart(xlLineMarkers).Chart
    
    Do While MyChart.SeriesCollection.Count > 0
    
    MyChart.FullSeriesCollection(1).Delete
    
    Loop
    
    MyChart.SeriesCollection.NewSeries
    MyChart.SeriesCollection(1).Name = ChartName
    MyChart.SeriesCollection(1).Values = ChartData
    
    'Scale
    MyChart.Axes(xlValue).MinimumScale = 530
    MyChart.Axes(xlValue).MaximumScale = 665
    
    'Zoom
    Sheet6.Activate
    ActiveWindow.Zoom = 125
    Sheet14.Activate
    
    'Create Image
    imageName = Application.DefaultFilePath & Application.PathSeparator & "TempChart.gif"
    MyChart.Export Filename:=imageName, FilterName:="GIF"
    Sheet6.ChartObjects(1).Delete
    Application.ScreenUpdating = True
    UserForm2.Image1.Picture = LoadPicture(imageName)
                   
    'Honda THR Assist: Case Depth
    Case 16
    
    'Using Find Function
    
    LastRow = Sheet7.Range("A" & Rows.Count).End(xlUp).Row
    FirstRow = LastRow - 25
    
    If FirstRow < 10 Then
        FirstRow = 10
    End If
        
        Set ChartData = Sheet7.Range("J" & FirstRow & ":J" & LastRow)
        ChartName = Sheet7.Range("J5")
        
    Application.ScreenUpdating = False
    
    ChartData.Replace What:="", Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    ChartData.Replace What:=ChartZero, Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    
    Set MyChart = Sheet7.Shapes.AddChart(xlLineMarkers).Chart
    
    Do While MyChart.SeriesCollection.Count > 0
    
    MyChart.FullSeriesCollection(1).Delete
    
    Loop
    
    MyChart.SeriesCollection.NewSeries
    MyChart.SeriesCollection(1).Name = ChartName
    MyChart.SeriesCollection(1).Values = ChartData
    
    'Scale
    MyChart.Axes(xlValue).MinimumScale = 4
    MyChart.Axes(xlValue).MaximumScale = 6.2
    
    'Zoom
    Sheet7.Activate
    ActiveWindow.Zoom = 125
    Sheet14.Activate
    
    'Create Image
    imageName = Application.DefaultFilePath & Application.PathSeparator & "TempChart.gif"
    MyChart.Export Filename:=imageName, FilterName:="GIF"
    Sheet7.ChartObjects(1).Delete
    Application.ScreenUpdating = True
    UserForm2.Image1.Picture = LoadPicture(imageName)
                   
    'Honda THR Sensor: Case Depth
    Case 17
    
    'Using Find Function
    
    LastRow = Sheet8.Range("A" & Rows.Count).End(xlUp).Row
    FirstRow = LastRow - 25
    
    If FirstRow < 10 Then
        FirstRow = 10
    End If
        
        Set ChartData = Sheet8.Range("J" & FirstRow & ":J" & LastRow)
        ChartName = Sheet8.Range("J6")
        
    Application.ScreenUpdating = False
    
    ChartData.Replace What:="", Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    ChartData.Replace What:=ChartZero, Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    
    Set MyChart = Sheet8.Shapes.AddChart(xlLineMarkers).Chart
    
    Do While MyChart.SeriesCollection.Count > 0
    
    MyChart.FullSeriesCollection(1).Delete
    
    Loop
    
    MyChart.SeriesCollection.NewSeries
    MyChart.SeriesCollection(1).Name = ChartName
    MyChart.SeriesCollection(1).Values = ChartData
    
    'Scale
    MyChart.Axes(xlValue).MinimumScale = 4
    MyChart.Axes(xlValue).MaximumScale = 6.2
    
    'Zoom
    Sheet8.Activate
    ActiveWindow.Zoom = 125
    Sheet14.Activate
    
    'Create Image
    imageName = Application.DefaultFilePath & Application.PathSeparator & "TempChart.gif"
    MyChart.Export Filename:=imageName, FilterName:="GIF"
    Sheet8.ChartObjects(1).Delete
    Application.ScreenUpdating = True
    UserForm2.Image1.Picture = LoadPicture(imageName)
                   
    'Honda THR Assist: Hardness
    Case 18
    
    'Using Find Function
    
    LastRow = Sheet7.Range("A" & Rows.Count).End(xlUp).Row
    FirstRow = LastRow - 25
    
    If FirstRow < 10 Then
        FirstRow = 10
    End If
        
        Set ChartData = Sheet7.Range("I" & FirstRow & ":I" & LastRow)
        ChartName = Sheet7.Range("H5")
        
    Application.ScreenUpdating = False
    
    ChartData.Replace What:="", Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    ChartData.Replace What:=ChartZero, Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    
    Set MyChart = Sheet7.Shapes.AddChart(xlLineMarkers).Chart
    
    Do While MyChart.SeriesCollection.Count > 0
    
    MyChart.FullSeriesCollection(1).Delete
    
    Loop
    
    MyChart.SeriesCollection.NewSeries
    MyChart.SeriesCollection(1).Name = ChartName
    MyChart.SeriesCollection(1).Values = ChartData
    
    'Scale
    MyChart.Axes(xlValue).MinimumScale = 580
    MyChart.Axes(xlValue).MaximumScale = 735
    
    'Zoom
    Sheet7.Activate
    ActiveWindow.Zoom = 125
    Sheet14.Activate
    
    'Create Image
    imageName = Application.DefaultFilePath & Application.PathSeparator & "TempChart.gif"
    MyChart.Export Filename:=imageName, FilterName:="GIF"
    Sheet7.ChartObjects(1).Delete
    Application.ScreenUpdating = True
    UserForm2.Image1.Picture = LoadPicture(imageName)
                   
    'Honda THR Sensor: Hardness
    Case 19
    
    'Using Find Function
    
    LastRow = Sheet8.Range("A" & Rows.Count).End(xlUp).Row
    LastRow = LastRow + 1
    FirstRow = LastRow - 25
    
    If FirstRow < 10 Then
        FirstRow = 10
    End If
             
        Set ChartData = Sheet8.Range("I" & FirstRow & ":I" & LastRow)
        ChartName = Sheet8.Range("H6")
        
    Application.ScreenUpdating = False
    
    ChartData.Replace What:="", Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    ChartData.Replace What:=ChartZero, Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    
    Set MyChart = Sheet8.Shapes.AddChart(xlLineMarkers).Chart
    
    Do While MyChart.SeriesCollection.Count > 0
    
    MyChart.FullSeriesCollection(1).Delete
    
    Loop
    
    MyChart.SeriesCollection.NewSeries
    MyChart.SeriesCollection(1).Name = ChartName
    MyChart.SeriesCollection(1).Values = ChartData
    
    'Scale
    MyChart.Axes(xlValue).MinimumScale = 580
    MyChart.Axes(xlValue).MaximumScale = 735
    
    'Zoom
    Sheet8.Activate
    ActiveWindow.Zoom = 125
    Sheet14.Activate
    
    'Create Image
    imageName = Application.DefaultFilePath & Application.PathSeparator & "TempChart.gif"
    MyChart.Export Filename:=imageName, FilterName:="GIF"
    Sheet8.ChartObjects(1).Delete
    Application.ScreenUpdating = True
    UserForm2.Image1.Picture = LoadPicture(imageName)
                   
    'Honda TJB Assist: Case Depth
    Case 20
    
    'Using Find Function
    LastRow = Sheet17.Range("A" & Rows.Count).End(xlUp).Row
    LastRow = LastRow + 1
    FirstRow = LastRow - 25
    
    If FirstRow < 10 Then
        FirstRow = 10
    End If
             
        Set ChartData = Sheet17.Range("J" & FirstRow & ":J" & LastRow)
        ChartName = Sheet17.Range("J5")
        
    Application.ScreenUpdating = False
    
    ChartData.Replace What:="", Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    ChartData.Replace What:=ChartZero, Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    
    Set MyChart = Sheet17.Shapes.AddChart(xlLineMarkers).Chart

    Do While MyChart.SeriesCollection.Count > 0
    
    MyChart.FullSeriesCollection(1).Delete
    
    Loop
    
    MyChart.SeriesCollection.NewSeries
    MyChart.SeriesCollection(1).Name = ChartName
    MyChart.SeriesCollection(1).Values = ChartData
    
    'Scale
    MyChart.Axes(xlValue).MinimumScale = 4
    MyChart.Axes(xlValue).MaximumScale = 6.2
    
    'Zoom
    Sheet17.Activate
    ActiveWindow.Zoom = 125
    Sheet14.Activate
    
    'Create Image
    imageName = Application.DefaultFilePath & Application.PathSeparator & "TempChart.gif"
    MyChart.Export Filename:=imageName, FilterName:="GIF"
    Sheet17.ChartObjects(1).Delete
    Application.ScreenUpdating = True
    UserForm2.Image1.Picture = LoadPicture(imageName)
                   
    'Honda TJB Assist: Hardness
    Case 21
    
    'Using Find Function
    LastRow = Sheet17.Range("A" & Rows.Count).End(xlUp).Row
    LastRow = LastRow + 1
    FirstRow = LastRow - 25
    
    If FirstRow < 10 Then
        FirstRow = 10
    End If
        
        Set ChartData = Sheet17.Range("I" & FirstRow & ":I" & LastRow)
        ChartName = Sheet17.Range("H5")
        
    Application.ScreenUpdating = False
    
    ChartData.Replace What:="", Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    ChartData.Replace What:=ChartZero, Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    
    Set MyChart = Sheet17.Shapes.AddChart(xlLineMarkers).Chart

    Do While MyChart.SeriesCollection.Count > 0
    
    MyChart.FullSeriesCollection(1).Delete
    
    Loop
        
    MyChart.SeriesCollection.NewSeries
    MyChart.SeriesCollection(1).Name = ChartName
    MyChart.SeriesCollection(1).Values = ChartData
    
    'scale
    MyChart.Axes(xlValue).MinimumScale = 580
    MyChart.Axes(xlValue).MaximumScale = 735
    
    'zoom
    Sheet17.Activate
    ActiveWindow.Zoom = 125
    Sheet14.Activate
    
    'Create Image
    imageName = Application.DefaultFilePath & Application.PathSeparator & "TempChart.gif"
    MyChart.Export Filename:=imageName, FilterName:="GIF"
    Sheet17.ChartObjects(1).Delete
    Application.ScreenUpdating = True
    UserForm2.Image1.Picture = LoadPicture(imageName)

    '09PL Ball Screw: Hardness A
    Case 22
    
    'Using Find Function
    LastRow = Sheet12.Range("A" & Rows.Count).End(xlUp).Row
    LastRow = LastRow + 1
    FirstRow = LastRow - 25
    
    If FirstRow < 10 Then
        FirstRow = 10
    End If
        
        Set ChartData = Sheet12.Range("S" & FirstRow & ":S" & LastRow)
        ChartName = Sheet12.Range("R5")
        
    Application.ScreenUpdating = False
    
    ChartData.Replace What:="", Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    ChartData.Replace What:=ChartZero, Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    
    Set MyChart = Sheet12.Shapes.AddChart(xlLineMarkers).Chart

    Do While MyChart.SeriesCollection.Count > 0
    
    MyChart.FullSeriesCollection(1).Delete
    
    Loop
        
    MyChart.SeriesCollection.NewSeries
    MyChart.SeriesCollection(1).Name = ChartName
    MyChart.SeriesCollection(1).Values = ChartData
    
    'scale
    MyChart.Axes(xlValue).MinimumScale = 630
    MyChart.Axes(xlValue).MaximumScale = 820
    
    'zoom
    Sheet12.Activate
    ActiveWindow.Zoom = 125
    Sheet14.Activate
    
    'Create Image
    imageName = Application.DefaultFilePath & Application.PathSeparator & "TempChart.gif"
    MyChart.Export Filename:=imageName, FilterName:="GIF"
    Sheet12.ChartObjects(1).Delete
    Application.ScreenUpdating = True
    UserForm2.Image1.Picture = LoadPicture(imageName)
    
    '09PL Ball Screw: Hardness B
    Case 23
    
    'Using Find Function
    LastRow = Sheet12.Range("A" & Rows.Count).End(xlUp).Row
    LastRow = LastRow + 1
    FirstRow = LastRow - 25
    
    If FirstRow < 10 Then
        FirstRow = 10
    End If
        
        Set ChartData = Sheet12.Range("V" & FirstRow & ":V" & LastRow)
        ChartName = Sheet12.Range("U5")
        
    Application.ScreenUpdating = False
    
    ChartData.Replace What:="", Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    ChartData.Replace What:=ChartZero, Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    
    Set MyChart = Sheet12.Shapes.AddChart(xlLineMarkers).Chart

    Do While MyChart.SeriesCollection.Count > 0
    
    MyChart.FullSeriesCollection(1).Delete
    
    Loop
        
    MyChart.SeriesCollection.NewSeries
    MyChart.SeriesCollection(1).Name = ChartName
    MyChart.SeriesCollection(1).Values = ChartData
    
    'scale
    MyChart.Axes(xlValue).MinimumScale = 630
    MyChart.Axes(xlValue).MaximumScale = 820
    
    'zoom
    Sheet12.Activate
    ActiveWindow.Zoom = 125
    Sheet14.Activate
    
    'Create Image
    imageName = Application.DefaultFilePath & Application.PathSeparator & "TempChart.gif"
    MyChart.Export Filename:=imageName, FilterName:="GIF"
    Sheet12.ChartObjects(1).Delete
    Application.ScreenUpdating = True
    UserForm2.Image1.Picture = LoadPicture(imageName)
    
    '09PL Ball Screw: Case Depth
    Case 24
    
    'Using Find Function
    LastRow = Sheet12.Range("A" & Rows.Count).End(xlUp).Row
    LastRow = LastRow + 1
    FirstRow = LastRow - 25
    
    If FirstRow < 10 Then
        FirstRow = 10
    End If
        
        Set ChartData = Sheet12.Range("T" & FirstRow & ":T" & LastRow)
        ChartName = Sheet12.Range("T5")
        
    Application.ScreenUpdating = False
    
    ChartData.Replace What:="", Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    ChartData.Replace What:=ChartZero, Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    
    Set MyChart = Sheet12.Shapes.AddChart(xlLineMarkers).Chart

    Do While MyChart.SeriesCollection.Count > 0
    
    MyChart.FullSeriesCollection(1).Delete
    
    Loop
        
    MyChart.SeriesCollection.NewSeries
    MyChart.SeriesCollection(1).Name = ChartName
    MyChart.SeriesCollection(1).Values = ChartData
    
    'scale
    MyChart.Axes(xlValue).MinimumScale = 0.7
    MyChart.Axes(xlValue).MaximumScale = 2
    
    'zoom
    Sheet12.Activate
    ActiveWindow.Zoom = 125
    Sheet14.Activate
    
    'Create Image
    imageName = Application.DefaultFilePath & Application.PathSeparator & "TempChart.gif"
    MyChart.Export Filename:=imageName, FilterName:="GIF"
    Sheet12.ChartObjects(1).Delete
    Application.ScreenUpdating = True
    UserForm2.Image1.Picture = LoadPicture(imageName)
                                        
    '09PL Rack Side: Tooth Hardness
    Case 25
    
    'Using Find Function
    LastRow = Sheet13.Range("A" & Rows.Count).End(xlUp).Row
    LastRow = LastRow + 1
    FirstRow = LastRow - 25
    
    If FirstRow < 10 Then
        FirstRow = 10
    End If
        
        Set ChartData = Sheet13.Range("I" & FirstRow & ":I" & LastRow)
        ChartName = Sheet13.Range("H5")
        
    Application.ScreenUpdating = False
    
    ChartData.Replace What:="", Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    ChartData.Replace What:=ChartZero, Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    
    Set MyChart = Sheet13.Shapes.AddChart(xlLineMarkers).Chart

    Do While MyChart.SeriesCollection.Count > 0
    
    MyChart.FullSeriesCollection(1).Delete
    
    Loop
        
    MyChart.SeriesCollection.NewSeries
    MyChart.SeriesCollection(1).Name = ChartName
    MyChart.SeriesCollection(1).Values = ChartData
    
    'scale
    MyChart.Axes(xlValue).MinimumScale = 580
    MyChart.Axes(xlValue).MaximumScale = 735
    
    'zoom
    Sheet13.Activate
    ActiveWindow.Zoom = 125
    Sheet14.Activate
    
    'Create Image
    imageName = Application.DefaultFilePath & Application.PathSeparator & "TempChart.gif"
    MyChart.Export Filename:=imageName, FilterName:="GIF"
    Sheet13.ChartObjects(1).Delete
    Application.ScreenUpdating = True
    UserForm2.Image1.Picture = LoadPicture(imageName)
    
    '09PL Rack Side: Case Depth
    Case 26
    
    'Using Find Function
    LastRow = Sheet13.Range("A" & Rows.Count).End(xlUp).Row
    LastRow = LastRow + 1
    FirstRow = LastRow - 25
    
    If FirstRow < 10 Then
        FirstRow = 10
    End If
        
        Set ChartData = Sheet13.Range("J" & FirstRow & ":J" & LastRow)
        ChartName = Sheet13.Range("J5")
        
    Application.ScreenUpdating = False
    
    ChartData.Replace What:="", Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    ChartData.Replace What:=ChartZero, Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    
    Set MyChart = Sheet13.Shapes.AddChart(xlLineMarkers).Chart

    Do While MyChart.SeriesCollection.Count > 0
    
    MyChart.FullSeriesCollection(1).Delete
    
    Loop
        
    MyChart.SeriesCollection.NewSeries
    MyChart.SeriesCollection(1).Name = ChartName
    MyChart.SeriesCollection(1).Values = ChartData
    
    'scale
    MyChart.Axes(xlValue).MinimumScale = 4#
    MyChart.Axes(xlValue).MaximumScale = 6.05
    
    'zoom
    Sheet13.Activate
    ActiveWindow.Zoom = 125
    Sheet14.Activate
    
    'Create Image
    imageName = Application.DefaultFilePath & Application.PathSeparator & "TempChart.gif"
    MyChart.Export Filename:=imageName, FilterName:="GIF"
    Sheet13.ChartObjects(1).Delete
    Application.ScreenUpdating = True
    UserForm2.Image1.Picture = LoadPicture(imageName)
    
    'BMW G2X Assist: Tooth Hardness
    Case 27
    
    'Using Find Function
    LastRow = Sheet18.Range("A" & Rows.Count).End(xlUp).Row
    LastRow = LastRow + 1
    FirstRow = LastRow - 25
    
    If FirstRow < 10 Then
        FirstRow = 10
    End If
    
        Set ChartData = Sheet18.Range("I" & FirstRow & ":I" & LastRow)
        ChartName = Sheet18.Range("H5")
        
    Application.ScreenUpdating = False
    
    ChartData.Replace What:="", Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    ChartData.Replace What:=ChartZero, Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    
    Set MyChart = Sheet18.Shapes.AddChart(xlLineMarkers).Chart

    Do While MyChart.SeriesCollection.Count > 0
    
    MyChart.FullSeriesCollection(1).Delete
    
    Loop
        
    MyChart.SeriesCollection.NewSeries
    MyChart.SeriesCollection(1).Name = ChartName
    MyChart.SeriesCollection(1).Values = ChartData
    
    'scale
    MyChart.Axes(xlValue).MinimumScale = 580
    MyChart.Axes(xlValue).MaximumScale = 735
    
    'zoom
    Sheet18.Activate
    ActiveWindow.Zoom = 125
    Sheet14.Activate
    
    'Create Image
    imageName = Application.DefaultFilePath & Application.PathSeparator & "TempChart.gif"
    MyChart.Export Filename:=imageName, FilterName:="GIF"
    Sheet18.ChartObjects(1).Delete
    Application.ScreenUpdating = True
    UserForm2.Image1.Picture = LoadPicture(imageName)
    
    'BMW G2X Assist: Case Depth
    Case 28
    
    'Using Find Function
    LastRow = Sheet18.Range("A" & Rows.Count).End(xlUp).Row
    LastRow = LastRow + 1
    FirstRow = LastRow - 25
    
    If FirstRow < 10 Then
        FirstRow = 10
    End If
    
        Set ChartData = Sheet18.Range("J" & FirstRow & ":J" & LastRow)
        ChartName = Sheet18.Range("J5")
        
    Application.ScreenUpdating = False
    
    ChartData.Replace What:="", Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    ChartData.Replace What:=ChartZero, Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    
    Set MyChart = Sheet18.Shapes.AddChart(xlLineMarkers).Chart

    Do While MyChart.SeriesCollection.Count > 0
    
    MyChart.FullSeriesCollection(1).Delete
    
    Loop
        
    MyChart.SeriesCollection.NewSeries
    MyChart.SeriesCollection(1).Name = ChartName
    MyChart.SeriesCollection(1).Values = ChartData
    
    'scale
    MyChart.Axes(xlValue).MinimumScale = 4.09
    MyChart.Axes(xlValue).MaximumScale = 6.05
    
    'zoom
    Sheet18.Activate
    ActiveWindow.Zoom = 125
    Sheet14.Activate
    
    'Create Image
    imageName = Application.DefaultFilePath & Application.PathSeparator & "TempChart.gif"
    MyChart.Export Filename:=imageName, FilterName:="GIF"
    Sheet18.ChartObjects(1).Delete
    Application.ScreenUpdating = True
    UserForm2.Image1.Picture = LoadPicture(imageName)

    'BMW G2X Sensor: Tooth Hardness
    Case 29
    
    'Using Find Function
    LastRow = Sheet19.Range("A" & Rows.Count).End(xlUp).Row
    LastRow = LastRow + 1
    FirstRow = LastRow - 25
    
    If FirstRow < 10 Then
        FirstRow = 10
    End If
    
        Set ChartData = Sheet19.Range("I" & FirstRow & ":I" & LastRow)
        ChartName = Sheet19.Range("H5")
        
    Application.ScreenUpdating = False
    
    ChartData.Replace What:="", Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    ChartData.Replace What:=ChartZero, Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    
    Set MyChart = Sheet19.Shapes.AddChart(xlLineMarkers).Chart

    Do While MyChart.SeriesCollection.Count > 0
    
    MyChart.FullSeriesCollection(1).Delete
    
    Loop
        
    MyChart.SeriesCollection.NewSeries
    MyChart.SeriesCollection(1).Name = ChartName
    MyChart.SeriesCollection(1).Values = ChartData
    
    'scale
    MyChart.Axes(xlValue).MinimumScale = 490
    MyChart.Axes(xlValue).MaximumScale = 720
    
    'zoom
    Sheet19.Activate
    ActiveWindow.Zoom = 125
    Sheet14.Activate
    
    'Create Image
    imageName = Application.DefaultFilePath & Application.PathSeparator & "TempChart.gif"
    MyChart.Export Filename:=imageName, FilterName:="GIF"
    Sheet19.ChartObjects(1).Delete
    Application.ScreenUpdating = True
    UserForm2.Image1.Picture = LoadPicture(imageName)
    
    'BMW G2X Sensor: Case Depth
    Case 30
    
    'Using Find Function
    LastRow = Sheet19.Range("A" & Rows.Count).End(xlUp).Row
    LastRow = LastRow + 1
    FirstRow = LastRow - 25
    
    If FirstRow < 10 Then
        FirstRow = 10
    End If
    
        Set ChartData = Sheet19.Range("J" & FirstRow & ":J" & LastRow)
        ChartName = Sheet19.Range("J5")
        
    Application.ScreenUpdating = False
    
    ChartData.Replace What:="", Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    ChartData.Replace What:=ChartZero, Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    
    Set MyChart = Sheet19.Shapes.AddChart(xlLineMarkers).Chart

    Do While MyChart.SeriesCollection.Count > 0
    
    MyChart.FullSeriesCollection(1).Delete
    
    Loop
        
    MyChart.SeriesCollection.NewSeries
    MyChart.SeriesCollection(1).Name = ChartName
    MyChart.SeriesCollection(1).Values = ChartData
    
    'scale
    MyChart.Axes(xlValue).MinimumScale = 4.09
    MyChart.Axes(xlValue).MaximumScale = 6.05
    
    'zoom
    Sheet19.Activate
    ActiveWindow.Zoom = 125
    Sheet14.Activate
    
    'Create Image
    imageName = Application.DefaultFilePath & Application.PathSeparator & "TempChart.gif"
    MyChart.Export Filename:=imageName, FilterName:="GIF"
    Sheet19.ChartObjects(1).Delete
    Application.ScreenUpdating = True
    UserForm2.Image1.Picture = LoadPicture(imageName)
    
    '15PL Rack Side: Tooth Hardness
    Case 31
    
    'Using Find Function
    LastRow = Sheet20.Range("A" & Rows.Count).End(xlUp).Row
    LastRow = LastRow + 1
    FirstRow = LastRow - 25
    
    If FirstRow < 10 Then
        FirstRow = 10
    End If
        
        Set ChartData = Sheet20.Range("I" & FirstRow & ":I" & LastRow)
        ChartName = Sheet20.Range("H5")
        
    Application.ScreenUpdating = False
    
    ChartData.Replace What:="", Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    ChartData.Replace What:=ChartZero, Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    
    Set MyChart = Sheet20.Shapes.AddChart(xlLineMarkers).Chart

    Do While MyChart.SeriesCollection.Count > 0
    
    MyChart.FullSeriesCollection(1).Delete
    
    Loop
        
    MyChart.SeriesCollection.NewSeries
    MyChart.SeriesCollection(1).Name = ChartName
    MyChart.SeriesCollection(1).Values = ChartData
    
    'scale
    MyChart.Axes(xlValue).MinimumScale = 580
    MyChart.Axes(xlValue).MaximumScale = 735
    
    'zoom
    Sheet20.Activate
    ActiveWindow.Zoom = 125
    Sheet14.Activate
    
    'Create Image
    imageName = Application.DefaultFilePath & Application.PathSeparator & "TempChart.gif"
    MyChart.Export Filename:=imageName, FilterName:="GIF"
    Sheet20.ChartObjects(1).Delete
    Application.ScreenUpdating = True
    UserForm2.Image1.Picture = LoadPicture(imageName)
    
    '015PL Rack Side: Case Depth
    Case 32
    
    'Using Find Function
    LastRow = Sheet20.Range("A" & Rows.Count).End(xlUp).Row
    LastRow = LastRow + 1
    FirstRow = LastRow - 25
    
    If FirstRow < 10 Then
        FirstRow = 10
    End If
        
        Set ChartData = Sheet20.Range("J" & FirstRow & ":J" & LastRow)
        ChartName = Sheet20.Range("J5")
        
    Application.ScreenUpdating = False
    
    ChartData.Replace What:="", Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    ChartData.Replace What:=ChartZero, Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    
    Set MyChart = Sheet20.Shapes.AddChart(xlLineMarkers).Chart

    Do While MyChart.SeriesCollection.Count > 0
    
    MyChart.FullSeriesCollection(1).Delete
    
    Loop
        
    MyChart.SeriesCollection.NewSeries
    MyChart.SeriesCollection(1).Name = ChartName
    MyChart.SeriesCollection(1).Values = ChartData
    
    'scale
    MyChart.Axes(xlValue).MinimumScale = 4#
    MyChart.Axes(xlValue).MaximumScale = 6.05
    
    'zoom
    Sheet20.Activate
    ActiveWindow.Zoom = 125
    Sheet14.Activate
    
    'Create Image
    imageName = Application.DefaultFilePath & Application.PathSeparator & "TempChart.gif"
    MyChart.Export Filename:=imageName, FilterName:="GIF"
    Sheet20.ChartObjects(1).Delete
    Application.ScreenUpdating = True
    UserForm2.Image1.Picture = LoadPicture(imageName)
    
    '15PL Ball Screw: Hardness A
    Case 33
    
    'Using Find Function
    LastRow = Sheet21.Range("A" & Rows.Count).End(xlUp).Row
    LastRow = LastRow + 1
    FirstRow = LastRow - 25
    
    If FirstRow < 10 Then
        FirstRow = 10
    End If
        
        Set ChartData = Sheet21.Range("S" & FirstRow & ":S" & LastRow)
        ChartName = Sheet21.Range("R5")
        
    Application.ScreenUpdating = False
    
    ChartData.Replace What:="", Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    ChartData.Replace What:=ChartZero, Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    
    Set MyChart = Sheet21.Shapes.AddChart(xlLineMarkers).Chart

    Do While MyChart.SeriesCollection.Count > 0
    
    MyChart.FullSeriesCollection(1).Delete
    
    Loop
        
    MyChart.SeriesCollection.NewSeries
    MyChart.SeriesCollection(1).Name = ChartName
    MyChart.SeriesCollection(1).Values = ChartData
    
    'scale
    MyChart.Axes(xlValue).MinimumScale = 630
    MyChart.Axes(xlValue).MaximumScale = 820
    
    'zoom
    Sheet21.Activate
    ActiveWindow.Zoom = 125
    Sheet14.Activate
    
    'Create Image
    imageName = Application.DefaultFilePath & Application.PathSeparator & "TempChart.gif"
    MyChart.Export Filename:=imageName, FilterName:="GIF"
    Sheet21.ChartObjects(1).Delete
    Application.ScreenUpdating = True
    UserForm2.Image1.Picture = LoadPicture(imageName)
    
    '15PL Ball Screw: Hardness B
    Case 34
    
    'Using Find Function
    LastRow = Sheet21.Range("A" & Rows.Count).End(xlUp).Row
    LastRow = LastRow + 1
    FirstRow = LastRow - 25
    
    If FirstRow < 10 Then
        FirstRow = 10
    End If
        
        Set ChartData = Sheet21.Range("V" & FirstRow & ":V" & LastRow)
        ChartName = Sheet21.Range("U5")
        
    Application.ScreenUpdating = False
    
    ChartData.Replace What:="", Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    ChartData.Replace What:=ChartZero, Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    
    Set MyChart = Sheet21.Shapes.AddChart(xlLineMarkers).Chart

    Do While MyChart.SeriesCollection.Count > 0
    
    MyChart.FullSeriesCollection(1).Delete
    
    Loop
        
    MyChart.SeriesCollection.NewSeries
    MyChart.SeriesCollection(1).Name = ChartName
    MyChart.SeriesCollection(1).Values = ChartData
    
    'scale
    MyChart.Axes(xlValue).MinimumScale = 630
    MyChart.Axes(xlValue).MaximumScale = 820
    
    'zoom
    Sheet21.Activate
    ActiveWindow.Zoom = 125
    Sheet14.Activate
    
    'Create Image
    imageName = Application.DefaultFilePath & Application.PathSeparator & "TempChart.gif"
    MyChart.Export Filename:=imageName, FilterName:="GIF"
    Sheet21.ChartObjects(1).Delete
    Application.ScreenUpdating = True
    UserForm2.Image1.Picture = LoadPicture(imageName)
    
    '15PL Ball Screw: Case Depth
    Case 35
    
    'Using Find Function
    LastRow = Sheet21.Range("A" & Rows.Count).End(xlUp).Row
    LastRow = LastRow + 1
    FirstRow = LastRow - 25
    
    If FirstRow < 10 Then
        FirstRow = 10
    End If
        
        Set ChartData = Sheet21.Range("T" & FirstRow & ":T" & LastRow)
        ChartName = Sheet21.Range("T5")
        
    Application.ScreenUpdating = False
    
    ChartData.Replace What:="", Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    ChartData.Replace What:=ChartZero, Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    
    Set MyChart = Sheet21.Shapes.AddChart(xlLineMarkers).Chart

    Do While MyChart.SeriesCollection.Count > 0
    
    MyChart.FullSeriesCollection(1).Delete
    
    Loop
        
    MyChart.SeriesCollection.NewSeries
    MyChart.SeriesCollection(1).Name = ChartName
    MyChart.SeriesCollection(1).Values = ChartData
    
    'scale
    MyChart.Axes(xlValue).MinimumScale = 0.7
    MyChart.Axes(xlValue).MaximumScale = 2
    
    'zoom
    Sheet21.Activate
    ActiveWindow.Zoom = 125
    Sheet14.Activate
    
    'Create Image
    imageName = Application.DefaultFilePath & Application.PathSeparator & "TempChart.gif"
    MyChart.Export Filename:=imageName, FilterName:="GIF"
    Sheet21.ChartObjects(1).Delete
    Application.ScreenUpdating = True
    UserForm2.Image1.Picture = LoadPicture(imageName)
    
    'Nissan L42P Assist: Root Hardness
    Case 36
    
    'Using Find Function
    LastRow = Sheet15.Range("A" & Rows.Count).End(xlUp).Row
    LastRow = LastRow + 1
    FirstRow = LastRow - 25
    
    If FirstRow < 10 Then
        FirstRow = 10
    End If
        
        Set ChartData = Sheet15.Range("M" & FirstRow & ":M" & LastRow)
        ChartName = Sheet15.Range("L5")
        
    Application.ScreenUpdating = False
    
    ChartData.Replace What:="", Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    ChartData.Replace What:=ChartZero, Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    
    Set MyChart = Sheet15.Shapes.AddChart(xlLineMarkers).Chart

    Do While MyChart.SeriesCollection.Count > 0
    
    MyChart.FullSeriesCollection(1).Delete
    
    Loop
        
    MyChart.SeriesCollection.NewSeries
    MyChart.SeriesCollection(1).Name = ChartName
    MyChart.SeriesCollection(1).Values = ChartData
    
    'scale
    MyChart.Axes(xlValue).MinimumScale = 580
    MyChart.Axes(xlValue).MaximumScale = 735
    
    'zoom
    Sheet15.Activate
    ActiveWindow.Zoom = 125
    Sheet14.Activate
    
    'Create Image
    imageName = Application.DefaultFilePath & Application.PathSeparator & "TempChart.gif"
    MyChart.Export Filename:=imageName, FilterName:="GIF"
    Sheet15.ChartObjects(1).Delete
    Application.ScreenUpdating = True
    UserForm2.Image1.Picture = LoadPicture(imageName)
    
    'Nissan L42P Sensor: Root Hardness
    Case 37
    
    'Using Find Function
    LastRow = Sheet16.Range("A" & Rows.Count).End(xlUp).Row
    LastRow = LastRow + 1
    FirstRow = LastRow - 25
    
    If FirstRow < 10 Then
        FirstRow = 10
    End If
        
        Set ChartData = Sheet16.Range("M" & FirstRow & ":M" & LastRow)
        ChartName = Sheet16.Range("L5")
        
    Application.ScreenUpdating = False
    
    ChartData.Replace What:="", Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    ChartData.Replace What:=ChartZero, Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    
    Set MyChart = Sheet16.Shapes.AddChart(xlLineMarkers).Chart

    Do While MyChart.SeriesCollection.Count > 0
    
    MyChart.FullSeriesCollection(1).Delete
    
    Loop
        
    MyChart.SeriesCollection.NewSeries
    MyChart.SeriesCollection(1).Name = ChartName
    MyChart.SeriesCollection(1).Values = ChartData
    
    'scale
    MyChart.Axes(xlValue).MinimumScale = 580
    MyChart.Axes(xlValue).MaximumScale = 735
    
    'zoom
    Sheet16.Activate
    ActiveWindow.Zoom = 125
    Sheet14.Activate
    
    'Create Image
    imageName = Application.DefaultFilePath & Application.PathSeparator & "TempChart.gif"
    MyChart.Export Filename:=imageName, FilterName:="GIF"
    Sheet16.ChartObjects(1).Delete
    Application.ScreenUpdating = True
    UserForm2.Image1.Picture = LoadPicture(imageName)
    
    'BMW LHD CGR Assist: Root Hardness
    Case 38
    
    'Using Find Function
    LastRow = Sheet1.Range("A" & Rows.Count).End(xlUp).Row
    LastRow = LastRow + 1
    FirstRow = LastRow - 25
    
    If FirstRow < 10 Then
        FirstRow = 10
    End If
        
        Set ChartData = Sheet1.Range("M" & FirstRow & ":M" & LastRow)
        ChartName = Sheet1.Range("L5")
        
    Application.ScreenUpdating = False
    
    ChartData.Replace What:="", Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    ChartData.Replace What:=ChartZero, Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    
    Set MyChart = Sheet1.Shapes.AddChart(xlLineMarkers).Chart

    Do While MyChart.SeriesCollection.Count > 0
    
    MyChart.FullSeriesCollection(1).Delete
    
    Loop
        
    MyChart.SeriesCollection.NewSeries
    MyChart.SeriesCollection(1).Name = ChartName
    MyChart.SeriesCollection(1).Values = ChartData
    
    'scale
    MyChart.Axes(xlValue).MinimumScale = 580
    MyChart.Axes(xlValue).MaximumScale = 735
    
    'zoom
    Sheet1.Activate
    ActiveWindow.Zoom = 125
    Sheet14.Activate
    
    'Create Image
    imageName = Application.DefaultFilePath & Application.PathSeparator & "TempChart.gif"
    MyChart.Export Filename:=imageName, FilterName:="GIF"
    Sheet1.ChartObjects(1).Delete
    Application.ScreenUpdating = True
    UserForm2.Image1.Picture = LoadPicture(imageName)
    
    'BMW LHD CGR Sensor: Root Hardness
    Case 39
    
    'Using Find Function
    LastRow = Sheet2.Range("A" & Rows.Count).End(xlUp).Row
    LastRow = LastRow + 1
    FirstRow = LastRow - 25
    
    If FirstRow < 10 Then
        FirstRow = 10
    End If
        
        Set ChartData = Sheet2.Range("M" & FirstRow & ":M" & LastRow)
        ChartName = Sheet2.Range("L5")
        
    Application.ScreenUpdating = False
    
    ChartData.Replace What:="", Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    ChartData.Replace What:=ChartZero, Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    
    Set MyChart = Sheet2.Shapes.AddChart(xlLineMarkers).Chart

    Do While MyChart.SeriesCollection.Count > 0
    
    MyChart.FullSeriesCollection(1).Delete
    
    Loop
        
    MyChart.SeriesCollection.NewSeries
    MyChart.SeriesCollection(1).Name = ChartName
    MyChart.SeriesCollection(1).Values = ChartData
    
    'scale
    MyChart.Axes(xlValue).MinimumScale = 580
    MyChart.Axes(xlValue).MaximumScale = 735
    
    'zoom
    Sheet2.Activate
    ActiveWindow.Zoom = 125
    Sheet14.Activate
    
    'Create Image
    imageName = Application.DefaultFilePath & Application.PathSeparator & "TempChart.gif"
    MyChart.Export Filename:=imageName, FilterName:="GIF"
    Sheet2.ChartObjects(1).Delete
    Application.ScreenUpdating = True
    UserForm2.Image1.Picture = LoadPicture(imageName)
    
    'BMW RHD CGR Assist: Root Hardness
    Case 40
    
    'Using Find Function
    LastRow = Sheet3.Range("A" & Rows.Count).End(xlUp).Row
    LastRow = LastRow + 1
    FirstRow = LastRow - 25
    
    If FirstRow < 10 Then
        FirstRow = 10
    End If
        
        Set ChartData = Sheet3.Range("M" & FirstRow & ":M" & LastRow)
        ChartName = Sheet3.Range("L5")
        
    Application.ScreenUpdating = False
    
    ChartData.Replace What:="", Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    ChartData.Replace What:=ChartZero, Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    
    Set MyChart = Sheet3.Shapes.AddChart(xlLineMarkers).Chart

    Do While MyChart.SeriesCollection.Count > 0
    
    MyChart.FullSeriesCollection(1).Delete
    
    Loop
        
    MyChart.SeriesCollection.NewSeries
    MyChart.SeriesCollection(1).Name = ChartName
    MyChart.SeriesCollection(1).Values = ChartData
    
    'scale
    MyChart.Axes(xlValue).MinimumScale = 580
    MyChart.Axes(xlValue).MaximumScale = 735
    
    'zoom
    Sheet3.Activate
    ActiveWindow.Zoom = 125
    Sheet14.Activate
    
    'Create Image
    imageName = Application.DefaultFilePath & Application.PathSeparator & "TempChart.gif"
    MyChart.Export Filename:=imageName, FilterName:="GIF"
    Sheet3.ChartObjects(1).Delete
    Application.ScreenUpdating = True
    UserForm2.Image1.Picture = LoadPicture(imageName)
    
    'BMW RHD CGR Sensor: Root Hardness
    Case 41
    
    'Using Find Function
    LastRow = Sheet4.Range("A" & Rows.Count).End(xlUp).Row
    LastRow = LastRow + 1
    FirstRow = LastRow - 25
    
    If FirstRow < 10 Then
        FirstRow = 10
    End If
        
        Set ChartData = Sheet4.Range("M" & FirstRow & ":M" & LastRow)
        ChartName = Sheet4.Range("L5")
        
    Application.ScreenUpdating = False
    
    ChartData.Replace What:="", Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    ChartData.Replace What:=ChartZero, Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    
    Set MyChart = Sheet4.Shapes.AddChart(xlLineMarkers).Chart

    Do While MyChart.SeriesCollection.Count > 0
    
    MyChart.FullSeriesCollection(1).Delete
    
    Loop
        
    MyChart.SeriesCollection.NewSeries
    MyChart.SeriesCollection(1).Name = ChartName
    MyChart.SeriesCollection(1).Values = ChartData
    
    'scale
    MyChart.Axes(xlValue).MinimumScale = 580
    MyChart.Axes(xlValue).MaximumScale = 735
    
    'zoom
    Sheet4.Activate
    ActiveWindow.Zoom = 125
    Sheet14.Activate
    
    'Create Image
    imageName = Application.DefaultFilePath & Application.PathSeparator & "TempChart.gif"
    MyChart.Export Filename:=imageName, FilterName:="GIF"
    Sheet4.ChartObjects(1).Delete
    Application.ScreenUpdating = True
    UserForm2.Image1.Picture = LoadPicture(imageName)
    
    'BMW LHD VGR Assist: Root Hardness
    Case 42
    
    'Using Find Function
    LastRow = Sheet5.Range("A" & Rows.Count).End(xlUp).Row
    LastRow = LastRow + 1
    FirstRow = LastRow - 25
    
    If FirstRow < 10 Then
        FirstRow = 10
    End If
        
        Set ChartData = Sheet5.Range("M" & FirstRow & ":M" & LastRow)
        ChartName = Sheet5.Range("L5")
        
    Application.ScreenUpdating = False
    
    ChartData.Replace What:="", Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    ChartData.Replace What:=ChartZero, Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    
    Set MyChart = Sheet5.Shapes.AddChart(xlLineMarkers).Chart

    Do While MyChart.SeriesCollection.Count > 0
    
    MyChart.FullSeriesCollection(1).Delete
    
    Loop
        
    MyChart.SeriesCollection.NewSeries
    MyChart.SeriesCollection(1).Name = ChartName
    MyChart.SeriesCollection(1).Values = ChartData
    
    'scale
    MyChart.Axes(xlValue).MinimumScale = 580
    MyChart.Axes(xlValue).MaximumScale = 735
    
    'zoom
    Sheet5.Activate
    ActiveWindow.Zoom = 125
    Sheet14.Activate
    
    'Create Image
    imageName = Application.DefaultFilePath & Application.PathSeparator & "TempChart.gif"
    MyChart.Export Filename:=imageName, FilterName:="GIF"
    Sheet5.ChartObjects(1).Delete
    Application.ScreenUpdating = True
    UserForm2.Image1.Picture = LoadPicture(imageName)
    
    'BMW RHD VGR Assist: Root Hardness
    Case 43
    
    'Using Find Function
    LastRow = Sheet6.Range("A" & Rows.Count).End(xlUp).Row
    LastRow = LastRow + 1
    FirstRow = LastRow - 25
    
    If FirstRow < 10 Then
        FirstRow = 10
    End If
        
        Set ChartData = Sheet6.Range("M" & FirstRow & ":M" & LastRow)
        ChartName = Sheet6.Range("L5")
        
    Application.ScreenUpdating = False
    
    ChartData.Replace What:="", Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    ChartData.Replace What:=ChartZero, Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    
    Set MyChart = Sheet6.Shapes.AddChart(xlLineMarkers).Chart

    Do While MyChart.SeriesCollection.Count > 0
    
    MyChart.FullSeriesCollection(1).Delete
    
    Loop
        
    MyChart.SeriesCollection.NewSeries
    MyChart.SeriesCollection(1).Name = ChartName
    MyChart.SeriesCollection(1).Values = ChartData
    
    'scale
    MyChart.Axes(xlValue).MinimumScale = 580
    MyChart.Axes(xlValue).MaximumScale = 735
    
    'zoom
    Sheet6.Activate
    ActiveWindow.Zoom = 125
    Sheet14.Activate
    
    'Create Image
    imageName = Application.DefaultFilePath & Application.PathSeparator & "TempChart.gif"
    MyChart.Export Filename:=imageName, FilterName:="GIF"
    Sheet6.ChartObjects(1).Delete
    Application.ScreenUpdating = True
    UserForm2.Image1.Picture = LoadPicture(imageName)
    
    'Honda THR Assist: Root Hardness
    Case 44
    
    'Using Find Function
    LastRow = Sheet7.Range("A" & Rows.Count).End(xlUp).Row
    LastRow = LastRow + 1
    FirstRow = LastRow - 25
    
    If FirstRow < 10 Then
        FirstRow = 10
    End If
        
        Set ChartData = Sheet7.Range("M" & FirstRow & ":M" & LastRow)
        ChartName = Sheet7.Range("L5")
        
    Application.ScreenUpdating = False
    
    ChartData.Replace What:="", Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    ChartData.Replace What:=ChartZero, Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    
    Set MyChart = Sheet7.Shapes.AddChart(xlLineMarkers).Chart

    Do While MyChart.SeriesCollection.Count > 0
    
    MyChart.FullSeriesCollection(1).Delete
    
    Loop
        
    MyChart.SeriesCollection.NewSeries
    MyChart.SeriesCollection(1).Name = ChartName
    MyChart.SeriesCollection(1).Values = ChartData
    
    'scale
    MyChart.Axes(xlValue).MinimumScale = 580
    MyChart.Axes(xlValue).MaximumScale = 735
    
    'zoom
    Sheet7.Activate
    ActiveWindow.Zoom = 125
    Sheet14.Activate
    
    'Create Image
    imageName = Application.DefaultFilePath & Application.PathSeparator & "TempChart.gif"
    MyChart.Export Filename:=imageName, FilterName:="GIF"
    Sheet7.ChartObjects(1).Delete
    Application.ScreenUpdating = True
    UserForm2.Image1.Picture = LoadPicture(imageName)

    'Honda THR Sensor: Root Hardness
    Case 45
    
    'Using Find Function
    LastRow = Sheet8.Range("A" & Rows.Count).End(xlUp).Row
    LastRow = LastRow + 1
    FirstRow = LastRow - 25
    
    If FirstRow < 10 Then
        FirstRow = 10
    End If
        
        Set ChartData = Sheet8.Range("M" & FirstRow & ":M" & LastRow)
        ChartName = Sheet8.Range("L5")
        
    Application.ScreenUpdating = False
    
    ChartData.Replace What:="", Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    ChartData.Replace What:=ChartZero, Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    
    Set MyChart = Sheet8.Shapes.AddChart(xlLineMarkers).Chart

    Do While MyChart.SeriesCollection.Count > 0
    
    MyChart.FullSeriesCollection(1).Delete
    
    Loop
        
    MyChart.SeriesCollection.NewSeries
    MyChart.SeriesCollection(1).Name = ChartName
    MyChart.SeriesCollection(1).Values = ChartData
    
    'scale
    MyChart.Axes(xlValue).MinimumScale = 580
    MyChart.Axes(xlValue).MaximumScale = 735
    
    'zoom
    Sheet8.Activate
    ActiveWindow.Zoom = 125
    Sheet14.Activate
    
    'Create Image
    imageName = Application.DefaultFilePath & Application.PathSeparator & "TempChart.gif"
    MyChart.Export Filename:=imageName, FilterName:="GIF"
    Sheet8.ChartObjects(1).Delete
    Application.ScreenUpdating = True
    UserForm2.Image1.Picture = LoadPicture(imageName)
    
    'Honda TJB Assist: Root Hardness
    Case 46
    
    'Using Find Function
    LastRow = Sheet17.Range("A" & Rows.Count).End(xlUp).Row
    LastRow = LastRow + 1
    FirstRow = LastRow - 25
    
    If FirstRow < 10 Then
        FirstRow = 10
    End If
        
        Set ChartData = Sheet17.Range("M" & FirstRow & ":M" & LastRow)
        ChartName = Sheet17.Range("L5")
        
    Application.ScreenUpdating = False
    
    ChartData.Replace What:="", Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    ChartData.Replace What:=ChartZero, Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    
    Set MyChart = Sheet17.Shapes.AddChart(xlLineMarkers).Chart

    Do While MyChart.SeriesCollection.Count > 0
    
    MyChart.FullSeriesCollection(1).Delete
    
    Loop
        
    MyChart.SeriesCollection.NewSeries
    MyChart.SeriesCollection(1).Name = ChartName
    MyChart.SeriesCollection(1).Values = ChartData
    
    'scale
    MyChart.Axes(xlValue).MinimumScale = 580
    MyChart.Axes(xlValue).MaximumScale = 735
    
    'zoom
    Sheet17.Activate
    ActiveWindow.Zoom = 125
    Sheet14.Activate
    
    'Create Image
    imageName = Application.DefaultFilePath & Application.PathSeparator & "TempChart.gif"
    MyChart.Export Filename:=imageName, FilterName:="GIF"
    Sheet17.ChartObjects(1).Delete
    Application.ScreenUpdating = True
    UserForm2.Image1.Picture = LoadPicture(imageName)
    
    '09PL Rack Side Assist: Root Hardness
    Case 47
    
    'Using Find Function
    LastRow = Sheet13.Range("A" & Rows.Count).End(xlUp).Row
    LastRow = LastRow + 1
    FirstRow = LastRow - 25
    
    If FirstRow < 10 Then
        FirstRow = 10
    End If
        
        Set ChartData = Sheet13.Range("M" & FirstRow & ":M" & LastRow)
        ChartName = Sheet13.Range("L5")
        
    Application.ScreenUpdating = False
    
    ChartData.Replace What:="", Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    ChartData.Replace What:=ChartZero, Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    
    Set MyChart = Sheet13.Shapes.AddChart(xlLineMarkers).Chart

    Do While MyChart.SeriesCollection.Count > 0
    
    MyChart.FullSeriesCollection(1).Delete
    
    Loop
        
    MyChart.SeriesCollection.NewSeries
    MyChart.SeriesCollection(1).Name = ChartName
    MyChart.SeriesCollection(1).Values = ChartData
    
    'scale
    MyChart.Axes(xlValue).MinimumScale = 580
    MyChart.Axes(xlValue).MaximumScale = 735
    
    'zoom
    Sheet13.Activate
    ActiveWindow.Zoom = 125
    Sheet14.Activate
    
    'Create Image
    imageName = Application.DefaultFilePath & Application.PathSeparator & "TempChart.gif"
    MyChart.Export Filename:=imageName, FilterName:="GIF"
    Sheet13.ChartObjects(1).Delete
    Application.ScreenUpdating = True
    UserForm2.Image1.Picture = LoadPicture(imageName)
    
    'BMW G2X Assist: Root Hardness
    Case 48
    
    'Using Find Function
    LastRow = Sheet18.Range("A" & Rows.Count).End(xlUp).Row
    LastRow = LastRow + 1
    FirstRow = LastRow - 25
    
    If FirstRow < 10 Then
        FirstRow = 10
    End If
        
        Set ChartData = Sheet18.Range("M" & FirstRow & ":M" & LastRow)
        ChartName = Sheet18.Range("L5")
        
    Application.ScreenUpdating = False
    
    ChartData.Replace What:="", Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    ChartData.Replace What:=ChartZero, Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    
    Set MyChart = Sheet18.Shapes.AddChart(xlLineMarkers).Chart

    Do While MyChart.SeriesCollection.Count > 0
    
    MyChart.FullSeriesCollection(1).Delete
    
    Loop
        
    MyChart.SeriesCollection.NewSeries
    MyChart.SeriesCollection(1).Name = ChartName
    MyChart.SeriesCollection(1).Values = ChartData
    
    'scale
    MyChart.Axes(xlValue).MinimumScale = 580
    MyChart.Axes(xlValue).MaximumScale = 735
    
    'zoom
    Sheet18.Activate
    ActiveWindow.Zoom = 125
    Sheet14.Activate
    
    'Create Image
    imageName = Application.DefaultFilePath & Application.PathSeparator & "TempChart.gif"
    MyChart.Export Filename:=imageName, FilterName:="GIF"
    Sheet18.ChartObjects(1).Delete
    Application.ScreenUpdating = True
    UserForm2.Image1.Picture = LoadPicture(imageName)
    
    'BMW G2X Sensor: Root Hardness
    Case 49
    
    'Using Find Function
    LastRow = Sheet19.Range("A" & Rows.Count).End(xlUp).Row
    LastRow = LastRow + 1
    FirstRow = LastRow - 25
    
    If FirstRow < 10 Then
        FirstRow = 10
    End If
        
        Set ChartData = Sheet19.Range("M" & FirstRow & ":M" & LastRow)
        ChartName = Sheet19.Range("L5")
        
    Application.ScreenUpdating = False
    
    ChartData.Replace What:="", Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    ChartData.Replace What:=ChartZero, Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    
    Set MyChart = Sheet19.Shapes.AddChart(xlLineMarkers).Chart

    Do While MyChart.SeriesCollection.Count > 0
    
    MyChart.FullSeriesCollection(1).Delete
    
    Loop
        
    MyChart.SeriesCollection.NewSeries
    MyChart.SeriesCollection(1).Name = ChartName
    MyChart.SeriesCollection(1).Values = ChartData
    
    'scale
    MyChart.Axes(xlValue).MinimumScale = 580
    MyChart.Axes(xlValue).MaximumScale = 735
    
    'zoom
    Sheet19.Activate
    ActiveWindow.Zoom = 125
    Sheet14.Activate
    
    'Create Image
    imageName = Application.DefaultFilePath & Application.PathSeparator & "TempChart.gif"
    MyChart.Export Filename:=imageName, FilterName:="GIF"
    Sheet19.ChartObjects(1).Delete
    Application.ScreenUpdating = True
    UserForm2.Image1.Picture = LoadPicture(imageName)
    
    '15PL Rack Side Assist: Root Hardness
    Case 50
    
    'Using Find Function
    LastRow = Sheet20.Range("A" & Rows.Count).End(xlUp).Row
    LastRow = LastRow + 1
    FirstRow = LastRow - 25
    
    If FirstRow < 10 Then
        FirstRow = 10
    End If
        
        Set ChartData = Sheet20.Range("M" & FirstRow & ":M" & LastRow)
        ChartName = Sheet20.Range("L5")
        
    Application.ScreenUpdating = False
    
    ChartData.Replace What:="", Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    ChartData.Replace What:=ChartZero, Replacement:="=NA()", LookAt:=xlWhole, MatchCase:=False
    
    Set MyChart = Sheet20.Shapes.AddChart(xlLineMarkers).Chart

    Do While MyChart.SeriesCollection.Count > 0
    
    MyChart.FullSeriesCollection(1).Delete
    
    Loop
        
    MyChart.SeriesCollection.NewSeries
    MyChart.SeriesCollection(1).Name = ChartName
    MyChart.SeriesCollection(1).Values = ChartData
    
    'scale
    MyChart.Axes(xlValue).MinimumScale = 580
    MyChart.Axes(xlValue).MaximumScale = 735
    
    'zoom
    Sheet20.Activate
    ActiveWindow.Zoom = 125
    Sheet14.Activate
    
    'Create Image
    imageName = Application.DefaultFilePath & Application.PathSeparator & "TempChart.gif"
    MyChart.Export Filename:=imageName, FilterName:="GIF"
    Sheet20.ChartObjects(1).Delete
    Application.ScreenUpdating = True
    UserForm2.Image1.Picture = LoadPicture(imageName)
                                        
End Select

End Sub

Private Sub CommandButton5_Click()

UserForm1.Show

End Sub

Private Sub CommandButton6_Click()

Dim ChartIndex As Integer
Dim LastRow As Long
Dim FirstRow As Long
Dim myResult As Long
Dim ChartData As Range
Dim Spread1 As Single
Dim Spread2 As Single
Dim Spread3 As Single
Dim Spread4 As Single
Dim Spread5 As Single
Dim Spread6 As Single
Dim i As Single
Dim j As Single
Dim m As Single
Dim XBar As Single
Dim YBar As Single

ChartIndex = ComboBox1.ListIndex

Select Case ChartIndex
    
    Case 0
'Nissan L42P Assist: Case Depth
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
'Nissan L42P Sensor: Case Depth
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
'Nissan L42P Assist: Hardness
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
'Nissan L42P Sensor: Hardness
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
'BMW LHD CGR Assist Case Depth
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
'BMW LHD CGR Sensor Case Depth
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
'BMW LHD CGR Assist Hardness
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
'BMW LHD CGR Sensor Hardness
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
'BMW RHD CGR Assist: Case Depth
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
'BMW RHD CGR Sensor: Case Depth
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
'BMW RHD CGR Assist: Hardness
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
'BMW RHD CGR Sensor: Hardness
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
'BMW LHD VGR Assist: Case Depth
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
'BMW LHD VGR Assist: Hardness
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
'BMW RHD VGR Assist: Case Depth
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
'BMW RHD VGR Assist: Hardness
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
'Honda THR Assist: Case Depth
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
'Honda THR Sensor: Case Depth
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
'Honda THR Assist: Hardness
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
'Honda THR Sensor: Hardness
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
'Honda TJB Assist: Case Depth
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
'Honda TJB Assist: Hardness
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
'09PL Ball Screw: Hardness A
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
'09PL Ball Screw: Hardness B
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
'09PL Ball Screw: Case Depth
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
'09PL Rack Side: Tooth Hardness
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
'09PL Rack Side: Case Depth
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
'BMW G2X Assist: Tooth Hardness
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
'BMW G2X Assist: Case Depth
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
'BMW G2X Sensor: Tooth Hardness
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
'BMW G2X Sensor: Case Depth
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
'15PL Rack Side: Tooth Hardness
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
'15PL Rack Side: Case Depth
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
'15PL Ball Screw: Hardness A
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
'015PL Ball Screw: Hardness B
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
'09PL Ball Screw: Case Depth
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
'Nissan L42P Assist: Root Hardness
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
'Nissan L42P Sensor: Root Hardness
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
'BMW LHD CGR Assist: Root Hardness
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
'BMW LHD CGR Sensor: Root Hardness
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
'BMW RHD CGR Assist: Root Hardness
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
'BMW RHD CGR Sensor: Root Hardness
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
'BMW LHD VGR Assist: Root Hardness
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
'BMW RHD VGR Assist: Root Hardness
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
'Honda THR Assist: Root Hardness
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
'Honda THR Sensor: Root Hardness
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
'Honda TJB Assist: Root Hardness
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
'09PL Rack Side: Root Hardness
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
'BMW G2X Assist: Root Hardness
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
'BMW G2X Sensor: Root Hardness
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
'015PL Rack Side: Root Hardness
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

If Abs(Spread1 - Spread2) > 30 Or Abs(Spread2 - Spread3) > 30 Or Abs(Spread3 - Spread4) > 30 Or Abs(Spread4 - Spread5) > 30 Or Abs(Spread5 - Spread6) > 30 Then
    MsgBox ("Abnormality: Hardness Variance")
End If

If Spread1 < 7 Or Spread2 < 7 Or Spread3 < 7 Or Spread4 < 7 Or Spread5 < 7 Or Spread6 < 7 Then
    If Abs(Spread1 - Spread2) > 0.3 Or Abs(Spread2 - Spread3) > 0.3 Or Abs(Spread3 - Spread4) > 0.3 Or Abs(Spread4 - Spread5) > 0.3 Or Abs(Spread5 - Spread6) > 0.3 Then
        MsgBox ("Abnormality: Case Depth Variance")
    End If
End If

XBar = (1 + 2 + 3 + 4 + 5 + 6) / 6
YBar = (Spread1 + Spread2 + Spread3 + Spread4 + Spread5 + Spread6) / 6
i = ((1 - XBar) * (Spread6 - YBar)) + ((2 - XBar) * (Spread5 - YBar)) + ((3 - XBar) * (Spread4 - YBar)) + ((4 - XBar) * (Spread3 - YBar)) + ((5 - XBar) * (Spread2 - YBar)) + ((6 - XBar) * (Spread1 - YBar))
j = (((1 - XBar) * (1 - XBar)) + ((2 - XBar) * (2 - XBar)) + ((3 - XBar) * (3 - XBar)) + ((4 - XBar) * (4 - XBar)) + ((5 - XBar) * (5 - XBar)) + ((6 - XBar) * (6 - XBar)))
m = i / j

If m > 0 Then
    MsgBox ("Upward Trend")
End If

If m < 0 Then
    MsgBox ("Downward Trend")
End If

End Sub

Private Sub CommandButton7_Click()

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

Private Sub UserForm_Initialize()
    
    '0
    ComboBox1.AddItem ("Nissan L42P Assist: Case Depth")
    '1
    ComboBox1.AddItem ("Nissan L42P Sensor: Case Depth")
    '2
    ComboBox1.AddItem ("Nissan L42P Assist: Tooth Hardness")
    '3
    ComboBox1.AddItem ("Nissan L42P Sensor: Tooth Hardness")
    '4
    ComboBox1.AddItem ("BMW LHD CGR Assist: Case Depth")
    '5
    ComboBox1.AddItem ("BMW LHD CGR Sensor: Case Depth")
    '6
    ComboBox1.AddItem ("BMW LHD CGR Assist: Tooth Hardness")
    '7
    ComboBox1.AddItem ("BMW LHD CGR Sensor: Tooth Hardness")
    '8
    ComboBox1.AddItem ("BMW RHD CGR Assist: Case Depth")
    '9
    ComboBox1.AddItem ("BMW RHD CGR Sensor: Case Depth")
    '10
    ComboBox1.AddItem ("BMW RHD CGR Assist: Tooth Hardness")
    '11
    ComboBox1.AddItem ("BMW RHD CGR Sensor: Tooth Hardness")
    '12
    ComboBox1.AddItem ("BMW LHD VGR Assist: Case Depth")
    '13
    ComboBox1.AddItem ("BMW LHD VGR Assist: Tooth Hardness")
    '14
    ComboBox1.AddItem ("BMW RHD VGR Assist: Case Depth")
    '15
    ComboBox1.AddItem ("BMW RHD VGR Assist: Tooth Hardness")
    '16
    ComboBox1.AddItem ("Honda THR Assist: Case Depth")
    '17
    ComboBox1.AddItem ("Honda THR Sensor: Case Depth")
    '18
    ComboBox1.AddItem ("Honda THR Assist: Tooth Hardness")
    '19
    ComboBox1.AddItem ("Honda THR Sensor: Tooth Hardness")
    '20
    ComboBox1.AddItem ("Honda TJB Assist: Case Depth")
    '21
    ComboBox1.AddItem ("Honda TJB Assist: Tooth Hardness")
    '22
    ComboBox1.AddItem ("09PL Ball Screw: Hardness A")
    '23
    ComboBox1.AddItem ("09PL Ball Screw: Hardness B")
    '24
    ComboBox1.AddItem ("09PL Ball Screw: Case Depth")
    '25
    ComboBox1.AddItem ("09PL Rack Side: Tooth Hardness")
    '26
    ComboBox1.AddItem ("09PL Rack Side: Case Depth")
    '27
    ComboBox1.AddItem ("BMW G2X Assist: Tooth Hardness")
    '28
    ComboBox1.AddItem ("BMW G2X Assist: Case Depth")
    '29
    ComboBox1.AddItem ("BMW G2X Sensor: Tooth Hardness")
    '30
    ComboBox1.AddItem ("BMW G2X Sensor: Case Depth")
    '31
    ComboBox1.AddItem ("15PL Rack Side: Tooth Hardness")
    '32
    ComboBox1.AddItem ("15PL Rack Side: Case Depth")
    '33
    ComboBox1.AddItem ("15PL Ball Screw: Hardness A")
    '34
    ComboBox1.AddItem ("15PL Ball Screw: Hardness B")
    '35
    ComboBox1.AddItem ("15PL Ball Screw: Case Depth")
    '36
    ComboBox1.AddItem ("Nissan L42P Assist: Root Hardness")
    '37
    ComboBox1.AddItem ("Nissan L42P Sensor: Root Hardness")
    '38
    ComboBox1.AddItem ("BMW LHD CGR Assist: Root Hardness")
    '39
    ComboBox1.AddItem ("BMW LHD CGR Sensor: Root Hardness")
    '40
    ComboBox1.AddItem ("BMW RHD CGR Assist: Root Hardness")
    '41
    ComboBox1.AddItem ("BMW RHD CGR Sensor: Root Hardness")
    '42
    ComboBox1.AddItem ("BMW LHD VGR Assist: Root Hardness")
    '43
    ComboBox1.AddItem ("BMW RHD VGR Assist: Root Hardness")
    '44
    ComboBox1.AddItem ("Honda THR Assist: Root Hardness")
    '45
    ComboBox1.AddItem ("Honda THR Sensor: Root Hardness")
    '46
    ComboBox1.AddItem ("Honda TJB Assist: Root Hardness")
    '47
    ComboBox1.AddItem ("09PL Rack Side: Root Hardness")
    '48
    ComboBox1.AddItem ("BMW G2X Assist: Root Hardness")
    '49
    ComboBox1.AddItem ("BMW G2X Sensor: Root Hardness")
    '50
    ComboBox1.AddItem ("15PL Rack Side: Root Hardness")
    
'Procedures

    '0
    'Procedures.AddItem ("Add Product to Trend Chart")
    '1
    'Procedures.AddItem ("Abnormality Reporting")
    
    
End Sub

