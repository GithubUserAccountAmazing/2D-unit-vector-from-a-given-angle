Attribute VB_Name = "Module1"
Public theTriangle, A, B, C, Calt, d, slice, angleslice, _
    line, line2, line3, line4, line180, line0 As ShapeRange
Public lawb As Workbook
Public laws As Worksheet
Public j As Long
Public animation As Boolean


Sub prettyvisual() 
'runs through all angles 0-359

DoEvents
animation = True
Application.Interactive = False

    Workbooks("theunitCircle").Worksheets("circle").Shapes.Range(Array("Partial Circle 3")).Visible = False
    Workbooks("theunitCircle").Worksheets("circle").Shapes.Range(Array("Partial Circle 4")).Visible = False
    DoEvents
    Workbooks("theunitCircle").Worksheets("circle").Range("N34").Select
    DoEvents
    
    For j = 1 To 2
    
        If j = 1 Then Workbooks("theunitCircle").Worksheets("circle").Range("R34") = 0
        
        For i = 1 To 360
        
            DoEvents
            
            If j = 2 And i = 360 Then
            
                slice.Visible = False
                angleslice.Visible = False
                
            End If
            
            Workbooks("theunitCircle").Worksheets("circle").Range("R34") = _
                Workbooks("theunitCircle").Worksheets("circle").Range("R34") + 1
                
            DoEvents
            
        Next i
        Workbooks("theunitCircle").Worksheets("circle").Range("R34") = 0
        
    Next j

    DoEvents

animation = False
slice.Visible = False
angleslice.Visible = False
Application.Interactive = True
Application.ScreenUpdating = True


End Sub


Sub triangle()

Application.ScreenUpdating = False
Application.EnableEvents = False

Set lawb = Workbooks("theunitCircle")
Set laws = lawb.Worksheets("circle")

Set theTriangle = laws.Shapes.Range(Array("triangleshape"))
Set line = laws.Shapes.Range(Array("line"))
Set line2 = laws.Shapes.Range(Array("line2"))
Set line3 = laws.Shapes.Range(Array("line3"))
Set line4 = laws.Shapes.Range(Array("line4"))
Set A = laws.Shapes.Range(Array("tA"))
Set B = laws.Shapes.Range(Array("tB"))
Set C = laws.Shapes.Range(Array("tC"))
Set d = laws.Shapes.Range(Array("d"))
Set slice = laws.Shapes.Range(Array("Partial Circle 3"))
Set angleslice = laws.Shapes.Range(Array("Partial Circle 4"))


theTriangle.LockAspectRatio = msoFalse
DoEvents
On Error Resume Next

    While laws.Range("D8") >= 360
        DoEvents
        laws.Range("R34") = laws.Range("D8") - 360
        DoEvents
    Wend

    DoEvents
   
    theTriangle.Width = Abs(laws.Range("F8").Value * 500)
    
    DoEvents

    theTriangle.Height = Abs(laws.Range("E8").Value * 500)

    DoEvents
    
    While theTriangle.Width > 500 Or theTriangle.Height > 500
        theTriangle.Width = theTriangle.Width * 0.999
        theTriangle.Height = theTriangle.Height * 0.999
        DoEvents
    Wend

    Set line180 = laws.Shapes.Range(Array("line180"))
    Set line0 = laws.Shapes.Range(Array("line0"))
    Set Calt = laws.Shapes.Range(Array("tC0"))
    
    If line180.Visible = True Or line0.Visible = True Then
        line180.Visible = False
        line0.Visible = False
        Calt.Visible = False
        C.Visible = True
    End If
    
    DoEvents
    
    If laws.Range("D8").Value < 90 Then
    
        If laws.Range("D8") <> 0 Then
            line.Visible = True
            line2.Visible = False
            line3.Visible = False
            line4.Visible = False
            line.Width = theTriangle.Width
            line.Height = theTriangle.Height
            line.Top = 1025 - Abs(laws.Range("E8").Value * 500)
        Else
            theTriangle.Width = 500
            line.Visible = False
            line2.Visible = False
            line3.Visible = False
            line4.Visible = False
            line180.Visible = False
            line0.Visible = True
            C.Visible = False
            Calt.Visible = True
        End If
            
    Else
        If laws.Range("D8").Value < 180 Then
            line.Visible = False
            line2.Visible = True  ' surprisingly, having several line shapes
            line3.Visible = False ' instead of 1 was much easier to work with
            line4.Visible = False
            line2.Width = theTriangle.Width
            line2.Height = theTriangle.Height
            line2.Top = 1025 - Abs(laws.Range("E8").Value * 500)
            line2.Left = 813.75 - Abs(laws.Range("F8").Value * 500)
        Else
            If laws.Range("D8").Value < 270 Then
            
                If laws.Range("D8").Value <> 180 Then
                    line.Visible = False
                    line2.Visible = False
                    line3.Visible = True
                    line4.Visible = False
                    line3.Width = theTriangle.Width
                    line3.Height = theTriangle.Height
                    line3.Top = 1025
                    line3.Left = 813.75 - Abs(laws.Range("F8").Value * 500)
                Else
                    theTriangle.Width = 500
                    line.Visible = False
                    line2.Visible = False
                    line3.Visible = False
                    line4.Visible = False
                    line180.Visible = True
                    Calt.Visible = True
                    C.Visible = False
                End If
            Else
                If laws.Range("D8").Value < 360 Then
                    line.Visible = False
                    line2.Visible = False
                    line3.Visible = False
                    line4.Visible = True
                    line4.Width = theTriangle.Width
                    line4.Height = theTriangle.Height
                    line4.Left = 813.75
                End If
            End If
        End If
    End If
    
    DoEvents

    theTriangle.LockAspectRatio = msoTrue

    theTriangle.Top = 1025 - Abs(laws.Range("E8").Value * 500)
    theTriangle.Left = 813.75
    
    A.Top = 955 - Abs(laws.Range("E8").Value * 250)
    A.Left = 865 + Abs(laws.Range("F8").Value * 500)
    
    B.Top = 1000 - laws.Range("E8").Value * 500 + (laws.Range("E8").Value * 500)
    B.Left = 854.5 + 0.425 * Abs(laws.Range("F8").Value * 500)

    C.Top = 965 - 0.5 * Abs(laws.Range("E8").Value * 500)
    C.Left = 813.75 + 0.503 * Abs(laws.Range("F8").Value * 500)
    
    If animation = False Then
        slice.Adjustments.Item(1) = -45 - laws.Range("D8").Value
        angleslice.Adjustments.Item(1) = -45 - laws.Range("D8").Value
    Else
        If j <> 2 Then
            slice.Adjustments.Item(1) = -45 - laws.Range("D8").Value
            angleslice.Adjustments.Item(1) = -45 - laws.Range("D8").Value
        Else
            slice.Adjustments.Item(2) = -45 - laws.Range("D8").Value
            angleslice.Adjustments.Item(2) = -45 - laws.Range("D8").Value
        End If
    End If

    DoEvents

        DoEvents
        If slice.Adjustments.Item(1) = -45 Then
            If animation = False Then
                slice.Visible = False
                angleslice.Visible = False
            End If
        Else
            slice.Visible = True
            angleslice.Visible = True
        End If

    DoEvents
    
    While theTriangle.Width > 500 Or theTriangle.Height > 500
    
        theTriangle.Width = theTriangle.Width * 0.999
        theTriangle.Height = theTriangle.Height * 0.999
        
        DoEvents
    
    Wend
    
    Application.EnableEvents = True
    
    DoEvents
    
    If animation = True Then
    
        With Application
            .Interactive = False
            .ScreenUpdating = True
        End With
    
    Else
        
        Application.ScreenUpdating = True
        
    End If

End Sub


