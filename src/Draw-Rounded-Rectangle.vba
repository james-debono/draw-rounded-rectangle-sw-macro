Option Explicit

' Global declarations (as per your original structure)
Dim swApp As SldWorks.SldWorks
Dim swModel As SldWorks.ModelDoc2
Dim swSketchMgr As SldWorks.SketchManager
Dim swSelMgr As SldWorks.SelectionMgr
Dim boolstatus As Boolean
Dim longstatus As Long, longwarnings As Long

Sub main()

    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    
    If swModel Is Nothing Then
        MsgBox "No SolidWorks document is open. Please open a part document and try again.", vbCritical
        Exit Sub
    End If

    ' Show the form
    UserForm1.Show
    
End Sub

Public Sub DrawAndDimensionSketch(swApp As SldWorks.SldWorks, swModel As SldWorks.ModelDoc2, H As Double, W As Double, R As Double)

    Set swSketchMgr = swModel.SketchManager
    Set swSelMgr = swModel.SelectionMgr

    ' Ensure a sketch is active on the Front Plane
    swModel.ClearSelection2 True
    boolstatus = swModel.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
    
    ' Enter or activate sketch on Front Plane.
    ' If a sketch is already active on Front Plane, this will re-enter it.
    ' If a sketch is active on a different plane, it will exit and re-enter on Front Plane.
    swSketchMgr.InsertSketch True

    ' Delete existing sketch entities in the current sketch to start clean
    If Not swSketchMgr.ActiveSketch Is Nothing Then
        Dim vSketchSegs As Variant
        vSketchSegs = swSketchMgr.ActiveSketch.GetSketchSegments
        If Not IsEmpty(vSketchSegs) Then
            Dim i As Long
            Dim swSketchSegment_Temp As SldWorks.SketchSegment ' Local declaration for loop
            For i = 0 To UBound(vSketchSegs)
                Set swSketchSegment_Temp = vSketchSegs(i)
                swSketchSegment_Temp.Select4 True, Nothing ' Use Select4 for robustness
            Next i
            swModel.EditDelete
        End If
        
        Dim vSketchPoints As Variant
        vSketchPoints = swSketchMgr.ActiveSketch.GetSketchPoints
        If Not IsEmpty(vSketchPoints) Then
            For i = 0 To UBound(vSketchPoints)
                vSketchPoints(i).Select4 True, Nothing ' Use Select4 for robustness
            Next i
            swModel.EditDelete
        End If
    End If
    
    swModel.ClearSelection2 True
    
    ' Maximize view for better visibility
    Dim myModelView As Object
    Set myModelView = swModel.ActiveView
    myModelView.FrameState = swWindowState_e.swWindowMaximized

    ' Define half dimensions for easier calculation
    Dim halfH As Double: halfH = H / 2
    Dim halfW As Double: halfW = W / 2
    
    ' Coordinates of the corners, considering the radius R
    ' These points define the ends of the straight lines and the centers/ends of the arcs
    Dim x_straight_end_left As Double: x_straight_end_left = -halfH + R
    Dim x_straight_end_right As Double: x_straight_end_right = halfH - R
    Dim y_straight_end_top As Double: y_straight_end_top = halfW - R
    Dim y_straight_end_bottom As Double: y_straight_end_bottom = -halfW + R
    
    ' Drawing the lines and arcs
    Dim skSegment1 As SldWorks.SketchSegment ' Top Line (left segment)
    Dim skSegment2 As SldWorks.SketchSegment ' Top-Right Arc
    Dim skSegment3 As SldWorks.SketchSegment ' Right Line
    Dim skSegment4 As SldWorks.SketchSegment ' Bottom-Right Arc
    Dim skSegment5 As SldWorks.SketchSegment ' Bottom Line (right segment)
    Dim skSegment6 As SldWorks.SketchSegment ' Bottom-Left Arc
    Dim skSegment7 As SldWorks.SketchSegment ' Left Line
    Dim skSegment8 As SldWorks.SketchSegment ' Top-Left Arc
    
    ' Order of creation ensures tangent connections
    ' Top line (left to right)
    Set skSegment1 = swSketchMgr.CreateLine(x_straight_end_left, halfW, 0#, x_straight_end_right, halfW, 0#)
    ' Top-right arc
    Set skSegment2 = swSketchMgr.CreateTangentArc(x_straight_end_right, halfW, 0#, halfH, y_straight_end_top, 0#, 1)
    ' Right line (top to bottom)
    Set skSegment3 = swSketchMgr.CreateTangentArc(halfH, y_straight_end_top, 0#, halfH, y_straight_end_bottom, 0#, 1) ' Start from end of arc2, go to y_straight_end_bottom
    ' Bottom-right arc
    Set skSegment4 = swSketchMgr.CreateTangentArc(halfH, y_straight_end_bottom, 0#, x_straight_end_right, -halfW, 0#, 1)
    ' Bottom line (right to left)
    Set skSegment5 = swSketchMgr.CreateTangentArc(x_straight_end_right, -halfW, 0#, x_straight_end_left, -halfW, 0#, 1)
    ' Bottom-left arc
    Set skSegment6 = swSketchMgr.CreateTangentArc(x_straight_end_left, -halfW, 0#, -halfH, y_straight_end_bottom, 0#, 1)
    ' Left line (bottom to top)
    Set skSegment7 = swSketchMgr.CreateTangentArc(-halfH, y_straight_end_bottom, 0#, -halfH, y_straight_end_top, 0#, 1)
    ' Top-left arc - connects back to the start of skSegment1
    Set skSegment8 = swSketchMgr.CreateTangentArc(-halfH, y_straight_end_top, 0#, x_straight_end_left, halfW, 0#, 1)
    
    ' Apply Explicit Tangent Constraints (usually inferred, but good for robustness)
    ' Between skSegment8 (Top-Left Arc) and skSegment1 (Top Line)
    swModel.ClearSelection2 True
    skSegment8.Select4 True, Nothing
    skSegment1.Select4 True, Nothing
    swSketchMgr.AddConstraint swSketchCONSTRAINTTYPE_e.swConstraintType_SketchTangent

    ' Between skSegment1 (Top Line) and skSegment2 (Top-Right Arc)
    swModel.ClearSelection2 True
    skSegment1.Select4 True, Nothing
    skSegment2.Select4 True, Nothing
    swSketchMgr.AddConstraint swSketchCONSTRAINTTYPE_e.swConstraintType_SketchTangent

    ' Between skSegment2 (Top-Right Arc) and skSegment3 (Right Line)
    swModel.ClearSelection2 True
    skSegment2.Select4 True, Nothing
    skSegment3.Select4 True, Nothing
    swSketchMgr.AddConstraint swSketchCONSTRAINTTYPE_e.swConstraintType_SketchTangent

    ' Between skSegment3 (Right Line) and skSegment4 (Bottom-Right Arc)
    swModel.ClearSelection2 True
    skSegment3.Select4 True, Nothing
    skSegment4.Select4 True, Nothing
    swSketchMgr.AddConstraint swSketchCONSTRAINTTYPE_e.swConstraintType_SketchTangent
    
    ' Between skSegment4 (Bottom-Right Arc) and skSegment5 (Bottom Line)
    swModel.ClearSelection2 True
    skSegment4.Select4 True, Nothing
    skSegment5.Select4 True, Nothing
    swSketchMgr.AddConstraint swSketchCONSTRAINTTYPE_e.swConstraintType_SketchTangent

    ' Between skSegment5 (Bottom Line) and skSegment6 (Bottom-Left Arc)
    swModel.ClearSelection2 True
    skSegment5.Select4 True, Nothing
    skSegment6.Select4 True, Nothing
    swSketchMgr.AddConstraint swSketchCONSTRAINTTYPE_e.swConstraintType_SketchTangent
    
    ' Between skSegment6 (Bottom-Left Arc) and skSegment7 (Left Line)
    swModel.ClearSelection2 True
    skSegment6.Select4 True, Nothing
    skSegment7.Select4 True, Nothing
    swSketchMgr.AddConstraint swSketchCONSTRAINTTYPE_e.swConstraintType_SketchTangent
    
    ' Between skSegment7 (Left Line) and skSegment8 (Top-Left Arc)
    swModel.ClearSelection2 True
    skSegment7.Select4 True, Nothing
    skSegment8.Select4 True, Nothing
    swSketchMgr.AddConstraint swSketchCONSTRAINTTYPE_e.swConstraintType_SketchTangent
    
    ' Add Dimensions
    Dim myDisplayDimH As SldWorks.DisplayDimension
    Dim myDisplayDimW As SldWorks.DisplayDimension
    Dim myDisplayDimR As SldWorks.DisplayDimension

    ' Horizontal Dimension (H)
    swModel.ClearSelection2 True
    skSegment1.Select4 True, Nothing ' Select the top line
    ' Position the dimension above the top line, with some offset
    Set myDisplayDimH = swModel.AddDimension2(0, halfW + (0.02 * swModel.GetUserUnit(swLengthUnit_e.swLengthUnit_Meters)), 0)
    If Not myDisplayDimH Is Nothing Then
        myDisplayDimH.Value = H
        myDisplayDimH.Name = "D1@Sketch1" ' Name the dimension D1
    End If
    
    ' Vertical Dimension (W)
    swModel.ClearSelection2 True
    skSegment3.Select4 True, Nothing ' Select the right line
    ' Position the dimension to the right of the right line, with some offset
    Set myDisplayDimW = swModel.AddDimension2(halfH + (0.02 * swModel.GetUserUnit(swLengthUnit_e.swLengthUnit_Meters)), 0, 0)
    If Not myDisplayDimW Is Nothing Then
        myDisplayDimW.Value = W
        myDisplayDimW.Name = "D2@Sketch1" ' Name the dimension D2
    End If
    
    ' Radius Dimension (R)
    swModel.ClearSelection2 True
    skSegment2.Select4 True, Nothing ' Select one of the arcs
    ' Position the dimension near the top-right corner, with some offset
    Set myDisplayDimR = swModel.AddDimension2(halfH + (0.02 * swModel.GetUserUnit(swLengthUnit_e.swLengthUnit_Meters)), halfW + (0.02 * swModel.GetUserUnit(swLengthUnit_e.swLengthUnit_Meters)), 0)
    If Not myDisplayDimR Is Nothing Then
        myDisplayDimR.Value = R
        myDisplayDimR.Name = "D3@Sketch1" ' Name the dimension D3
    End If
    
    ' Make all arcs equal radius (since we're controlling with one R dimension)
    swModel.ClearSelection2 True
    skSegment2.Select4 True, Nothing
    skSegment4.Select4 True, Nothing
    skSegment6.Select4 True, Nothing
    skSegment8.Select4 True, Nothing
    swSketchMgr.AddConstraint swSketchCONSTRAINTTYPE_e.swConstraintType_SketchEqualRadius
    
    ' Ensure lines are horizontal/vertical where appropriate
    swModel.ClearSelection2 True
    skSegment1.Select4 True, Nothing
    swSketchMgr.AddConstraint swSketchCONSTRAINTTYPE_e.swConstraintType_SketchHorizontal
    
    swModel.ClearSelection2 True
    skSegment3.Select4 True, Nothing
    swSketchMgr.AddConstraint swSketchCONSTRAINTTYPE_e.swConstraintType_SketchVertical
    
    swModel.ClearSelection2 True
    skSegment5.Select4 True, Nothing
    swSketchMgr.AddConstraint swSketchCONSTRAINTTYPE_e.swConstraintType_SketchHorizontal
    
    swModel.ClearSelection2 True
    skSegment7.Select4 True, Nothing
    swSketchMgr.AddConstraint swSketchCONSTRAINTTYPE_e.swConstraintType_SketchVertical
    
    swModel.EditRebuild3 ' Rebuild to apply changes

    swSketchMgr.InsertSketch True ' Exit the sketch
    
End Sub