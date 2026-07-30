Option Explicit

'==================================================================
'  Draw Squound 0.2.1  -  module "Draw_Squound1"
'
'  Draws a dimensioned "squound" (rounded rectangle) into the sketch
'  that is currently open for editing, centred on that sketch's
'  origin. Whichever plane or face the sketch sits on is the plane
'  the shape lands on.
'
'  Every length passed to DrawAndDimensionSketch is in METRES,
'  which is the SolidWorks internal unit system.
'==================================================================

Dim swApp As SldWorks.SldWorks
Dim swModel As SldWorks.ModelDoc2
Dim swSketchMgr As SldWorks.SketchManager

Sub main()

    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc

    If swModel Is Nothing Then
        MsgBox "No SolidWorks document is open. Please open a document and try again.", vbCritical
        Exit Sub
    End If

    Set swSketchMgr = swModel.SketchManager

    If swSketchMgr.ActiveSketch Is Nothing Then
        MsgBox "No sketch is open for editing. Start or edit a sketch, then run the macro again.", vbCritical
        Exit Sub
    End If

    UserForm1.Show

End Sub


'------------------------------------------------------------------
'  W = width  (along sketch X)
'  H = height (along sketch Y)
'  R = corner radius
'  All in metres.
'------------------------------------------------------------------
' The object parameters are ByVal so that VBA converts the interface for us.
' Passed ByRef, handing a SketchLine to a SketchSegment parameter (or an
' Object to a typed one) is a "ByRef argument type mismatch" compile error.
Public Sub DrawAndDimensionSketch(ByVal swAppIn As SldWorks.SldWorks, _
                                  ByVal swModelIn As SldWorks.ModelDoc2, _
                                  ByVal W As Double, ByVal H As Double, ByVal R As Double)

    Set swApp = swAppIn
    Set swModel = swModelIn
    Set swSketchMgr = swModel.SketchManager

    If W <= 0 Or H <= 0 Or R <= 0 Then
        MsgBox "Width, height and corner radius must all be greater than zero.", vbCritical
        Exit Sub
    End If

    If R >= W / 2 Or R >= H / 2 Then
        MsgBox "The corner radius must be smaller than half the width and half the height.", vbCritical
        Exit Sub
    End If

    ' --- draw into the sketch that is already open ------------------
    ' All coordinates below are sketch coordinates, so the shape lands
    ' on whatever plane or face this sketch belongs to.
    If swSketchMgr.ActiveSketch Is Nothing Then
        MsgBox "No sketch is open for editing. Start or edit a sketch, then run the macro again.", vbCritical
        Exit Sub
    End If

    swModel.ClearSelection2 True

    ' Put the geometry straight into the model database. Left at the default
    ' False, SketchManager draws through the user interface, and silently
    ' returns Nothing whenever the geometry falls outside the visible
    ' graphics area - which is the usual reason a working sketch macro
    ' suddenly starts failing.
    Dim bAddToDBOrig As Boolean
    bAddToDBOrig = swSketchMgr.AddToDB

    On Error GoTo CleanUp
    swSketchMgr.AddToDB = True

    ' --- corner coordinates ---------------------------------------
    Dim halfW As Double: halfW = W / 2
    Dim halfH As Double: halfH = H / 2
    Dim xL As Double: xL = -halfW + R      ' where the top/bottom straights begin
    Dim xR As Double: xR = halfW - R
    Dim yB As Double: yB = -halfH + R      ' where the left/right straights begin
    Dim yT As Double: yT = halfH - R

    Dim swTopLn As SldWorks.SketchLine
    Dim swRightLn As SldWorks.SketchLine
    Dim swBottomLn As SldWorks.SketchLine
    Dim swLeftLn As SldWorks.SketchLine
    Dim swArcTR As SldWorks.SketchArc
    Dim swArcBR As SldWorks.SketchArc
    Dim swArcBL As SldWorks.SketchArc
    Dim swArcTL As SldWorks.SketchArc

    ' Walked clockwise from the top-left, so every corner turns clockwise (-1).
    ' CreateArc takes centre, then start, then end.
    Set swTopLn = swSketchMgr.CreateLine(xL, halfH, 0#, xR, halfH, 0#)
    Set swArcTR = swSketchMgr.CreateArc(xR, yT, 0#, xR, halfH, 0#, halfW, yT, 0#, -1)
    Set swRightLn = swSketchMgr.CreateLine(halfW, yT, 0#, halfW, yB, 0#)
    Set swArcBR = swSketchMgr.CreateArc(xR, yB, 0#, halfW, yB, 0#, xR, -halfH, 0#, -1)
    Set swBottomLn = swSketchMgr.CreateLine(xR, -halfH, 0#, xL, -halfH, 0#)
    Set swArcBL = swSketchMgr.CreateArc(xL, yB, 0#, xL, -halfH, 0#, -halfW, yB, 0#, -1)
    Set swLeftLn = swSketchMgr.CreateLine(-halfW, yB, 0#, -halfW, yT, 0#)
    Set swArcTL = swSketchMgr.CreateArc(xL, yT, 0#, -halfW, yT, 0#, xL, halfH, 0#, -1)

    swSketchMgr.AddToDB = bAddToDBOrig

    ' --- close the loop -------------------------------------------
    ' AddToDB = True means nothing is inferred, so every relation is
    ' added by hand below.
    JoinAt swTopLn, swArcTR, xR, halfH
    JoinAt swArcTR, swRightLn, halfW, yT
    JoinAt swRightLn, swArcBR, halfW, yB
    JoinAt swArcBR, swBottomLn, xR, -halfH
    JoinAt swBottomLn, swArcBL, xL, -halfH
    JoinAt swArcBL, swLeftLn, -halfW, yB
    JoinAt swLeftLn, swArcTL, -halfW, yT
    JoinAt swArcTL, swTopLn, xL, halfH

    ' --- tangency at each corner ----------------------------------
    AddRelation2 swTopLn, swArcTR, "sgTANGENT"
    AddRelation2 swArcTR, swRightLn, "sgTANGENT"
    AddRelation2 swRightLn, swArcBR, "sgTANGENT"
    AddRelation2 swArcBR, swBottomLn, "sgTANGENT"
    AddRelation2 swBottomLn, swArcBL, "sgTANGENT"
    AddRelation2 swArcBL, swLeftLn, "sgTANGENT"
    AddRelation2 swLeftLn, swArcTL, "sgTANGENT"
    AddRelation2 swArcTL, swTopLn, "sgTANGENT"

    ' --- keep the straights axis-aligned --------------------------
    AddRelation1 swTopLn, "sgHORIZONTAL"
    AddRelation1 swBottomLn, "sgHORIZONTAL"
    AddRelation1 swRightLn, "sgVERTICAL"
    AddRelation1 swLeftLn, "sgVERTICAL"

    ' --- one radius drives all four corners -----------------------
    swModel.ClearSelection2 True
    swArcTR.Select4 True, Nothing
    swArcBR.Select4 True, Nothing
    swArcBL.Select4 True, Nothing
    swArcTL.Select4 True, Nothing
    swModel.SketchAddConstraints "sgSAMELENGTH"
    swModel.ClearSelection2 True

    ' --- dimensions -----------------------------------------------
    ' Measured between the two opposing straight sides, so the value is
    ' the overall size rather than the shortened straight run.
    Dim gap As Double
    gap = R + 0.01                                   ' clear of the profile

    AddLinearDim swLeftLn, swRightLn, 0#, -halfH - gap, W, "Width"
    AddLinearDim swTopLn, swBottomLn, halfW + gap, 0#, H, "Height"

    Dim swRadDim As SldWorks.DisplayDimension
    swModel.ClearSelection2 True
    swArcTR.Select4 True, Nothing
    Set swRadDim = swModel.AddDimension2(halfW + gap, halfH + gap, 0#)
    SetDimValue swRadDim, R, "CornerRadius"
    swModel.ClearSelection2 True

    ' The sketch is left open so you can carry on working in it.
    ' AddToDB bypasses the graphics layer, so ask for a redraw.
    swModel.GraphicsRedraw2

    Exit Sub

CleanUp:
    swSketchMgr.AddToDB = bAddToDBOrig
    MsgBox "Draw Squound failed: " & Err.Description & " (error " & Err.Number & ")", vbCritical

End Sub


'------------------------------------------------------------------
'  Helpers
'------------------------------------------------------------------

' Merges the endpoints of two segments that meet near (x, y).
' Picking the nearer endpoint avoids depending on how SolidWorks
' orders an arc's start and end points.
Private Sub JoinAt(ByVal swSeg1 As SldWorks.SketchSegment, ByVal swSeg2 As SldWorks.SketchSegment, _
                   ByVal x As Double, ByVal y As Double)

    Dim swPt1 As SldWorks.SketchPoint
    Dim swPt2 As SldWorks.SketchPoint

    Set swPt1 = EndPointNear(swSeg1, x, y)
    Set swPt2 = EndPointNear(swSeg2, x, y)

    swModel.ClearSelection2 True
    swPt1.Select4 True, Nothing
    swPt2.Select4 True, Nothing
    swModel.SketchAddConstraints "sgMERGEPOINTS"
    swModel.ClearSelection2 True

End Sub

Private Function EndPointNear(ByVal swSeg As SldWorks.SketchSegment, _
                              ByVal x As Double, ByVal y As Double) As SldWorks.SketchPoint

    Dim swPt1 As SldWorks.SketchPoint
    Dim swPt2 As SldWorks.SketchPoint

    If swSeg.GetType() = swSketchSegments_e.swSketchLINE Then
        Dim swLine As SldWorks.SketchLine
        Set swLine = swSeg
        Set swPt1 = swLine.GetStartPoint2
        Set swPt2 = swLine.GetEndPoint2
    Else
        Dim swArc As SldWorks.SketchArc
        Set swArc = swSeg
        Set swPt1 = swArc.GetStartPoint2
        Set swPt2 = swArc.GetEndPoint2
    End If

    If (swPt1.x - x) ^ 2 + (swPt1.y - y) ^ 2 <= (swPt2.x - x) ^ 2 + (swPt2.y - y) ^ 2 Then
        Set EndPointNear = swPt1
    Else
        Set EndPointNear = swPt2
    End If

End Function

Private Sub AddRelation1(ByVal swSeg As SldWorks.SketchSegment, ByVal constraintName As String)

    swModel.ClearSelection2 True
    swSeg.Select4 True, Nothing
    swModel.SketchAddConstraints constraintName
    swModel.ClearSelection2 True

End Sub

Private Sub AddRelation2(ByVal swSeg1 As SldWorks.SketchSegment, ByVal swSeg2 As SldWorks.SketchSegment, _
                         ByVal constraintName As String)

    swModel.ClearSelection2 True
    swSeg1.Select4 True, Nothing
    swSeg2.Select4 True, Nothing
    swModel.SketchAddConstraints constraintName
    swModel.ClearSelection2 True

End Sub

Private Sub AddLinearDim(ByVal swSeg1 As SldWorks.SketchSegment, ByVal swSeg2 As SldWorks.SketchSegment, _
                         ByVal x As Double, ByVal y As Double, _
                         ByVal dimValue As Double, ByVal dimName As String)

    Dim swDispDim As SldWorks.DisplayDimension

    swModel.ClearSelection2 True
    swSeg1.Select4 True, Nothing
    swSeg2.Select4 True, Nothing
    Set swDispDim = swModel.AddDimension2(x, y, 0#)
    SetDimValue swDispDim, dimValue, dimName
    swModel.ClearSelection2 True

End Sub

' AddDimension2 hands back a DisplayDimension, which is only the
' annotation. The driving value lives on the Dimension underneath it,
' and SystemValue is always in metres.
Private Sub SetDimValue(ByVal swDispDim As SldWorks.DisplayDimension, _
                        ByVal dimValue As Double, ByVal dimName As String)

    If swDispDim Is Nothing Then Exit Sub

    Dim swDim As SldWorks.Dimension
    Set swDim = swDispDim.GetDimension
    If swDim Is Nothing Then Exit Sub

    swDim.SystemValue = dimValue
    swDim.Name = dimName

End Sub
