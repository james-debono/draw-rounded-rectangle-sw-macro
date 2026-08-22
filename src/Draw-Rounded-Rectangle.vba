'==============================================================================
' Draw Rounded Rectangle
'
' Draws a rounded rectangle into the sketch you already have open for editing,
' centred on that sketch's origin. Whichever plane or face the sketch sits on is
' where the shape lands.
'
' Type one number for the scale and the corner radius comes out at one tenth of it.
'
' No dimensions are added. The shape is built from sketch relations only, so it
' arrives under-defined and ready to be dragged or dimensioned by hand. Every
' relation is added explicitly - eight merged endpoints, eight tangents, four
' horizontal and vertical, and three equal-radius - which leaves exactly five
' degrees of freedom: position in x and y, plus width, height and radius.
'
' That is the right number. Drag a straight side and the shape stays a rounded
' rectangle rather than coming apart, and if you add your own width, height and
' radius dimensions later the sketch becomes fully defined with no redundancy.
'
' To use, open or start a sketch, then run the macro.
'
'   Version   0.4.0
'   Date      2026-08-21
'   Author    James Debono
'   Licence   MIT - full text below
'   Source    https://github.com/james-debono/draw-rounded-rectangle-sw-macro
'
'------------------------------------------------------------------------------
' CHANGELOG (summary - see CHANGELOG.md for the full history)
'
'   0.4.0   Renamed from "Draw Squound". Now has its own repository.
'   0.3.3   Version shown in brackets in the title bar. Source URL updated.
'   0.3.2   Version shown in the form's title bar.
'   0.3.1   Licence and header.
'   0.3.0   One scale input instead of three, and no driving dimensions.
'   0.2.1   Draws into the active sketch rather than creating one.
'   0.2.0   Drawing routine rewritten against the verified API.
'   0.1.0   Initial version.
'
'------------------------------------------------------------------------------
' MIT Licence
' SPDX-License-Identifier: MIT
'
' Copyright (c) 2026 James Debono
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.
'==============================================================================

Option Explicit

'--- Notes for maintenance ----------------------------------------------------
'
' Every length passed to DrawRoundedRectangle is in METRES, the SolidWorks internal unit
' system. The form converts from millimetres.
'
' DrawRoundedRectangle stays general and knows nothing about "scale" - a future form with
' separate width and height inputs needs no change here.
'
' ISketchManager.AddToDB is set True before creating geometry and restored after.
' Left at the default, SketchManager draws through the user interface and
' silently returns Nothing when geometry falls outside the visible graphics area.
' That is the classic cause of a sketch macro that "used to work".
'------------------------------------------------------------------------------

' Must match the Version line in the header block above. build-library.ps1 checks
' that they agree and fails the build if they drift. Shown in the form's title bar.
Public Const MACRO_VERSION As String = "0.4.0"

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
Public Sub DrawRoundedRectangle(ByVal swAppIn As SldWorks.SldWorks, _
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

    ' No dimensions are added. The relations above leave five degrees of
    ' freedom - position x and y, width, height and corner radius - so the
    ' shape is drawn at the requested size but stays free to adjust.

    ' The sketch is left open so you can carry on working in it.
    ' AddToDB bypasses the graphics layer, so ask for a redraw.
    swModel.GraphicsRedraw2

    Exit Sub

CleanUp:
    swSketchMgr.AddToDB = bAddToDBOrig
    MsgBox "Draw Rounded Rectangle failed: " & Err.Description & " (error " & Err.Number & ")", vbCritical

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

