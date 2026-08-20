'==============================================================================
' Draw Squound - user interface
'
' Collects the scale and starts the drawing. All geometry work is done in the
' main module.
'
'   Version   0.3.3
'   Date      2026-08-20
'   Author    James Debono
'   Licence   MIT - full text in the header of the main module
'   Source    https://github.com/james-debono/draw-squound-sw-macro
'
' Controls
'   TextBox1        scale, entered in millimetres
'   CommandButton1  Draw
'==============================================================================

Option Explicit

'--- Notes for maintenance ----------------------------------------------------
'
' This form owns the scale-to-dimensions rule and nothing else.
'
' The scale sets both the height and the width, and the corner radius comes out at
' CORNER_RADIUS_FRACTION of it. Height and width are passed as two separate
' numbers rather than tied together by an equal relation, so they can be pulled
' apart afterwards.
'------------------------------------------------------------------------------

Private Const CORNER_RADIUS_FRACTION As Double = 0.1

Private Sub CommandButton1_Click()

    Dim swApp As SldWorks.SldWorks
    Dim swModel As SldWorks.ModelDoc2

    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc

    If swModel Is Nothing Then
        MsgBox "No SolidWorks document is open. Please open a document and try again.", vbCritical
        Exit Sub
    End If

    Dim scaleValue As Double

    If Not TryReadMM(TextBox1, scaleValue) Then
        MsgBox "Please enter a valid number for the scale.", vbCritical
        TextBox1.SetFocus
        Exit Sub
    End If

    If scaleValue <= 0 Then
        MsgBox "The scale must be greater than zero.", vbCritical
        TextBox1.SetFocus
        Exit Sub
    End If

    ' DrawSquound takes width, then height, then corner radius, all in metres
    DrawSquound swApp, swModel, scaleValue, scaleValue, scaleValue * CORNER_RADIUS_FRACTION

    Me.Hide

End Sub

' Reads a millimetre value from a text box and converts it to metres.
' Returns False if the text is not a number.
Private Function TryReadMM(ByVal swTextBox As MSForms.TextBox, ByRef value As Double) As Boolean

    Dim txt As String
    txt = Trim$(swTextBox.Text)

    If Not IsNumeric(txt) Then
        TryReadMM = False
        Exit Function
    End If

    value = CDbl(txt) / 1000#
    TryReadMM = True

End Function

Private Sub UserForm_Initialize()

    ' Shown in the form's title bar, so the version is visible whenever the macro
    ' is used. Overrides whatever caption the designer holds.
    Me.Caption = "Draw Squound (" & MACRO_VERSION & ")"

    TextBox1.Text = "100"    ' scale, mm

End Sub
