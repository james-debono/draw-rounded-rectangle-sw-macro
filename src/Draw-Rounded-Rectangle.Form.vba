Option Explicit

'==================================================================
'  Draw Squound 0.2.1  -  code behind "UserForm1"
'
'  TextBox1 = Height, TextBox2 = Width, TextBox3 = Corner Radius,
'  all entered in millimetres to match the form labels.
'==================================================================

Private Sub CommandButton1_Click()

    Dim swApp As SldWorks.SldWorks
    Dim swModel As SldWorks.ModelDoc2

    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc

    If swModel Is Nothing Then
        MsgBox "No SolidWorks document is open. Please open a document and try again.", vbCritical
        Exit Sub
    End If

    Dim H As Double
    Dim W As Double
    Dim R As Double

    If Not TryReadMM(TextBox1, H) Then
        MsgBox "Please enter a valid number for the height.", vbCritical
        TextBox1.SetFocus
        Exit Sub
    End If

    If Not TryReadMM(TextBox2, W) Then
        MsgBox "Please enter a valid number for the width.", vbCritical
        TextBox2.SetFocus
        Exit Sub
    End If

    If Not TryReadMM(TextBox3, R) Then
        MsgBox "Please enter a valid number for the corner radius.", vbCritical
        TextBox3.SetFocus
        Exit Sub
    End If

    ' DrawAndDimensionSketch takes width first, then height, both in metres
    DrawAndDimensionSketch swApp, swModel, W, H, R

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

    TextBox1.Text = "100"    ' height, mm
    TextBox2.Text = "100"    ' width, mm
    TextBox3.Text = "10"     ' corner radius, mm

End Sub
