Option Explicit On
Option Strict Off

Imports System.Windows.Forms
Imports Inventor

Namespace ToolInventor2020.Drawing.Buttons
    Public Module Draw_8
        Public Sub OnExecute(ByVal Context As NameValueMap)
            Try
                Dim invApp As Inventor.Application = g_inventorApplication
                If invApp Is Nothing Then
                    MessageBox.Show("Không tìm thấy Inventor Application.", "Scale View", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                If invApp.ActiveDocument Is Nothing OrElse invApp.ActiveDocument.DocumentType <> DocumentTypeEnum.kDrawingDocumentObject Then
                    MessageBox.Show("Chức năng này chỉ dùng cho bản vẽ Drawing.", "Scale View", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Exit Sub
                End If

                Dim drawingDocument As DrawingDocument = CType(invApp.ActiveDocument, DrawingDocument)
                Dim scale As Double
                If Not PromptForScale(scale) Then Exit Sub

                Dim choice As DialogResult = MessageBox.Show(
                    "Có thay đổi tỷ lệ cho tất cả view của sheet hiện tại không?" & vbCrLf & vbCrLf &
                    "Yes: tất cả view" & vbCrLf &
                    "No: chọn một view", "Scale View", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)

                If choice = DialogResult.Cancel Then Exit Sub

                If choice = DialogResult.Yes Then
                    For Each drawingView As DrawingView In drawingDocument.ActiveSheet.DrawingViews
                        If Not drawingView.ScaleFromBase Then drawingView.Scale = scale
                    Next
                Else
                    Dim pickedObject As Object = invApp.CommandManager.Pick(SelectionFilterEnum.kDrawingViewFilter, "Chọn Drawing View cần đổi tỷ lệ")
                    If pickedObject Is Nothing Then Exit Sub

                    Dim drawingView As DrawingView = CType(pickedObject, DrawingView)
                    drawingView.ScaleFromBase = False
                    drawingView.Scale = scale
                End If

                drawingDocument.Update()
            Catch ex As Exception
                MessageBox.Show("Không thể đổi tỷ lệ view:" & vbCrLf & ex.Message, "Scale View", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        Private Function PromptForScale(ByRef scale As Double) As Boolean
            Const allowedScales As String = "5/1, 2/1, 1, 1/2, 1/4, 1/5, 1/6, 1/8, 1/10, 1/15, 1/20, 1/25, 1/30, 1/35, 1/40, 1/45, 1/50, 1/60, 1/75, 1/80, 1/100, 1/125, 1/150"

            Do
                Dim input As String = Microsoft.VisualBasic.Interaction.InputBox(
                    "Nhập tỷ lệ view (ví dụ: 1/2 hoặc 2)." & vbCrLf & allowedScales,
                    "Scale View", "1")

                If String.IsNullOrWhiteSpace(input) Then Return False
                If TryParseScale(input, scale) Then Return True

                MessageBox.Show("Tỷ lệ không hợp lệ. Ví dụ hợp lệ: 1/2, 1, 2 hoặc 5/1.", "Scale View", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Loop
        End Function

        Private Function TryParseScale(ByVal text As String, ByRef scale As Double) As Boolean
            Dim value As String = text.Trim().Replace(" ", "")
            Dim separator As Integer = value.IndexOf("/"c)

            If separator >= 0 Then
                Dim numerator As Double
                Dim denominator As Double
                If Not Double.TryParse(value.Substring(0, separator), numerator) OrElse
                   Not Double.TryParse(value.Substring(separator + 1), denominator) OrElse denominator = 0 Then Return False
                scale = numerator / denominator
            ElseIf Not Double.TryParse(value, scale) Then
                Return False
            End If

            Return scale > 0
        End Function
    End Module
End Namespace
