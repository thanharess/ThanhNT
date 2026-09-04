Option Explicit On
Option Strict Off

Imports System.Windows.Forms
Imports Inventor

Namespace ToolInventor2020.Assembly.Buttons.caclenhlapghep

    'Gán module này cho một nút Inventor. Các Ass_11, Ass_12, Ass_13 vẫn giữ nguyên.
    Public Module ass_5

        Public Sub OnExecute(ByVal Context As NameValueMap)
            Select Case ShowSheetMetalMenu()
                Case 1
                    'Chỉ Sheet Metal ở Assembly cấp hiện tại.
                    Ass_11.OnExecute(Context)

                Case 2
                    'Quét cả Sub Assembly, mỗi file Part chỉ thêm một lần.
                    Ass_12.OnExecute(Context)

                Case 3
                    'Quét cả Sub Assembly, giữ số lần xuất hiện thực tế của Part.
                    Ass_13.OnExecute(Context)
            End Select
        End Sub

        Private Function ShowSheetMetalMenu() As Integer
            Dim result As Integer = 0

            Using form As New Form()
                form.Text = "Sheet Metal Unfold"
                form.Width = 560
                form.Height = 275
                form.StartPosition = FormStartPosition.CenterScreen
                form.FormBorderStyle = FormBorderStyle.FixedDialog
                form.MaximizeBox = False
                form.MinimizeBox = False

                Dim title As New Label() With {
                    .Text = "CHỌN CÁCH TẠO ASSEMBLY SHEET METAL", .Left = 20, .Top = 15,
                    .Width = 500, .Height = 28
                }
                form.Controls.Add(title)

                form.Tag = 0
                AddMenuButton(form, "1. Chỉ lấy Sheet Metal ở Assembly hiện tại", 55, 1)
                AddMenuButton(form, "2. Quét toàn bộ Sub Assembly — loại Part trùng", 105, 2)
                AddMenuButton(form, "3. Quét toàn bộ Sub Assembly — giữ Part trùng", 155, 3)

                Dim cancelButton As New Button() With {
                    .Text = "HỦY", .Left = 20, .Top = 207, .Width = 500, .Height = 32
                }
                AddHandler cancelButton.Click, Sub() form.Close()
                form.Controls.Add(cancelButton)

                form.ShowDialog()
                result = CInt(form.Tag)
            End Using

            Return result
        End Function

        Private Sub AddMenuButton(ByVal form As Form, ByVal text As String, ByVal top As Integer, ByVal value As Integer)
            Dim button As New Button() With {
                .Text = text, .Left = 20, .Top = top, .Width = 500, .Height = 42
            }
            AddHandler button.Click, Sub()
                                         form.Tag = value
                                         form.Close()
                                     End Sub
            form.Controls.Add(button)
        End Sub

    End Module

End Namespace