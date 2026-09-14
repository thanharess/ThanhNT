Option Explicit On
Option Strict Off

Imports System.Windows.Forms
Imports Inventor

Namespace ToolInventor2020.Drawing.Buttons.Drawdim

    Public Module draw_2

        Public Sub OnExecute(ByVal Context As NameValueMap)
            Select Case ShowSheetMetalMenu()

                Case 1
                    Draw_2a.OnExecute(Context)
                Case 2
                    draw_2b.OnExecute(Context)
                Case 3
                    draw_2c.OnExecute(Context)
                Case 4
                    draw_2d.OnExecute(Context)
                Case 5
                    draw_2e.OnExecute(Context)
                Case 6
                    Draw_dim_hole.OnExecute(Context)
            End Select
        End Sub

        Private Function ShowSheetMetalMenu() As Integer
            Dim result As Integer = 0

            Using form As New Form()
                form.Text = "Ghi Partlist"
                form.Width = 550
                form.Height = 540
                form.StartPosition = FormStartPosition.CenterScreen
                form.FormBorderStyle = FormBorderStyle.FixedDialog
                form.MaximizeBox = False
                form.MinimizeBox = False

                Dim title As New Label() With {
                    .Text = "Ghi thêm, thay thông tin vào Partlist ENG,VIE", .Left = 20, .Top = 15,
                    .Width = 500, .Height = 28
                }
                form.Controls.Add(title)

                form.Tag = 0
                AddMenuButton(form, "Dim lỗ Dim về cạnh (Dim ít có bỏ qua lỗ nếu trùng, bỏ qua lỗ array)", 45, 1) 'ok
                AddMenuButton(form, "Dim lỗ Dim về cạnh (Dim tương đối có bỏ qua lỗ nếu trùng, bỏ qua lỗ array)", 95, 2) 'ok
                AddMenuButton(form, "Dim lỗ Dim về cạnh (Nhiều dim ko bỏ qua lỗ, dim tất cả các lỗ)", 145, 3) 'ok
                AddMenuButton(form, "Dim lỗ Base Dimline Set về cạnh", 195, 4) 'ok
                AddMenuButton(form, "Dim lỗ Base Dimline về cạnh", 245, 5) 'ok
                AddMenuButton(form, "Dim kích thước lỗ", 295, 6) ' ok
                ' AddMenuButton(form, "All lever lọc các loại tấm. lọc part", 345, 7) 'ok
                '  AddMenuButton(form, "All lever Bóc tách số lượng tổng cho PL, vật tư mua,.. lọc part", 395, 8) 'ok

                Dim cancelButton As New Button() With {
                    .Text = "HỦY", .Left = 20, .Top = 445, .Width = 500, .Height = 32
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