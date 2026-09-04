Option Explicit On
Option Strict Off

Imports System.Windows.Forms
Imports Inventor

Namespace ToolInventor2020.Assembly.Buttons.caclenhboctach.part

    'Gán module này cho một nút Inventor. Các Ass_11, Ass_12, Ass_13 vẫn giữ nguyên.
    Public Module Ass_boctach_part_1

        Public Sub OnExecute(ByVal Context As NameValueMap)
            Select Case ShowSheetMetalMenu()
                Case 1
                    'Chỉ Sheet Metal ở Assembly cấp hiện tại.
                    Ass_boctach_part_1a.OnExecute(Context)
                Case 2
                    Ass_boctach_part_1d.OnExecute(Context)
                Case 3
                    Ass_boctach_part_1f.OnExecute(Context)
                Case 4
                    'Quét cả Sub Assembly, giữ số lần xuất hiện thực tế của Part.
                    Ass_boctach_part_1c.OnExecute(Context)
                Case 5
                    'Quét cả Sub Assembly, mỗi file Part chỉ thêm một lần.
                    Ass_boctach_part_1b.OnExecute(Context)
                Case 6
                    Ass_boctach_part_1e.OnExecute(Context)
                Case 7
                    Ass_boctach_part_1fpart.OnExecute(Context)

            End Select
        End Sub

        Private Function ShowSheetMetalMenu() As Integer
            Dim result As Integer = 0

            Using form As New Form()
                form.Text = "Sheet Metal Unfold"
                form.Width = 560
                form.Height = 450
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
                AddMenuButton(form, "1 Assembly top lever Bóc tách số lượng tổng tấm lọc part", 55, 1)
                AddMenuButton(form, "2 Assembly top lever tách số loại sp cho các vật tư mua, tiêu chuẩn lọc part", 105, 2)
                AddMenuButton(form, "3 Assembly + part top lever Bóc tách số lượng tổng cho các PL & vật tư mua,.. lọc ASS & part", 155, 3)
                AddMenuButton(form, "4 All Assembly Bóc tách số lượng tổng tấm lọc part", 205, 4)
                AddMenuButton(form, "5 All Assembly lọc các loại tấm lọc part", 255, 5)
                AddMenuButton(form, "6 All Assembly Bóc tách số lượng tổng cho các PL & vật tư mua,.. lọc part", 305, 6)
                AddMenuButton(form, "7 Assembly top lever Bóc tách số lượng tổng cho các PL & vật tư mua,.. lọc part", 355, 7)

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