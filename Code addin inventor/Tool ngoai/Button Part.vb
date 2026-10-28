Option Explicit On
Option Strict Off

Imports System.Windows.Forms
Imports Inventor

Namespace ToolInventor2020.Toolngoai

    Public Module Button_Part

        Public Sub OnExecute(ByVal Context As NameValueMap)
            Select Case ShowMenue()

                Case 1
                    Part.Buttons.solid.Part_Solid_1.OnExecute(Context)          ' Thêm dung sai vào dim sketch
                Case 2
                    Part.Buttons.solid.Part_Solid_2.OnExecute(Context)          ' Xóa fix all sketch
                Case 3
                    Part.Buttons.solid.Part_Solid_3.OnExecute(Context)          ' Tạo coil line 3d
                Case 4
                    Part.Buttons.solid.Part_Solid_4.OnExecute(Context)          ' Thay tên body
                Case 5
                    Part.Buttons.solid.Part_Solid_5.OnExecute(Context)          ' Export từng body
                Case 6
                    Part.Buttons.solid.Part_Solid_6.OnExecute(Context)          ' Thay tên body theo thứ tự
                Case 7
                    Assembly.Buttons.Lenhngoaicumlap.Im_EX_step_part.OnExecute(Context) ' Import step to part
                Case 8
                    Part.Buttons.Part_Solid_7e.OnExecute(Context)               ' Thay màu part
                Case 9
                    Toolngoai.Vitrifile.Vitrifile()                      ' Mở nơi lưu File
                Case 10
                    Toolngoai.delete_link_OLE.OnExecute(Context)                ' Xóa link OLE

            End Select
        End Sub

        Private Function ShowMenue() As Integer
            Dim result As Integer = 0

            Using form As New Form()
                form.Text = "Part Tools"
                form.Width = 520
                form.Height = 520
                form.StartPosition = FormStartPosition.CenterScreen
                form.FormBorderStyle = FormBorderStyle.FixedDialog
                form.MaximizeBox = False
                form.MinimizeBox = False
                form.Tag = 0

                Dim title As New Label() With {
                    .Text = "Chọn lệnh Part (10 lệnh)",
                    .Left = 20, .Top = 12, .Width = 460, .Height = 26,
                    .Font = New System.Drawing.Font("Segoe UI", 11.0!, System.Drawing.FontStyle.Bold)
                }
                form.Controls.Add(title)

                AddBtn(form, "1.  Thêm dung sai vào dim sketch", 50, 1)
                AddBtn(form, "2.  Xóa fix all sketch", 95, 2)
                AddBtn(form, "3.  Tạo coil line 3d", 140, 3)
                AddBtn(form, "4.  Thay tên body", 185, 4)
                AddBtn(form, "5.  Export từng body sang các định dạng khác", 230, 5)
                AddBtn(form, "6.  Thay tên body part theo thứ tự", 275, 6)
                AddBtn(form, "7.  Import step to part", 320, 7)
                AddBtn(form, "8.  Thay màu part", 365, 8)
                AddBtn(form, "9.  Mở nơi lưu File", 410, 9)
                AddBtn(form, "10. Xóa link OLE", 455, 10)

                Dim cancel As New Button() With {
                    .Text = "HỦY", .Left = 20, .Top = 510, .Width = 460, .Height = 32
                }
                AddHandler cancel.Click, Sub() form.Close()
                form.Controls.Add(cancel)

                form.Height = 590
                form.ShowDialog()
                result = CInt(form.Tag)
            End Using

            Return result
        End Function

        Private Sub AddBtn(form As Form, text As String, top As Integer, value As Integer)
            Dim b As New Button() With {
                .Text = text,
                .Left = 20,
                .Top = top,
                .Width = 460,
                .Height = 38,
                .TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            }
            AddHandler b.Click, Sub()
                                    form.Tag = value
                                    form.Close()
                                End Sub
            form.Controls.Add(b)
        End Sub

    End Module

End Namespace