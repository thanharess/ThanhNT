Option Explicit On
Option Strict Off

Imports System.Windows.Forms
Imports Inventor

Namespace ToolInventor2020.Toolngoai

    ''' <summary>
    ''' 1 nút Drawing → hiện form chọn 20 lệnh
    ''' </summary>
    Public Module Button_Drawing

        Public Sub OnExecute(ByVal Context As NameValueMap)
            Select Case ShowDrawingMenu()

                Case 1
                    Drawing.Buttons.Draw_1.OnExecute(Context)                                   ' Sửa số thập phân dim
                Case 2
                    Drawing.Buttons.Drawdim.draw_2.OnExecute(Context)                           ' Auto Dim
                Case 3
                    Drawing.Buttons.Drawdim.draw_3.OnExecute(Context)                           ' Xóa, căn chỉnh dim
                Case 4
                    Drawing.Buttons.Draw_4.OnExecute(Context)                                   ' Tìm Dim bị edit
                Case 5
                    Drawing.Buttons.Draw_ResetPartsList.OnExecute(Context)                      ' Reset part list
                Case 6
                    Drawing.Buttons.DrawPartList.Draw_Partlist_1.OnExecute(Context)             ' Ghi Partlist ENG,VIE
                Case 7
                    Drawing.Buttons.Draw_7.OnExecute(Context)                                   ' Nút chuyển Sheet
                Case 8
                    Drawing.Buttons.DrawView.Draw_Change_scale_view.OnExecute(Context)          ' Đổi scale view
                Case 9
                    Drawing.Buttons.Draw_9.OnExecute(Context)                                   ' Xóa Part,ASS trùng lặp
                Case 10
                    Drawing.Buttons.Draw_5.OnExecute(Context)                                   ' Xóa centermark, centerline
                Case 11
                    Drawing.Buttons.DrawView.Draw_Up_date_view.OnExecute(Context)               ' Update Views
                Case 12
                    Drawing.Buttons.Draw_Rename_Sheet.OnExecute(Context)                        ' Rename sheet
                Case 13
                    Drawing.Buttons.DrawView.Draw_Rename_View.OnExecute(Context)                ' Thay chữ cái view sheet
                Case 14
                    Drawing.Buttons.draw_14.OnExecute(Context)                                  ' Đổi scale Hatch
                Case 15
                    Drawing.Buttons.Drawdelete.Draw_delete.OnExecute(Context)                   ' Delete
                Case 16
                    Toolngoai.Vitrifile.Vitrifile()                                   ' Mở nơi lưu File
                Case 17
                    Drawing.Buttons.Drawdim.Dim_Linear_Total.OnExecute(Context)                 ' Dim Kích thước tổng
                Case 18
                    Drawing.Buttons.DrawPartList.Create_PartsList.OnExecute(Context)            ' Create_PartsList
                Case 19
                    Drawing.Buttons.Drawtext.draw_text_list_1.OnExecute(Context)                ' Text Edit
                Case 20
                    Drawing.Buttons.DrawSheet.Drawing_TitleBlock_Border.OnExecute(Context)      ' Drawing_TitleBlock_Border

            End Select
        End Sub

        Private Function ShowDrawingMenu() As Integer
            Dim result As Integer = 0

            Using form As New Form()
                form.Text = "Drawing Tools - ToolInventor2020"
                form.Width = 780
                form.Height = 650
                form.StartPosition = FormStartPosition.CenterScreen
                form.FormBorderStyle = FormBorderStyle.FixedDialog
                form.MaximizeBox = False
                form.MinimizeBox = False
                form.Tag = 0

                Dim title As New Label() With {
                    .Text = "Chọn lệnh Drawing (20 lệnh)",
                    .Left = 20,
                    .Top = 12,
                    .Width = 720,
                    .Height = 26,
                    .Font = New System.Drawing.Font("Segoe UI", 11.0!, System.Drawing.FontStyle.Bold)
                }
                form.Controls.Add(title)

                '========== CỘT 1 ==========
                AddMenuButton(form, "1.  Sửa số thập phân dim", 50, 20, 1)
                AddMenuButton(form, "2.  Auto Dim", 95, 20, 2)
                AddMenuButton(form, "3.  Xóa / căn chỉnh dim", 140, 20, 3)
                AddMenuButton(form, "4.  Tìm Dim bị edit", 185, 20, 4)
                AddMenuButton(form, "5.  Reset part list", 230, 20, 5)
                AddMenuButton(form, "6.  Ghi Partlist ENG,VIE", 275, 20, 6)
                AddMenuButton(form, "7.  Nút chuyển Sheet", 320, 20, 7)
                AddMenuButton(form, "8.  Đổi scale view", 365, 20, 8)
                AddMenuButton(form, "9.  Xóa Part/ASS trùng lặp", 410, 20, 9)
                AddMenuButton(form, "10. Xóa centermark, centerline", 455, 20, 10)

                '========== CỘT 2 ==========
                AddMenuButton(form, "11. Update Views", 50, 400, 11)
                AddMenuButton(form, "12. Rename sheet", 95, 400, 12)
                AddMenuButton(form, "13. Thay chữ cái view sheet", 140, 400, 13)
                AddMenuButton(form, "14. Đổi scale Hatch", 185, 400, 14)
                AddMenuButton(form, "15. Delete", 230, 400, 15)
                AddMenuButton(form, "16. Mở nơi lưu File", 275, 400, 16)
                AddMenuButton(form, "17. Dim Kích thước tổng", 320, 400, 17)
                AddMenuButton(form, "18. Create_PartsList", 365, 400, 18)
                AddMenuButton(form, "19. Text Edit", 410, 400, 19)
                AddMenuButton(form, "20. TitleBlock / Border", 455, 400, 20)

                ' Nút Hủy
                Dim cancelButton As New Button() With {
                    .Text = "HỦY",
                    .Left = 20,
                    .Top = 515,
                    .Width = 720,
                    .Height = 32
                }
                AddHandler cancelButton.Click, Sub() form.Close()
                form.Controls.Add(cancelButton)

                form.ShowDialog()
                result = CInt(form.Tag)
            End Using

            Return result
        End Function

        Private Sub AddMenuButton(ByVal form As Form,
                                  ByVal text As String,
                                  ByVal top As Integer,
                                  ByVal left As Integer,
                                  ByVal value As Integer)

            Dim btn As New Button() With {
                .Text = text,
                .Left = left,
                .Top = top,
                .Width = 350,
                .Height = 38,
                .TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            }

            AddHandler btn.Click, Sub()
                                      form.Tag = value
                                      form.Close()
                                  End Sub

            form.Controls.Add(btn)
        End Sub

    End Module

End Namespace