Option Explicit On
Option Strict Off

Imports System.Windows.Forms
Imports Inventor
Imports ToolInventor2020.ToolInventor2020.Assembly.Buttons

Namespace ToolInventor2020.Toolngoai

    Public Module Button_Assembly

        Public Sub OnExecute(ByVal Context As NameValueMap)
            Select Case ShowMenu()

                Case 1  ' Constrain, Ground, Delete (combo sẵn)
                    Assembly.Buttons.caclenhlapghep.constraint.Ass_LG_C_1.OnExecute(Context)
                Case 2  ' Hiện gốc tọa độ để Constrain
                    Assembly.Buttons.caclenhlapghep.constraint.Ass_LG_C_2.OnExecute(Context)
                Case 3  ' Tắt ALL Adaptive cụm LG
                    Assembly.Buttons.caclenhlapghep.Ass_LG_1.OnExecute(Context)
                Case 4  ' UPDATE DESIGN STANDARD
                    Assembly.Buttons.caclenhlapghep.Ass_LG_2.OnExecute(Context)
                Case 5  ' Ẩn file
                    Assembly.Buttons.caclenhlapghep.Ass_LG_3.OnExecute(Context)
                Case 6  ' Covert to sheetmetal
                    Assembly.Buttons.Part.Ass_Part_1.OnExecute(Context)
                Case 7  ' Combo part 1
                    Assembly.Buttons.Part.Ass_Part_list_1.OnExecute(Context)
                Case 8  ' Trải ALL Sheetmetal
                    Assembly.Buttons.Part.Ass_Part_4.OnExecute(Context)
                Case 9  ' Save copy to replace part
                    Assembly.Buttons.Part.Ass_part_5.OnExecute(Context)
                Case 10 ' Ẩn all plane part
                    Assembly.Buttons.Part.Ass_Part_7.OnExecute(Context)
                Case 11 ' Lọc tấm, mua, Standard
                    Assembly.Buttons.caclenhboctach.part.Ass_boctach_part_1.OnExecute(Context)
                Case 12 ' Auto Drawing
                    Assembly.Buttons.AutoCreateDrawing.ASS_Auto_Drawing.OnExecute(Context)
                Case 13 ' Xem lỗi cắt Frame
                    Assembly.Buttons.Frame.Ass_Frame_1.OnExecute(Context)
                Case 14 ' Import, EX step & part
                    Assembly.Buttons.Lenhngoaicumlap.Im_EX_step_part.OnExecute(Context)
                Case 15 ' Design Assistant
                    Assembly.Buttons.Lenhngoaicumlap.Design_Assistant.OnExecute(Context)
                Case 16 ' Mở nơi lưu File
                    Toolngoai.Vitrifile.Vitrifile()
                Case 17 ' Xóa link OLE
                    Toolngoai.delete_link_OLE.OnExecute(Context)
                Case 18 ' Change View Representation
                    ASS_View.View_Representation.OnExecute(Context)

            End Select
        End Sub

        Private Function ShowMenu() As Integer
            Dim result As Integer = 0

            Using form As New Form()
                form.Text = "Assembly Tools"
                form.Width = 780
                form.Height = 560
                form.StartPosition = FormStartPosition.CenterScreen
                form.FormBorderStyle = FormBorderStyle.FixedDialog
                form.MaximizeBox = False
                form.MinimizeBox = False
                form.Tag = 0

                Dim title As New Label() With {
                    .Text = "Chọn lệnh Assembly (18 lệnh)",
                    .Left = 20, .Top = 10, .Width = 720, .Height = 26,
                    .Font = New System.Drawing.Font("Segoe UI", 11.0!, System.Drawing.FontStyle.Bold)
                }
                form.Controls.Add(title)

                ' Cột 1
                AddBtn(form, "1.  Constrain, Ground, Delete", 45, 20, 1)
                AddBtn(form, "2.  Hiện gốc tọa độ để Constrain", 85, 20, 2)
                AddBtn(form, "3.  Tắt ALL Adaptive cụm LG", 125, 20, 3)
                AddBtn(form, "4.  UPDATE DESIGN STANDARD", 165, 20, 4)
                AddBtn(form, "5.  Ẩn file", 205, 20, 5)
                AddBtn(form, "6.  Covert to sheetmetal", 245, 20, 6)
                AddBtn(form, "7.  Combo part 1", 285, 20, 7)
                AddBtn(form, "8.  Trải ALL Sheetmetal", 325, 20, 8)
                AddBtn(form, "9.  Save copy to replace part", 365, 20, 9)

                ' Cột 2
                AddBtn(form, "10. Ẩn all plane part", 45, 400, 10)
                AddBtn(form, "11. Lọc tấm, mua, Standard", 85, 400, 11)
                AddBtn(form, "12. Auto Drawing", 125, 400, 12)
                AddBtn(form, "13. Xem lỗi cắt Frame", 165, 400, 13)
                AddBtn(form, "14. Import / Export step & part", 205, 400, 14)
                AddBtn(form, "15. Design Assistant", 245, 400, 15)
                AddBtn(form, "16. Mở nơi lưu File", 285, 400, 16)
                AddBtn(form, "17. Xóa link OLE", 325, 400, 17)
                AddBtn(form, "18. Change View Representation", 365, 400, 18)

                Dim cancel As New Button() With {
                    .Text = "HỦY", .Left = 20, .Top = 420, .Width = 720, .Height = 32
                }
                AddHandler cancel.Click, Sub() form.Close()
                form.Controls.Add(cancel)

                form.ShowDialog()
                result = CInt(form.Tag)
            End Using

            Return result
        End Function

        Private Sub AddBtn(form As Form, text As String, top As Integer, left As Integer, value As Integer)
            Dim b As New Button() With {
                .Text = text,
                .Left = left,
                .Top = top,
                .Width = 350,
                .Height = 34,
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