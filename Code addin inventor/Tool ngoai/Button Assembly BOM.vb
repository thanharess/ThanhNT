Option Explicit On
Option Strict Off

Imports System.Windows.Forms
Imports Inventor

Namespace ToolInventor2020.Toolngoai

    Public Module Button_Assembly2

        Public Sub OnExecute(ByVal Context As NameValueMap)
            Select Case ShowMenuq()

                Case 1
                    Assembly2.Buttons.BOMcode.Ass_Bom_1.OnExecute(Context)      ' Thay tên ,STT BOM
                Case 2
                    Assembly2.Buttons.BOMcode.Ass_Bom_2.OnExecute(Context)      ' Item1 Buy Top lever về cuối
                Case 3
                    Assembly2.Buttons.BOMcode.Ass_Bom_3.OnExecute(Context)      ' Ghi item1,t,SL
                Case 4
                    Assembly2.Buttons.BOMcode.Ass_Bom_4.OnExecute(Context)      ' Ghi PL part,stocknumber
                Case 5
                    Assembly2.Buttons.BOMcode.ass_bom_6.OnExecute(Context)     ' Xếp Model Browser ASS
                Case 6
                    Assembly2.Buttons.BOMcode.ass_bom_7.OnExecute(Context)     ' Thay đổi đuôi tên BOM

            End Select
        End Sub

        Private Function ShowMenuq() As Integer
            Dim result As Integer = 0

            Using form As New Form()
                form.Text = "BOM Tools"
                form.Width = 500
                form.Height = 420
                form.StartPosition = FormStartPosition.CenterScreen
                form.FormBorderStyle = FormBorderStyle.FixedDialog
                form.MaximizeBox = False
                form.MinimizeBox = False
                form.Tag = 0

                Dim title As New Label() With {
                    .Text = "Chọn lệnh BOM (6 lệnh)",
                    .Left = 20, .Top = 12, .Width = 440, .Height = 26,
                    .Font = New System.Drawing.Font("Segoe UI", 11.0!, System.Drawing.FontStyle.Bold)
                }
                form.Controls.Add(title)

                AddBtn(form, "1. Thay tên , STT BOM", 50, 1)
                AddBtn(form, "2. Item1 Buy Top lever về cuối", 95, 2)
                AddBtn(form, "3. Ghi item1, t, SL", 140, 3)
                AddBtn(form, "4. Ghi PL part, stocknumber", 185, 4)
                AddBtn(form, "5. Xếp Model Browser ASS", 230, 5)
                AddBtn(form, "6. Thay đổi đuôi tên BOM", 275, 6)

                Dim cancel As New Button() With {
                    .Text = "HỦY", .Left = 20, .Top = 330, .Width = 440, .Height = 32
                }
                AddHandler cancel.Click, Sub() form.Close()
                form.Controls.Add(cancel)

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
                .Width = 440,
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