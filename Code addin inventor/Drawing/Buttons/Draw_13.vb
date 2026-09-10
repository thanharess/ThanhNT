Option Explicit On
Option Strict Off

Imports Inventor
Imports System.Windows.Forms
Imports System.Drawing

Namespace ToolInventor2020.Drawing.Buttons

    Public Module draw_13

        Public Sub OnExecute(ByVal Context As NameValueMap)

            Try
                Dim oDrawDoc As DrawingDocument =
                    TryCast(g_inventorApplication.ActiveDocument, DrawingDocument)

                If oDrawDoc Is Nothing Then
                    MessageBox.Show("Document hiện tại không phải Drawing.",
                                    "Rename View Label",
                                    MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return
                End If

                Dim startLetter As String = ""
                Dim scope As String = ""
                Dim skipProjected As Boolean = False

                If Not ShowInputDialog(startLetter, scope, skipProjected) Then Return

                startLetter = startLetter.Trim().ToUpper()

                If startLetter.Length <> 1 OrElse Not Char.IsLetter(startLetter.Chars(0)) Then
                    MessageBox.Show("Chữ cái bắt đầu phải là 1 ký tự A-Z.",
                                    "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return
                End If

                Dim count As Integer = 0

                If scope = "Active Sheet" Then
                    count = RenameViewsOnSheet(oDrawDoc.ActiveSheet, startLetter, skipProjected)
                Else
                    count = RenameViewsAllSheets(oDrawDoc, startLetter, skipProjected)
                End If

                oDrawDoc.Update2(True)

                Dim msg As String = "Đã đặt lại tên " & count.ToString() & " View(s)" & vbCrLf &
                                    "Bắt đầu từ chữ: " & startLetter & vbCrLf &
                                    "Đã bỏ qua Base View"

                If skipProjected Then
                    msg &= vbCrLf & "Đã bỏ qua Projected View"
                End If

                MessageBox.Show(msg, "Rename View Label", MessageBoxButtons.OK, MessageBoxIcon.Information)

            Catch ex As Exception
                MessageBox.Show("Lỗi: " & ex.Message,
                                "Rename View Label",
                                MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try

        End Sub


        '==========================================================
        ' RENAME TRÊN 1 SHEET
        '==========================================================
        Private Function RenameViewsOnSheet(ByVal oSheet As Sheet, ByVal startLetter As String, ByVal skipProjected As Boolean) As Integer

            If oSheet Is Nothing Then Return 0

            Dim currentChar As Integer = Asc(startLetter)
            Dim count As Integer = 0

            For Each oView As DrawingView In oSheet.DrawingViews

                Try
                    ' Bỏ qua Base View
                    If oView.ViewType = DrawingViewTypeEnum.kStandardDrawingViewType Then Continue For

                    ' Bỏ qua Overlay View
                    If oView.ViewType = DrawingViewTypeEnum.kOverlayDrawingViewType Then Continue For

                    ' Bỏ qua Projected View (nếu được chọn)
                    If skipProjected AndAlso oView.ViewType = DrawingViewTypeEnum.kProjectedDrawingViewType Then
                        Continue For
                    End If

                    oView.Name = Chr(currentChar)

                    currentChar += 1
                    count += 1

                    If currentChar > Asc("Z") Then
                        currentChar = Asc("A")
                    End If

                Catch
                End Try

            Next

            Return count

        End Function


        '==========================================================
        ' RENAME TẤT CẢ SHEET
        '==========================================================
        Private Function RenameViewsAllSheets(ByVal oDrawDoc As DrawingDocument, ByVal startLetter As String, ByVal skipProjected As Boolean) As Integer

            Dim total As Integer = 0
            Dim currentChar As Integer = Asc(startLetter)

            For Each oSheet As Sheet In oDrawDoc.Sheets

                For Each oView As DrawingView In oSheet.DrawingViews
                    Try
                        ' Bỏ qua Base View
                        If oView.ViewType = DrawingViewTypeEnum.kStandardDrawingViewType Then Continue For

                        ' Bỏ qua Overlay View
                        If oView.ViewType = DrawingViewTypeEnum.kOverlayDrawingViewType Then Continue For

                        ' Bỏ qua Projected View (nếu được chọn)
                        If skipProjected AndAlso oView.ViewType = DrawingViewTypeEnum.kProjectedDrawingViewType Then
                            Continue For
                        End If

                        oView.Name = Chr(currentChar)

                        currentChar += 1
                        total += 1

                        If currentChar > Asc("Z") Then
                            currentChar = Asc("A")
                        End If

                    Catch
                    End Try
                Next
            Next

            Return total

        End Function


        '==========================================================
        ' FORM NHẬP
        '==========================================================
        Private Function ShowInputDialog(ByRef startLetter As String, ByRef scope As String, ByRef skipProjected As Boolean) As Boolean

            Dim frm As New Form()
            frm.Text = "Đặt tên View theo chữ cái"
            frm.Size = New Size(380, 260)
            frm.StartPosition = FormStartPosition.CenterScreen
            frm.FormBorderStyle = FormBorderStyle.FixedDialog
            frm.MaximizeBox = False
            frm.MinimizeBox = False
            frm.Font = New Font("Segoe UI", 9)

            Dim lbl1 As New Label()
            lbl1.Text = "Chữ cái bắt đầu (A, B, C...):"
            lbl1.Location = New System.Drawing.Point(20, 20)
            lbl1.AutoSize = True

            Dim txtLetter As New System.Windows.Forms.TextBox()
            txtLetter.Text = "A"
            txtLetter.Location = New System.Drawing.Point(20, 45)
            txtLetter.Width = 60
            txtLetter.MaxLength = 1
            txtLetter.CharacterCasing = CharacterCasing.Upper

            Dim lbl2 As New Label()
            lbl2.Text = "Phạm vi:"
            lbl2.Location = New System.Drawing.Point(20, 85)
            lbl2.AutoSize = True

            Dim cboScope As New ComboBox()
            cboScope.Location = New System.Drawing.Point(20, 110)
            cboScope.Width = 320
            cboScope.DropDownStyle = ComboBoxStyle.DropDownList
            cboScope.Items.Add("Active Sheet")
            cboScope.Items.Add("All Sheets")
            cboScope.SelectedIndex = 0

            Dim chkSkipProjected As New CheckBox()
            chkSkipProjected.Text = "Bỏ qua Projected View"
            chkSkipProjected.Location = New System.Drawing.Point(20, 150)
            chkSkipProjected.AutoSize = True
            chkSkipProjected.Checked = True

            Dim btnOK As New Button()
            btnOK.Text = "OK"
            btnOK.Location = New System.Drawing.Point(180, 185)
            btnOK.Width = 70
            btnOK.DialogResult = DialogResult.OK

            Dim btnCancel As New Button()
            btnCancel.Text = "Cancel"
            btnCancel.Location = New System.Drawing.Point(260, 185)
            btnCancel.Width = 70
            btnCancel.DialogResult = DialogResult.Cancel

            frm.Controls.AddRange({lbl1, txtLetter, lbl2, cboScope, chkSkipProjected, btnOK, btnCancel})
            frm.AcceptButton = btnOK
            frm.CancelButton = btnCancel

            If frm.ShowDialog() = DialogResult.OK Then
                startLetter = txtLetter.Text
                scope = cboScope.SelectedItem.ToString()
                skipProjected = chkSkipProjected.Checked
                Return True
            End If

            Return False

        End Function

    End Module

End Namespace