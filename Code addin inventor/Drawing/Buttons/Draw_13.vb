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
                Dim skipBase As Boolean = True
                Dim skipProjected As Boolean = False
                Dim skipSection As Boolean = False
                Dim skipDetail As Boolean = False
                Dim skipAuxiliary As Boolean = False
                Dim skipOverlay As Boolean = True

                If Not ShowInputDialog(startLetter, scope,
                                       skipBase, skipProjected, skipSection,
                                       skipDetail, skipAuxiliary, skipOverlay) Then Return

                startLetter = startLetter.Trim().ToUpper()

                If String.IsNullOrEmpty(startLetter) OrElse Not IsValidLetter(startLetter) Then
                    MessageBox.Show("Chữ cái bắt đầu không hợp lệ (chỉ dùng A-Z, AA, AB...).",
                                    "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return
                End If

                Dim count As Integer = 0

                If scope = "Active Sheet" Then
                    count = RenameViewsOnSheet(oDrawDoc.ActiveSheet, startLetter,
                                               skipBase, skipProjected, skipSection,
                                               skipDetail, skipAuxiliary, skipOverlay)
                Else
                    count = RenameViewsAllSheets(oDrawDoc, startLetter,
                                                 skipBase, skipProjected, skipSection,
                                                 skipDetail, skipAuxiliary, skipOverlay)
                End If

                oDrawDoc.Update2(True)

                Dim msg As String = "Đã đặt lại tên " & count.ToString() & " View(s)" & vbCrLf &
                                    "Bắt đầu từ: " & startLetter & vbCrLf

                If skipBase Then msg &= "• Bỏ qua Base View" & vbCrLf
                If skipProjected Then msg &= "• Bỏ qua Projected View" & vbCrLf
                If skipSection Then msg &= "• Bỏ qua Section View" & vbCrLf
                If skipDetail Then msg &= "• Bỏ qua Detail View" & vbCrLf
                If skipAuxiliary Then msg &= "• Bỏ qua Auxiliary View" & vbCrLf
                If skipOverlay Then msg &= "• Bỏ qua Overlay View" & vbCrLf

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
        Private Function RenameViewsOnSheet(ByVal oSheet As Sheet,
                                            ByVal startLetter As String,
                                            ByVal skipBase As Boolean,
                                            ByVal skipProjected As Boolean,
                                            ByVal skipSection As Boolean,
                                            ByVal skipDetail As Boolean,
                                            ByVal skipAuxiliary As Boolean,
                                            ByVal skipOverlay As Boolean) As Integer

            If oSheet Is Nothing Then Return 0

            Dim currentName As String = startLetter
            Dim count As Integer = 0

            For Each oView As DrawingView In oSheet.DrawingViews

                Try
                    If ShouldSkipView(oView, skipBase, skipProjected, skipSection,
                                      skipDetail, skipAuxiliary, skipOverlay) Then
                        Continue For
                    End If

                    oView.Name = currentName
                    currentName = GetNextLetter(currentName)
                    count += 1

                Catch
                End Try

            Next

            Return count

        End Function


        '==========================================================
        ' RENAME TẤT CẢ SHEET
        '==========================================================
        Private Function RenameViewsAllSheets(ByVal oDrawDoc As DrawingDocument,
                                              ByVal startLetter As String,
                                              ByVal skipBase As Boolean,
                                              ByVal skipProjected As Boolean,
                                              ByVal skipSection As Boolean,
                                              ByVal skipDetail As Boolean,
                                              ByVal skipAuxiliary As Boolean,
                                              ByVal skipOverlay As Boolean) As Integer

            Dim total As Integer = 0
            Dim currentName As String = startLetter

            For Each oSheet As Sheet In oDrawDoc.Sheets
                For Each oView As DrawingView In oSheet.DrawingViews
                    Try
                        If ShouldSkipView(oView, skipBase, skipProjected, skipSection,
                                          skipDetail, skipAuxiliary, skipOverlay) Then
                            Continue For
                        End If

                        oView.Name = currentName
                        currentName = GetNextLetter(currentName)
                        total += 1

                    Catch
                    End Try
                Next
            Next

            Return total

        End Function


        '==========================================================
        ' KIỂM TRA BỎ QUA VIEW
        '==========================================================
        Private Function ShouldSkipView(ByVal oView As DrawingView,
                                        ByVal skipBase As Boolean,
                                        ByVal skipProjected As Boolean,
                                        ByVal skipSection As Boolean,
                                        ByVal skipDetail As Boolean,
                                        ByVal skipAuxiliary As Boolean,
                                        ByVal skipOverlay As Boolean) As Boolean

            Select Case oView.ViewType
                Case DrawingViewTypeEnum.kStandardDrawingViewType
                    Return skipBase
                Case DrawingViewTypeEnum.kProjectedDrawingViewType
                    Return skipProjected
                Case DrawingViewTypeEnum.kSectionDrawingViewType
                    Return skipSection
                Case DrawingViewTypeEnum.kDetailDrawingViewType
                    Return skipDetail
                Case DrawingViewTypeEnum.kAuxiliaryDrawingViewType
                    Return skipAuxiliary
                Case DrawingViewTypeEnum.kOverlayDrawingViewType
                    Return skipOverlay
                Case Else
                    Return False
            End Select

        End Function


        '==========================================================
        ' TĂNG CHỮ CÁI: A → B → ... → Z → AA → AB → ...
        '==========================================================
        Private Function GetNextLetter(ByVal current As String) As String

            If String.IsNullOrEmpty(current) Then Return "A"

            Dim chars() As Char = current.ToUpper().ToCharArray()
            Dim i As Integer = chars.Length - 1

            Do While i >= 0
                If chars(i) < "Z"c Then
                    chars(i) = Chr(Asc(chars(i)) + 1)
                    Return New String(chars)
                Else
                    chars(i) = "A"c
                    i -= 1
                End If
            Loop

            ' Đã vượt Z, ZZ... → thêm 1 chữ A ở đầu
            Return "A" & New String(chars)

        End Function


        '==========================================================
        ' KIỂM TRA CHỮ HỢP LỆ
        '==========================================================
        Private Function IsValidLetter(ByVal text As String) As Boolean
            For Each c As Char In text
                If Not Char.IsLetter(c) Then Return False
            Next
            Return True
        End Function


        '==========================================================
        ' FORM
        '==========================================================
        Private Function ShowInputDialog(ByRef startLetter As String,
                                         ByRef scope As String,
                                         ByRef skipBase As Boolean,
                                         ByRef skipProjected As Boolean,
                                         ByRef skipSection As Boolean,
                                         ByRef skipDetail As Boolean,
                                         ByRef skipAuxiliary As Boolean,
                                         ByRef skipOverlay As Boolean) As Boolean

            Dim frm As New Form()
            frm.Text = "Đặt tên View theo chữ cái"
            frm.Size = New Size(400, 430)
            frm.StartPosition = FormStartPosition.CenterScreen
            frm.FormBorderStyle = FormBorderStyle.FixedDialog
            frm.MaximizeBox = False
            frm.MinimizeBox = False
            frm.Font = New Font("Segoe UI", 9)

            Dim lbl1 As New Label()
            lbl1.Text = "Chữ cái bắt đầu (A, B, AA, AB...):"
            lbl1.Location = New System.Drawing.Point(20, 10)
            lbl1.AutoSize = True

            Dim txtLetter As New System.Windows.Forms.TextBox()
            txtLetter.Text = "A"
            txtLetter.Location = New System.Drawing.Point(20, 40)
            txtLetter.Width = 80
            txtLetter.CharacterCasing = CharacterCasing.Upper

            Dim lbl2 As New Label()
            lbl2.Text = "Phạm vi:"
            lbl2.Location = New System.Drawing.Point(20, 70)
            lbl2.AutoSize = True

            Dim cboScope As New ComboBox()
            cboScope.Location = New System.Drawing.Point(20, 98)
            cboScope.Width = 340
            cboScope.DropDownStyle = ComboBoxStyle.DropDownList
            cboScope.Items.Add("Active Sheet")
            cboScope.Items.Add("All Sheets")
            cboScope.SelectedIndex = 0

            Dim lbl3 As New Label()
            lbl3.Text = "Bỏ qua loại View:"
            lbl3.Location = New System.Drawing.Point(20, 140)
            lbl3.AutoSize = True

            Dim chkBase As New CheckBox() With {.Text = "Base View", .Location = New System.Drawing.Point(20, 165), .AutoSize = True, .Checked = True}
            Dim chkProjected As New CheckBox() With {.Text = "Projected View", .Location = New System.Drawing.Point(20, 190), .AutoSize = True, .Checked = True}
            Dim chkSection As New CheckBox() With {.Text = "Section View", .Location = New System.Drawing.Point(20, 215), .AutoSize = True, .Checked = False}
            Dim chkDetail As New CheckBox() With {.Text = "Detail View", .Location = New System.Drawing.Point(20, 240), .AutoSize = True, .Checked = False}
            Dim chkAuxiliary As New CheckBox() With {.Text = "Auxiliary View", .Location = New System.Drawing.Point(20, 265), .AutoSize = True, .Checked = False}
            Dim chkOverlay As New CheckBox() With {.Text = "Overlay View", .Location = New System.Drawing.Point(20, 290), .AutoSize = True, .Checked = True}

            Dim btnOK As New Button() With {.Text = "OK", .Location = New System.Drawing.Point(200, 345), .Width = 70, .DialogResult = DialogResult.OK}
            Dim btnCancel As New Button() With {.Text = "Cancel", .Location = New System.Drawing.Point(280, 345), .Width = 70, .DialogResult = DialogResult.Cancel}

            frm.Controls.AddRange({
                lbl1, txtLetter, lbl2, cboScope, lbl3,
                chkBase, chkProjected, chkSection, chkDetail, chkAuxiliary, chkOverlay,
                btnOK, btnCancel
            })

            frm.AcceptButton = btnOK
            frm.CancelButton = btnCancel

            If frm.ShowDialog() = DialogResult.OK Then
                startLetter = txtLetter.Text
                scope = cboScope.SelectedItem.ToString()
                skipBase = chkBase.Checked
                skipProjected = chkProjected.Checked
                skipSection = chkSection.Checked
                skipDetail = chkDetail.Checked
                skipAuxiliary = chkAuxiliary.Checked
                skipOverlay = chkOverlay.Checked
                Return True
            End If

            Return False

        End Function

    End Module

End Namespace