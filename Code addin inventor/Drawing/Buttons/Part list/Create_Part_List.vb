Option Explicit On
Option Strict Off

Imports System.Windows.Forms
Imports Inventor
Imports System.Collections.Generic

Namespace ToolInventor2020.Drawing.Buttons.DrawPartList

    Public Module Create_PartsList

        Public Sub OnExecute(ByVal Context As NameValueMap)

            Dim app As Inventor.Application = g_inventorApplication

            Try
                '=====================================================
                ' 1. KIỂM TRA DRAWING
                '=====================================================
                If app.ActiveDocument Is Nothing OrElse
                   app.ActiveDocument.DocumentType <> DocumentTypeEnum.kDrawingDocumentObject Then

                    MessageBox.Show(
                        "Vui lòng mở file Drawing (.idw)!",
                        "Lỗi",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error)

                    Exit Sub
                End If

                Dim oDrawDoc As DrawingDocument =
                    CType(app.ActiveDocument, DrawingDocument)

                Dim oSheet As Sheet = oDrawDoc.ActiveSheet
                Dim tg As TransientGeometry = app.TransientGeometry

                '=====================================================
                ' 2. CHỌN VIEW CỦA ASSEMBLY
                '=====================================================
                Dim pickedObj As Object = Nothing

                Try
                    oDrawDoc.SelectSet.Clear()
                Catch
                End Try

                Try
                    pickedObj = app.CommandManager.Pick(
                        SelectionFilterEnum.kDrawingViewFilter,
                        "Chọn view của Assembly cần tạo Parts List")
                Catch
                    Exit Sub
                End Try

                If pickedObj Is Nothing Then Exit Sub

                Dim oView As DrawingView = TryCast(pickedObj, DrawingView)

                If oView Is Nothing Then
                    MessageBox.Show(
                        "Không phải DrawingView.",
                        "Lỗi",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning)

                    Exit Sub
                End If

                '=====================================================
                ' 3. KIỂM TRA VIEW CÓ PHẢI ASSEMBLY
                '=====================================================
                Dim refDoc As Document = Nothing

                Try
                    refDoc = oView.ReferencedDocumentDescriptor.ReferencedDocument
                Catch
                End Try

                If refDoc Is Nothing Then
                    MessageBox.Show(
                        "View không có model tham chiếu.",
                        "Lỗi",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning)

                    Exit Sub
                End If

                If refDoc.DocumentType <>
                   DocumentTypeEnum.kAssemblyDocumentObject Then

                    MessageBox.Show(
                        "View này không tham chiếu Assembly." &
                        vbCrLf & vbCrLf &
                        "Parts List chỉ tạo được từ view của Assembly.",
                        "Lỗi",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning)

                    Exit Sub
                End If

                '=====================================================
                ' 4. CHỌN STYLE PARTS LIST
                '=====================================================
                Dim styleName As String = PickPartsListStyle(oDrawDoc)

                ' Hủy chọn Style = dừng toàn bộ
                If styleName = "" Then Exit Sub

                '=====================================================
                ' 5. NHẬP OFFSET
                '=====================================================
                Dim sOffset As String = InputBox(
                    "Khoảng cách từ view (mm) — đặt Parts List ở góc trên phải view:" &
                    vbCrLf & vbCrLf &
                    "Ví dụ: 20",
                    "Tạo Parts List",
                    "20")

                If String.IsNullOrWhiteSpace(sOffset) Then Exit Sub

                Dim offsetMM As Double = 20.0

                If Not Double.TryParse(
                    sOffset.Replace(",", ".").Trim(),
                    Globalization.NumberStyles.Any,
                    Globalization.CultureInfo.InvariantCulture,
                    offsetMM) Then

                    MessageBox.Show(
                        "Khoảng cách không hợp lệ.",
                        "Lỗi",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning)

                    Exit Sub
                End If

                If offsetMM < 0 Then
                    MessageBox.Show(
                        "Khoảng cách phải lớn hơn hoặc bằng 0.",
                        "Lỗi",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning)

                    Exit Sub
                End If

                ' Inventor dùng cm
                Dim offsetCM As Double = offsetMM / 10.0

                '=====================================================
                ' 6. TÍNH VỊ TRÍ MONG MUỐN
                '=====================================================

                Dim posX As Double =
                    oView.Position.X +
                    oView.Width / 2.0 +
                    offsetCM

                Dim posY As Double =
                    oView.Position.Y +
                    oView.Height / 2.0

                '=====================================================
                ' 7. KIỂM TRA SHEET ĐANG ACTIVE
                '=====================================================
                Try
                    oSheet.Activate()
                    If oSheet.PartsLists.Count > 0 Then

                        MessageBox.Show(
                            "Sheet này đã có Parts List." &
                            vbCrLf & vbCrLf &
                            "Không tạo thêm Parts List.",
                            "Tạo Parts List",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning)

                        Exit Sub

                    End If
                Catch
                End Try

                '=====================================================
                ' 8. TẠO PARTS LIST TẠI VỊ TRÍ AN TOÀN
                '
                ' Không tạo trực tiếp tại mép view vì có thể
                ' gây E_INVALIDARG nếu điểm nằm ngoài Sheet.
                '=====================================================

                Dim safeX As Double = oSheet.Width / 2.0
                Dim safeY As Double = oSheet.Height / 2.0

                Dim safePos As Point2d =
                    tg.CreatePoint2d(safeX, safeY)

                Dim oPL As PartsList = Nothing

                Try

                    oPL = oSheet.PartsLists.Add(
                        oView,
                        safePos)

                Catch ex As Exception

                    MessageBox.Show(
                        "Không tạo được Parts List:" &
                        vbCrLf & vbCrLf &
                        ex.Message &
                        vbCrLf & vbCrLf &
                        "View: " & oView.Name &
                        vbCrLf &
                        "Sheet: " & oSheet.Name,
                        "Lỗi",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error)

                    Exit Sub

                End Try

                If oPL Is Nothing Then

                    MessageBox.Show(
                        "Parts List trả về Nothing.",
                        "Lỗi",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error)

                    Exit Sub

                End If

                '=====================================================
                ' 9. GÁN STYLE
                '=====================================================
                Try

                    oPL.Style =
                        oDrawDoc.StylesManager.PartsListStyles.Item(styleName)

                Catch ex As Exception

                    MessageBox.Show(
                        "Không thể áp dụng Style:" &
                        vbCrLf & vbCrLf &
                        styleName &
                        vbCrLf & vbCrLf &
                        ex.Message,
                        "Cảnh báo",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning)

                End Try

                '=====================================================
                ' 10. SORT
                '=====================================================
                Try
                    oPL.Sort("ITEM")
                Catch
                    Try
                        oPL.Sort("Item")
                    Catch
                    End Try
                End Try

                '=====================================================
                ' 11. UPDATE
                '=====================================================
                Try
                    oPL.Update()
                Catch
                End Try

                Try
                    oDrawDoc.Update()
                Catch
                End Try

                '=====================================================
                ' 12. SAU KHI TẠO XONG MỚI DI CHUYỂN
                '=====================================================


                Dim finalPos As Point2d =
                        tg.CreatePoint2d(posX, posY)

                    oPL.Position = finalPos

                Catch
                    ' Nếu vị trí mong muốn không hợp lệ
                    ' thì giữ Parts List tại vị trí an toàn.
                End Try

        End Sub


        '=============================================================
        ' CHỌN STYLE PARTS LIST
        '
        ' Trả về:
        '   Tên Style = OK
        '   ""        = Hủy
        '=============================================================
        Private Function PickPartsListStyle(
            ByVal oDrawDoc As DrawingDocument) As String

            Try
                Dim styles As Inventor.PartsListStylesEnumerator =
                    oDrawDoc.StylesManager.PartsListStyles

                If styles Is Nothing OrElse styles.Count = 0 Then
                    Return ""
                End If

                '=====================================================
                ' LẤY DANH SÁCH STYLE
                '=====================================================
                Dim names As New List(Of String)

                For i As Integer = 1 To styles.Count

                    Try
                        Dim style As Inventor.PartsListStyle =
                            styles.Item(i)

                        If style IsNot Nothing Then
                            names.Add(style.Name)
                        End If

                    Catch
                    End Try

                Next

                If names.Count = 0 Then Return ""

                ' Chỉ có 1 Style → dùng luôn
                If names.Count = 1 Then
                    Return names(0)
                End If

                '=====================================================
                ' TẠO FORM
                '=====================================================
                Dim frm As New Form

                frm.Text = "Chọn Style Parts List"
                frm.ClientSize =
                    New System.Drawing.Size(420, 320)

                frm.StartPosition =
                    FormStartPosition.CenterScreen

                frm.FormBorderStyle =
                    FormBorderStyle.FixedDialog

                frm.MaximizeBox = False
                frm.MinimizeBox = False
                frm.ShowInTaskbar = False

                '=====================================================
                ' LIST
                '=====================================================
                Dim lst As New ListBox

                lst.Bounds =
                    New System.Drawing.Rectangle(
                        12,
                        12,
                        396,
                        250)

                lst.Font =
                    New System.Drawing.Font(
                        "Segoe UI",
                        9.0F)

                For Each n As String In names
                    lst.Items.Add(n)
                Next

                If lst.Items.Count > 0 Then
                    lst.SelectedIndex = 0
                End If

                '=====================================================
                ' BUTTON OK
                '=====================================================
                Dim btnOK As New Button

                btnOK.Text = "OK"

                btnOK.Bounds =
                    New System.Drawing.Rectangle(
                        230,
                        275,
                        85,
                        28)

                btnOK.DialogResult =
                    DialogResult.OK

                '=====================================================
                ' BUTTON HỦY
                '=====================================================
                Dim btnCancel As New Button

                btnCancel.Text = "Hủy"

                btnCancel.Bounds =
                    New System.Drawing.Rectangle(
                        323,
                        275,
                        85,
                        28)

                btnCancel.DialogResult =
                    DialogResult.Cancel

                '=====================================================
                ' ENTER / ESC
                '=====================================================
                frm.AcceptButton = btnOK
                frm.CancelButton = btnCancel

                '=====================================================
                ' DOUBLE CLICK = OK
                '=====================================================
                AddHandler lst.DoubleClick,
                    Sub(sender As Object, e As EventArgs)

                        If lst.SelectedItem IsNot Nothing Then
                            frm.DialogResult = DialogResult.OK
                            frm.Close()
                        End If

                    End Sub

                '=====================================================
                ' ADD CONTROL
                '=====================================================
                frm.Controls.Add(lst)
                frm.Controls.Add(btnOK)
                frm.Controls.Add(btnCancel)

                '=====================================================
                ' HIỂN THỊ MODAL
                '=====================================================
                Dim result As DialogResult =
                    frm.ShowDialog()

                '=====================================================
                ' KIỂM TRA KẾT QUẢ
                '=====================================================
                If result <> DialogResult.OK Then
                    Return ""
                End If

                If lst.SelectedItem Is Nothing Then
                    Return ""
                End If

                Return lst.SelectedItem.ToString()

            Catch ex As Exception

                MessageBox.Show(
                    "Lỗi khi chọn Style:" &
                    vbCrLf & vbCrLf &
                    ex.Message,
                    "Tạo Parts List",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning)

                Return ""

            End Try

        End Function

    End Module

End Namespace