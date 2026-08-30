
Imports System
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Microsoft.VisualBasic ' For Interaction.InputBox
Imports Inventor
Imports System.Collections.Generic


Namespace ThanhN.Drawing.Buttons
    Public Module Draw_6
        Public Sub OnExecute(ByVal Context As NameValueMap)



            Try
                    ' Kết nối tới Inventor (lấy instance đang chạy)
                    Dim invApp As Inventor.Application = Nothing
                    Try
                        invApp = CType(Marshal.GetActiveObject("Inventor.Application"), Inventor.Application)
                    Catch ex As COMException
                        MessageBox.Show("Không tìm thấy Inventor đang chạy. Vui lòng mở Inventor và một Drawing trước khi chạy.", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Return
                    End Try

                    ' Lấy document hiện hành và kiểm tra là Drawing
                    Dim doc As Document = invApp.ActiveDocument
                    If doc Is Nothing OrElse doc.DocumentType <> DocumentTypeEnum.kDrawingDocumentObject Then
                        MessageBox.Show("Vui lòng mở một Drawing document trong Inventor trước khi chạy.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        Return
                    End If

                    Dim oDrawDoc As DrawingDocument = CType(doc, DrawingDocument)
                    Dim oSheet As Sheet = oDrawDoc.ActiveSheet
                    If oSheet Is Nothing Then
                        MessageBox.Show("Không thể truy cập ActiveSheet.", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Return
                    End If

                    If oSheet.PartsLists.Count < 1 Then
                        MessageBox.Show("Không tìm thấy PartsList trên sheet hiện hành.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        Return
                    End If

                    Dim oPartList As PartsList = oSheet.PartsLists.Item(1)

                    ' Nhập vật liệu (InputBox)
                    Dim aff As String = Interaction.InputBox("Vật Liệu", "Tên", "SS400")

                ' Hiển thị dialog chọn (mô phỏng InputListBox)
                Dim names As New List(Of String) From {"Ten chi tiet - Don vi", "Ten goi - Don vi"}

                Dim myparam As String = ShowSelectionDialog("Vị trí cột STT", "Lựa Chọn", names)
                    If String.IsNullOrEmpty(myparam) Then
                        MessageBox.Show("Bạn đã hủy lựa chọn.", "Hủy", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        Return
                    End If

                    ' Bắt transaction (undo wrapper) nếu có
                    Dim txn As Transaction = Nothing
                    Dim txnName As String = "Update PartsList Cells"
                    Try
                        txn = invApp.TransactionManager.StartTransaction(oDrawDoc, txnName)
                    Catch
                        txn = Nothing
                    End Try

                    ' Các biến dùng trong vòng lặp
                    Dim assssd As String = ""
                    Dim assssd2 As String = ""
                    Dim assssd3 As String = ""

                    If myparam = "Ten chi tiet - Don vi" Then
                        assssd = "Đơn vị"
                        assssd2 = "Vật liệu"
                        assssd3 = "Tên chi tiết"

                        ' Xử lý assembly rows: đặt Đơn vị = "Bộ", xóa vật liệu
                        For i As Integer = 1 To oPartList.PartsListRows.Count
                            Try
                                Dim refRow = oPartList.PartsListRows(i).ReferencedRows(1)
                                Dim bomRow = refRow.BOMRow
                                Dim compDef = bomRow.ComponentDefinitions(1)
                                If bomRow.BOMStructure <> BOMStructureEnum.kPurchasedBOMStructure AndAlso compDef.Document.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                                    Dim oCell2 = oPartList.PartsListRows.Item(i).Item(assssd)
                                    Dim oCell3 = oPartList.PartsListRows.Item(i).Item(assssd2)
                                    Try
                                        oCell2.Static = True
                                        oCell2.Value = "Bộ"
                                    Catch
                                    End Try
                                    Try
                                        oCell3.Static = True
                                        If oCell3.Value <> "" Then oCell3.Value = ""
                                    Catch
                                    End Try
                                End If
                            Catch
                                ' Bỏ qua nếu row không có ReferencedRows hoặc cấu trúc khác
                            End Try
                        Next

                        ' Index part (part documents)
                        For j As Integer = 1 To oPartList.PartsListRows.Count
                            Try
                                Dim refRow = oPartList.PartsListRows(j).ReferencedRows(1)
                                Dim bomRow = refRow.BOMRow
                                Dim compDef = bomRow.ComponentDefinitions(1)
                                If bomRow.BOMStructure <> BOMStructureEnum.kPurchasedBOMStructure AndAlso compDef.Document.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                                    Dim oCell2 = oPartList.PartsListRows.Item(j).Item(assssd)
                                    Dim oCell3 = oPartList.PartsListRows.Item(j).Item(assssd2)
                                    Dim oCell5 = oPartList.PartsListRows.Item(j).Item(assssd3)
                                    Try
                                        oCell2.Value = "Cái"
                                        oCell2.Static = True
                                    Catch
                                    End Try

                                    Dim val5 As String = ""
                                    Try
                                        val5 = If(oCell5.Value IsNot Nothing, oCell5.Value.ToString(), "")
                                    Catch
                                        val5 = ""
                                    End Try

                                    ' Các điều kiện chuỗi giống iLogic
                                    If (val5.Length >= 2 AndAlso (val5.Substring(0, 2).ToUpper() = "TH" OrElse val5.Substring(0, 2).ToUpper() = "TR" OrElse val5.Substring(0, 2).ToUpper() = "TH" OrElse val5.Substring(0, 2).ToUpper() = "XG" OrElse val5.Substring(0, 2).ToUpper() = "TR")) AndAlso val5.EndsWith("L", StringComparison.OrdinalIgnoreCase) Then
                                        Try : oCell2.Value = "Thanh" : Catch : End Try
                                    End If

                                    If val5.Length >= 1 AndAlso ("TPVLIHZCU".IndexOf(val5.Substring(0, 1).ToUpper()) >= 0) AndAlso val5.EndsWith("L", StringComparison.OrdinalIgnoreCase) Then
                                        Try : oCell2.Value = "Thanh" : Catch : End Try
                                    ElseIf val5.Length >= 2 AndAlso (val5.Substring(0, 2) = "PL" OrElse val5.Substring(0, 2) = "Tô" OrElse val5.Substring(0, 2) = "Tấ" OrElse val5.Substring(0, 2) = "Mã" OrElse val5.Substring(0, 2) = "Bi") Then
                                        Try : oCell2.Value = "Tấm" : Catch : End Try
                                    End If

                                    Try
                                        oCell3.Static = True
                                        oCell3.Value = aff
                                    Catch
                                    End Try
                                End If
                            Catch
                                ' Bỏ qua
                            End Try
                        Next

                        ' Index standard part (purchased)
                        For k As Integer = 1 To oPartList.PartsListRows.Count
                            Try
                                Dim refRow = oPartList.PartsListRows(k).ReferencedRows(1)
                                Dim bomRow = refRow.BOMRow
                                If bomRow.BOMStructure = BOMStructureEnum.kPurchasedBOMStructure Then
                                    Dim oCell2 = oPartList.PartsListRows.Item(k).Item(assssd)
                                    Dim oCell3 = oPartList.PartsListRows.Item(k).Item(assssd2)
                                    Try
                                        oCell2.Static = True
                                        oCell3.Static = True
                                        oCell2.Value = "Cái"
                                        oCell3.Value = ""
                                    Catch
                                    End Try
                                End If
                            Catch
                            End Try
                        Next

                    ElseIf myparam = "Ten goi - Don vi" Then
                        assssd = "Đơn vị"
                        assssd2 = "Vật liệu"
                        assssd3 = "Tên gọi"

                        ' Xử lý assembly rows
                        For i As Integer = 1 To oPartList.PartsListRows.Count
                            Try
                                Dim refRow = oPartList.PartsListRows(i).ReferencedRows(1)
                                Dim bomRow = refRow.BOMRow
                                Dim compDef = bomRow.ComponentDefinitions(1)
                                If bomRow.BOMStructure <> BOMStructureEnum.kPurchasedBOMStructure AndAlso compDef.Document.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                                    Dim oCell2 = oPartList.PartsListRows.Item(i).Item(assssd)
                                    Dim oCell3 = oPartList.PartsListRows.Item(i).Item(assssd2)
                                    Try
                                        oCell2.Static = True
                                        oCell2.Value = "Bộ"
                                    Catch
                                    End Try
                                    Try
                                        oCell3.Static = True
                                        If oCell3.Value <> "" Then oCell3.Value = ""
                                    Catch
                                    End Try
                                End If
                            Catch
                            End Try
                        Next

                        ' Index part
                        For j As Integer = 1 To oPartList.PartsListRows.Count
                            Try
                                Dim refRow = oPartList.PartsListRows(j).ReferencedRows(1)
                                Dim bomRow = refRow.BOMRow
                                Dim compDef = bomRow.ComponentDefinitions(1)
                                If bomRow.BOMStructure <> BOMStructureEnum.kPurchasedBOMStructure AndAlso compDef.Document.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                                    Dim oCell2 = oPartList.PartsListRows.Item(j).Item(assssd)
                                    Dim oCell3 = oPartList.PartsListRows.Item(j).Item(assssd2)
                                    Dim oCell5 = oPartList.PartsListRows.Item(j).Item(assssd3)
                                    Try
                                        oCell2.Value = "Cái"
                                        oCell2.Static = True
                                    Catch
                                    End Try

                                    Dim val5 As String = ""
                                    Try
                                        val5 = If(oCell5.Value IsNot Nothing, oCell5.Value.ToString(), "")
                                    Catch
                                        val5 = ""
                                    End Try

                                    If (val5.Length >= 2 AndAlso (val5.Substring(0, 2).ToUpper() = "TH" OrElse val5.Substring(0, 2).ToUpper() = "TR" OrElse val5.Substring(0, 2).ToUpper() = "TH" OrElse val5.Substring(0, 2).ToUpper() = "XG" OrElse val5.Substring(0, 2).ToUpper() = "TR")) AndAlso val5.EndsWith("L", StringComparison.OrdinalIgnoreCase) Then
                                        Try : oCell2.Value = "Thanh" : Catch : End Try
                                    End If

                                    If val5.Length >= 1 AndAlso ("TPVLIHZCU".IndexOf(val5.Substring(0, 1).ToUpper()) >= 0) AndAlso val5.EndsWith("L", StringComparison.OrdinalIgnoreCase) Then
                                        Try : oCell2.Value = "Thanh" : Catch : End Try
                                    ElseIf val5.Length >= 2 AndAlso (val5.Substring(0, 2) = "PL" OrElse val5.Substring(0, 2) = "Tô" OrElse val5.Substring(0, 2) = "Tấ" OrElse val5.Substring(0, 2) = "Mã" OrElse val5.Substring(0, 2) = "Bi") Then
                                        Try : oCell2.Value = "Tấm" : Catch : End Try
                                    End If

                                    Try
                                        oCell3.Static = True
                                        oCell3.Value = aff
                                    Catch
                                    End Try
                                End If
                            Catch
                            End Try
                        Next

                        ' Index standard part
                        For k As Integer = 1 To oPartList.PartsListRows.Count
                            Try
                                Dim refRow = oPartList.PartsListRows(k).ReferencedRows(1)
                                Dim bomRow = refRow.BOMRow
                                If bomRow.BOMStructure = BOMStructureEnum.kPurchasedBOMStructure Then
                                    Dim oCell2 = oPartList.PartsListRows.Item(k).Item(assssd)
                                    Dim oCell3 = oPartList.PartsListRows.Item(k).Item(assssd2)
                                    Try
                                        oCell2.Static = True
                                        oCell3.Static = True
                                        oCell2.Value = "Cái"
                                        oCell3.Value = ""
                                    Catch
                                    End Try
                                End If
                            Catch
                            End Try
                        Next

                    End If

                    ' Cập nhật document
                    Try
                        oDrawDoc.Update()
                    Catch
                        Try
                            oDrawDoc.Update2()
                        Catch
                        End Try
                    End Try

                    ' Kết thúc transaction
                    Try
                        If txn IsNot Nothing Then txn.End()
                    Catch
                    End Try

                    MessageBox.Show("Hoàn tất xử lý PartsList.", "Kết quả", MessageBoxButtons.OK, MessageBoxIcon.Information)

                Catch ex As Exception
                    MessageBox.Show("Lỗi: " & ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End Sub

            ' Helper: hiển thị dialog chọn 1 mục từ danh sách (mô phỏng InputListBox)
            Private Function ShowSelectionDialog(title As String, prompt As String, items As List(Of String)) As String
                Dim result As String = ""
                Using frm As New Form()
                    frm.Text = title
                    frm.StartPosition = FormStartPosition.CenterScreen
                    frm.FormBorderStyle = FormBorderStyle.FixedDialog
                    frm.MinimizeBox = False
                    frm.MaximizeBox = False
                    frm.Width = 400
                    frm.Height = 220

                    Dim lbl As New Label()
                    lbl.Text = prompt
                    lbl.AutoSize = True
                    lbl.Top = 10
                    lbl.Left = 10
                    frm.Controls.Add(lbl)

                    Dim lb As New ListBox()
                    lb.Top = 35
                    lb.Left = 10
                    lb.Width = frm.ClientSize.Width - 20
                    lb.Height = 100
                    lb.SelectionMode = SelectionMode.One
                    For Each it In items
                        lb.Items.Add(it)
                    Next
                    If lb.Items.Count > 0 Then lb.SelectedIndex = 0
                    frm.Controls.Add(lb)

                    Dim btnOk As New Button()
                    btnOk.Text = "OK"
                    btnOk.DialogResult = DialogResult.OK
                    btnOk.Top = lb.Bottom + 10
                    btnOk.Left = frm.ClientSize.Width - 180
                    frm.Controls.Add(btnOk)

                    Dim btnCancel As New Button()
                    btnCancel.Text = "Cancel"
                    btnCancel.DialogResult = DialogResult.Cancel
                    btnCancel.Top = lb.Bottom + 10
                    btnCancel.Left = frm.ClientSize.Width - 90
                    frm.Controls.Add(btnCancel)

                    frm.AcceptButton = btnOk
                    frm.CancelButton = btnCancel

                    If frm.ShowDialog() = DialogResult.OK Then
                        If lb.SelectedItem IsNot Nothing Then
                            result = lb.SelectedItem.ToString()
                        End If
                    Else
                        result = ""
                    End If
                End Using
                Return result
            End Function

        End Module

End Namespace
