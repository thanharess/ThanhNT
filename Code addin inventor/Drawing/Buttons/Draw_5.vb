Imports System
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Inventor


Namespace ThanhN.Drawing.Buttons
    Public Module Draw_5
        Public Sub OnExecute(ByVal Context As NameValueMap)

            Try
                ' 1. Kết nối tới Inventor (lấy instance đang chạy)
                Dim invApp As Inventor.Application = Nothing
                Try
                    invApp = CType(Marshal.GetActiveObject("Inventor.Application"), Inventor.Application)
                Catch ex As COMException
                    MessageBox.Show("Không tìm thấy Inventor đang chạy. Vui lòng mở Inventor và một Drawing trước khi chạy.", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End Try

                ' 2. Lấy document hiện hành và kiểm tra là Drawing
                Dim doc As Document = invApp.ActiveDocument
                If doc Is Nothing OrElse doc.DocumentType <> DocumentTypeEnum.kDrawingDocumentObject Then
                    MessageBox.Show("Vui lòng mở một Drawing document trong Inventor trước khi chạy.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Return
                End If

                Dim oDrawDoc As DrawingDocument = CType(doc, DrawingDocument)

                ' 3. Lấy ActiveSheet
                Dim oSheet As Sheet = oDrawDoc.ActiveSheet
                If oSheet Is Nothing Then
                    MessageBox.Show("Không thể truy cập ActiveSheet.", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End If

                ' 4. Kiểm tra PartsLists tồn tại
                If oSheet.PartsLists.Count < 1 Then
                    MessageBox.Show("Không tìm thấy PartsList trên sheet hiện hành.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Return
                End If

                ' 5. Lấy PartsList đầu tiên
                Dim oPartsList As PartsList = oSheet.PartsLists(1)

                ' 6. Bắt transaction (undo wrapper) — tùy phiên bản Inventor, TransactionManager có thể khác
                Dim txn As Transaction = Nothing
                Dim txnName As String = "Set PartsList Cells Static = False"
                Try
                    txn = invApp.TransactionManager.StartTransaction(oDrawDoc, txnName)
                Catch
                    ' Nếu không thể bắt transaction, tiếp tục nhưng không có undo wrapper
                    txn = Nothing
                End Try

                ' 7. Duyệt từng row và column, đặt Static = False
                Dim rowCount As Integer = 0
                Try
                    For Each oRow As PartsListRow In oPartsList.PartsListRows
                        rowCount += 1
                        For Each oColumn As PartsListColumn In oPartsList.PartsListColumns
                            Try
                                ' Một số ô có thể không hỗ trợ gán, nên bọc try/catch
                                oRow.Item(oColumn).Static = False
                            Catch exCell As Exception
                                ' Bỏ qua ô gây lỗi, tiếp tục
                            End Try
                        Next
                    Next
                Catch exRows As Exception
                    MessageBox.Show("Lỗi khi duyệt PartsListRows/Columns: " & exRows.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    ' Kết thúc transaction nếu đang mở
                    Try
                        If txn IsNot Nothing Then txn.End()
                    Catch
                    End Try
                    Return
                End Try

                ' 8. Cập nhật document và kết thúc transaction
                Try
                    oDrawDoc.Update()
                Catch
                    Try
                        oDrawDoc.Update2()
                    Catch
                    End Try
                End Try

                Try
                    If txn IsNot Nothing Then txn.End()
                Catch
                End Try

                MessageBox.Show("Hoàn tất. Đã xử lý " & rowCount & " hàng trong PartsList.", "Kết quả", MessageBoxButtons.OK, MessageBoxIcon.Information)

            Catch ex As Exception
                MessageBox.Show("Lỗi không mong muốn: " & ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub
    End Module

End Namespace