Imports System
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Inventor


Namespace ToolInventor2020.Drawing.Buttons
    Public Module Draw_3
        Public Sub OnExecute(ByVal Context As NameValueMap)



            Try
                ' 1. Kết nối tới Inventor (lấy instance đang chạy hoặc khởi tạo mới)
                Dim invApp As Inventor.Application = Nothing
                Try
                    invApp = CType(Marshal.GetActiveObject("Inventor.Application"), Inventor.Application)
                Catch ex As COMException
                    Dim invType As Type = Type.GetTypeFromProgID("Inventor.Application")
                    If invType Is Nothing Then
                        MessageBox.Show("Không tìm thấy Autodesk Inventor trên hệ thống.", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Return
                    End If
                    invApp = CType(Activator.CreateInstance(invType), Inventor.Application)
                    invApp.Visible = True
                End Try

                ' 2. Kiểm tra document hiện hành là Drawing
                Dim doc As Document = invApp.ActiveDocument
                If doc Is Nothing OrElse doc.DocumentType <> DocumentTypeEnum.kDrawingDocumentObject Then
                    MessageBox.Show("Vui lòng mở một Drawing document trong Inventor trước khi chạy.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Return
                End If

                Dim drawDoc As DrawingDocument = CType(doc, DrawingDocument)
                Dim activeSheet As Sheet = drawDoc.ActiveSheet
                If activeSheet Is Nothing Then
                    MessageBox.Show("Không thể truy cập ActiveSheet.", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End If

                ' 3. Center text cho LinearGeneralDimension và AngularGeneralDimension
                Dim centeredCount As Integer = 0
                Try
                    For Each dd As DrawingDimension In activeSheet.DrawingDimensions
                        Try
                            If TypeOf dd Is LinearGeneralDimension OrElse TypeOf dd Is AngularGeneralDimension Then
                                dd.CenterText()
                                centeredCount += 1
                            End If
                        Catch exDim As Exception
                            ' Bỏ qua dimension gây lỗi, tiếp tục
                            Console.WriteLine("Không center được 1 dimension: " & exDim.Message)
                        End Try
                    Next
                Catch exIter As Exception
                    MessageBox.Show("Lỗi khi duyệt DrawingDimensions: " & exIter.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End Try

                ' 4. Chọn tất cả dimension trên sheet và gọi lệnh Arrange Dimensions
                Try
                    Dim sel As SelectSet = drawDoc.SelectSet
                    sel.Clear()

                    For Each dd As DrawingDimension In activeSheet.DrawingDimensions
                        Try
                            sel.Select(dd)
                        Catch
                            ' Nếu không thể chọn 1 đối tượng, bỏ qua
                        End Try
                    Next

                    Try
                        invApp.CommandManager.ControlDefinitions.Item("DrawingArrangeDimensionsCmd").Execute()
                    Catch exCmd As Exception
                        ' Nếu lệnh không thực thi được, thông báo nhưng không dừng chương trình
                        MessageBox.Show("Không thể thực thi lệnh Arrange Dimensions: " & exCmd.Message, "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    End Try
                Catch exSelect As Exception
                    MessageBox.Show("Lỗi khi chọn dimensions: " & exSelect.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try

                ' 5. Lưu document (tùy chọn)
                Try
                    drawDoc.Save2(True)
                Catch
                    ' Bỏ qua nếu không muốn lưu tự động
                End Try

                MessageBox.Show($"Hoàn tất. Centered: {centeredCount} dimensions.", "Kết quả", MessageBoxButtons.OK, MessageBoxIcon.Information)

            Catch ex As Exception
                MessageBox.Show("Lỗi không mong muốn: " & ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub
    End Module

End Namespace