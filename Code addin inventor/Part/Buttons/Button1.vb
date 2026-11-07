Imports System.Collections.Generic
Imports System.Windows.Forms
Imports Inventor

Namespace ToolInventor2020.Part.Buttons
    Public Module Button1

        Public Sub OnExecute(ByVal Context As NameValueMap)

            Try
                Dim invApp As Inventor.Application = g_inventorApplication

                '=====================================================
                ' KIỂM TRA DOCUMENT
                '=====================================================
                If invApp.ActiveDocumentType <> DocumentTypeEnum.kPartDocumentObject Then
                    MessageBox.Show("Vui lòng mở file Part (.ipt) trước!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Exit Sub
                End If

                Dim oPartDoc As PartDocument = CType(invApp.ActiveDocument, PartDocument)
                Dim oSelectSet As SelectSet = oPartDoc.SelectSet

                '=====================================================
                ' KIỂM TRA CÓ OBJECT NÀO ĐƯỢC CHỌN KHÔNG
                '=====================================================
                If oSelectSet.Count = 0 Then
                    MessageBox.Show("Hãy chọn solid body cần xóa trước!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Exit Sub
                End If

                '=====================================================
                ' LẤY CÁC SOLID BODY ĐANG ĐƯỢC CHỌN
                '=====================================================
                Dim solidsToDelete As New List(Of SurfaceBody)

                For Each obj As Object In oSelectSet

                    ' Kiểm tra object có phải SurfaceBody không
                    If TypeOf obj Is SurfaceBody Then
                        Dim oBody As SurfaceBody = CType(obj, SurfaceBody)

                        ' Chỉ thêm solid body (không lấy surface)
                        If oBody.IsSolid Then
                            solidsToDelete.Add(oBody)
                        End If
                    End If

                Next

                '=====================================================
                ' KIỂM TRA CÓ SOLID HỢP LỆ KHÔNG
                '=====================================================
                If solidsToDelete.Count = 0 Then
                    MessageBox.Show("Object đang chọn không phải Solid Body!" & vbCrLf &
                                    "Hãy chọn solid body trong Browser hoặc trên màn hình.",
                                    "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Exit Sub
                End If

                '=====================================================
                ' XÁC NHẬN TRƯỚC KHI XÓA
                '=====================================================
                Dim msg As String = "Bạn có chắc muốn xóa " & solidsToDelete.Count.ToString() & " solid body đang chọn?" & vbCrLf & vbCrLf

                For Each b As SurfaceBody In solidsToDelete
                    msg &= "• " & b.Name & vbCrLf
                Next

                Dim result As DialogResult = MessageBox.Show(msg, "Xác nhận xóa Solid",
                                                             MessageBoxButtons.YesNo,
                                                             MessageBoxIcon.Warning)

                If result <> DialogResult.Yes Then
                    Exit Sub
                End If

                '=====================================================
                ' THỰC HIỆN XÓA
                '=====================================================
                Dim deletedCount As Integer = 0

                For Each oBody As SurfaceBody In solidsToDelete
                    Try
                        oBody.Delete()
                        deletedCount += 1
                    Catch ex As Exception
                        MessageBox.Show("Không thể xóa solid: " & oBody.Name & vbCrLf & ex.Message,
                                        "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End Try
                Next

                '=====================================================
                ' CẬP NHẬT DOCUMENT
                '=====================================================
                oPartDoc.Update()
                oPartDoc.Save()

                invApp.UserInterfaceManager.UserInteractionManager.PostStatus(
                    "Đã xóa thành công " & deletedCount.ToString() & " solid body.")

            Catch ex As Exception
                Try
                    g_inventorApplication.UserInterfaceManager.UserInteractionManager.PostStatus(
                        "Error: " & ex.Message)
                Catch
                End Try

                MessageBox.Show("Lỗi: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try

        End Sub

    End Module
End Namespace