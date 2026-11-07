Imports Inventor
Imports System.Windows.Forms

Namespace ToolInventor2020.Part.Buttons.Tsketch
    Public Module part_Sketch_4

        Public Sub OnExecute(ByVal Context As NameValueMap)

            Dim invApp As Inventor.Application = g_inventorApplication

            Try
                Dim doc As Document = invApp.ActiveDocument
                If doc Is Nothing Then
                    MessageBox.Show("Không có tài liệu nào đang mở!", "Thông báo")
                    Exit Sub
                End If

                '=====================================================
                ' KIỂM TRA LOẠI TÀI LIỆU
                '=====================================================
                Dim docTypeName As String = ""
                Select Case doc.DocumentType
                    Case DocumentTypeEnum.kPartDocumentObject
                        docTypeName = "Part"
                    Case DocumentTypeEnum.kAssemblyDocumentObject
                        docTypeName = "Assembly"
                    Case DocumentTypeEnum.kDrawingDocumentObject
                        docTypeName = "Drawing"
                    Case Else
                        MessageBox.Show("Loại tài liệu không được hỗ trợ!" & vbCrLf &
                                        "Chỉ hỗ trợ Part (.ipt), Assembly (.iam), Drawing (.idw/.dwg).",
                                        "Thông báo")
                        Exit Sub
                End Select

                '=====================================================
                ' LẤY ĐỐI TƯỢNG ĐANG ĐƯỢC EDIT (CÁCH DUY NHẤT ĐÁNG TIN)
                '=====================================================
                Dim editObj As Object = Nothing
                Try
                    editObj = invApp.ActiveEditObject
                Catch
                    editObj = Nothing
                End Try

                If editObj Is Nothing Then
                    MessageBox.Show("Không có đối tượng nào đang được chỉnh sửa trong " & docTypeName & "." & vbCrLf &
                                    "Hãy double-click vào Sketch để vào chế độ edit trước!",
                                    "Thông báo")
                    Exit Sub
                End If

                '=====================================================
                ' XÁC MINH ĐỐI TƯỢNG LÀ SKETCH (Part/Assembly/Drawing)
                '=====================================================
                Dim oSketch As Sketch = Nothing

                If TypeOf editObj Is DrawingSketch Then
                    ' Drawing: DrawingSketch kế thừa từ Sketch
                    oSketch = CType(editObj, DrawingSketch)
                ElseIf TypeOf editObj Is Sketch Then
                    ' Part / Assembly
                    oSketch = CType(editObj, Sketch)
                Else
                    MessageBox.Show("Đối tượng đang mở không phải là Sketch." & vbCrLf &
                                    "Loại hiện tại: " & editObj.[GetType]().Name,
                                    "Thông báo")
                    Exit Sub
                End If

                '=====================================================
                ' XÁC MINH SKETCH ĐANG THỰC SỰ TRONG CHẾ ĐỘ EDIT
                ' (Deactivated = False nghĩa là đang edit)
                '=====================================================
                Try
                    If oSketch.Deactivated Then
                        MessageBox.Show("Sketch '" & oSketch.Name & "' không ở trong chế độ chỉnh sửa." & vbCrLf &
                                        "Hãy double-click vào Sketch để vào chế độ edit trước!",
                                        "Thông báo")
                        Exit Sub
                    End If
                Catch
                    ' Nếu không truy cập được thuộc tính Deactivated, bỏ qua bước này
                    ' (vẫn tiếp tục vì đã có ActiveEditObject xác nhận)
                End Try

                '=====================================================
                ' KIỂM TRA DIMENSION
                '=====================================================
                Dim oDims As DimensionConstraints = oSketch.DimensionConstraints
                Dim dimCount As Integer = oDims.Count

                If dimCount = 0 Then
                    MessageBox.Show("Sketch '" & oSketch.Name & "' không có dimension nào để xóa!",
                                    "Thông báo")
                    Exit Sub
                End If

                '=====================================================
                ' XÁC NHẬN (ghi rõ tên sketch để tránh nhầm lẫn)
                '=====================================================
                Dim msg As String = "Tài liệu: " & docTypeName & vbCrLf &
                                    "Sketch ĐANG MỞ: '" & oSketch.Name & "'" & vbCrLf &
                                    "Số dimension sẽ xóa: " & dimCount.ToString() & vbCrLf & vbCrLf &
                                    "CHỈ sketch '" & oSketch.Name & "' bị ảnh hưởng. Tiếp tục?"
                If MessageBox.Show(msg, "Xác nhận xóa", MessageBoxButtons.YesNo,
                                   MessageBoxIcon.Warning) <> DialogResult.Yes Then
                    Exit Sub
                End If

                '=====================================================
                ' XÓA NGƯỢC TỪ CUỐI
                '=====================================================
                Dim deletedCount As Integer = 0

                For i As Integer = dimCount To 1 Step -1
                    Try
                        oDims.Item(i).Delete()
                        deletedCount += 1
                    Catch
                        ' Bỏ qua dimension không xóa được (driven/projected/reference)
                    End Try
                Next

                '=====================================================
                ' CẬP NHẬT
                '=====================================================
                doc.Update2(True)

                invApp.StatusBarText = "Đã xóa " & deletedCount.ToString() & " / " &
                                       dimCount.ToString() & " dimension trong sketch '" &
                                       oSketch.Name & "' (" & docTypeName & ")."

            Catch ex As Exception
                invApp.StatusBarText = "Error: " & ex.Message
                MessageBox.Show("Lỗi: " & ex.Message, "Error")
            End Try

        End Sub

    End Module
End Namespace