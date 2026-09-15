Imports Inventor
Imports System.Collections.Generic

Namespace ToolInventor2020.Part.Buttons
    Public Module Part_Solid_7

        ' ==========================================================
        ' CẤU HÌNH DANH SÁCH VẬT LIỆU & NGOẠI QUAN
        ' ==========================================================
        Private ReadOnly MaterialNames As String() = BuildMaterialNames()
        Private ReadOnly AppearanceNames As String() = BuildAppearanceNames()

        Private Function BuildMaterialNames() As String()
            Dim list As New List(Of String) From {"Steel, Mild"}
            For i As Integer = 1 To 23
                list.Add("Steel, Mild " & i)
            Next
            Return list.ToArray()
        End Function

        Private Function BuildAppearanceNames() As String()
            Dim list As New List(Of String) From {"Semi-Polished"}
            For i As Integer = 1 To 23
                list.Add("Semi-Polished(" & i & ")")
            Next
            Return list.ToArray()
        End Function

        ' ==========================================================
        ' HÀM CHÍNH - ĐƯỢC GỌI KHI BẤM NÚT
        ' ==========================================================
        Public Sub OnExecute(ByVal Context As NameValueMap)
            Try
                ' 1. Kiểm tra tài liệu hiện tại
                Dim oDoc As Document = g_inventorApplication.ActiveDocument
                If oDoc Is Nothing Then
                    PostStatus("⚠ Không có tài liệu nào đang mở.")
                    Return
                End If

                ' 2. Ép kiểu sang PartDocument
                Dim oPartDoc As PartDocument = TryCast(oDoc, PartDocument)
                If oPartDoc Is Nothing Then
                    PostStatus("⚠ Tài liệu đang mở không phải là Part (.ipt).")
                    Return
                End If

                ' 3. Chọn ngẫu nhiên Vật liệu & Ngoại quan
                Dim rand As New Random()
                Dim chosenMaterial As String = MaterialNames(rand.Next(0, MaterialNames.Length))
                Dim chosenAppearance As String = AppearanceNames(rand.Next(0, AppearanceNames.Length))

                ' 4. Tìm & gán Vật liệu
                Dim matAsset As Asset = FindAsset(oPartDoc.MaterialAssets, chosenMaterial)
                If matAsset IsNot Nothing Then
                    oPartDoc.ActiveMaterial = matAsset
                Else
                    PostStatus("⚠ Không tìm thấy vật liệu: " & chosenMaterial)
                End If

                ' 5. Tìm & gán Ngoại quan
                Dim appAsset As Asset = FindAsset(oPartDoc.AppearanceAssets, chosenAppearance)
                If appAsset IsNot Nothing Then
                    oPartDoc.ActiveAppearance = appAsset
                Else
                    PostStatus("⚠ Không tìm thấy ngoại quan: " & chosenAppearance)
                End If

                ' 6. Cập nhật tài liệu
                oPartDoc.Update()

                ' 7. Thông báo kết quả
                PostStatus($"✔ Đã đổi → Vật liệu: {chosenMaterial} | Ngoại quan: {chosenAppearance}")

            Catch ex As Exception
                Try
                    g_inventorApplication.UserInterfaceManager.UserInteractionManager.PostStatus("Error in Part Action 7: " & ex.Message)
                Catch
                End Try
            End Try
        End Sub

        ' ==========================================================
        ' HELPER: Tìm Asset theo DisplayName (không phân biệt hoa/thường)
        ' ==========================================================
        Private Function FindAsset(assets As Assets, displayName As String) As Asset
            If assets Is Nothing Then Return Nothing

            For Each asset As Asset In assets
                If String.Equals(asset.DisplayName, displayName, StringComparison.OrdinalIgnoreCase) Then
                    Return asset
                End If
            Next

            Return Nothing
        End Function

        ' ==========================================================
        ' HELPER: Gửi thông báo lên thanh trạng thái Inventor
        ' ==========================================================
        Private Sub PostStatus(msg As String)
            Try
                g_inventorApplication.UserInterfaceManager.UserInteractionManager.PostStatus(msg)
            Catch
                ' Bỏ qua nếu không post được
            End Try
        End Sub

    End Module
End Namespace