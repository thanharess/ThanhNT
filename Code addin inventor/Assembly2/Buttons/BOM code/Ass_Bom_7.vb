Imports Inventor
Imports System.Windows.Forms
Imports System.Text.RegularExpressions
Imports System.Runtime.InteropServices

Namespace ToolInventor2020.Assembly2.Buttons.BOMcode

    Public Module ass_bom_7

        Public Sub OnExecute(ByVal Context As NameValueMap)
            ' để trống
        End Sub

        '=====================================================
        ' HÀM 1: SỬA PART NUMBER (có hộp thoại Find / Replace)
        '=====================================================
        Public Sub UpdatePartNumber_BOM()
            Dim findText As String = InputBox("Nhập chữ cần tìm:", "Part Number - Find", "mm")
            If String.IsNullOrEmpty(findText) Then Exit Sub

            Dim replaceText As String = InputBox("Nhập chữ cần thay thế:", "Part Number - Replace", "L")
            If replaceText Is Nothing Then Exit Sub   ' bấm Cancel

            RunOnStructureBOM(updatePartNumber:=True, updateStockNumber:=False, findText, replaceText)
        End Sub

        '=====================================================
        ' HÀM 2: SỬA STOCK NUMBER (có hộp thoại Find / Replace)
        '=====================================================
        Public Sub UpdateStockNumber_BOM()
            Dim findText As String = InputBox("Nhập chữ cần tìm:", "Stock Number - Find", "mm")
            If String.IsNullOrEmpty(findText) Then Exit Sub

            Dim replaceText As String = InputBox("Nhập chữ cần thay thế:", "Stock Number - Replace", "L")
            If replaceText Is Nothing Then Exit Sub

            RunOnStructureBOM(updatePartNumber:=False, updateStockNumber:=True, findText, replaceText)
        End Sub

        '=====================================================
        ' LẤY INVENTOR APPLICATION
        '=====================================================
        Private Function GetInventorApp() As Inventor.Application
            Try
                Return DirectCast(Marshal.GetActiveObject("Inventor.Application"), Inventor.Application)
            Catch
                MessageBox.Show("Không tìm thấy Inventor đang chạy!")
                Return Nothing
            End Try
        End Function

        '=====================================================
        ' HÀM CHÍNH – CHẠY TRÊN STRUCTURE BOM
        '=====================================================
        Private Sub RunOnStructureBOM(updatePartNumber As Boolean,
                                      updateStockNumber As Boolean,
                                      findText As String,
                                      replaceText As String)

            Dim invApp As Inventor.Application = GetInventorApp()
            If invApp Is Nothing Then Exit Sub

            If invApp.ActiveDocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
                MessageBox.Show("Hãy mở Assembly trước!")
                Exit Sub
            End If

            Dim asmDoc As AssemblyDocument = invApp.ActiveDocument
            Dim bom As BOM = asmDoc.ComponentDefinition.BOM

            bom.StructuredViewEnabled = True
            Dim structuredView As BOMView = bom.BOMViews.Item("Structured")

            Dim countChanged As Integer = 0

            For Each row As BOMRow In structuredView.BOMRows
                ProcessBOMRow(row, updatePartNumber, updateStockNumber, findText, replaceText, countChanged)
            Next

            asmDoc.Update2(True)
            asmDoc.Save2(True)

            MessageBox.Show("Hoàn tất!" & vbCrLf &
                            "Số lượng item đã sửa: " & countChanged.ToString() & vbCrLf &
                            "Tìm: """ & findText & """ → Thay bằng: """ & replaceText & """",
                            "BOM Update")
        End Sub

        '=====================================================
        ' XỬ LÝ TỪNG HÀNG BOM (ĐỆ QUY)
        '=====================================================
        Private Sub ProcessBOMRow(row As BOMRow,
                                  updatePartNumber As Boolean,
                                  updateStockNumber As Boolean,
                                  findText As String,
                                  replaceText As String,
                                  ByRef countChanged As Integer)

            Try
                Dim def As ComponentDefinition = row.ComponentDefinitions.Item(1)
                Dim doc As Document = def.Document

                Dim designProps As PropertySet = Nothing
                Try
                    designProps = doc.PropertySets.Item("Design Tracking Properties")
                Catch
                    Exit Sub
                End Try

                Dim changed As Boolean = False

                '----- Sửa Part Number -----
                If updatePartNumber Then
                    Try
                        Dim prop As Inventor.Property = designProps.Item("Part Number")
                        Dim oldVal As String = prop.Value.ToString()

                        If oldVal.IndexOf(findText, StringComparison.OrdinalIgnoreCase) >= 0 Then
                            Dim newVal As String = Regex.Replace(oldVal, Regex.Escape(findText), replaceText, RegexOptions.IgnoreCase)
                            prop.Value = newVal
                            changed = True
                        End If
                    Catch
                    End Try
                End If

                '----- Sửa Stock Number -----
                If updateStockNumber Then
                    Try
                        Dim prop As Inventor.Property = designProps.Item("Stock Number")
                        Dim oldVal As String = prop.Value.ToString()

                        If oldVal.IndexOf(findText, StringComparison.OrdinalIgnoreCase) >= 0 Then
                            Dim newVal As String = Regex.Replace(oldVal, Regex.Escape(findText), replaceText, RegexOptions.IgnoreCase)
                            prop.Value = newVal
                            changed = True
                        End If
                    Catch
                    End Try
                End If

                If changed Then
                    doc.Save2(True)
                    countChanged += 1
                End If

                ' Đệ quy vào các hàng con
                If row.ChildRows IsNot Nothing Then
                    For Each childRow As BOMRow In row.ChildRows
                        ProcessBOMRow(childRow, updatePartNumber, updateStockNumber, findText, replaceText, countChanged)
                    Next
                End If

            Catch
            End Try
        End Sub

    End Module

End Namespace