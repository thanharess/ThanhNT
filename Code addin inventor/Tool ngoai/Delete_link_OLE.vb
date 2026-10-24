Imports System.Windows.Forms
Imports Inventor

Namespace ToolInventor2020.Toolngoai
    Public Module delete_link_OLE

        Public Sub OnExecute(ByVal Context As NameValueMap)

            ' Lấy ứng dụng Inventor đang chạy
            Dim invApp As Inventor.Application = System.Runtime.InteropServices.Marshal.GetActiveObject("Inventor.Application")

            Dim oDoc As Document = invApp.ActiveDocument

            ' Kiểm tra tài liệu hiện tại
            If oDoc Is Nothing Then
                MessageBox.Show("Không có tài liệu đang mở!", "Visual Studio")
                Return
            End If

            If oDoc.DocumentType <> DocumentTypeEnum.kPartDocumentObject AndAlso
               oDoc.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
                MessageBox.Show("Chỉ hỗ trợ file Part (.ipt) hoặc Assembly (.iam)!", "Visual Studio")
                Return
            End If

            '=================================================
            ' KHAI BÁO BIẾN ĐẾM (đưa lên đầu để tránh lỗi ẩn biến)
            '=================================================
            Dim countExcel As Integer = 0
            Dim countDerived As Integer = 0
            Dim countOLE As Integer = 0
            Dim countDocs As Integer = 0
            Dim modeText As String = ""

            Dim processTopOnly As Boolean = False
            Dim processOnlyAssemblies As Boolean = False
            Dim processOnlyParts As Boolean = False
            Dim processAll As Boolean = False

            '=================================================
            ' NẾU LÀ PART → XỬ LÝ NGAY, KHÔNG HỎI
            '=================================================
            If oDoc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                invApp.SilentOperation = True

                Try
                    Call RemoveAllParameterLinks(oDoc, countExcel, countDerived, countOLE)
                    countDocs = 1
                    oDoc.Update2(True)
                    ' oDoc.Save2(True)

                    MessageBox.Show("Đã xử lý xong Part!" & vbCrLf & vbCrLf &
                                    "Excel Tables đã xóa: " & countExcel & vbCrLf &
                                    "Derived Parameters đã xóa: " & countDerived & vbCrLf &
                                    "OLE Links đã xóa: " & countOLE, "Visual Studio")
                Catch ex As Exception
                    MessageBox.Show("Lỗi: " & ex.Message, "Visual Studio")
                Finally
                    invApp.SilentOperation = False
                End Try
                Return
            End If

            '=================================================
            ' NẾU LÀ ASSEMBLY → HỎI CHỌN CHẾ ĐỘ
            '=================================================
            Dim result As DialogResult
            result = MessageBox.Show(
                "Chọn chế độ xử lý cho Assembly:" & vbCrLf & vbCrLf &
                "YES      = Chỉ Top Level (Assembly tổng)" & vbCrLf &
                "NO       = Xử lý đệ quy (Part + Assembly)" & vbCrLf &
                "CANCEL   = Chọn chế độ nâng cao...",
                "Chọn chế độ xóa liên kết Parameters",
                MessageBoxButtons.YesNoCancel,
                MessageBoxIcon.Question)

            If result = DialogResult.Yes Then
                processTopOnly = True
            ElseIf result = DialogResult.No Then
                processAll = True
            Else
                '----- Chế độ nâng cao -----
                Dim result2 As DialogResult
                result2 = MessageBox.Show(
                    "Chọn chế độ nâng cao:" & vbCrLf & vbCrLf &
                    "YES  = Chỉ áp dụng cho các Assembly (all level)" & vbCrLf &
                    "NO   = Chỉ áp dụng cho các Part (all level)" & vbCrLf &
                    "CANCEL = Hủy",
                    "Chế độ nâng cao",
                    MessageBoxButtons.YesNoCancel,
                    MessageBoxIcon.Question)

                If result2 = DialogResult.Yes Then
                    processOnlyAssemblies = True
                ElseIf result2 = DialogResult.No Then
                    processOnlyParts = True
                Else
                    Return
                End If
            End If

            invApp.SilentOperation = True

            Try
                Dim oAsmDoc As AssemblyDocument = CType(oDoc, AssemblyDocument)

                '=================================================
                ' 1. CHỈ TOP LEVEL
                '=================================================
                If processTopOnly Then
                    modeText = "Chỉ Top Level (Assembly tổng)"
                    Call RemoveAllParameterLinks(oAsmDoc, countExcel, countDerived, countOLE)
                    countDocs += 1
                    oAsmDoc.Update2(True)
                    ' oAsmDoc.Save2(True)

                    '=================================================
                    ' 2. ĐỆ QUY TẤT CẢ (Part + Assembly)
                    '=================================================
                ElseIf processAll Then
                    modeText = "Đệ quy toàn bộ (Part + Assembly)"

                    For Each oRefDoc As Document In oAsmDoc.AllReferencedDocuments
                        Try
                            If oRefDoc.DocumentType = DocumentTypeEnum.kPartDocumentObject OrElse
                               oRefDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then

                                Call RemoveAllParameterLinks(oRefDoc, countExcel, countDerived, countOLE)
                                countDocs += 1

                                Try
                                    oRefDoc.Update2(True)
                                    ' oRefDoc.Save2(True)
                                Catch
                                End Try
                            End If
                        Catch
                        End Try
                    Next

                    ' Xử lý lại Assembly tổng
                    Call RemoveAllParameterLinks(oAsmDoc, countExcel, countDerived, countOLE)
                    countDocs += 1
                    oAsmDoc.Update2(True)
                    ' oAsmDoc.Save2(True)

                    '=================================================
                    ' 3. CHỈ CÁC ASSEMBLY (ALL LEVEL)
                    '=================================================
                ElseIf processOnlyAssemblies Then
                    modeText = "Chỉ các Assembly (all level)"

                    For Each oRefDoc As Document In oAsmDoc.AllReferencedDocuments
                        Try
                            If oRefDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                                Call RemoveAllParameterLinks(oRefDoc, countExcel, countDerived, countOLE)
                                countDocs += 1

                                Try
                                    oRefDoc.Update2(True)
                                    ' oRefDoc.Save2(True)
                                Catch
                                End Try
                            End If
                        Catch
                        End Try
                    Next

                    ' Xử lý Assembly tổng
                    Call RemoveAllParameterLinks(oAsmDoc, countExcel, countDerived, countOLE)
                    countDocs += 1
                    oAsmDoc.Update2(True)
                    '  oAsmDoc.Save2(True)

                    '=================================================
                    ' 4. CHỈ CÁC PART (ALL LEVEL)
                    '=================================================
                ElseIf processOnlyParts Then
                    modeText = "Chỉ các Part (all level)"

                    For Each oRefDoc As Document In oAsmDoc.AllReferencedDocuments
                        Try
                            If oRefDoc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                                Call RemoveAllParameterLinks(oRefDoc, countExcel, countDerived, countOLE)
                                countDocs += 1

                                Try
                                    oRefDoc.Update2(True)
                                    ' oRefDoc.Save2(True)
                                Catch
                                End Try
                            End If
                        Catch
                        End Try
                    Next
                End If

                MessageBox.Show("Đã xử lý xong!" & vbCrLf & vbCrLf &
                                "Chế độ: " & modeText & vbCrLf &
                                "Số file đã xử lý: " & countDocs & vbCrLf &
                                "Excel Tables đã xóa: " & countExcel & vbCrLf &
                                "Derived Parameters đã xóa: " & countDerived & vbCrLf &
                                "OLE Links đã xóa: " & countOLE, "Visual Studio")

            Catch ex As Exception
                MessageBox.Show("Lỗi: " & ex.Message, "Visual Studio")
            Finally
                invApp.SilentOperation = False
            End Try

        End Sub

        ''' <summary>
        ''' Xóa liên kết Excel + liên kết Parameters giữa các file
        ''' </summary>
        Private Sub RemoveAllParameterLinks(ByVal oDoc As Document,
                                            ByRef countExcel As Integer,
                                            ByRef countDerived As Integer,
                                            ByRef countOLE As Integer)

            Dim oParams As Parameters = Nothing

            Try
                oParams = oDoc.ComponentDefinition.Parameters
            Catch
                Return
            End Try

            If oParams Is Nothing Then Return

            '=================================================
            ' 1. XÓA LIÊN KẾT EXCEL (ParameterTables)
            '=================================================
            Try
                Dim oTables As ParameterTables = oParams.ParameterTables

                For i As Integer = oTables.Count To 1 Step -1
                    Try
                        oTables.Item(i).Delete()
                        countExcel += 1
                    Catch
                    End Try
                Next
            Catch
            End Try

            '=================================================
            ' 2. XÓA LIÊN KẾT PARAMETERS GIỮA CÁC FILE
            '=================================================
            Try
                Dim oDerivedTables As DerivedParameterTables = oParams.DerivedParameterTables

                For i As Integer = oDerivedTables.Count To 1 Step -1
                    Try
                        Dim oTable As DerivedParameterTable = oDerivedTables.Item(i)

                        If oTable.HasReferenceComponent = False Then
                            oTable.Delete()
                            countDerived += 1
                        End If
                    Catch
                    End Try
                Next
            Catch
            End Try

            '=================================================
            ' 3. XÓA OLE LINKS EXCEL
            '=================================================
            Try
                For Each oOLE As ReferencedOLEFileDescriptor In oDoc.ReferencedOLEFileDescriptors
                    Try
                        Dim sName As String = LCase(oOLE.DisplayName)

                        If InStr(sName, ".xls") > 0 OrElse
                           InStr(sName, "excel") > 0 OrElse
                           InStr(sName, "embedding") > 0 Then

                            oOLE.Delete()
                            countOLE += 1
                        End If
                    Catch
                    End Try
                Next
            Catch
            End Try

        End Sub

    End Module
End Namespace