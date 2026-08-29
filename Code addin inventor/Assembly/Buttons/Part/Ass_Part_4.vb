Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Inventor

Namespace ThanhN.Assembly.Buttons.part
    Public Module Ass_Part_4


        Public Sub OnExecute(ByVal Context As NameValueMap)

            Try

                CreateAllFlatPatterns()

            Catch ex As Exception

                MessageBox.Show(
                    "Error: " & ex.Message,
                    "Assembly2 - Create Flat Pattern",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error)

            End Try

        End Sub

        Private Sub CreateAllFlatPatterns()

                '==========================================================
                ' Kiểm tra Active Document
                '==========================================================
                Dim oActiveDoc As Document =
                g_inventorApplication.ActiveDocument

                If oActiveDoc Is Nothing Then
                    Return
                End If


                If oActiveDoc.DocumentType <>
                DocumentTypeEnum.kAssemblyDocumentObject Then

                    MessageBox.Show(
                    "Vui lòng mở Assembly.",
                    "Create Flat Pattern",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning)

                    Return

                End If


                '==========================================================
                ' Lấy Assembly
                '==========================================================
                Dim oAsmDoc As AssemblyDocument =
                TryCast(oActiveDoc, AssemblyDocument)

                If oAsmDoc Is Nothing Then
                    Return
                End If


                '==========================================================
                ' Duyệt toàn bộ Reference Documents
                '==========================================================
                Dim oRefDocs As DocumentsEnumerator =
                oAsmDoc.AllReferencedDocuments

                Dim successCount As Integer = 0
                Dim skipCount As Integer = 0


                For Each oRefDoc As Document In oRefDocs

                    Try

                        '==================================================
                        ' Chỉ xử lý Part
                        '==================================================
                        If oRefDoc.DocumentType <>
                        DocumentTypeEnum.kPartDocumentObject Then

                            Continue For

                        End If


                    '==================================================
                    ' Lấy đường dẫn file IPT
                    '==================================================
                    Dim iptPathName As String = oRefDoc.FullFileName

                    If String.IsNullOrEmpty(iptPathName) Then
                        Continue For
                    End If

                    If Not System.IO.File.Exists(iptPathName) Then
                        Continue For
                    End If

                    Dim oPartDoc As PartDocument =
    TryCast(
        g_inventorApplication.Documents.Open(
            iptPathName,
            False),
        PartDocument)


                    If oPartDoc Is Nothing Then
                            Continue For
                        End If


                        Try

                            '==============================================
                            ' Kiểm tra có phải Sheet Metal hay không
                            '==============================================
                            Dim oCompDef As SheetMetalComponentDefinition =
                            TryCast(
                                oPartDoc.ComponentDefinition,
                                SheetMetalComponentDefinition)


                            If oCompDef Is Nothing Then

                                skipCount += 1
                                Continue For

                            End If


                            '==============================================
                            ' Tạo Flat Pattern
                            '==============================================
                            If oCompDef.HasFlatPattern = False Then

                                oCompDef.Unfold()

                            Else

                                oCompDef.FlatPattern.Edit()

                            End If


                            '==============================================
                            ' Thoát Flat Pattern Edit
                            '==============================================
                            oCompDef.FlatPattern.ExitEdit()


                            '==============================================
                            ' Update Part
                            '==============================================
                            oPartDoc.Update()

                            successCount += 1


                        Catch

                            skipCount += 1

                        Finally

                            '==============================================
                            ' Đóng Part
                            '==============================================
                            Try

                                oPartDoc.Close(
                                True)

                            Catch

                            End Try

                        End Try


                    Catch

                        skipCount += 1

                    End Try

                Next


                '==========================================================
                ' Update Assembly
                '==========================================================
                Try
                    oAsmDoc.Update()
                Catch
                End Try


                '==========================================================
                ' Thông báo
                '==========================================================
                MessageBox.Show(
                "Hoàn thành tạo Flat Pattern." &
                vbCrLf & vbCrLf &
                "Đã xử lý: " & successCount.ToString() &
                vbCrLf &
                "Bỏ qua/Lỗi: " & skipCount.ToString(),
                "Assembly2 - Create Flat Pattern",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information)

            End Sub



        End Module
End Namespace
