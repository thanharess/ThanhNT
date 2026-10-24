Imports System.Windows.Forms
Imports Inventor

Namespace ToolInventor2020.Toolngoai

    Public Module delete_link_OLE

        Public Sub OnExecute(ByVal Context As NameValueMap)

            Dim invApp As Inventor.Application = System.Runtime.InteropServices.Marshal.GetActiveObject("Inventor.Application")
            Dim oDoc As Document = invApp.ActiveDocument

            If oDoc Is Nothing Then
                MessageBox.Show("Không có tài liệu đang mở!", "Visual Studio")
                Return
            End If

            If oDoc.DocumentType <> DocumentTypeEnum.kPartDocumentObject AndAlso
               oDoc.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
                MessageBox.Show("Chỉ hỗ trợ file Part (.ipt) hoặc Assembly (.iam)!", "Visual Studio")
                Return
            End If

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
            ' NẾU LÀ PART → XỬ LÝ NGAY
            '=================================================
            If oDoc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                invApp.SilentOperation = True
                Try
                    Call RemoveAllParameterLinks(oDoc, countExcel, countDerived, countOLE)
                    countDocs = 1
                    oDoc.Update2(True)
                    'oDoc.Save2(True)

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
            ' NẾU LÀ ASSEMBLY → HIỆN FORM CHỌN
            '=================================================
            Dim selectedMode As Integer = ShowModeSelectionForm()

            If selectedMode = 0 Then Return   ' Người dùng đóng form / Cancel

            Select Case selectedMode
                Case 1
                    processTopOnly = True
                Case 2
                    processAll = True
                Case 3
                    processOnlyAssemblies = True
                Case 4
                    processOnlyParts = True
            End Select

            invApp.SilentOperation = True

            Try
                Dim oAsmDoc As AssemblyDocument = CType(oDoc, AssemblyDocument)

                If processTopOnly Then
                    modeText = "Chỉ Top Level (Assembly tổng)"
                    Call RemoveAllParameterLinks(oAsmDoc, countExcel, countDerived, countOLE)
                    countDocs += 1
                    oAsmDoc.Update2(True)
                    'oAsmDoc.Save2(True)

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
                                    'oRefDoc.Save2(True)
                                Catch
                                End Try
                            End If
                        Catch
                        End Try
                    Next

                    Call RemoveAllParameterLinks(oAsmDoc, countExcel, countDerived, countOLE)
                    countDocs += 1
                    oAsmDoc.Update2(True)
                    'oAsmDoc.Save2(True)

                ElseIf processOnlyAssemblies Then
                    modeText = "Chỉ các Assembly (all level)"

                    For Each oRefDoc As Document In oAsmDoc.AllReferencedDocuments
                        Try
                            If oRefDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                                Call RemoveAllParameterLinks(oRefDoc, countExcel, countDerived, countOLE)
                                countDocs += 1
                                Try
                                    oRefDoc.Update2(True)
                                    'oRefDoc.Save2(True)
                                Catch
                                End Try
                            End If
                        Catch
                        End Try
                    Next

                    Call RemoveAllParameterLinks(oAsmDoc, countExcel, countDerived, countOLE)
                    countDocs += 1
                    oAsmDoc.Update2(True)
                    'oAsmDoc.Save2(True)

                ElseIf processOnlyParts Then
                    modeText = "Chỉ các Part (all level)"

                    For Each oRefDoc As Document In oAsmDoc.AllReferencedDocuments
                        Try
                            If oRefDoc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                                Call RemoveAllParameterLinks(oRefDoc, countExcel, countDerived, countOLE)
                                countDocs += 1
                                Try
                                    oRefDoc.Update2(True)
                                    'oRefDoc.Save2(True)
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
        ''' Form chọn chế độ bằng nút bấm
        ''' </summary>
        Private Function ShowModeSelectionForm() As Integer

            Dim selected As Integer = 0

            Dim frm As New Form()
            frm.Text = "Xóa liên kết Parameters / OLE"
            frm.Size = New System.Drawing.Size(380, 280)
            frm.StartPosition = FormStartPosition.CenterScreen
            frm.FormBorderStyle = FormBorderStyle.FixedDialog
            frm.MaximizeBox = False
            frm.MinimizeBox = False
            frm.TopMost = True

            Dim lbl As New Label()
            lbl.Text = "Chọn chế độ xử lý:"
            lbl.Location = New System.Drawing.Point(20, 15)
            lbl.Size = New System.Drawing.Size(320, 25)
            lbl.Font = New System.Drawing.Font("Segoe UI", 10, System.Drawing.FontStyle.Bold)
            frm.Controls.Add(lbl)

            Dim btn1 As New Button()
            btn1.Text = "1. Chỉ Top Level (Assembly tổng)"
            btn1.Location = New System.Drawing.Point(30, 50)
            btn1.Size = New System.Drawing.Size(300, 35)
            btn1.Font = New System.Drawing.Font("Segoe UI", 9)
            AddHandler btn1.Click, Sub()
                                       selected = 1
                                       frm.Close()
                                   End Sub
            frm.Controls.Add(btn1)

            Dim btn2 As New Button()
            btn2.Text = "2. Đệ quy toàn bộ (Part + Assembly)"
            btn2.Location = New System.Drawing.Point(30, 95)
            btn2.Size = New System.Drawing.Size(300, 35)
            btn2.Font = New System.Drawing.Font("Segoe UI", 9)
            AddHandler btn2.Click, Sub()
                                       selected = 2
                                       frm.Close()
                                   End Sub
            frm.Controls.Add(btn2)

            Dim btn3 As New Button()
            btn3.Text = "3. Chỉ các Assembly (all level)"
            btn3.Location = New System.Drawing.Point(30, 140)
            btn3.Size = New System.Drawing.Size(300, 35)
            btn3.Font = New System.Drawing.Font("Segoe UI", 9)
            AddHandler btn3.Click, Sub()
                                       selected = 3
                                       frm.Close()
                                   End Sub
            frm.Controls.Add(btn3)

            Dim btn4 As New Button()
            btn4.Text = "4. Chỉ các Part (all level)"
            btn4.Location = New System.Drawing.Point(30, 185)
            btn4.Size = New System.Drawing.Size(300, 35)
            btn4.Font = New System.Drawing.Font("Segoe UI", 9)
            AddHandler btn4.Click, Sub()
                                       selected = 4
                                       frm.Close()
                                   End Sub
            frm.Controls.Add(btn4)

            frm.ShowDialog()

            Return selected

        End Function

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

            ' 1. Xóa Excel Tables
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

            ' 2. Xóa Derived Parameters
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

            ' 3. Xóa OLE Links Excel
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