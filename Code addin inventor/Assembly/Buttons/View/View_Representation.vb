
Imports System.Windows.Forms
    Imports Inventor

Namespace ToolInventor2020.Assembly.Buttons.ASS_View

    Public Module View_Representation

        Public Sub OnExecute(ByVal Context As NameValueMap)

            Dim invApp As Inventor.Application = System.Runtime.InteropServices.Marshal.GetActiveObject("Inventor.Application")
            Dim oDoc As Document = invApp.ActiveDocument

            ' Chỉ chạy trên Assembly
            If oDoc Is Nothing OrElse oDoc.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
                MessageBox.Show("Chỉ hỗ trợ file Assembly (.iam)!", "Visual Studio")
                Return
            End If

            Dim oAsmDoc As AssemblyDocument = CType(oDoc, AssemblyDocument)

            '=================================================
            ' BƯỚC 1: CHỌN CHỨC NĂNG
            '=================================================
            Dim selectedFunc As Integer = ShowFunctionForm()
            If selectedFunc = 0 Then Return

            '=================================================
            ' BƯỚC 2: CHỌN PHẠM VI (Top Level / All Level)
            '=================================================
            Dim selectedScope As Integer = ShowScopeForm()
            If selectedScope = 0 Then Return

            Dim isAllLevel As Boolean = (selectedScope = 2)

            invApp.SilentOperation = True

            Try
                Select Case selectedFunc

                    Case 1
                        '----- Kích hoạt View theo tên -----
                        Dim viewName As String = InputBox("Nhập tên View Representation cần kích hoạt:", "Kích hoạt View", "Master")
                        If String.IsNullOrEmpty(viewName) Then Return

                        If isAllLevel Then
                            Call ActivateView_AllLevel(oAsmDoc, viewName)
                        Else
                            Call ActivateView_TopLevel(oAsmDoc, viewName)
                        End If

                    Case 2
                        '----- Chuyển về View Master -----
                        If isAllLevel Then
                            Call ActivateView_AllLevel(oAsmDoc, "Master")
                        Else
                            Call ActivateView_TopLevel(oAsmDoc, "Master")
                        End If

                    Case 3
                        '----- Chuyển về View Default -----
                        If isAllLevel Then
                            Call ActivateView_AllLevel(oAsmDoc, "Default")
                        Else
                            Call ActivateView_TopLevel(oAsmDoc, "Default")
                        End If

                End Select

            Catch ex As Exception
                MessageBox.Show("Lỗi: " & ex.Message, "Visual Studio")
            Finally
                invApp.SilentOperation = False
            End Try

        End Sub

        '=================================================
        ' KÍCH HOẠT VIEW - TOP LEVEL
        '=================================================
        Private Sub ActivateView_TopLevel(ByVal oAsmDoc As AssemblyDocument, ByVal viewName As String)

            Dim oRepMgr As RepresentationsManager = oAsmDoc.ComponentDefinition.RepresentationsManager
            Dim found As Boolean = False

            For Each oView As DesignViewRepresentation In oRepMgr.DesignViewRepresentations
                If String.Compare(oView.Name, viewName, True) = 0 Then
                    oView.Activate()
                    found = True
                    MessageBox.Show("Đã kích hoạt View """ & oView.Name & """ (Top Level)", "Visual Studio")
                    Exit For
                End If
            Next

            ' Nếu tìm Default không có thì thử Master
            If Not found AndAlso String.Compare(viewName, "Default", True) = 0 Then
                For Each oView As DesignViewRepresentation In oRepMgr.DesignViewRepresentations
                    If oView.Name = "Master" Then
                        oView.Activate()
                        found = True
                        MessageBox.Show("Không có View Default → đã chuyển về Master (Top Level)", "Visual Studio")
                        Exit For
                    End If
                Next
            End If

            If Not found Then
                MessageBox.Show("Không tìm thấy View: " & viewName, "Visual Studio")
            End If

        End Sub

        '=================================================
        ' KÍCH HOẠT VIEW - ALL LEVEL (đệ quy)
        '=================================================
        Private Sub ActivateView_AllLevel(ByVal oAsmDoc As AssemblyDocument, ByVal viewName As String)

            Dim countSuccess As Integer = 0
            Dim countFail As Integer = 0

            ' Xử lý Assembly tổng trước
            If ActivateViewInDoc(oAsmDoc, viewName) Then
                countSuccess += 1
            Else
                countFail += 1
            End If

            ' Xử lý tất cả Sub-Assembly
            For Each oRefDoc As Document In oAsmDoc.AllReferencedDocuments
                Try
                    If oRefDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                        If ActivateViewInDoc(CType(oRefDoc, AssemblyDocument), viewName) Then
                            countSuccess += 1
                        Else
                            countFail += 1
                        End If
                    End If
                Catch
                    countFail += 1
                End Try
            Next

            MessageBox.Show("Đã xử lý xong (All Level)!" & vbCrLf & vbCrLf &
                            "Thành công: " & countSuccess & vbCrLf &
                            "Không tìm thấy / Lỗi: " & countFail, "Visual Studio")

        End Sub

        '=================================================
        ' HÀM PHỤ: KÍCH HOẠT VIEW TRONG 1 ASSEMBLY
        '=================================================
        Private Function ActivateViewInDoc(ByVal oAsmDoc As AssemblyDocument, ByVal viewName As String) As Boolean

            Try
                Dim oRepMgr As RepresentationsManager = oAsmDoc.ComponentDefinition.RepresentationsManager

                For Each oView As DesignViewRepresentation In oRepMgr.DesignViewRepresentations
                    If String.Compare(oView.Name, viewName, True) = 0 Then
                        oView.Activate()
                        Return True
                    End If
                Next

                ' Fallback: nếu tìm Default không có thì dùng Master
                If String.Compare(viewName, "Default", True) = 0 Then
                    For Each oView As DesignViewRepresentation In oRepMgr.DesignViewRepresentations
                        If oView.Name = "Master" Then
                            oView.Activate()
                            Return True
                        End If
                    Next
                End If

            Catch
            End Try

            Return False

        End Function

        '=================================================
        ' FORM CHỌN CHỨC NĂNG
        '=================================================
        Private Function ShowFunctionForm() As Integer

            Dim selected As Integer = 0

            Dim frm As New Form()
            frm.Text = "View Representation"
            frm.Size = New System.Drawing.Size(360, 230)
            frm.StartPosition = FormStartPosition.CenterScreen
            frm.FormBorderStyle = FormBorderStyle.FixedDialog
            frm.MaximizeBox = False
            frm.MinimizeBox = False
            frm.TopMost = True

            Dim lbl As New Label()
            lbl.Text = "Chọn chức năng:"
            lbl.Location = New System.Drawing.Point(20, 15)
            lbl.Size = New System.Drawing.Size(300, 25)
            lbl.Font = New System.Drawing.Font("Segoe UI", 10, System.Drawing.FontStyle.Bold)
            frm.Controls.Add(lbl)

            Dim btn1 As New Button()
            btn1.Text = "1. Kích hoạt View theo tên"
            btn1.Location = New System.Drawing.Point(30, 50)
            btn1.Size = New System.Drawing.Size(280, 35)
            AddHandler btn1.Click, Sub()
                                       selected = 1
                                       frm.Close()
                                   End Sub
            frm.Controls.Add(btn1)

            Dim btn2 As New Button()
            btn2.Text = "2. Chuyển về View Master"
            btn2.Location = New System.Drawing.Point(30, 95)
            btn2.Size = New System.Drawing.Size(280, 35)
            AddHandler btn2.Click, Sub()
                                       selected = 2
                                       frm.Close()
                                   End Sub
            frm.Controls.Add(btn2)

            Dim btn3 As New Button()
            btn3.Text = "3. Chuyển về View Default"
            btn3.Location = New System.Drawing.Point(30, 140)
            btn3.Size = New System.Drawing.Size(280, 35)
            AddHandler btn3.Click, Sub()
                                       selected = 3
                                       frm.Close()
                                   End Sub
            frm.Controls.Add(btn3)

            frm.ShowDialog()
            Return selected

        End Function

        '=================================================
        ' FORM CHỌN PHẠM VI
        '=================================================
        Private Function ShowScopeForm() As Integer

            Dim selected As Integer = 0

            Dim frm As New Form()
            frm.Text = "Chọn phạm vi áp dụng"
            frm.Size = New System.Drawing.Size(340, 180)
            frm.StartPosition = FormStartPosition.CenterScreen
            frm.FormBorderStyle = FormBorderStyle.FixedDialog
            frm.MaximizeBox = False
            frm.MinimizeBox = False
            frm.TopMost = True

            Dim lbl As New Label()
            lbl.Text = "Áp dụng cho:"
            lbl.Location = New System.Drawing.Point(20, 15)
            lbl.Size = New System.Drawing.Size(280, 25)
            lbl.Font = New System.Drawing.Font("Segoe UI", 10, System.Drawing.FontStyle.Bold)
            frm.Controls.Add(lbl)

            Dim btn1 As New Button()
            btn1.Text = "1. Chỉ Top Level"
            btn1.Location = New System.Drawing.Point(30, 50)
            btn1.Size = New System.Drawing.Size(260, 35)
            AddHandler btn1.Click, Sub()
                                       selected = 1
                                       frm.Close()
                                   End Sub
            frm.Controls.Add(btn1)

            Dim btn2 As New Button()
            btn2.Text = "2. All Level (đệ quy toàn bộ)"
            btn2.Location = New System.Drawing.Point(30, 95)
            btn2.Size = New System.Drawing.Size(260, 35)
            AddHandler btn2.Click, Sub()
                                       selected = 2
                                       frm.Close()
                                   End Sub
            frm.Controls.Add(btn2)

            frm.ShowDialog()
            Return selected

        End Function

    End Module
End Namespace