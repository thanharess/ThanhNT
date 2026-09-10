
Option Explicit On

Imports Inventor
Imports System.Windows.Forms
Imports System.Drawing

Namespace ToolInventor2020.Assembly.Buttons.caclenhlapghep

    Public Module Ass_LG_3

        Public Sub OnExecute(ByVal Context As NameValueMap)

            Dim invApp As Inventor.Application = System.Runtime.InteropServices.Marshal.GetActiveObject("Inventor.Application")

            If invApp.ActiveDocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
                MessageBox.Show("Vui lòng mở Assembly trước!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Exit Sub
            End If

            Dim asmDoc As AssemblyDocument = invApp.ActiveEditDocument
            Dim oDef As AssemblyComponentDefinition = asmDoc.ComponentDefinition

            '=====================================================
            ' TẠO FORM CHỌN CHỨC NĂNG
            '=====================================================
            Dim frm As New Form()
            frm.Text = "Ẩn Component - Inventor 2020"
            frm.Size = New Size(380, 420)
            frm.StartPosition = FormStartPosition.CenterScreen
            frm.FormBorderStyle = FormBorderStyle.FixedDialog
            frm.MaximizeBox = False
            frm.MinimizeBox = False
            frm.Font = New Font("Segoe UI", 9)

            '----- Tạo các nút -----
            Dim btn1 As New Button() With {.Text = "1. Ẩn Referent", .Location = New System.Drawing.Point(40, 20), .Size = New Size(280, 35)}
            Dim btn2 As New Button() With {.Text = "2. Ẩn Phantom", .Location = New System.Drawing.Point(40, 60), .Size = New Size(280, 35)}
            Dim btn3 As New Button() With {.Text = "3. Ẩn Purchased (đồ mua)", .Location = New System.Drawing.Point(40, 100), .Size = New Size(280, 35)}
            Dim btn4 As New Button() With {.Text = "4. Ẩn Weldment (cụm hàn)", .Location = New System.Drawing.Point(40, 140), .Size = New Size(280, 35)}
            Dim btn5 As New Button() With {.Text = "5. Ẩn Part", .Location = New System.Drawing.Point(40, 180), .Size = New Size(280, 35)}
            Dim btn6 As New Button() With {.Text = "6. Ẩn Sheet Metal", .Location = New System.Drawing.Point(40, 220), .Size = New Size(280, 35)}
            Dim btn7 As New Button() With {.Text = "7. Ẩn TẤT CẢ các loại trên", .Location = New System.Drawing.Point(40, 260), .Size = New Size(280, 35)}
            Dim btn8 As New Button() With {.Text = "8. Hiện tất cả + chỉ ẩn Referent", .Location = New System.Drawing.Point(40, 300), .Size = New Size(280, 35), .BackColor = System.Drawing.Color.LightYellow}
            Dim btn0 As New Button() With {.Text = "0. Hiện lại tất cả", .Location = New System.Drawing.Point(40, 340), .Size = New Size(280, 35), .BackColor = System.Drawing.Color.LightGreen}

            '----- Gán sự kiện -----
            AddHandler btn1.Click, Sub()
                                       HideByType(oDef, "Referent")
                                       frm.Close()
                                   End Sub

            AddHandler btn2.Click, Sub()
                                       HideByType(oDef, "Phantom")
                                       frm.Close()
                                   End Sub

            AddHandler btn3.Click, Sub()
                                       HideByType(oDef, "Purchased")
                                       frm.Close()
                                   End Sub

            AddHandler btn4.Click, Sub()
                                       HideByType(oDef, "Weldment")
                                       frm.Close()
                                   End Sub

            AddHandler btn5.Click, Sub()
                                       HideByType(oDef, "Part")
                                       frm.Close()
                                   End Sub

            AddHandler btn6.Click, Sub()
                                       HideByType(oDef, "SheetMetal")
                                       frm.Close()
                                   End Sub

            AddHandler btn7.Click, Sub()
                                       HideByType(oDef, "All")
                                       frm.Close()
                                   End Sub

            AddHandler btn8.Click, Sub()
                                       ShowAll(oDef)
                                       HideByType(oDef, "Referent")
                                       frm.Close()
                                   End Sub

            AddHandler btn0.Click, Sub()
                                       ShowAll(oDef)
                                       MessageBox.Show("Đã hiện lại tất cả component!", "Hoàn tất", MessageBoxButtons.OK, MessageBoxIcon.Information)
                                       frm.Close()
                                   End Sub

            '----- Thêm nút vào form -----
            frm.Controls.AddRange({btn1, btn2, btn3, btn4, btn5, btn6, btn7, btn8, btn0})

            '----- Hiện form -----
            frm.ShowDialog()

            asmDoc.Update2(True)

        End Sub

        '=====================================================
        ' HÀM ẨN THEO LOẠI
        '=====================================================
        Private Sub HideByType(ByVal oDef As AssemblyComponentDefinition, ByVal mode As String)

            For Each occ As ComponentOccurrence In oDef.Occurrences

                Try
                    If occ.Suppressed Then Continue For

                    Dim shouldHide As Boolean = False

                    Select Case mode

                        Case "Referent"
                            If IsReferent(occ) Then shouldHide = True

                        Case "Phantom"
                            If IsPhantom(occ) Then shouldHide = True

                        Case "Purchased"
                            If IsPurchased(occ) Then shouldHide = True

                        Case "Weldment"
                            If IsWeldment(occ) Then shouldHide = True

                        Case "Part"
                            If IsPart(occ) Then shouldHide = True

                        Case "SheetMetal"
                            If IsSheetMetal(occ) Then shouldHide = True

                        Case "All"
                            If IsReferent(occ) OrElse
                               IsPhantom(occ) OrElse
                               IsPurchased(occ) OrElse
                               IsWeldment(occ) OrElse
                               IsPart(occ) OrElse
                               IsSheetMetal(occ) Then
                                shouldHide = True
                            End If

                    End Select

                    If shouldHide Then
                        occ.Visible = False
                        ' occ.Suppress()   ' Bỏ comment nếu muốn Suppress
                    End If

                Catch
                End Try

            Next

            MessageBox.Show("Đã ẩn xong theo chế độ: " & mode, "Hoàn tất", MessageBoxButtons.OK, MessageBoxIcon.Information)

        End Sub

        '=====================================================
        ' HIỆN LẠI TẤT CẢ
        '=====================================================
        Private Sub ShowAll(ByVal oDef As AssemblyComponentDefinition)

            For Each occ As ComponentOccurrence In oDef.Occurrences
                Try
                    occ.Visible = True
                    If occ.Suppressed Then occ.Unsuppress()
                Catch
                End Try
            Next

        End Sub

        '=====================================================
        ' KIỂM TRA TỪNG LOẠI
        '=====================================================
        Private Function IsReferent(ByVal occ As ComponentOccurrence) As Boolean
            Try
                If occ.IsContentMember Then Return True

                If occ.DefinitionDocumentType = DocumentTypeEnum.kPartDocumentObject Then
                    Dim pDoc As PartDocument = occ.Definition.Document
                    If pDoc.ComponentDefinition.IsReferencePart Then Return True
                End If
            Catch
            End Try
            Return False
        End Function

        Private Function IsPhantom(ByVal occ As ComponentOccurrence) As Boolean
            Try
                If occ.BOMStructure = BOMStructureEnum.kPhantomBOMStructure Then Return True
            Catch
            End Try
            Return False
        End Function

        Private Function IsPurchased(ByVal occ As ComponentOccurrence) As Boolean
            Try
                Dim doc As Document = occ.Definition.Document
                Dim designProps As PropertySet = doc.PropertySets.Item("Design Tracking Properties")
                Dim desc As String = ""
                Try
                    desc = designProps.Item("Description").Value.ToString()
                Catch
                End Try

                If LCase(desc).Contains("purchased") OrElse LCase(desc).Contains("đồ mua") Then Return True
                If occ.BOMStructure = BOMStructureEnum.kPurchasedBOMStructure Then Return True

            Catch
            End Try
            Return False
        End Function

        Private Function IsWeldment(ByVal occ As ComponentOccurrence) As Boolean
            Try
                If occ.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                    Dim aDoc As AssemblyDocument = occ.Definition.Document
                    If aDoc.ComponentDefinition.IsWeldment Then Return True
                End If
            Catch
            End Try
            Return False
        End Function

        Private Function IsPart(ByVal occ As ComponentOccurrence) As Boolean
            Try
                If occ.DefinitionDocumentType = DocumentTypeEnum.kPartDocumentObject Then
                    Dim pDoc As PartDocument = occ.Definition.Document
                    If Not pDoc.ComponentDefinition.IsSheetMetal AndAlso
                       Not pDoc.ComponentDefinition.IsReferencePart Then
                        Return True
                    End If
                End If
            Catch
            End Try
            Return False
        End Function

        Private Function IsSheetMetal(ByVal occ As ComponentOccurrence) As Boolean
            Try
                If occ.DefinitionDocumentType = DocumentTypeEnum.kPartDocumentObject Then
                    Dim pDoc As PartDocument = occ.Definition.Document
                    If pDoc.ComponentDefinition.IsSheetMetal Then Return True
                End If
            Catch
            End Try
            Return False
        End Function

    End Module

End Namespace