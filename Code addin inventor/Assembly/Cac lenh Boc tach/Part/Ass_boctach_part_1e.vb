Option Explicit On
Option Strict Off
Imports System.Collections.Generic
Imports System.Windows.Forms
Imports Inventor
Imports IO = System.IO
Imports ToolInventor2020   ' CommonKeywords

Namespace ToolInventor2020.Assembly.Buttons.caclenhboctach.part
    Public Module Ass_boctach_part_1e

        Public Sub OnExecute(ByVal Context As NameValueMap)
            Dim oApp As Inventor.Application = g_inventorApplication
            Dim selectedFile As String = ""
            Dim activeDoc As Document = Nothing

            Try
                Try
                    activeDoc = oApp.ActiveDocument
                Catch
                    activeDoc = Nothing
                End Try

                If activeDoc IsNot Nothing AndAlso activeDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                    Dim result As DialogResult = MessageBox.Show("Sử dụng file lắp ghép đang mở?", "Chọn file", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                    If result = DialogResult.Yes Then
                        selectedFile = activeDoc.FullFileName
                    Else
                        selectedFile = SelectAssemblyFile()
                    End If
                Else
                    selectedFile = SelectAssemblyFile()
                End If

                If String.IsNullOrWhiteSpace(selectedFile) Then Exit Sub

                Dim oOrigAsmDoc As AssemblyDocument = Nothing
                Dim openedByCode As Boolean = False

                If activeDoc IsNot Nothing AndAlso String.Equals(selectedFile, activeDoc.FullFileName, StringComparison.OrdinalIgnoreCase) Then
                    oOrigAsmDoc = CType(activeDoc, AssemblyDocument)
                Else
                    oOrigAsmDoc = CType(oApp.Documents.Open(selectedFile, False), AssemblyDocument)
                    openedByCode = True
                End If

                ' Chọn loại chi tiết
                Dim partType As Integer = SimpleSelectPartType()
                If partType = 0 Then
                    If openedByCode Then oOrigAsmDoc.Close(False)
                    Exit Sub
                End If

                ' Tạo Assembly đích
                Dim targetDoc As AssemblyDocument = CType(oApp.Documents.Add(DocumentTypeEnum.kAssemblyDocumentObject), AssemblyDocument)
                Dim baseName As String = IO.Path.GetFileNameWithoutExtension(selectedFile)
                Dim displayName As String = baseName & "_Aggregated_Parts"
                targetDoc.DisplayName = displayName

                ' Parts Only BOM
                Dim oBOM As BOM = oOrigAsmDoc.ComponentDefinition.BOM
                Try : oBOM.PartsOnlyViewEnabled = True : Catch : End Try
                Try : oBOM.Update() : Catch : End Try

                Dim partsView As BOMView = Nothing
                Try
                    partsView = oBOM.BOMViews.Item("Parts Only")
                Catch
                    partsView = Nothing
                End Try

                If partsView Is Nothing Then
                    MessageBox.Show("Không lấy được Parts-Only BOM view.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    If openedByCode Then oOrigAsmDoc.Close(False)
                    targetDoc.Close(False)
                    Exit Sub
                End If

                Dim oTG As TransientGeometry = oApp.TransientGeometry
                Dim oMatrix As Matrix = oTG.CreateMatrix()
                Dim added As Integer = 0

                '==========================================================
                ' UPDATE ASSEMBLY ĐÍCH + HIỆN ITEM QTY TRONG BOM
                '==========================================================
                Try
                    targetDoc.Update2(True)
                Catch
                End Try

                Try
                    Dim targetBOM As BOM = targetDoc.ComponentDefinition.BOM

                    '------------------------------------------------------
                    ' BẬT PARTS ONLY
                    '------------------------------------------------------
                    Try
                        targetBOM.PartsOnlyViewEnabled = True
                    Catch
                    End Try

                    Try
                        targetBOM.Update()
                    Catch
                    End Try

                    '------------------------------------------------------
                    ' LẤY PARTS ONLY VIEW
                    '------------------------------------------------------
                    Dim targetPartsView As BOMView = Nothing

                    Try
                        targetPartsView = targetBOM.BOMViews.Item("Parts Only")
                    Catch
                        targetPartsView = Nothing
                    End Try

                    '------------------------------------------------------
                    ' HIỆN CỘT QTY
                    ' KHÔNG DÙNG BOMColumn
                    '------------------------------------------------------
                    If targetPartsView IsNot Nothing Then

                        Try
                            Dim cols As Object = targetPartsView.BOMColumns

                            For i As Integer = 1 To cols.Count

                                Dim col As Object = Nothing

                                Try
                                    col = cols.Item(i)
                                Catch
                                    col = Nothing
                                End Try

                                If col Is Nothing Then Continue For

                                Dim title As String = ""
                                Dim propName As String = ""

                                Try
                                    title = CStr(col.Title)
                                Catch
                                    title = ""
                                End Try

                                Try
                                    propName = CStr(col.PropertyName)
                                Catch
                                    propName = ""
                                End Try

                                '------------------------------------------
                                ' QTY / ITEM QTY / QUANTITY
                                '------------------------------------------
                                If String.Equals(title.Trim(), "QTY", StringComparison.OrdinalIgnoreCase) OrElse
                   String.Equals(title.Trim(), "Item Qty", StringComparison.OrdinalIgnoreCase) OrElse
                   String.Equals(title.Trim(), "Quantity", StringComparison.OrdinalIgnoreCase) OrElse
                   String.Equals(propName.Trim(), "Quantity", StringComparison.OrdinalIgnoreCase) Then

                                    Try
                                        col.Visible = True
                                    Catch
                                    End Try

                                End If

                            Next

                        Catch
                        End Try

                    End If

                Catch
                    ' Không để lỗi BOM làm dừng chương trình
                End Try

                MessageBox.Show(
    "Done." & vbCrLf &
    "Added occurrences: " & added.ToString(),
    "Info",
    MessageBoxButtons.OK,
    MessageBoxIcon.Information
)



                For Each row As BOMRow In partsView.BOMRows
                    Try
                        If row Is Nothing OrElse row.ComponentDefinitions Is Nothing OrElse row.ComponentDefinitions.Count = 0 Then Continue For

                        Dim doc As Document = row.ComponentDefinitions.Item(1).Document
                        If doc Is Nothing Then Continue For

                        Dim include As Boolean = False

                        ' Lấy Part Number (dùng chung cho filter 3 & 4)
                        Dim pn As String = ""
                        Try
                            pn = GetProperty(doc, "Part Number")
                        Catch
                        End Try

                        Select Case partType
                            Case 1   ' SHEET METAL
                                If doc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                                    Dim part As PartDocument = CType(doc, PartDocument)
                                    Dim sm As SheetMetalComponentDefinition = TryCast(part.ComponentDefinition, SheetMetalComponentDefinition)
                                    include = (sm IsNot Nothing AndAlso sm.Features.Count > 0)
                                End If

                            Case 2   ' PURCHASED
                                include = (row.BOMStructure = BOMStructureEnum.kPurchasedBOMStructure)

                            Case 3   ' STANDARD LIBRARY  (Bearing + Fastener + Standard)
                                include = IsLibraryPart(pn)

                            Case 4   ' PURCHASED + LIBRARY
                                Dim purchased As Boolean = (row.BOMStructure = BOMStructureEnum.kPurchasedBOMStructure)
                                include = purchased OrElse IsLibraryPart(pn)
                        End Select

                        If Not include Then Continue For

                        Dim qty As Integer = 1
                        Try
                            qty = CInt(row.ItemQuantity)
                        Catch
                            qty = 1
                        End Try


                        For i As Integer = 1 To Math.Max(1, qty)
                            Try
                                targetDoc.ComponentDefinition.Occurrences.Add(doc.FullFileName, oMatrix)
                                added += 1
                            Catch
                            End Try
                        Next

                    Catch
                    End Try
                Next

                targetDoc.Update2(True)

                MessageBox.Show("Done. Added occurrences: " & added.ToString(), "Info", MessageBoxButtons.OK, MessageBoxIcon.Information)

                If openedByCode Then oOrigAsmDoc.Close(False)

            Catch ex As Exception
                MessageBox.Show("Error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        '=====================================================
        ' Kiểm tra chi tiết thuộc thư viện tiêu chuẩn
        ' BearingKeywords + FastenerKeywords + StandardKeywords
        '=====================================================
        '  Private Function IsLibraryPart(pn As String) As Boolean
        ' If String.IsNullOrEmpty(pn) Then Return False
        'Return CommonKeywords.IsBearing(pn) OrElse
        '          CommonKeywords.IsFastener(pn) OrElse
        '         CommonKeywords.IsStandardKeyword(pn)
        'End Function




        '==========================================================
        ' KIỂM TRA KEYWORD TỪ KÝ TỰ ĐẦU TIÊN
        ' SO SÁNH TỪ TRÁI SANG PHẢI
        '==========================================================
        Private Function IsLibraryPart(ByVal pn As String) As Boolean

            If String.IsNullOrWhiteSpace(pn) Then
                Return False
            End If

            pn = pn.Trim().ToUpperInvariant()

            Try
                If ToolInventor2020.CommonKeywords.IsBearing(pn) Then
                    Return True
                End If
            Catch
            End Try

            Try
                If ToolInventor2020.CommonKeywords.IsFastener(pn) Then
                    Return True
                End If
            Catch
            End Try

            Try
                If ToolInventor2020.CommonKeywords.IsStandardKeyword(pn) Then
                    Return True
                End If
            Catch
            End Try

            Return False

        End Function


        '=====================================================
        ' Dialog chọn loại
        '=====================================================
        Private Function SimpleSelectPartType() As Integer
            Dim result As Integer = 0

            Using dlg As New Form()
                dlg.Text = "CHỌN DẠNG CHI TIẾT"
                dlg.StartPosition = FormStartPosition.CenterScreen
                dlg.Size = New System.Drawing.Size(420, 220)
                dlg.FormBorderStyle = FormBorderStyle.FixedDialog
                dlg.MaximizeBox = False
                dlg.MinimizeBox = False

                Dim b1 As New Button() With {.Text = "SHEET METAL", .Left = 20, .Top = 20, .Width = 160, .Height = 60}
                Dim b2 As New Button() With {.Text = "PURCHASED", .Left = 220, .Top = 20, .Width = 160, .Height = 60}
                Dim b3 As New Button() With {.Text = "STANDARD LIBRARY", .Left = 20, .Top = 100, .Width = 160, .Height = 60}
                Dim b4 As New Button() With {.Text = "PURCHASED + LIBRARY", .Left = 220, .Top = 100, .Width = 160, .Height = 60}

                AddHandler b1.Click, Sub()
                                         result = 1
                                         dlg.DialogResult = DialogResult.OK
                                     End Sub
                AddHandler b2.Click, Sub()
                                         result = 2
                                         dlg.DialogResult = DialogResult.OK
                                     End Sub
                AddHandler b3.Click, Sub()
                                         result = 3
                                         dlg.DialogResult = DialogResult.OK
                                     End Sub
                AddHandler b4.Click, Sub()
                                         result = 4
                                         dlg.DialogResult = DialogResult.OK
                                     End Sub

                dlg.Controls.Add(b1)
                dlg.Controls.Add(b2)
                dlg.Controls.Add(b3)
                dlg.Controls.Add(b4)

                dlg.ShowDialog()
            End Using

            Return result
        End Function

        Private Function SelectAssemblyFile() As String
            Try
                Using dlg As New OpenFileDialog()
                    dlg.Filter = "Inventor Assembly (*.iam)|*.iam|All files (*.*)|*.*"
                    If dlg.ShowDialog() = DialogResult.OK Then
                        Return dlg.FileName
                    End If
                End Using
            Catch
            End Try
            Return ""
        End Function

        Private Function GetProperty(doc As Document, propName As String) As String
            Try
                If doc Is Nothing Then Return ""
                Dim ps As PropertySet = doc.PropertySets.Item("Design Tracking Properties")
                Dim prop As Inventor.Property = ps.Item(propName)
                If prop Is Nothing OrElse prop.Value Is Nothing Then Return ""
                Return CStr(prop.Value).Trim()
            Catch
                Return ""
            End Try
        End Function

    End Module
End Namespace