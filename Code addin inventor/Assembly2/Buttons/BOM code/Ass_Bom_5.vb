


Option Explicit On
Option Strict Off

Imports Inventor
Imports System.Windows.Forms
Imports System.Collections.Generic
Imports System.Runtime.InteropServices


Namespace ToolInventor2020.Assembly2.Buttons.BOMcode

    Public Module Ass_Bom_5

        '===============================================================
        ' DANH SÁCH TỪ KHÓA – CHỈ KIỂM TRA ĐOẠN ĐẦU (StartsWith)
        ' Bạn có thể thêm / sửa ở đây
        '===============================================================

        ' Nhóm vòng bi / gối bi (xếp trước bulong trong nhóm Purchased)
        Private ReadOnly BearingKeywords As String() = {
            "vòng bi", "vong bi", "bearing", "motor",
            "gối bi", "goi bi", "gối đỡ", "goi do",
            "pillow", "plummer", "ucp", "ucf", "ucfl", "khóa trục", "khoa truc"
        }

        ' Nhóm bulong / ốc / vít / tiêu chuẩn (xếp DƯỚI CÙNG trong nhóm Purchased)
        Private ReadOnly FastenerKeywords As String() = {
            "bulong", "bu lông", "bu long", "ốc", "oc", "đai ốc", "dai oc", "vít", "vit", "ecu", "êcu", "then", "then chốt", "long đen", "long den",
            "long đen", "long den", "washer", "iso", "din", "jis", "m3", "m4", "m5", "m6", "m8", "lock collar", "locknut", "lock nut",
            "m10", "m12", "m16", "m20", "m24", "m30", "m36", "m42", "m48", "ss 2", "iso 4", "din 125", "din 127", "din 933", "din 934", "din 6912"
        }


        '===============================================================
        ' LẤY INVENTOR APPLICATION
        '===============================================================
        Private Function GetInventorApplication() As Inventor.Application
            Try
                Return CType(Marshal.GetActiveObject("Inventor.Application"), Inventor.Application)
            Catch ex As Exception
                MessageBox.Show("Không lấy được Inventor đang chạy." & vbCrLf & vbCrLf & ex.Message,
                                "BOM", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return Nothing
            End Try
        End Function


        Public Sub OnExecute(ByVal Context As NameValueMap)
            Dim invApp As Inventor.Application = Nothing
            Try
                invApp = GetInventorApplication()
                If invApp Is Nothing Then Exit Sub
                Main(invApp)
            Catch ex As Exception
                MessageBox.Show("Lỗi BOM:" & vbCrLf & vbCrLf & ex.Message,
                                "BOM", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub


        Private Sub Main(ByVal invApp As Inventor.Application)

            Dim oAsm As AssemblyDocument = Nothing

            Try
                oAsm = TryCast(invApp.ActiveDocument, AssemblyDocument)
                If oAsm Is Nothing Then
                    MessageBox.Show("Vui lòng mở Assembly (.iam) trước khi chạy.",
                                    "BOM", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Exit Sub
                End If

                Dim oBOM As BOM = oAsm.ComponentDefinition.BOM
                Try : oBOM.StructuredViewEnabled = True : Catch : End Try
                Try : oBOM.StructuredViewFirstLevelOnly = False : Catch : End Try

                Dim oBOMView As BOMView = Nothing
                Try
                    oBOMView = oBOM.BOMViews.Item("Structured")
                Catch
                    MessageBox.Show("Không tìm thấy Structured BOM.", "BOM",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End Try
                If oBOMView Is Nothing Then Exit Sub


                '-------------------------------------------------------
                ' NHẬP CHỮ
                '-------------------------------------------------------
                Dim baseText As String = InputBox(
                    "Nhập chữ dùng cho Part Number / Stock Number." & vbCrLf & vbCrLf &
                    "Để trống = chỉ Sort + đánh STT (không sửa PN/SN).",
                    "Part Number / Stock Number", "")

                If baseText Is Nothing Then baseText = ""
                baseText = baseText.Trim()


                Dim mode As Integer = 1
                Dim applyPN As Boolean = False
                Dim applySN As Boolean = False

                If baseText <> "" Then

                    Dim modeIdx As Integer = PickFromList(
                        "Chọn cách ghi",
                        New String() {
                            "2 - Thay toàn bộ (xóa cũ → chữ + STT)",
                            "3 - Thêm chữ phía SAU",
                            "4 - Thêm chữ phía TRƯỚC"
                        }, 0)

                    If modeIdx < 0 Then Exit Sub
                    mode = modeIdx + 2

                    Dim targetIdx As Integer = PickFromList(
                        "Áp dụng cho",
                        New String() {
                            "Chỉ Part Number",
                            "Chỉ Stock Number",
                            "Cả Part Number và Stock Number"
                        }, 2)

                    If targetIdx < 0 Then Exit Sub

                    Select Case targetIdx
                        Case 0 : applyPN = True : applySN = False
                        Case 1 : applyPN = False : applySN = True
                        Case 2 : applyPN = True : applySN = True
                    End Select
                End If


                '-------------------------------------------------------
                ' XÁC NHẬN
                '-------------------------------------------------------
                Dim confirm As String =
                    "CHUẨN BỊ CHẠY" & vbCrLf &
                    "------------------" & vbCrLf &
                    "Sort + đánh STT" & vbCrLf &
                    "• Purchased luôn nằm dưới Normal" & vbCrLf &
                    "• Trong Purchased: thường → vòng bi → bulong/ốc" & vbCrLf &
                    "• Từ khóa chỉ kiểm tra đoạn ĐẦU" & vbCrLf

                If baseText = "" Then
                    confirm &= "PN / SN: KHÔNG SỬA"
                Else
                    confirm &= "Chữ: " & baseText & vbCrLf
                    Select Case mode
                        Case 2 : confirm &= "Cách: Thay toàn bộ + STT" & vbCrLf
                        Case 3 : confirm &= "Cách: Thêm phía sau" & vbCrLf
                        Case 4 : confirm &= "Cách: Thêm phía trước" & vbCrLf
                    End Select

                    If applyPN AndAlso applySN Then
                        confirm &= "Áp dụng: PN + SN"
                    ElseIf applyPN Then
                        confirm &= "Áp dụng: Chỉ Part Number"
                    Else
                        confirm &= "Áp dụng: Chỉ Stock Number"
                    End If
                End If

                confirm &= vbCrLf & vbCrLf & "Tiếp tục?"

                If MessageBox.Show(confirm, "Xác nhận",
                                   MessageBoxButtons.OKCancel,
                                   MessageBoxIcon.Question) <> DialogResult.OK Then
                    Exit Sub
                End If


                '-------------------------------------------------------
                ' SORT
                '-------------------------------------------------------
                Dim sortedRows As List(Of BOMRow) = SortRows(oBOMView.BOMRows)
                If sortedRows Is Nothing OrElse sortedRows.Count = 0 Then
                    MessageBox.Show("Không có dòng BOM nào để xử lý.", "BOM")
                    Exit Sub
                End If


                '-------------------------------------------------------
                ' XỬ LÝ
                '-------------------------------------------------------
                Dim stt As Integer = 1
                Dim changedPN As Integer = 0
                Dim changedSN As Integer = 0
                Dim totalRows As Integer = 0
                Dim listDocs As New List(Of Document)

                For Each row As BOMRow In sortedRows

                    System.Windows.Forms.Application.DoEvents()
                    If row Is Nothing Then Continue For

                    Dim refDoc As Document = Nothing
                    Try
                        If row.ComponentDefinitions Is Nothing OrElse
                           row.ComponentDefinitions.Count = 0 Then Continue For
                        refDoc = row.ComponentDefinitions.Item(1).Document
                    Catch
                        Continue For
                    End Try

                    If refDoc Is Nothing Then Continue For

                    Try
                        If String.Equals(refDoc.FullFileName, oAsm.FullFileName,
                                         StringComparison.OrdinalIgnoreCase) Then
                            Continue For
                        End If
                    Catch
                    End Try

                    ' STT
                    Try
                        row.ItemNumber = stt.ToString()
                    Catch
                    End Try

                    ' Ghi PN / SN
                    If baseText <> "" AndAlso mode >= 2 Then
                        Dim curPN As String = GetProperty(refDoc, "Part Number")
                        Dim curSN As String = GetProperty(refDoc, "Stock Number")

                        Dim newPN As String = BuildValue(curPN, baseText, stt, mode)
                        Dim newSN As String = BuildValue(curSN, baseText, stt, mode)

                        If applyPN AndAlso newPN <> "" Then
                            If SetDesignProperty(refDoc, "Part Number", newPN) Then
                                changedPN += 1
                                If Not listDocs.Contains(refDoc) Then listDocs.Add(refDoc)
                            End If
                        End If

                        If applySN AndAlso newSN <> "" Then
                            If SetDesignProperty(refDoc, "Stock Number", newSN) Then
                                changedSN += 1
                                If Not listDocs.Contains(refDoc) Then listDocs.Add(refDoc)
                            End If
                        End If
                    End If

                    totalRows += 1
                    stt += 1
                Next


                ' Save + Update
                For Each d As Document In listDocs
                    Try
                        If d.IsModifiable Then
                            d.Update()
                            ''''''''''''''''''''''''''''''''''''''''''''''''''''  d.Save2(True)
                        End If
                    Catch
                    End Try
                Next

                Try : oBOM.Update() : Catch : End Try
                Try : oAsm.Update2(True) : Catch : End Try
                Try
                    '''''''''''''''''''''''''''''''''''''''''''  If oAsm.IsModifiable Then oAsm.Save2(True)
                Catch
                End Try


                Dim msg As String =
                    "HOÀN TẤT" & vbCrLf &
                    "====================" & vbCrLf &
                    "Tổng dòng xử lý : " & totalRows.ToString() & vbCrLf &
                    "STT cuối        : " & (stt - 1).ToString() & vbCrLf & vbCrLf

                If baseText = "" Then
                    msg &= "Part Number / Stock Number: KHÔNG SỬA"
                Else
                    msg &= "Part Number đã ghi : " & changedPN.ToString() & vbCrLf &
                           "Stock Number đã ghi: " & changedSN.ToString()
                End If

                MessageBox.Show(msg, "BOM", MessageBoxButtons.OK, MessageBoxIcon.Information)

            Catch ex As Exception
                MessageBox.Show("Lỗi:" & vbCrLf & vbCrLf & ex.Message,
                                "BOM", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub


        '===============================================================
        ' BUILD VALUE
        '===============================================================
        Private Function BuildValue(current As String, baseText As String, stt As Integer, mode As Integer) As String
            If current Is Nothing Then current = ""
            current = current.Trim()

            Select Case mode
                Case 2 : Return baseText & stt.ToString()
                Case 3
                    If current = "" Then Return baseText
                    Return current & " " & baseText
                Case 4
                    If current = "" Then Return baseText
                    Return baseText & " " & current
            End Select
            Return ""
        End Function


        '===============================================================
        ' SET / GET PROPERTY
        '===============================================================
        Private Function SetDesignProperty(doc As Document, propName As String, value As String) As Boolean
            Try
                If doc Is Nothing OrElse Not doc.IsModifiable Then Return False

                Dim designProps As PropertySet = Nothing
                Try
                    designProps = doc.PropertySets.Item("Design Tracking Properties")
                Catch
                    Return False
                End Try

                Try
                    Dim prop As Inventor.Property = designProps.Item(propName)
                    prop.Value = value
                    Return True
                Catch
                    Try
                        designProps.Add(value, propName)
                        Return True
                    Catch
                        Return False
                    End Try
                End Try
            Catch
                Return False
            End Try
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


        '===============================================================
        ' LẤY TEXT ĐỂ KIỂM TRA (ưu tiên Part Number → Stock Number → Description)
        '===============================================================
        Private Function GetSearchText(row As BOMRow) As String
            Try
                If row Is Nothing OrElse row.ComponentDefinitions Is Nothing OrElse
                   row.ComponentDefinitions.Count = 0 Then Return ""

                Dim doc As Document = row.ComponentDefinitions.Item(1).Document

                Dim pn As String = GetProperty(doc, "Part Number")
                If pn <> "" Then Return pn.Trim().ToLowerInvariant()

                Dim sn As String = GetProperty(doc, "Stock Number")
                If sn <> "" Then Return sn.Trim().ToLowerInvariant()

                Dim desc As String = GetProperty(doc, "Description")
                Return desc.Trim().ToLowerInvariant()

            Catch
                Return ""
            End Try
        End Function


        '===============================================================
        ' KIỂM TRA TỪ KHÓA – CHỈ ĐOẠN ĐẦU (StartsWith)
        '===============================================================
        Private Function IsBearing(text As String) As Boolean
            If String.IsNullOrEmpty(text) Then Return False
            For Each kw As String In BearingKeywords
                If text.StartsWith(kw.ToLowerInvariant()) Then Return True
            Next
            Return False
        End Function

        Private Function IsFastener(text As String) As Boolean
            If String.IsNullOrEmpty(text) Then Return False
            For Each kw As String In FastenerKeywords
                If text.StartsWith(kw.ToLowerInvariant()) Then Return True
            Next
            Return False
        End Function


        '===============================================================
        ' SORT
        '===============================================================
        ' Thứ tự cố định:
        ' 1. Assembly thường
        ' 2. Assembly Purchased
        ' 3. Part thường                ← KHÔNG ép gì xuống
        ' 4. Part Purchased
        '      (trong nhóm này: thường → vòng bi → bulong/ốc)
        ' 5. Phantom Assembly
        ' 6. Phantom Part
        ' 7. Reference
        '===============================================================
        Private Function SortRows(ByVal bomRows As BOMRowsEnumerator) As List(Of BOMRow)

            Dim normalAsm As New List(Of Tuple(Of BOMRow, Double))
            Dim purchasedAsm As New List(Of Tuple(Of BOMRow, Integer, String))
            Dim normalPart As New List(Of Tuple(Of BOMRow, Double))
            Dim purchasedPart As New List(Of Tuple(Of BOMRow, Integer, String))
            Dim phantomAsm As New List(Of Tuple(Of BOMRow, Double))
            Dim phantomPart As New List(Of Tuple(Of BOMRow, Double))
            Dim reference As New List(Of BOMRow)

            If bomRows Is Nothing Then Return New List(Of BOMRow)

            For Each row As BOMRow In bomRows
                If row Is Nothing Then Continue For

                Dim doc As Document = Nothing
                Try
                    If row.ComponentDefinitions Is Nothing OrElse
                       row.ComponentDefinitions.Count = 0 Then
                        If row.BOMStructure = BOMStructureEnum.kReferenceBOMStructure Then
                            reference.Add(row)
                        End If
                        Continue For
                    End If
                    doc = row.ComponentDefinitions.Item(1).Document
                Catch
                    Continue For
                End Try
                If doc Is Nothing Then Continue For

                Dim isAsm As Boolean = False
                Dim isPart As Boolean = False
                Try
                    isAsm = (doc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject)
                    isPart = (doc.DocumentType = DocumentTypeEnum.kPartDocumentObject)
                Catch
                End Try

                Dim searchText As String = GetSearchText(row)
                Dim isFast As Boolean = IsFastener(searchText)
                Dim isBear As Boolean = IsBearing(searchText)

                '----- REFERENCE -----
                If row.BOMStructure = BOMStructureEnum.kReferenceBOMStructure Then
                    reference.Add(row)
                    Continue For
                End If

                '----- PHANTOM -----
                If row.BOMStructure = BOMStructureEnum.kPhantomBOMStructure Then
                    Dim m As Double = GetMass(doc)
                    If isAsm Then
                        phantomAsm.Add(Tuple.Create(row, m))
                    ElseIf isPart Then
                        phantomPart.Add(Tuple.Create(row, m))
                    Else
                        reference.Add(row)
                    End If
                    Continue For
                End If

                '----- PURCHASED -----
                If row.BOMStructure = BOMStructureEnum.kPurchasedBOMStructure Then
                    Dim prio As Integer = 0
                    If isFast Then prio = 2
                    If isBear Then prio = 1

                    Dim pn As String = GetPartNumber(row)

                    If isAsm Then
                        purchasedAsm.Add(Tuple.Create(row, prio, pn))
                    Else
                        purchasedPart.Add(Tuple.Create(row, prio, pn))
                    End If
                    Continue For
                End If

                '----- NORMAL (không ép xuống) -----
                If isAsm Then
                    normalAsm.Add(Tuple.Create(row, GetMass(doc)))
                ElseIf isPart Then
                    normalPart.Add(Tuple.Create(row, GetMass(doc)))
                Else
                    reference.Add(row)
                End If
            Next


            '----- SORT -----
            normalAsm.Sort(Function(a, b) b.Item2.CompareTo(a.Item2))
            normalPart.Sort(Function(a, b) b.Item2.CompareTo(a.Item2))
            phantomAsm.Sort(Function(a, b) b.Item2.CompareTo(a.Item2))
            phantomPart.Sort(Function(a, b) b.Item2.CompareTo(a.Item2))

            ' Trong Purchased: priority 0 → 1 (vòng bi) → 2 (bulong/ốc)
            purchasedAsm.Sort(Function(a, b)
                                  Dim c = a.Item2.CompareTo(b.Item2)
                                  If c <> 0 Then Return c
                                  Return String.Compare(a.Item3, b.Item3, StringComparison.OrdinalIgnoreCase)
                              End Function)

            purchasedPart.Sort(Function(a, b)
                                   Dim c = a.Item2.CompareTo(b.Item2)
                                   If c <> 0 Then Return c
                                   Return String.Compare(a.Item3, b.Item3, StringComparison.OrdinalIgnoreCase)
                               End Function)

            reference.Sort(Function(a, b) String.Compare(GetPartNumber(a), GetPartNumber(b), StringComparison.OrdinalIgnoreCase))


            '----- GỘP -----
            Dim result As New List(Of BOMRow)

            For Each x In normalAsm : result.Add(x.Item1) : Next
            For Each x In purchasedAsm : result.Add(x.Item1) : Next
            For Each x In normalPart : result.Add(x.Item1) : Next
            For Each x In purchasedPart : result.Add(x.Item1) : Next
            For Each x In phantomAsm : result.Add(x.Item1) : Next
            For Each x In phantomPart : result.Add(x.Item1) : Next
            For Each x In reference : result.Add(x) : Next

            Return result
        End Function


        Private Function GetMass(ByVal doc As Document) As Double
            Try
                If doc Is Nothing Then Return 0
                If doc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                    Return CType(doc, AssemblyDocument).ComponentDefinition.MassProperties.Mass
                ElseIf doc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                    Return CType(doc, PartDocument).ComponentDefinition.MassProperties.Mass
                End If
            Catch
            End Try
            Return 0
        End Function


        Private Function GetPartNumber(ByVal row As BOMRow) As String
            Try
                If row Is Nothing OrElse row.ComponentDefinitions Is Nothing OrElse
                   row.ComponentDefinitions.Count = 0 Then Return ""
                Return GetProperty(row.ComponentDefinitions.Item(1).Document, "Part Number")
            Catch
                Return ""
            End Try
        End Function


        '===============================================================
        ' PICK LIST
        '===============================================================
        Private Function PickFromList(ByVal title As String,
                                      ByVal items As String(),
                                      Optional ByVal defaultIndex As Integer = 0) As Integer

            Dim frm As New Form()
            Try
                frm.Text = title
                frm.StartPosition = FormStartPosition.CenterScreen
                frm.FormBorderStyle = FormBorderStyle.FixedDialog
                frm.MaximizeBox = False
                frm.MinimizeBox = False
                frm.ShowInTaskbar = False
                frm.Width = 460
                frm.Height = 310

                Dim lst As New ListBox()
                lst.Left = 12 : lst.Top = 12
                lst.Width = 420 : lst.Height = 200
                lst.Font = New System.Drawing.Font("Segoe UI", 10)

                For Each s As String In items
                    lst.Items.Add(s)
                Next

                If lst.Items.Count > 0 Then
                    If defaultIndex >= 0 AndAlso defaultIndex < lst.Items.Count Then
                        lst.SelectedIndex = defaultIndex
                    Else
                        lst.SelectedIndex = 0
                    End If
                End If

                Dim btnOK As New Button()
                btnOK.Text = "OK"
                btnOK.Left = 250 : btnOK.Top = 225
                btnOK.Width = 85 : btnOK.Height = 30
                btnOK.DialogResult = DialogResult.OK

                Dim btnCancel As New Button()
                btnCancel.Text = "Hủy"
                btnCancel.Left = 345 : btnCancel.Top = 225
                btnCancel.Width = 85 : btnCancel.Height = 30
                btnCancel.DialogResult = DialogResult.Cancel

                frm.Controls.Add(lst)
                frm.Controls.Add(btnOK)
                frm.Controls.Add(btnCancel)
                frm.AcceptButton = btnOK
                frm.CancelButton = btnCancel
                frm.KeyPreview = True

                AddHandler lst.DoubleClick,
                    Sub(s, e)
                        frm.DialogResult = DialogResult.OK
                        frm.Close()
                    End Sub

                AddHandler frm.KeyDown,
                    Sub(s, e)
                        If e.KeyCode = Keys.Escape Then
                            e.Handled = True
                            frm.DialogResult = DialogResult.Cancel
                            frm.Close()
                        End If
                    End Sub

                If frm.ShowDialog() <> DialogResult.OK Then Return -1
                If lst.SelectedIndex < 0 Then Return -1
                Return lst.SelectedIndex

            Finally
                If frm IsNot Nothing Then
                    Try : frm.Dispose() : Catch : End Try
                End If
            End Try
        End Function

    End Module

End Namespace