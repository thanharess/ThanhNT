Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Windows.Forms
Imports Inventor

Namespace ToolInventor2020.Drawing.Buttons.Drawtext
    Public Module ThayChuTrongTextModule

        '═══════════════════════════════════════════════════════════
        ' STATE cho InteractionEvents (giữ giữa các callback)
        '═══════════════════════════════════════════════════════════
        Private _inventorApp As Inventor.Application
        Private _oDrawDoc As DrawingDocument
        Private _pendingPairs As List(Of ReplacePair)
        Private _interactEvents As InteractionEvents
        Private _selectEvents As SelectEvents
        Private _collectedObjects As New List(Of Object)()


        '═══════════════════════════════════════════════════════════
        ' LỆNH CHÍNH
        '═══════════════════════════════════════════════════════════
        Public Sub OnExecute(ByVal Context As NameValueMap)

            ' ── 1. Lấy Inventor ──
            Try
                _inventorApp = CType(Marshal.GetActiveObject("Inventor.Application"), Inventor.Application)
            Catch ex As Exception
                Console.WriteLine("Không tìm thấy Inventor: " & ex.Message)
                Return
            End Try

            If _inventorApp.ActiveDocument Is Nothing OrElse
               _inventorApp.ActiveDocument.DocumentType <> DocumentTypeEnum.kDrawingDocumentObject Then
                Console.WriteLine("Vui lòng mở một bản vẽ (IDW/DWG) trước.")
                Return
            End If

            _oDrawDoc = CType(_inventorApp.ActiveDocument, DrawingDocument)

            ' ── 2. Form nhập cặp TÌM / THAY ──
            Dim preSelectionCount As Integer = 0
            Try : preSelectionCount = _oDrawDoc.SelectSet.Count : Catch : End Try

            Using form As New ReplaceForm(preSelectionCount)
                If form.ShowDialog() <> DialogResult.OK Then
                    Console.WriteLine("✖ Đã hủy lệnh.")
                    Return
                End If
                _pendingPairs = form.Pairs
            End Using

            If _pendingPairs Is Nothing OrElse _pendingPairs.Count = 0 Then
                Console.WriteLine("⚠️ Không có cặp tìm/thay nào.")
                Return
            End If

            Console.WriteLine("=== THAYCHUTRONGTEXT ===")
            For Each p In _pendingPairs
                Console.WriteLine($"  • ""{p.Find}""  →  ""{p.Replace}""")
            Next

            ' ── 3. Xác định danh sách đối tượng ──
            If preSelectionCount > 0 Then
                ' ▶ Dùng các đối tượng đã chọn trước
                Dim targets As New List(Of Object)()
                For i As Integer = 1 To preSelectionCount
                    targets.Add(_oDrawDoc.SelectSet.Item(i))
                Next

                Console.WriteLine($"Sử dụng {targets.Count} đối tượng đã chọn.")
                ApplyAll(targets, _pendingPairs)

                Try
                    _oDrawDoc.SelectSet.Clear()
                    _inventorApp.ActiveView.Update()
                Catch
                End Try

                _pendingPairs = Nothing
            Else
                ' ▶ Bật chế độ chọn tương tác (quét chuột)
                StartInteractiveSelect()
                ' ⚠️ Không làm gì thêm — sự kiện OnTerminate sẽ tiếp tục xử lý
            End If
        End Sub


        '═══════════════════════════════════════════════════════════
        ' BẮT ĐẦU CHẾ ĐỘ CHỌN TƯƠNG TÁC
        '═══════════════════════════════════════════════════════════
        Private Sub StartInteractiveSelect()
            _collectedObjects.Clear()

            Try : _oDrawDoc.SelectSet.Clear() : Catch : End Try

            Try
                _interactEvents = _inventorApp.CommandManager.CreateInteractionEvents()
                _interactEvents.InteractionDisabled = False
                ' ── ĐÃ BỎ dòng StopOnCommand ──

                _selectEvents = _interactEvents.SelectEvents
                _selectEvents.AddSelectionFilter(SelectionFilterEnum.kAllEntitiesFilter)
                _selectEvents.WindowSelectEnabled = True

                AddHandler _selectEvents.OnSelect, AddressOf HandleOnSelect
                AddHandler _interactEvents.OnTerminate, AddressOf HandleOnTerminate

                _interactEvents.Start()
            Catch ex As Exception
                MessageBox.Show("Không khởi động được chế độ chọn: " & ex.Message,
                        "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error)
                _interactEvents = Nothing
                _selectEvents = Nothing
                Return
            End Try

            Console.WriteLine("─────────────────────────────────────────────────")
            Console.WriteLine("▶ Chọn đối tượng bằng cách CLICK hoặc QUÉT CHUỘT.")
            Console.WriteLine("▶ Nhấn ESC (hoặc chuột phải → Done) khi chọn xong.")
            Console.WriteLine("─────────────────────────────────────────────────")

            ' Giữ OnExecute không return để InteractionEvents không bị kill
            Do While _interactEvents IsNot Nothing
                System.Windows.Forms.Application.DoEvents()
                System.Threading.Thread.Sleep(20)
            Loop

            Console.WriteLine("▶ Đã kết thúc chế độ chọn.")
        End Sub



        '═══════════════════════════════════════════════════════════
        ' CALLBACK: Mỗi khi user chọn 1 đối tượng (hoặc quét chuột)
        '═══════════════════════════════════════════════════════════
        Private Sub HandleOnSelect(ByVal JustSelectedEntities As ObjectsEnumerator,
                            ByVal SelectionDevice As SelectionDeviceEnum,
                            ByVal ModelPosition As Inventor.Point,
                            ByVal ViewPosition As Point2d,
                            ByVal View As Inventor.View)
            Try
                For Each obj As Object In JustSelectedEntities
                    If obj Is Nothing Then Continue For
                    If Not _collectedObjects.Contains(obj) Then
                        _collectedObjects.Add(obj)
                        Console.WriteLine($"  + {obj.Type}")
                    End If
                Next
            Catch ex As Exception
                Console.WriteLine("  ⚠️ OnSelect: " & ex.Message)
            End Try
        End Sub


        '═══════════════════════════════════════════════════════════
        ' CALLBACK: User kết thúc chọn (ESC / Done)
        '═══════════════════════════════════════════════════════════
        Private Sub HandleOnTerminate()
            ' Gỡ event handlers
            Try
                If _selectEvents IsNot Nothing Then
                    RemoveHandler _selectEvents.OnSelect, AddressOf HandleOnSelect
                End If
                If _interactEvents IsNot Nothing Then
                    RemoveHandler _interactEvents.OnTerminate, AddressOf HandleOnTerminate
                End If
            Catch
            End Try

            ' Áp dụng
            Try
                If _collectedObjects.Count > 0 Then
                    Console.WriteLine($"Đã chọn {_collectedObjects.Count} đối tượng. Đang áp dụng...")
                    ApplyAll(_collectedObjects, _pendingPairs)
                Else
                    Console.WriteLine("⚠️ Không có đối tượng nào được chọn.")
                End If
            Catch ex As Exception
                Console.WriteLine("  ⚠️ Apply: " & ex.Message)
            End Try

            ' Reset
            Try
                If _oDrawDoc IsNot Nothing Then
                    _oDrawDoc.SelectSet.Clear()
                    _inventorApp.ActiveView.Update()
                End If
            Catch
            End Try

            _collectedObjects.Clear()
            _pendingPairs = Nothing
            _selectEvents = Nothing
            _interactEvents = Nothing      ' ← Cờ để DoEvents loop thoát
        End Sub


        '═══════════════════════════════════════════════════════════
        ' ÁP DỤNG CẶP TÌM/THAY CHO DANH SÁCH ĐỐI TƯỢNG
        '═══════════════════════════════════════════════════════════
        Private Sub ApplyAll(ByVal targets As List(Of Object), ByVal pairs As List(Of ReplacePair))
            If targets Is Nothing OrElse pairs Is Nothing Then Return

            Dim count As Integer = 0
            For Each tObj In targets
                Dim oldTxt As String = GetTextFromEntity(tObj)
                If String.IsNullOrEmpty(oldTxt) Then Continue For

                Dim newTxt As String = oldTxt
                For Each p In pairs
                    If String.IsNullOrEmpty(p.Find) Then Continue For
                    newTxt = ReplaceIgnoreCaseSkipTags(newTxt, p.Find, If(p.Replace, ""))
                Next

                If newTxt <> oldTxt Then
                    If SetTextToEntity(tObj, newTxt) Then count += 1
                End If
            Next

            Console.WriteLine($"✅ Hoàn tất – đã xử lý {count}/{targets.Count} đối tượng.")

            Try
                _inventorApp.ActiveView.Update()
            Catch
            End Try
        End Sub


        '═══════════════════════════════════════════════════════════
        ' ĐỌC TEXT
        '═══════════════════════════════════════════════════════════
        Private Function GetTextFromEntity(ByVal obj As Object) As String
            Try
                If TypeOf obj Is GeneralNote Then Return CType(obj, GeneralNote).FormattedText
                If TypeOf obj Is LeaderNote Then Return CType(obj, LeaderNote).FormattedText
                If TypeOf obj Is Inventor.TextBox Then Return CType(obj, Inventor.TextBox).FormattedText

                If TypeOf obj Is DrawingDimension Then
                    Return CType(obj, DrawingDimension).Text.FormattedText
                End If

                If TypeOf obj Is DrawingViewLabel Then
                    Return CType(obj, DrawingViewLabel).FormattedText
                End If

                If TypeOf obj Is ModelGeneralNote Then
                    Return CType(obj, ModelGeneralNote).Definition.Text.FormattedText
                End If
                If TypeOf obj Is ModelLeaderNote Then
                    Return CType(obj, ModelLeaderNote).Definition.Text.FormattedText
                End If
            Catch
            End Try
            Return Nothing
        End Function


        '═══════════════════════════════════════════════════════════
        ' GHI TEXT
        '═══════════════════════════════════════════════════════════
        Private Function SetTextToEntity(ByVal obj As Object, ByVal newText As String) As Boolean
            Try
                If TypeOf obj Is GeneralNote Then
                    CType(obj, GeneralNote).FormattedText = newText
                    Return True
                End If
                If TypeOf obj Is LeaderNote Then
                    CType(obj, LeaderNote).FormattedText = newText
                    Return True
                End If
                If TypeOf obj Is Inventor.TextBox Then
                    CType(obj, Inventor.TextBox).FormattedText = newText
                    Return True
                End If
                If TypeOf obj Is DrawingDimension Then
                    CType(obj, DrawingDimension).Text.FormattedText = newText
                    Return True
                End If
                If TypeOf obj Is DrawingViewLabel Then
                    CType(obj, DrawingViewLabel).FormattedText = newText
                    Return True
                End If
                If TypeOf obj Is ModelGeneralNote Then
                    CType(obj, ModelGeneralNote).Definition.Text.FormattedText = newText
                    Return True
                End If
                If TypeOf obj Is ModelLeaderNote Then
                    CType(obj, ModelLeaderNote).Definition.Text.FormattedText = newText
                    Return True
                End If
            Catch ex As Exception
                Console.WriteLine("  ⚠️ " & ex.Message)
            End Try
            Return False
        End Function


        '═══════════════════════════════════════════════════════════
        ' REPLACE AN TOÀN — bỏ qua mọi tag <...>
        '═══════════════════════════════════════════════════════════
        Private Function ReplaceIgnoreCaseSkipTags(ByVal s As String, ByVal find As String, ByVal replace As String) As String
            If String.IsNullOrEmpty(s) OrElse String.IsNullOrEmpty(find) Then Return s

            Dim sb As New StringBuilder(s.Length)
            Dim i As Integer = 0

            While i < s.Length
                If s(i) = "<"c Then
                    Dim j As Integer = s.IndexOf(">"c, i)
                    If j < 0 Then
                        sb.Append(s.Substring(i))
                        Exit While
                    End If
                    sb.Append(s.Substring(i, j - i + 1))
                    i = j + 1
                Else
                    Dim j As Integer = s.IndexOf("<"c, i)
                    Dim chunkEnd As Integer = If(j < 0, s.Length, j)
                    sb.Append(ReplaceIgnoreCase(s.Substring(i, chunkEnd - i), find, replace))
                    i = chunkEnd
                End If
            End While

            Return sb.ToString()
        End Function


        Private Function ReplaceIgnoreCase(ByVal str As String, ByVal oldValue As String, ByVal newValue As String) As String
            If String.IsNullOrEmpty(oldValue) OrElse String.IsNullOrEmpty(str) Then Return str

            Dim sb As New StringBuilder()
            Dim i As Integer = 0
            While i < str.Length
                If i + oldValue.Length <= str.Length AndAlso
                   String.Compare(str, i, oldValue, 0, oldValue.Length,
                       StringComparison.OrdinalIgnoreCase) = 0 Then
                    sb.Append(newValue)
                    i += oldValue.Length
                Else
                    sb.Append(str(i))
                    i += 1
                End If
            End While
            Return sb.ToString()
        End Function

    End Module


    '═══════════════════════════════════════════════════════════
    ' Model
    '═══════════════════════════════════════════════════════════
    Public Class ReplacePair
        Public Property Find As String
        Public Property Replace As String
        Public Sub New(ByVal f As String, ByVal r As String)
            Me.Find = f
            Me.Replace = r
        End Sub
    End Class


    '═══════════════════════════════════════════════════════════
    ' FORM
    '═══════════════════════════════════════════════════════════
    Public Class ReplaceForm
        Inherits Form

        Private dgv As DataGridView
        Private btnAdd As Button
        Private btnRemove As Button
        Private btnOK As Button
        Private btnCancel As Button

        Public Property Pairs As List(Of ReplacePair)

        Public Sub New(ByVal preSelectionCount As Integer)
            Me.Text = "Thay chữ trong Text (Inventor)"
            Me.Width = 620
            Me.Height = 500
            Me.FormBorderStyle = FormBorderStyle.FixedDialog
            Me.StartPosition = FormStartPosition.CenterScreen
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.Font = New Font("Segoe UI", 9.0F)

            Dim lblTitle As New Label With {
                .Text = "Nhập các cặp TÌM / THAY THẾ (không phân biệt hoa thường):",
                .Left = 12, .Top = 10, .Width = 590,
                .Font = New Font(Me.Font, FontStyle.Bold)
            }

            Dim lblHint As New Label With {
                .Text = "• Cột 'Thay bằng' để trống = XÓA cụm từ đó" & vbCrLf &
                        "• Áp dụng an toàn với tag XML (<SCALE>, <VIEW>, <StyleOverride>...)",
                .Left = 12, .Top = 32, .Width = 590, .Height = 32,
                .ForeColor = System.Drawing.Color.DimGray
            }

            ' ── Thông báo trạng thái selection ──
            Dim lblSelection As New Label With {
                .Left = 12, .Top = 68, .Width = 590, .Height = 20
            }
            If preSelectionCount > 0 Then
                lblSelection.Text = $"✓ Đã có {preSelectionCount} đối tượng được chọn trên bản vẽ — sẽ áp dụng cho các đối tượng này."
                lblSelection.ForeColor = System.Drawing.Color.FromArgb(0, 120, 0)
            Else
                lblSelection.Text = "ℹ Chưa có đối tượng nào được chọn — sau khi OK, bạn sẽ chọn trên bản vẽ (quét chuột được)."
                lblSelection.ForeColor = System.Drawing.Color.FromArgb(0, 80, 160)
            End If

            ' ── Grid ──
            dgv = New DataGridView With {
                .Left = 12, .Top = 96, .Width = 590, .Height = 320,
                .AllowUserToAddRows = False,
                .AllowUserToDeleteRows = False,
                .AllowUserToResizeRows = False,
                .RowHeadersVisible = False,
                .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                .EditMode = DataGridViewEditMode.EditOnEnter,
                .SelectionMode = DataGridViewSelectionMode.CellSelect,
                .MultiSelect = False
            }

            dgv.Columns.Add(New DataGridViewTextBoxColumn With {
                .Name = "colFind", .HeaderText = "Tìm", .FillWeight = 50
            })
            dgv.Columns.Add(New DataGridViewTextBoxColumn With {
                .Name = "colReplace", .HeaderText = "Thay bằng  (để trống = xóa)", .FillWeight = 50
            })

            dgv.Rows.Add("", "")

            ' ── Buttons ──
            btnAdd = New Button With {.Text = "Thêm dòng", .Left = 12, .Top = 426, .Width = 110, .Height = 28}
            btnRemove = New Button With {.Text = "Xóa dòng", .Left = 130, .Top = 426, .Width = 110, .Height = 28}
            btnOK = New Button With {
                .Text = "OK", .Left = 406, .Top = 426, .Width = 90, .Height = 28,
                .DialogResult = DialogResult.OK
            }
            btnCancel = New Button With {
                .Text = "Cancel", .Left = 502, .Top = 426, .Width = 90, .Height = 28,
                .DialogResult = DialogResult.Cancel
            }

            AddHandler btnAdd.Click, Sub(s, e)
                                         Dim idx As Integer = dgv.Rows.Add("", "")
                                         dgv.CurrentCell = dgv.Rows(idx).Cells(0)
                                         dgv.BeginEdit(True)
                                     End Sub

            AddHandler btnRemove.Click, Sub(s, e)
                                            If dgv.CurrentRow IsNot Nothing Then dgv.Rows.Remove(dgv.CurrentRow)
                                            If dgv.Rows.Count = 0 Then dgv.Rows.Add("", "")
                                        End Sub

            AddHandler btnOK.Click, Sub(s, e)
                                        dgv.EndEdit()
                                        Pairs = New List(Of ReplacePair)()
                                        For Each row As DataGridViewRow In dgv.Rows
                                            Dim f As String = If(row.Cells(0).Value?.ToString(), "")
                                            Dim r As String = If(row.Cells(1).Value?.ToString(), "")
                                            If Not String.IsNullOrEmpty(f) Then
                                                Pairs.Add(New ReplacePair(f, r))
                                            End If
                                        Next
                                        If Pairs.Count = 0 Then
                                            MessageBox.Show(
                                                "Vui lòng nhập ít nhất 1 cặp TÌM/THAY.",
                                                "Chưa có dữ liệu",
                                                MessageBoxButtons.OK, MessageBoxIcon.Warning)
                                            Me.DialogResult = DialogResult.None
                                        End If
                                    End Sub

            Me.Controls.Add(lblTitle)
            Me.Controls.Add(lblHint)
            Me.Controls.Add(lblSelection)
            Me.Controls.Add(dgv)
            Me.Controls.Add(btnAdd)
            Me.Controls.Add(btnRemove)
            Me.Controls.Add(btnOK)
            Me.Controls.Add(btnCancel)

            Me.AcceptButton = btnOK
            Me.CancelButton = btnCancel
        End Sub
    End Class
End Namespace