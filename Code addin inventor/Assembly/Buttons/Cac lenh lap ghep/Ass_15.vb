
Option Explicit On

Imports Inventor
Imports System.Windows.Forms
Imports System.Collections.Generic
Imports IO = System.IO

Namespace ThanhN.Assembly.Buttons.caclenhlapghep

    Public Module ass_15

        Public Sub OnExecute(ByVal Context As NameValueMap)

            Dim app As Inventor.Application = g_inventorApplication
            Dim sourceDoc As AssemblyDocument = Nothing
            Dim targetDoc As AssemblyDocument = Nothing
            Dim sourceOpened As Boolean = False

            Try
                Dim partType As Integer = SelectPartType()
                If partType = 0 Then Exit Sub

                Dim selectedFile As String = GetSourceFile(app)
                If String.IsNullOrEmpty(selectedFile) Then Exit Sub

                Dim activeDoc As Document = Nothing

                Try
                    activeDoc = app.ActiveDocument
                Catch
                End Try

                If activeDoc IsNot Nothing AndAlso
                   activeDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject AndAlso
                   String.Equals(activeDoc.FullFileName,
                                 selectedFile,
                                 StringComparison.OrdinalIgnoreCase) Then

                    sourceDoc = CType(activeDoc, AssemblyDocument)

                Else
                    sourceDoc = CType(app.Documents.Open(selectedFile, False),
                                      AssemblyDocument)
                    sourceOpened = True
                End If

                Dim targetOption As Integer = SelectTargetType()

                If targetOption = 0 Then
                    If sourceOpened Then sourceDoc.Close(False)
                    Exit Sub
                End If

                If targetOption = 1 Then

                    targetDoc =
                        CType(app.Documents.Add(
                            DocumentTypeEnum.kAssemblyDocumentObject),
                            AssemblyDocument)

                    Dim baseName As String =
                        IO.Path.GetFileNameWithoutExtension(selectedFile)

                    Select Case partType
                        Case 1
                            targetDoc.DisplayName =
                                baseName & "_SheetMetal_Unfold"

                        Case 2
                            targetDoc.DisplayName =
                                baseName & "_Purchased_Parts"

                        Case 3
                            targetDoc.DisplayName =
                                baseName & "_Library_Parts"

                        Case 4
                            targetDoc.DisplayName =
                                baseName & "_Standard_Parts"
                    End Select

                Else

                    Dim targetFile As String =
                        SelectAssemblyFile("Chọn Assembly đích")

                    If String.IsNullOrEmpty(targetFile) Then

                        If sourceOpened Then
                            sourceDoc.Close(False)
                        End If

                        Exit Sub
                    End If

                    targetDoc =
                        CType(app.Documents.Open(targetFile, True),
                              AssemblyDocument)

                End If

                Dim parts As New List(Of Tuple(Of String, String))

                CollectParts(
                    sourceDoc.ComponentDefinition,
                    parts,
                    partType)

                If parts.Count = 0 Then

                    MessageBox.Show(
                        "Không tìm thấy chi tiết phù hợp.",
                        "Thông báo",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information)

                    If sourceOpened Then
                        sourceDoc.Close(False)
                    End If

                    targetDoc.Close(False)
                    Exit Sub
                End If

                parts.Sort(
                    Function(a, b)
                        Return StringComparer.OrdinalIgnoreCase.Compare(
                            a.Item2,
                            b.Item2)
                    End Function)

                Dim matrix As Matrix =
                    app.TransientGeometry.CreateMatrix()

                Dim index As Integer = 1

                For Each p As Tuple(Of String, String) In parts

                    Try

                        Dim occ As ComponentOccurrence =
                            targetDoc.ComponentDefinition.Occurrences.Add(
                                p.Item1,
                                matrix)

                        Try
                            occ.Grounded = True
                        Catch
                        End Try

                        Try
                            occ.Name =
                                index.ToString("000") &
                                " - " &
                                p.Item2
                        Catch
                        End Try

                        index += 1

                    Catch
                    End Try

                Next

                targetDoc.Update2(True)

                Dim typeName As String =
                    GetTypeName(partType)

                If targetOption = 1 Then

                    Dim defaultName As String = ""

                    Select Case partType

                        Case 1
                            defaultName = "SheetMetal_Unfold.iam"

                        Case 2
                            defaultName = "Purchased_Parts.iam"

                        Case 3
                            defaultName = "Library_Parts.iam"

                        Case 4
                            defaultName = "Standard_Parts.iam"

                    End Select

                    Dim saveFile As String =
                        SaveAssemblyFile(defaultName)

                    If Not String.IsNullOrEmpty(saveFile) Then

                        targetDoc.SaveAs(
                            saveFile,
                            False)

                        MessageBox.Show(
                            "Đã tạo Assembly mới." &
                            vbCrLf &
                            "Loại: " & typeName &
                            vbCrLf &
                            "Số lượng: " &
                            parts.Count.ToString() &
                            vbCrLf &
                            "File: " & saveFile,
                            "Hoàn tất",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information)

                    End If

                Else

                    targetDoc.Save()

                    MessageBox.Show(
                        "Đã thêm " &
                        parts.Count.ToString() &
                        " chi tiết " &
                        typeName &
                        " vào Assembly đích.",
                        "Hoàn tất",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information)

                End If

                If sourceOpened Then
                    sourceDoc.Close(False)
                End If

                targetDoc.Activate()

            Catch ex As Exception

                MessageBox.Show(
                    "Lỗi:" &
                    vbCrLf &
                    ex.Message,
                    "Select Parts",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error)

                If sourceOpened AndAlso
                   sourceDoc IsNot Nothing Then

                    Try
                        sourceDoc.Close(False)
                    Catch
                    End Try

                End If

            End Try

        End Sub

        Private Function SelectPartType() As Integer

            Dim result As Integer = 0

            Using f As New Form()

                f.Text = "CHỌN DẠNG CHI TIẾT"
                f.StartPosition =
                    FormStartPosition.CenterScreen

                f.Size =
                    New System.Drawing.Size(620, 375)

                f.FormBorderStyle =
                    FormBorderStyle.FixedDialog

                f.MaximizeBox = False
                f.MinimizeBox = False

                Dim title As New Label()

                title.Text =
                    "CHỌN DẠNG CHI TIẾT"

                title.Font =
                    New System.Drawing.Font(
                        "Tahoma",
                        18,
                        System.Drawing.FontStyle.Bold)

                title.AutoSize = True

                title.Location =
                    New System.Drawing.Point(135, 20)

                f.Controls.Add(title)

                Dim b1 As Button =
                    MakePartButton(
                        "SHEET METAL",
                        "Tấm sheet metal",
                        25,
                        70)

                Dim b2 As Button =
                    MakePartButton(
                        "PURCHASED",
                        "Chi tiết mua",
                        315,
                        70)

                Dim b3 As Button =
                    MakePartButton(
                        "STANDARD LIBRARY",
                        "ISO / DIN / SKF",
                        25,
                        175)

                Dim b4 As Button =
                    MakePartButton(
                        "PURCHASED + LIBRARY",
                        "Cả hai loại",
                        315,
                        175)

                AddHandler b1.Click,
                    Sub(sender As Object, e As EventArgs)

                        result = 1
                        f.Close()

                    End Sub

                AddHandler b2.Click,
                    Sub(sender As Object, e As EventArgs)

                        result = 2
                        f.Close()

                    End Sub

                AddHandler b3.Click,
                    Sub(sender As Object, e As EventArgs)

                        result = 3
                        f.Close()

                    End Sub

                AddHandler b4.Click,
                    Sub(sender As Object, e As EventArgs)

                        result = 4
                        f.Close()

                    End Sub

                f.Controls.Add(b1)
                f.Controls.Add(b2)
                f.Controls.Add(b3)
                f.Controls.Add(b4)

                Dim cancel As New Button()

                cancel.Text = "HỦY"

                cancel.Size =
                    New System.Drawing.Size(100, 30)

                cancel.Location =
                    New System.Drawing.Point(260, 275)

                cancel.DialogResult =
                    DialogResult.Cancel

                f.Controls.Add(cancel)

                f.CancelButton = cancel

                f.ShowDialog()

            End Using

            Return result

        End Function

        Private Function MakePartButton(
            text As String,
            subText As String,
            x As Integer,
            y As Integer) As Button

            Dim b As New Button()

            b.Text =
                text &
                vbCrLf &
                subText

            b.Font =
                New System.Drawing.Font(
                    "Tahoma",
                    10,
                    System.Drawing.FontStyle.Bold)

            b.Size =
                New System.Drawing.Size(270, 85)

            b.Location =
                New System.Drawing.Point(x, y)

            b.FlatStyle =
                FlatStyle.Flat

            Return b

        End Function

        Private Function SelectTargetType() As Integer

            Dim result As Integer = 0

            Using f As New Form()

                f.Text = "CHỌN ASSEMBLY ĐÍCH"

                f.StartPosition =
                    FormStartPosition.CenterScreen

                f.Size =
                    New System.Drawing.Size(620, 280)

                f.FormBorderStyle =
                    FormBorderStyle.FixedDialog

                f.MaximizeBox = False
                f.MinimizeBox = False

                Dim title As New Label()

                title.Text =
                    "Chọn cách thêm chi tiết"

                title.Font =
                    New System.Drawing.Font(
                        "Tahoma",
                        13,
                        System.Drawing.FontStyle.Bold)

                title.AutoSize = True

                title.Location =
                    New System.Drawing.Point(205, 20)

                f.Controls.Add(title)

                Dim newBtn As New Button()

                newBtn.Text =
                    "TẠO ASSEMBLY MỚI" &
                    vbCrLf &
                    "Tạo file .iam mới"

                newBtn.Font =
                    New System.Drawing.Font(
                        "Tahoma",
                        10,
                        System.Drawing.FontStyle.Bold)

                newBtn.Size =
                    New System.Drawing.Size(270, 90)

                newBtn.Location =
                    New System.Drawing.Point(25, 70)

                Dim oldBtn As New Button()

                oldBtn.Text =
                    "ASSEMBLY CÓ SẴN" &
                    vbCrLf &
                    "Chọn file .iam để thêm"

                oldBtn.Font =
                    New System.Drawing.Font(
                        "Tahoma",
                        10,
                        System.Drawing.FontStyle.Bold)

                oldBtn.Size =
                    New System.Drawing.Size(270, 90)

                oldBtn.Location =
                    New System.Drawing.Point(315, 70)

                AddHandler newBtn.Click,
                    Sub(sender As Object, e As EventArgs)

                        result = 1
                        f.Close()

                    End Sub

                AddHandler oldBtn.Click,
                    Sub(sender As Object, e As EventArgs)

                        result = 2
                        f.Close()

                    End Sub

                f.Controls.Add(newBtn)
                f.Controls.Add(oldBtn)

                Dim cancel As New Button()

                cancel.Text = "HỦY"

                cancel.Size =
                    New System.Drawing.Size(100, 30)

                cancel.Location =
                    New System.Drawing.Point(255, 180)

                cancel.DialogResult =
                    DialogResult.Cancel

                f.Controls.Add(cancel)

                f.CancelButton = cancel

                f.ShowDialog()

            End Using

            Return result

        End Function

        Private Function GetSourceFile(
            app As Inventor.Application) As String

            Dim doc As Document = Nothing

            Try
                doc = app.ActiveDocument
            Catch
            End Try

            If doc IsNot Nothing AndAlso
               doc.DocumentType =
               DocumentTypeEnum.kAssemblyDocumentObject Then

                Dim r As DialogResult =
                    MessageBox.Show(
                        "Sử dụng Assembly đang mở làm nguồn?",
                        "Assembly nguồn",
                        MessageBoxButtons.YesNoCancel,
                        MessageBoxIcon.Question)

                If r = DialogResult.Yes Then
                    Return doc.FullFileName
                End If

                If r = DialogResult.Cancel Then
                    Return ""
                End If

            End If

            Return SelectAssemblyFile(
                "Chọn Assembly nguồn")

        End Function

        Private Function SelectAssemblyFile(
            title As String) As String

            Using dlg As New OpenFileDialog()

                dlg.Title = title
                dlg.Filter =
                    "Assembly Files (*.iam)|*.iam"

                dlg.Multiselect = False

                If dlg.ShowDialog() =
                   DialogResult.OK Then

                    Return dlg.FileName

                End If

            End Using

            Return ""

        End Function

        Private Function SaveAssemblyFile(
            defaultName As String) As String

            Using dlg As New SaveFileDialog()

                dlg.Title =
                    "Lưu Assembly mới"

                dlg.Filter =
                    "Assembly Files (*.iam)|*.iam"

                dlg.DefaultExt = "iam"
                dlg.AddExtension = True
                dlg.FileName = defaultName

                If dlg.ShowDialog() =
                   DialogResult.OK Then

                    Return dlg.FileName

                End If

            End Using

            Return ""

        End Function

        Private Function GetTypeName(
            partType As Integer) As String

            Select Case partType

                Case 1
                    Return "Sheet Metal"

                Case 2
                    Return "Purchased"

                Case 3
                    Return "Standard Library"

                Case 4
                    Return "Purchased + Library"

            End Select

            Return "Part"

        End Function

        Private Sub CollectParts(
            asmDef As AssemblyComponentDefinition,
            ByRef parts As List(Of Tuple(Of String, String)),
            partType As Integer)

            For Each occ As ComponentOccurrence In asmDef.Occurrences

                Try

                    If occ.Suppressed Then
                        Continue For
                    End If

                    Dim doc As Document = Nothing

                    Try
                        doc =
                            occ.ReferencedDocumentDescriptor.
                            ReferencedDocument
                    Catch

                        Try
                            doc = occ.Definition.Document
                        Catch
                        End Try

                    End Try

                    If doc Is Nothing Then
                        Continue For
                    End If

                    If String.IsNullOrEmpty(
                        doc.FullFileName) Then

                        Continue For
                    End If

                    If doc.DocumentType =
                       DocumentTypeEnum.kPartDocumentObject Then

                        Dim part As PartDocument =
                            TryCast(doc, PartDocument)

                        If part Is Nothing Then
                            Continue For
                        End If

                        Dim selected As Boolean = False
                        Dim partNum As String = ""

                        Try

                            partNum =
                                CStr(
                                    part.PropertySets(
                                        "Design Tracking Properties").
                                    Item("Part Number").Value).Trim()

                        Catch
                        End Try

                        Dim pn As String =
                            partNum.ToUpperInvariant()

                        Select Case partType

                            Case 1

                                Dim sm As SheetMetalComponentDefinition =
                                    TryCast(
                                        part.ComponentDefinition,
                                        SheetMetalComponentDefinition)

                                selected =
                                    sm IsNot Nothing AndAlso
                                    sm.Features.Count > 0

                            Case 2

                                selected =
                                    occ.BOMStructure =
                                    BOMStructureEnum.
                                    kPurchasedBOMStructure

                            Case 3

                                selected =
                                    pn.Contains("ISO") OrElse
                                    pn.Contains("DIN") OrElse
                                    pn.Contains("SKF")

                            Case 4

                                Dim purchased As Boolean =
                                    occ.BOMStructure =
                                    BOMStructureEnum.
                                    kPurchasedBOMStructure

                                Dim library As Boolean =
                                    pn.Contains("ISO") OrElse
                                    pn.Contains("DIN") OrElse
                                    pn.Contains("SKF") OrElse
                                    pn.Contains("SS") OrElse
                                    pn.Contains("GB") OrElse
                                    pn.Contains("JIS") OrElse
                                    pn.Contains("ANSI") OrElse
                                    pn.Contains("BSI") OrElse
                                    pn.Contains("GOST") OrElse
                                    pn.Contains("ASTM")

                                selected =
                                    purchased OrElse library

                        End Select

                        If selected Then

                            Dim name As String =
                                partNum

                            If String.IsNullOrWhiteSpace(name) Then

                                name =
                                    IO.Path.
                                    GetFileNameWithoutExtension(
                                        part.FullFileName)

                            End If

                            parts.Add(
                                Tuple.Create(
                                    part.FullFileName,
                                    name))

                        End If

                    ElseIf doc.DocumentType =
                           DocumentTypeEnum.
                           kAssemblyDocumentObject Then

                        Dim subAsm As AssemblyDocument =
                            TryCast(doc, AssemblyDocument)

                        If subAsm IsNot Nothing Then

                            CollectParts(
                                subAsm.ComponentDefinition,
                                parts,
                                partType)

                        End If

                    End If

                Catch
                End Try

            Next

        End Sub

    End Module

End Namespace

