Imports System.Collections.Generic
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Inventor
Imports IO = System.IO

Namespace ToolInventor2020.Part.Buttons
    Public Module button15

        Public Sub OnExecute(ByVal Context As NameValueMap)

            Dim invApp As Inventor.Application = CType(Marshal.GetActiveObject("Inventor.Application"), Inventor.Application)

            '=====================================================
            ' 1. Kiểm tra đang ở Assembly
            '=====================================================
            If invApp.ActiveDocument Is Nothing OrElse
               invApp.ActiveDocument.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then

                MessageBox.Show("Mở Assembly và chọn Part / Sub-Assembly trước!", "Thông báo",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Exit Sub
            End If

            Dim asmDoc As AssemblyDocument = CType(invApp.ActiveDocument, AssemblyDocument)

            '=====================================================
            ' 2. Lấy các component đang được chọn
            '=====================================================
            Dim selectedDocs As New Dictionary(Of String, Document)

            Dim selSet As SelectSet = asmDoc.SelectSet

            If selSet.Count = 0 Then
                MessageBox.Show("Hãy chọn ít nhất 1 Part hoặc Sub-Assembly!", "Thông báo",
                                MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If

            For Each obj As Object In selSet
                Try
                    Dim occ As ComponentOccurrence = Nothing

                    If TypeOf obj Is ComponentOccurrence Then
                        occ = CType(obj, ComponentOccurrence)
                    ElseIf TypeOf obj Is ComponentOccurrenceProxy Then
                        occ = CType(obj, ComponentOccurrenceProxy).NativeObject
                    End If

                    If occ Is Nothing OrElse occ.Suppressed Then Continue For
                    If occ.DefinitionDocumentType <> DocumentTypeEnum.kPartDocumentObject AndAlso
                       occ.DefinitionDocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then Continue For

                    Dim doc As Document = occ.Definition.Document
                    If doc Is Nothing OrElse String.IsNullOrEmpty(doc.FullFileName) Then Continue For

                    Dim fullPath As String = IO.Path.GetFullPath(doc.FullFileName)
                    If Not selectedDocs.ContainsKey(fullPath) Then
                        selectedDocs.Add(fullPath, doc)
                    End If
                Catch
                End Try
            Next

            If selectedDocs.Count = 0 Then
                MessageBox.Show("Không có Part / Sub-Assembly hợp lệ trong lựa chọn.", "Thông báo",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Exit Sub
            End If

            '=====================================================
            ' 3. Nhập PREFIX + SUFFIX để tự sửa tên file
            '=====================================================
            Dim PREFIX As String = InputBox("Nhập PREFIX (để trống nếu không cần)", "PREFIX", "")
            Dim SUFFIX As String = InputBox("Nhập SUFFIX (để trống nếu không cần)", "SUFFIX", "")

            '=====================================================
            ' 4. Hỏi chế độ xuất cho Sub-Assembly
            '=====================================================
            Dim hasAssembly As Boolean = False
            For Each d As Document In selectedDocs.Values
                If d.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                    hasAssembly = True
                    Exit For
                End If
            Next

            Dim exportAsmAsPart As Boolean = False

            If hasAssembly Then
                Dim ans As DialogResult = MessageBox.Show(
                    "Trong lựa chọn có Sub-Assembly." & vbCrLf & vbCrLf &
                    "Bạn muốn xuất Sub-Assembly dưới dạng nào?" & vbCrLf & vbCrLf &
                    "YES  = Xuất dạng PART (Derived single body)" & vbCrLf &
                    "NO   = Xuất dạng CỤM LẮP (Assembly)",
                    "Chế độ xuất Sub-Assembly",
                    MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)

                If ans = DialogResult.Cancel Then Exit Sub
                exportAsmAsPart = (ans = DialogResult.Yes)
            End If

            '=====================================================
            ' 5. Chọn thư mục lưu
            '=====================================================
            Dim settingFile As String = ""
            Dim lastFolder As String = ""
            Dim outputFolder As String = ""

            Try
                Dim appData As String = System.Environment.GetFolderPath(System.Environment.SpecialFolder.ApplicationData)
                settingFile = IO.Path.Combine(appData, "Inventor_iLogic_STEP_Export_Selected.txt")

                If IO.File.Exists(settingFile) Then
                    Dim lines() As String = IO.File.ReadAllLines(settingFile)
                    If lines.Length > 0 Then lastFolder = lines(0)
                End If
            Catch
            End Try

            Try
                Dim folderDlg As New FolderBrowserDialog()
                folderDlg.Description = "Chọn thư mục lưu STEP AP214"
                folderDlg.ShowNewFolderButton = True

                If lastFolder <> "" AndAlso IO.Directory.Exists(lastFolder) Then
                    folderDlg.SelectedPath = lastFolder
                End If

                If folderDlg.ShowDialog() <> DialogResult.OK Then Exit Sub

                outputFolder = folderDlg.SelectedPath

                Try
                    IO.File.WriteAllText(settingFile, outputFolder)
                Catch
                End Try
            Catch ex As Exception
                MessageBox.Show("Lỗi chọn folder: " & ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End Try

            '=====================================================
            ' 6. Lấy STEP Translator
            '=====================================================
            Dim stepTranslator As TranslatorAddIn = Nothing
            Try
                stepTranslator = CType(invApp.ApplicationAddIns.ItemById("{90AF7F40-0C01-11D5-8E83-0010B541CD80}"), TranslatorAddIn)
            Catch
            End Try

            If stepTranslator Is Nothing Then
                MessageBox.Show("Không tìm thấy STEP Translator Add-in.", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            '=====================================================
            ' 7. Xuất từng file + áp dụng PREFIX / SUFFIX
            '=====================================================
            Dim successCount As Integer = 0
            Dim failCount As Integer = 0
            Dim tempFiles As New List(Of String)

            For Each kvp As KeyValuePair(Of String, Document) In selectedDocs
                Dim doc As Document = kvp.Value
                Dim srcPath As String = kvp.Key
                Dim exportDoc As Document = doc
                Dim isTemp As Boolean = False

                Try
                    ' Nếu là Assembly và chọn chế độ Part → tạo Derived tạm
                    If doc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject AndAlso exportAsmAsPart Then
                        Dim tempPartPath As String = IO.Path.Combine(IO.Path.GetTempPath(),
                            "TEMP_DERIVED_" & IO.Path.GetFileNameWithoutExtension(srcPath) & "_" & Guid.NewGuid().ToString("N").Substring(0, 8) & ".ipt")

                        Dim templatePath As String = invApp.FileManager.GetTemplateFile(DocumentTypeEnum.kPartDocumentObject)
                        Dim tempPart As PartDocument = CType(invApp.Documents.Add(DocumentTypeEnum.kPartDocumentObject, templatePath, True), PartDocument)
                        tempPart.SaveAs(tempPartPath, False)

                        Dim derivedDef As DerivedAssemblyDefinition = tempPart.ComponentDefinition.ReferenceComponents.DerivedAssemblyComponents.CreateDefinition(srcPath)
                        derivedDef.DeriveStyle = DerivedComponentStyleEnum.kDeriveAsSingleBodyNoSeams
                        tempPart.ComponentDefinition.ReferenceComponents.DerivedAssemblyComponents.Add(derivedDef)

                        tempPart.Update()
                        tempPart.Save2(True)

                        exportDoc = tempPart
                        isTemp = True
                        tempFiles.Add(tempPartPath)
                    End If

                    ' ===== TỰ SỬA TÊN FILE =====
                    Dim originalName As String = IO.Path.GetFileNameWithoutExtension(srcPath)
                    Dim newName As String = PREFIX & originalName & SUFFIX
                    If isTemp Then newName &= "_PART"

                    Dim stepFullPath As String = IO.Path.Combine(outputFolder, newName & ".stp")

                    If IO.File.Exists(stepFullPath) Then
                        Dim ans As DialogResult = MessageBox.Show(
                            "File đã tồn tại:" & vbCrLf & stepFullPath & vbCrLf & vbCrLf & "Ghi đè?",
                            "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                        If ans = DialogResult.No Then
                            failCount += 1
                            Continue For
                        End If
                    End If

                    ' Xuất STEP AP214
                    Dim oContext As TranslationContext = invApp.TransientObjects.CreateTranslationContext()
                    oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism

                    Dim oOptions As NameValueMap = invApp.TransientObjects.CreateNameValueMap()

                    If stepTranslator.HasSaveCopyAsOptions(exportDoc, oContext, oOptions) Then
                        oOptions.Value("ApplicationProtocolType") = 3
                    End If

                    Dim oData As DataMedium = invApp.TransientObjects.CreateDataMedium()
                    oData.FileName = stepFullPath

                    stepTranslator.SaveCopyAs(exportDoc, oContext, oOptions, oData)

                    successCount += 1

                Catch ex As Exception
                    failCount += 1
                    MessageBox.Show("Lỗi xuất:" & vbCrLf & srcPath & vbCrLf & vbCrLf & ex.Message,
                                    "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Finally
                    If isTemp AndAlso exportDoc IsNot Nothing Then
                        Try : exportDoc.Close(True) : Catch : End Try
                    End If
                End Try
            Next

            ' Xóa file tạm
            For Each tmp As String In tempFiles
                Try
                    If IO.File.Exists(tmp) Then IO.File.Delete(tmp)
                Catch
                End Try
            Next

            '=====================================================
            ' 8. Kết quả
            '=====================================================
            Dim modeText As String = If(exportAsmAsPart, "Part (Derived)", "Cụm lắp (Assembly)")
            MessageBox.Show(
                "XUẤT STEP AP214 HOÀN TẤT!" & vbCrLf & vbCrLf &
                "Số lượng chọn : " & selectedDocs.Count.ToString() & vbCrLf &
                "Thành công    : " & successCount.ToString() & vbCrLf &
                "Thất bại      : " & failCount.ToString() & vbCrLf &
                "Chế độ cụm    : " & modeText & vbCrLf &
                "PREFIX        : " & PREFIX & vbCrLf &
                "SUFFIX        : " & SUFFIX & vbCrLf & vbCrLf &
                "Folder lưu    : " & outputFolder,
                "Kết quả", MessageBoxButtons.OK, MessageBoxIcon.Information)

        End Sub
    End Module
End Namespace