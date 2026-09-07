Imports System.Windows.Forms
Imports System.Collections.Generic
Imports System.IO
Imports Inventor

Namespace ToolInventor2020.Assembly.Buttons.Part

    Public Class Ass_Part_9

        ' ========== CẤU HÌNH DỄ SỬA ==========
        Private Const USE_3_CHARS As Boolean = True
        Private Const DEFAULT_MATERIAL As String = "Steel"          ' <-- Sửa vật liệu mặc định tại đây
        Private Const SHEETMETAL_MATERIAL As String = "Steel"       ' <-- Sửa vật liệu Sheet Metal tại đây
        Private Const SKIP_SUPPRESSED As Boolean = True
        Private Const SKIP_CONTENT_CENTER As Boolean = True

        '=====================================================
        ' ENTRY POINT - Gọi từ button
        '=====================================================
        Public Shared Sub OnExecute(ByVal Context As NameValueMap)

            Try
                ' Lấy Application từ Add-in (cách an toàn nhất)
                Dim invApp As Inventor.Application = GetInventorApp()

                If invApp Is Nothing Then
                    MessageBox.Show("Không lấy được Inventor Application!", "Change Material")
                    Return
                End If

                If invApp.ActiveDocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
                    MessageBox.Show("Mở file Assembly trước", "Change Material")
                    Return
                End If

                Dim asmDoc As AssemblyDocument = CType(invApp.ActiveDocument, AssemblyDocument)

                ' ========== BẢNG LỌC PREFIX → VẬT LIỆU ==========
                Dim matMap As New Dictionary(Of String, String)

                matMap.Add("ST", "Steel")
                matMap.Add("AL", "Aluminum 6061")
                matMap.Add("SS", "Stainless Steel")
                matMap.Add("CU", "Copper")
                matMap.Add("BR", "Brass")

                matMap.Add("MID", "Mild Steel")
                matMap.Add("STL", "Steel")
                matMap.Add("S45", "Steel")
                matMap.Add("SUS", "Stainless Steel")
                matMap.Add("AL6", "Aluminum 6061")
                matMap.Add("PVC", "PVC")
                matMap.Add("NYL", "Nylon")
                matMap.Add("ABS", "ABS Plastic")

                ' Thêm dòng mới tại đây nếu cần
                ' matMap.Add("XXX", "Tên vật liệu")

                invApp.SilentOperation = True

                Dim countChanged As Integer = 0
                Dim countSkipped As Integer = 0

                ProcessAssembly(invApp, asmDoc, matMap, countChanged, countSkipped)

                asmDoc.Update2(True)
                asmDoc.Save2(True)

                invApp.SilentOperation = False

                MessageBox.Show(
                    "Hoàn tất!" & vbCrLf &
                    "Đã thay vật liệu: " & countChanged & vbCrLf &
                    "Bỏ qua: " & countSkipped,
                    "Change Material")

            Catch ex As Exception
                MessageBox.Show("Lỗi: " & ex.Message, "Change Material")
            End Try

        End Sub

        '=====================================================
        ' LẤY INVENTOR APPLICATION (an toàn cho Add-in)
        '=====================================================
        Private Shared Function GetInventorApp() As Inventor.Application
            Try
                ' Cách 1: Từ Add-in server (nếu bạn có biến global)
                ' Return StandardAddInServer.m_inventorApplication

                ' Cách 2: Lấy từ Running Object Table (luôn hoạt động)
                Dim invApp As Inventor.Application = Nothing
                invApp = CType(System.Runtime.InteropServices.Marshal.GetActiveObject("Inventor.Application"), Inventor.Application)
                Return invApp
            Catch
                Return Nothing
            End Try
        End Function

        '=====================================================
        ' PROCESS ASSEMBLY
        '=====================================================
        Private Shared Sub ProcessAssembly(
            invApp As Inventor.Application,
            asmDoc As AssemblyDocument,
            matMap As Dictionary(Of String, String),
            ByRef countChanged As Integer,
            ByRef countSkipped As Integer)

            For Each occ As ComponentOccurrence In asmDoc.ComponentDefinition.Occurrences

                Try
                    If SKIP_SUPPRESSED AndAlso occ.Suppressed Then
                        countSkipped += 1
                        Continue For
                    End If

                    If SKIP_CONTENT_CENTER AndAlso occ.IsContentMember Then
                        countSkipped += 1
                        Continue For
                    End If

                    Dim doc As Document = occ.Definition.Document

                    If doc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then

                        Dim partDoc As PartDocument = CType(doc, PartDocument)
                        Dim isSheetMetal As Boolean = False

                        Try
                            If partDoc.SubType = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}" OrElse
                               TypeOf partDoc.ComponentDefinition Is SheetMetalComponentDefinition Then
                                isSheetMetal = True
                            End If
                        Catch
                        End Try

                        Dim targetMat As String = ""

                        If isSheetMetal Then
                            targetMat = SHEETMETAL_MATERIAL
                        Else
                            Dim pn As String = ""
                            Try
                                Dim designProps As PropertySet = partDoc.PropertySets.Item("Design Tracking Properties")
                                pn = designProps.Item("Part Number").Value.ToString().Trim().ToUpper()
                            Catch
                                pn = System.IO.Path.GetFileNameWithoutExtension(partDoc.FullFileName).ToUpper()
                            End Try

                            Dim prefix As String = ""
                            If USE_3_CHARS Then
                                If pn.Length >= 3 Then
                                    prefix = pn.Substring(0, 3)
                                ElseIf pn.Length >= 2 Then
                                    prefix = pn.Substring(0, 2)
                                End If
                            Else
                                If pn.Length >= 2 Then
                                    prefix = pn.Substring(0, 2)
                                End If
                            End If

                            If matMap.ContainsKey(prefix) Then
                                targetMat = matMap(prefix)
                            Else
                                targetMat = DEFAULT_MATERIAL
                            End If
                        End If

                        If targetMat <> "" Then
                            If SetMaterial(invApp, partDoc, targetMat) Then
                                countChanged += 1
                            Else
                                countSkipped += 1
                            End If
                        End If

                    ElseIf doc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then

                        ProcessAssembly(invApp, CType(doc, AssemblyDocument), matMap, countChanged, countSkipped)

                    End If

                Catch
                    countSkipped += 1
                End Try

            Next

        End Sub

        '=====================================================
        ' SET MATERIAL
        '=====================================================
        Private Shared Function SetMaterial(
            invApp As Inventor.Application,
            partDoc As PartDocument,
            materialName As String) As Boolean

            Try
                If partDoc.ActiveMaterial IsNot Nothing Then
                    If String.Compare(partDoc.ActiveMaterial.DisplayName, materialName, True) = 0 Then
                        Return True
                    End If
                End If

                Dim matAsset As Inventor.MaterialAsset = Nothing

                ' 1. Tìm trong document
                Try
                    Dim localAssets As Inventor.AssetsEnumerator = partDoc.MaterialAssets
                    Dim i As Integer
                    For i = 1 To localAssets.Count
                        Dim a As Inventor.Asset = localAssets.Item(i)
                        If String.Compare(a.DisplayName, materialName, True) = 0 Then
                            matAsset = CType(a, Inventor.MaterialAsset)
                            Exit For
                        End If
                    Next
                Catch
                End Try

                ' 2. Tìm trong Asset Libraries
                If matAsset Is Nothing Then
                    Try
                        Dim libs As Inventor.AssetLibraries = invApp.AssetLibraries
                        Dim j As Integer
                        For j = 1 To libs.Count
                            Dim assetLib As Inventor.AssetLibrary = libs.Item(j)

                            Try
                                Dim matAssets As Inventor.AssetsEnumerator = assetLib.MaterialAssets
                                Dim k As Integer
                                For k = 1 To matAssets.Count
                                    Dim a As Inventor.Asset = matAssets.Item(k)
                                    If String.Compare(a.DisplayName, materialName, True) = 0 Then
                                        matAsset = CType(a.CopyTo(partDoc), Inventor.MaterialAsset)
                                        Exit For
                                    End If
                                Next
                            Catch
                            End Try

                            If matAsset IsNot Nothing Then Exit For
                        Next
                    Catch
                    End Try
                End If

                If matAsset Is Nothing Then
                    Return False
                End If

                partDoc.ActiveMaterial = matAsset
                partDoc.Update2(True)
                partDoc.Save2(True)

                Return True

            Catch
                Return False
            End Try

        End Function

    End Class

End Namespace