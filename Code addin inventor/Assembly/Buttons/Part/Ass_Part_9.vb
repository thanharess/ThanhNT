Imports System.Windows.Forms
Imports System.Collections.Generic
Imports System.IO
Imports Inventor

Namespace ToolInventor2020.Assembly.Buttons.Part

    Public Class Ass_Part_9

        ' ========== CẤU HÌNH DỄ SỬA ==========
        Private Const USE_3_CHARS As Boolean = True
        Private Const DEFAULT_MATERIAL As String = "Steel"
        Private Const SHEETMETAL_MATERIAL As String = "Steel"
        Private Const SKIP_SUPPRESSED As Boolean = True
        Private Const SKIP_CONTENT_CENTER As Boolean = True

        ' ⚠️ Có đổi Appearance không?
        Private Const CHANGE_APPEARANCE As Boolean = True

        ' ⚠️ Nếu Material không có trong bảng map → dùng appearance mặc định
        Private Const DEFAULT_APPEARANCE As String = "Semi-Polished"

        '=====================================================
        ' ENTRY POINT
        '=====================================================
        Public Shared Sub OnExecute(ByVal Context As NameValueMap)
            Try
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

                ' ========== BẢNG MAP VẬT LIỆU → MÀU (APPEARANCE) ==========
                ' Sửa theo bộ vật liệu/ngoại quan thực tế của bạn
                Dim matToAppearance As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
                matToAppearance("Steel") = "Semi-Polished"
                matToAppearance("Mild Steel") = "Semi-Polished(1)"
                matToAppearance("Stainless Steel") = "Semi-Polished(2)"
                matToAppearance("Aluminum 6061") = "Semi-Polished(3)"
                matToAppearance("Copper") = "Semi-Polished(4)"
                matToAppearance("Brass") = "Semi-Polished(5)"
                matToAppearance("PVC") = "Semi-Polished(6)"
                matToAppearance("Nylon") = "Semi-Polished(7)"
                matToAppearance("ABS Plastic") = "Semi-Polished(8)"


                ' ========== MAP CHO TỪNG STEEL MILD 1→25 ==========
                ' Nếu muốn mỗi Mild có màu riêng thì bật khối này:
                For i As Integer = 1 To 25
                    ' Ví dụ: Steel, Mild 1 → Semi-Polished(9)
                    '       Steel, Mild 2 → Semi-Polished(10)
                    '       ... đến hết 23, rồi quay lại từ đầu
                    Dim appIdx As Integer = 8 + i   ' 9, 10, ..., 33
                    If appIdx > 23 Then appIdx = ((appIdx - 1) Mod 23) + 1
                    matToAppearance("Steel, Mild " & i) = "Semi-Polished(" & appIdx & ")"
                Next
                matToAppearance("Steel, Mild") = "Semi-Polished"

                invApp.SilentOperation = True

                Dim countChanged As Integer = 0
                Dim countSkipped As Integer = 0
                Dim countNotFound As Integer = 0
                Dim notFoundList As New List(Of String)

                ProcessAssembly(invApp, asmDoc, matMap, matToAppearance,
                                countChanged, countSkipped, countNotFound, notFoundList)

                asmDoc.Update2(True)
                ' asmDoc.Save2(True)

                invApp.SilentOperation = False

                Dim msg As String = "Hoàn tất!" & vbCrLf &
                                    "Đã thay vật liệu + màu: " & countChanged & vbCrLf &
                                    "Bỏ qua (suppress/content): " & countSkipped & vbCrLf &
                                    "Không tìm thấy vật liệu: " & countNotFound

                If countNotFound > 0 AndAlso notFoundList.Count > 0 Then
                    msg &= vbCrLf & vbCrLf & "Danh sách part bị bỏ qua (tối đa 20):" & vbCrLf
                    Dim show As Integer = Math.Min(20, notFoundList.Count)
                    For i As Integer = 0 To show - 1
                        msg &= "  • " & notFoundList(i) & vbCrLf
                    Next
                    If notFoundList.Count > 20 Then
                        msg &= "  ... và " & (notFoundList.Count - 20) & " part khác"
                    End If
                End If

                MessageBox.Show(msg, "Change Material")

            Catch ex As Exception
                MessageBox.Show("Lỗi: " & ex.Message, "Change Material")
            End Try
        End Sub

        '=====================================================
        ' LẤY INVENTOR APPLICATION
        '=====================================================
        Private Shared Function GetInventorApp() As Inventor.Application
            Try
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
            matToAppearance As Dictionary(Of String, String),
            ByRef countChanged As Integer,
            ByRef countSkipped As Integer,
            ByRef countNotFound As Integer,
            ByRef notFoundList As List(Of String))

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
                            ' ===== XÁC ĐỊNH APPEARANCE MỤC TIÊU =====
                            Dim targetAppearance As String = ""
                            If CHANGE_APPEARANCE Then
                                If matToAppearance.ContainsKey(targetMat) Then
                                    targetAppearance = matToAppearance(targetMat)
                                Else
                                    targetAppearance = DEFAULT_APPEARANCE
                                End If
                            End If

                            Dim result As Integer = SetMaterialAndAppearance(
                                invApp, partDoc, targetMat, targetAppearance)

                            If result = 1 Then
                                countChanged += 1
                            ElseIf result = 0 Then
                                countNotFound += 1
                                Try
                                    Dim name As String = System.IO.Path.GetFileNameWithoutExtension(partDoc.FullFileName)
                                    If String.IsNullOrEmpty(name) Then name = occ.Name
                                    notFoundList.Add(name & " → " & targetMat)
                                Catch
                                    notFoundList.Add(occ.Name & " → " & targetMat)
                                End Try
                            Else
                                countSkipped += 1
                            End If
                        Else
                            countSkipped += 1
                        End If

                    ElseIf doc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                        ProcessAssembly(invApp, CType(doc, AssemblyDocument), matMap, matToAppearance,
                                        countChanged, countSkipped, countNotFound, notFoundList)
                    End If

                Catch
                    countSkipped += 1
                End Try
            Next
        End Sub

        '=====================================================
        ' SET MATERIAL + APPEARANCE
        ' Trả về:  1 = OK
        '          0 = không tìm thấy vật liệu (bỏ qua)
        '         -1 = lỗi khác
        '=====================================================
        Private Shared Function SetMaterialAndAppearance(
            invApp As Inventor.Application,
            partDoc As PartDocument,
            materialName As String,
            appearanceName As String) As Integer

            Try
                Dim changed As Boolean = False

                ' ---------- 1. ĐỔI VẬT LIỆU ----------
                Dim needMaterial As Boolean = True
                If partDoc.ActiveMaterial IsNot Nothing Then
                    If String.Compare(partDoc.ActiveMaterial.DisplayName, materialName, True) = 0 Then
                        needMaterial = False
                    End If
                End If

                If needMaterial Then
                    Dim matAsset As Inventor.MaterialAsset = Nothing

                    ' Tìm trong document
                    Try
                        For Each a As Inventor.Asset In partDoc.MaterialAssets
                            If String.Compare(a.DisplayName, materialName, True) = 0 Then
                                matAsset = TryCast(a, Inventor.MaterialAsset)
                                If matAsset IsNot Nothing Then Exit For
                            End If
                        Next
                    Catch
                    End Try

                    ' Tìm trong Libraries
                    If matAsset Is Nothing Then
                        Try
                            Dim libs As Inventor.AssetLibraries = invApp.AssetLibraries
                            For j As Integer = 1 To libs.Count
                                Dim assetLib As Inventor.AssetLibrary = libs.Item(j)
                                Try
                                    For Each a As Inventor.Asset In assetLib.MaterialAssets
                                        If String.Compare(a.DisplayName, materialName, True) = 0 Then
                                            Dim copied As Inventor.Asset = a.CopyTo(partDoc)
                                            matAsset = TryCast(copied, Inventor.MaterialAsset)
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

                    ' ⚠️ KHÔNG tìm thấy vật liệu → bỏ qua
                    If matAsset Is Nothing Then
                        Return 0
                    End If

                    partDoc.ActiveMaterial = matAsset
                    changed = True
                End If

                ' ---------- 2. ĐỔI APPEARANCE (MÀU) ----------
                If CHANGE_APPEARANCE AndAlso Not String.IsNullOrEmpty(appearanceName) Then
                    Dim needAppearance As Boolean = True
                    Try
                        If partDoc.ActiveAppearance IsNot Nothing Then
                            If String.Compare(partDoc.ActiveAppearance.DisplayName, appearanceName, True) = 0 Then
                                needAppearance = False
                            End If
                        End If
                    Catch
                    End Try

                    If needAppearance Then
                        Dim appAsset As Inventor.Asset = Nothing

                        ' Tìm trong document
                        Try
                            For Each a As Inventor.Asset In partDoc.AppearanceAssets
                                If String.Compare(a.DisplayName, appearanceName, True) = 0 Then
                                    appAsset = a
                                    Exit For
                                End If
                            Next
                        Catch
                        End Try

                        ' Tìm trong Libraries
                        If appAsset Is Nothing Then
                            Try
                                Dim libs As Inventor.AssetLibraries = invApp.AssetLibraries
                                For j As Integer = 1 To libs.Count
                                    Dim assetLib As Inventor.AssetLibrary = libs.Item(j)
                                    Try
                                        For Each a As Inventor.Asset In assetLib.AppearanceAssets
                                            If String.Compare(a.DisplayName, appearanceName, True) = 0 Then
                                                appAsset = a.CopyTo(partDoc)
                                                Exit For
                                            End If
                                        Next
                                    Catch
                                    End Try
                                    If appAsset IsNot Nothing Then Exit For
                                Next
                            Catch
                            End Try
                        End If

                        ' Không tìm thấy appearance → bỏ qua im lặng (vẫn coi như OK nếu material đổi được)
                        If appAsset IsNot Nothing Then
                            partDoc.ActiveAppearance = appAsset
                            changed = True
                        End If
                    End If
                End If

                ' ---------- 3. UPDATE ----------
                If changed Then
                    partDoc.Update2(True)
                    ' partDoc.Save2(True)
                End If

                Return 1

            Catch
                Return -1
            End Try
        End Function

    End Class

End Namespace