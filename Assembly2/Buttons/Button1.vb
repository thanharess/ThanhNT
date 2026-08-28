Imports Inventor
Imports System.Windows.Forms
Imports System.Runtime.InteropServices

Namespace ThanhN.Assembly2.Buttons
    Public Module Button1
        Public Sub OnExecute(ByVal Context As NameValueMap)
            Try
                ' Converted from iLogic macro: create sheet metal parts from an Excel sheet
                Dim i As Integer
                Dim staaad As Integer = 0
                Dim Quydoi As Double
                Dim myparam As String
                myparam = Microsoft.VisualBasic.Interaction.InputBox("file part", "folder", g_inventorApplication.DesignProjectManager.ActiveDesignProject.WorkspacePath)
                If myparam = "" Then Exit Sub

                ' Show standard OpenFileDialog to pick the Excel file
                Dim ofd As New OpenFileDialog()
                ofd.Filter = "Excel (*.xls;*.xlsx)|*.xls;*.xlsx"
                Try
                    ofd.InitialDirectory = g_inventorApplication.DesignProjectManager.ActiveDesignProject.WorkspacePath
                Catch
                End Try
                If ofd.ShowDialog() <> DialogResult.OK Then Return
                Dim ClintBrown3D As String = ofd.FileName
                If String.IsNullOrEmpty(ClintBrown3D) Then Return

                ' Variables used when reading Excel
                Dim cdcd As Double
                Dim crcr As Double
                Dim cdaycday As Double
                Dim thongtin As String
                Dim PL As String
                Dim vlvl As String
                Dim ten As String
                Dim SLSLSl As Double

                For i = 3 To 10 Step 1
                    PL = GoExcel.CellValue(ClintBrown3D, "bth", "b" & i)
                    SLSLSl = GoExcel.CellValue(ClintBrown3D, "bth", "c" & i)
                    cdcd = GoExcel.CellValue(ClintBrown3D, "bth", "d" & i)
                    crcr = GoExcel.CellValue(ClintBrown3D, "bth", "e" & i)
                    cdaycday = GoExcel.CellValue(ClintBrown3D, "bth", "f" & i)
                    thongtin = GoExcel.CellValue(ClintBrown3D, "bth", "g" & i)
                    vlvl = GoExcel.CellValue(ClintBrown3D, "bth", "h" & i)
                    ten = GoExcel.CellValue(ClintBrown3D, "bth", "j" & i)

                    If thongtin = "mm" Then
                        Quydoi = 10
                    ElseIf thongtin = "cm" Then
                        Quydoi = 1
                    ElseIf thongtin = "m" Then
                        Quydoi = 0.01
                    Else
                        Quydoi = 1
                    End If

                    If cdcd = 0 Or SLSLSl = 0 Then Exit For

                    ' Create a new sheet metal document using default sheet metal template
                    Dim oSheetMetalDoc As PartDocument
                    oSheetMetalDoc = g_inventorApplication.Documents.Add(Inventor.DocumentTypeEnum.kPartDocumentObject,
                        g_inventorApplication.FileManager.GetTemplateFile(Inventor.DocumentTypeEnum.kPartDocumentObject, , , "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}"), True)

                    Dim oCompDef As SheetMetalComponentDefinition = oSheetMetalDoc.ComponentDefinition
                    Dim oSheetMetalFeatures As SheetMetalFeatures = oCompDef.Features

                    ' Create sketch on X-Y work plane
                    Dim oSketch As PlanarSketch = oCompDef.Sketches.Add(oCompDef.WorkPlanes.Item(3))

                    ' Override the thickness for the document
                    oSheetMetalDoc.ComponentDefinition.UseSheetMetalStyleThickness = False
                    Dim oThicknessParam As Parameter = oSheetMetalDoc.ComponentDefinition.Thickness
                    oThicknessParam.Value = cdaycday / Quydoi

                    If PL = "HT" Then
                        ' Draw circle by center point
                        Dim dkkk As SketchCircle
                        dkkk = oSketch.SketchCircles.AddByCenterRadius(g_inventorApplication.TransientGeometry.CreatePoint2d(0, 0), cdcd / (2 * Quydoi))
                        Dim arcConstraint As DiameterDimConstraint = oSketch.DimensionConstraints.AddDiameter(dkkk, g_inventorApplication.TransientGeometry.CreatePoint2d(-0.5 / Quydoi, -0.5 / Quydoi))
                        Dim oProfile As Profile = oSketch.Profiles.AddForSolid
                        Dim oFaceFeatureDefinition As FaceFeatureDefinition = oSheetMetalFeatures.FaceFeatures.CreateFaceFeatureDefinition(oProfile)
                        Dim oFaceFeature As FaceFeature = oSheetMetalFeatures.FaceFeatures.Add(oFaceFeatureDefinition)

                        If crcr <> 0 Then
                            Dim oCutDefinition As CutDefinition
                            Dim oFrontFace As Face = oFaceFeature.Faces.Item(3)
                            oSketch = oCompDef.Sketches.AddWithOrientation(oFrontFace, oCompDef.WorkAxes.Item(1), True, True, oCompDef.WorkPoints(1))
                            dkkk = oSketch.SketchCircles.AddByCenterRadius(g_inventorApplication.TransientGeometry.CreatePoint2d(0, 0), crcr / (2 * Quydoi))
                            arcConstraint = oSketch.DimensionConstraints.AddDiameter(dkkk, g_inventorApplication.TransientGeometry.CreatePoint2d(-0.5 / Quydoi, -0.5 / Quydoi))
                            oProfile = oSketch.Profiles.AddForSolid
                            oCutDefinition = oSheetMetalFeatures.CutFeatures.CreateCutDefinition(oProfile)
                            oSheetMetalFeatures.CutFeatures.Add(oCutDefinition)
                        End If

                    ElseIf PL = "HCN" Then
                        ' Draw centered rectangle
                        Dim rectangleLines As SketchEntitiesEnumerator
                        rectangleLines = oSketch.SketchLines.AddAsTwoPointCenteredRectangle(g_inventorApplication.TransientGeometry.CreatePoint2d(0, 0), g_inventorApplication.TransientGeometry.CreatePoint2d(cdcd / (2 * Quydoi), crcr / (2 * Quydoi)))
                        Dim recSketchLine As SketchLine = rectangleLines.Item(1)
                        oSketch.DimensionConstraints.AddTwoPointDistance(recSketchLine.StartSketchPoint, recSketchLine.EndSketchPoint, DimensionOrientationEnum.kHorizontalDim, g_inventorApplication.TransientGeometry.CreatePoint2d(0, -crcr / (2 * Quydoi) - 3 / Quydoi))
                        Dim recSketchLine2 As SketchLine = rectangleLines.Item(4)
                        oSketch.DimensionConstraints.AddTwoPointDistance(recSketchLine2.StartSketchPoint, recSketchLine2.EndSketchPoint, DimensionOrientationEnum.kVerticalDim, g_inventorApplication.TransientGeometry.CreatePoint2d(-cdcd / (2 * Quydoi) - 3 / Quydoi, 0))
                        Dim oProfile As Profile = oSketch.Profiles.AddForSolid
                        Dim oFaceFeatureDefinition As FaceFeatureDefinition = oSheetMetalFeatures.FaceFeatures.CreateFaceFeatureDefinition(oProfile)
                        Dim oFaceFeature As FaceFeature = oSheetMetalFeatures.FaceFeatures.Add(oFaceFeatureDefinition)
                    Else
                        ' Default: centered rectangle using cdcd x crcr
                        Dim rectangleLines As SketchEntitiesEnumerator
                        rectangleLines = oSketch.SketchLines.AddAsTwoPointCenteredRectangle(g_inventorApplication.TransientGeometry.CreatePoint2d(0, 0), g_inventorApplication.TransientGeometry.CreatePoint2d(cdcd / (2 * Quydoi), crcr / (2 * Quydoi)))
                        Dim oProfile As Profile = oSketch.Profiles.AddForSolid
                        Dim oFaceFeatureDefinition As FaceFeatureDefinition = oSheetMetalFeatures.FaceFeatures.CreateFaceFeatureDefinition(oProfile)
                        Dim oFaceFeature As FaceFeature = oSheetMetalFeatures.FaceFeatures.Add(oFaceFeatureDefinition)
                    End If

                    ' Optionally save the document using the 'ten' from Excel (if provided)
                    If Not String.IsNullOrEmpty(ten) Then
                        Try
                            Dim savePath As String = System.IO.Path.Combine(myparam, ten & ".ipt")
                            oSheetMetalDoc.SaveAs(savePath, False)
                        Catch
                        End Try
                    End If
                Next

                MessageBox.Show("Operation completed.", "Button1", MessageBoxButtons.OK, MessageBoxIcon.Information)

            Catch ex As Exception
                Try
                    g_inventorApplication.UserInterfaceManager.UserInteractionManager.PostStatus("Error in Button1: " & ex.Message)
                Catch
                End Try
            End Try
        End Sub
    End Module

    ' Late-bound Excel helper to read single cell values without requiring an Interop reference
    Public Module GoExcel
        Public Function CellValue(ByVal workbookPath As String, ByVal sheetName As String, ByVal cellAddress As String) As Object
            Dim excel As Object = Nothing
            Dim wb As Object = Nothing
            Dim ws As Object = Nothing
            Dim result As Object = Nothing
            Try
                excel = CreateObject("Excel.Application")
                excel.Visible = False
                wb = excel.Workbooks.Open(workbookPath, ReadOnly:=True)
                ws = wb.Worksheets(sheetName)
                result = ws.Range(cellAddress).Value
            Catch ex As Exception
                result = Nothing
            Finally
                Try
                    If wb IsNot Nothing Then
                        wb.Close(False)
                        Marshal.ReleaseComObject(wb)
                    End If
                    If excel IsNot Nothing Then
                        excel.Quit()
                        Marshal.ReleaseComObject(excel)
                    End If
                Catch
                End Try
                ws = Nothing
                wb = Nothing
                excel = Nothing
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try
            Return result
        End Function
    End Module

End Namespace
