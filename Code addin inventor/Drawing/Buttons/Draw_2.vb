Option Explicit On
Option Strict Off

Imports System.Windows.Forms
Imports Inventor

Namespace ThanhN.Drawing.Buttons
    Public Module Draw_2
        Public Sub OnExecute(ByVal Context As NameValueMap)


            Dim app As Inventor.Application = g_inventorApplication

            Try
                If app.ActiveDocument Is Nothing OrElse
                       app.ActiveDocument.DocumentType <> Inventor.DocumentTypeEnum.kDrawingDocumentObject Then
                    MessageBox.Show("Vui lòng mở file Drawing (.idw)!", "Lỗi",
                                        MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                Dim oDrawDoc As Inventor.DrawingDocument =
                        CType(app.ActiveDocument, Inventor.DrawingDocument)

                Dim oSheet As Inventor.Sheet = oDrawDoc.ActiveSheet
                Dim tg As Inventor.TransientGeometry = app.TransientGeometry

                Dim countDia As Integer = 0
                Dim countThread As Integer = 0
                Dim countFail As Integer = 0

                For Each oView As Inventor.DrawingView In oSheet.DrawingViews
                    Try
                        If oView.DrawingCurves Is Nothing Then Continue For

                        For Each oCurve As Inventor.DrawingCurve In oView.DrawingCurves
                            Try
                                If oCurve.CurveType <> Inventor.CurveTypeEnum.kCircleCurve Then Continue For

                                Dim oIntent As Inventor.GeometryIntent =
                                        oSheet.CreateGeometryIntent(oCurve, Inventor.PointIntentEnum.kCircularLeftPointIntent)

                                Dim oPoint As Inventor.Point2d = oIntent.PointOnSheet.Copy()
                                Dim oVector As Inventor.Vector2d = oCurve.CenterPoint.VectorTo(oPoint)
                                oVector.ScaleBy(0.3)
                                oVector.AddVector(tg.CreateVector2d(oVector.X, System.Math.Abs(oVector.X)))
                                oPoint.TranslateBy(oVector)

                                Dim isThread As Boolean = IsThreadedHoleCurve(oCurve)

                                If isThread Then
                                    ' Lỗ ren → Hole / Thread Note
                                    Try
                                        oSheet.DrawingNotes.HoleThreadNotes.Add(oPoint, oIntent)
                                        countThread += 1
                                    Catch
                                        ' Fallback diameter
                                        Try
                                            oSheet.DrawingDimensions.GeneralDimensions.AddDiameter(
                                                    oPoint, oIntent, True, False, False)
                                            countDia += 1
                                        Catch
                                            countFail += 1
                                        End Try
                                    End Try
                                Else
                                    ' Lỗ thường → Diameter
                                    Try
                                        oSheet.DrawingDimensions.GeneralDimensions.AddDiameter(
                                                oPoint, oIntent, True, False, False)
                                        countDia += 1
                                    Catch
                                        countFail += 1
                                    End Try
                                End If

                            Catch
                                countFail += 1
                            End Try
                        Next
                    Catch
                    End Try
                Next

                oDrawDoc.Update()

                MessageBox.Show(
                        "Hoàn tất!" & vbCrLf & vbCrLf &
                        "Diameter (lỗ thường): " & countDia.ToString() & vbCrLf &
                        "Hole/Thread Note (lỗ ren): " & countThread.ToString() & vbCrLf &
                        "Lỗi / bỏ qua: " & countFail.ToString(),
                        "Dim lỗ",
                        MessageBoxButtons.OK, MessageBoxIcon.Information)

            Catch ex As Exception
                MessageBox.Show("Lỗi:" & vbCrLf & ex.Message, "Dim lỗ",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try

        End Sub

        '=================================================
        ' Kiểm tra DrawingCurve có phải lỗ REN không
        '=================================================
        Private Function IsThreadedHoleCurve(oCurve As Inventor.DrawingCurve) As Boolean
            Try
                Dim modelGeom As Object = oCurve.ModelGeometry
                If modelGeom Is Nothing Then Return False

                If TypeOf modelGeom Is Inventor.Edge Then
                    Dim ed As Inventor.Edge = CType(modelGeom, Inventor.Edge)

                    For Each fc As Inventor.Face In ed.Faces
                        ' ThreadInfos trên face
                        Try
                            If fc.ThreadInfos IsNot Nothing AndAlso fc.ThreadInfos.Count > 0 Then
                                Return True
                            End If
                        Catch
                        End Try

                        ' Feature tạo face
                        Try
                            Dim feat As Inventor.PartFeature = fc.CreatedByFeature
                            If feat Is Nothing Then Continue For

                            Dim tName As String = TypeName(feat).ToUpperInvariant()

                            If tName.Contains("THREAD") Then Return True

                            If tName.Contains("HOLE") Then
                                Try
                                    Dim hf As Inventor.HoleFeature = CType(feat, Inventor.HoleFeature)
                                    ' Tapped / threaded hole
                                    If hf.Tapped Then Return True
                                Catch
                                End Try
                            End If
                        Catch
                        End Try
                    Next
                End If

            Catch
            End Try

            Return False
        End Function

    End Module

End Namespace
