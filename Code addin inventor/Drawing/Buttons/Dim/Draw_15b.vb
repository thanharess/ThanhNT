Option Explicit On
Option Strict Off

Imports System.Windows.Forms
Imports Inventor
Imports System.Collections.Generic
Imports System.Linq

Namespace ToolInventor2020.Drawing.Buttons

    Public Module draw_15b

        Public Sub OnExecute(ByVal Context As NameValueMap)

            Dim app As Inventor.Application = g_inventorApplication

            Try
                If app.ActiveDocument Is Nothing OrElse
                   app.ActiveDocument.DocumentType <> DocumentTypeEnum.kDrawingDocumentObject Then
                    MessageBox.Show("Vui lòng mở file Drawing (.idw)!", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                Dim oDrawDoc As DrawingDocument = CType(app.ActiveDocument, DrawingDocument)
                Dim oSheet As Sheet = oDrawDoc.ActiveSheet
                Dim tg As TransientGeometry = app.TransientGeometry

                Dim countOK As Integer = 0
                Dim countFail As Integer = 0
                Dim countZero As Integer = 0

                Const TOL As Double = 0.3
                Const RATIO As Double = 0.55
                Const ARRAY_RADIUS As Double = 30.0

                For Each oView As DrawingView In oSheet.DrawingViews

                    '===== 1. Lấy đường tròn + lọc ARRAY =====
                    Dim allHoles As New List(Of HoleInfo)

                    For Each oCurve As DrawingCurve In oView.DrawingCurves
                        Try
                            If oCurve.CurveType <> CurveTypeEnum.kCircleCurve Then Continue For
                            Dim c As Point2d = oCurve.CenterPoint
                            If c Is Nothing Then Continue For

                            ' --- Kiểm tra có phải Circular Pattern không ---
                            Dim isArray As Boolean = False
                            Try
                                Dim modelGeom As Object = oCurve.ModelGeometry
                                If TypeOf modelGeom Is Edge Then
                                    Dim ed As Edge = CType(modelGeom, Edge)
                                    For Each fc As Face In ed.Faces
                                        Dim feat As PartFeature = fc.CreatedByFeature
                                        If feat IsNot Nothing Then
                                            Dim tName As String = TypeName(feat).ToUpperInvariant()
                                            If tName.Contains("CIRCULARPATTERN") OrElse tName.Contains("PATTERN") Then
                                                isArray = True
                                                Exit For
                                            End If
                                        End If
                                    Next
                                End If
                            Catch
                            End Try
                            If isArray Then Continue For

                            Dim hi As New HoleInfo
                            hi.Curve = oCurve
                            hi.Center = c
                            hi.X = c.X
                            hi.Y = c.Y
                            Try : hi.Radius = oCurve.Radius : Catch : hi.Radius = 1.5 : End Try
                            allHoles.Add(hi)
                        Catch
                        End Try
                    Next

                    ' Lọc thêm bằng mật độ (phòng trường hợp không lấy được feature)
                    Dim cleanHoles As New List(Of HoleInfo)
                    For Each h In allHoles
                        Dim nearby = allHoles.Where(Function(o) o IsNot h AndAlso
                            Math.Sqrt((o.X - h.X) ^ 2 + (o.Y - h.Y) ^ 2) < ARRAY_RADIUS).Count()
                        If nearby >= 3 Then Continue For
                        cleanHoles.Add(h)
                    Next
                    If cleanHoles.Count = 0 Then cleanHoles = allHoles

                    If cleanHoles.Count = 0 Then Continue For

                    '===== 2. Tìm 4 cạnh =====
                    Dim leftEdge, rightEdge, topEdge, bottomEdge As DrawingCurve
                    Dim minX As Double = Double.MaxValue, maxX As Double = Double.MinValue
                    Dim maxY As Double = Double.MinValue, minY As Double = Double.MaxValue

                    For Each oCurve As DrawingCurve In oView.DrawingCurves
                        Try
                            If oCurve.CurveType <> CurveTypeEnum.kLineCurve AndAlso
                               oCurve.CurveType <> CurveTypeEnum.kLineSegmentCurve Then Continue For
                            Dim p1 = oCurve.StartPoint, p2 = oCurve.EndPoint
                            If Math.Abs(p1.X - p2.X) < 0.03 Then
                                If p1.X < minX Then minX = p1.X : leftEdge = oCurve
                                If p1.X > maxX Then maxX = p1.X : rightEdge = oCurve
                            End If
                            If Math.Abs(p1.Y - p2.Y) < 0.03 Then
                                If p1.Y > maxY Then maxY = p1.Y : topEdge = oCurve
                                If p1.Y < minY Then minY = p1.Y : bottomEdge = oCurve
                            End If
                        Catch
                        End Try
                    Next

                    Dim viewW = maxX - minX
                    Dim viewH = maxY - minY
                    Dim limitX = minX + viewW * RATIO
                    Dim limitY = maxY - viewH * RATIO

                    '===== 3. Dim tổng =====
                    If leftEdge IsNot Nothing AndAlso rightEdge IsNot Nothing Then
                        Try
                            Dim tp = tg.CreatePoint2d((minX + maxX) / 2, maxY + 5.5)
                            oSheet.DrawingDimensions.GeneralDimensions.AddLinear(tp,
                                oSheet.CreateGeometryIntent(leftEdge), oSheet.CreateGeometryIntent(rightEdge),
                                DimensionTypeEnum.kHorizontalDimensionType)
                            countOK += 1
                        Catch : countFail += 1 : End Try
                    End If
                    If topEdge IsNot Nothing AndAlso bottomEdge IsNot Nothing Then
                        Try
                            Dim tp = tg.CreatePoint2d(minX - 5.5, (maxY + minY) / 2)
                            oSheet.DrawingDimensions.GeneralDimensions.AddLinear(tp,
                                oSheet.CreateGeometryIntent(topEdge), oSheet.CreateGeometryIntent(bottomEdge),
                                DimensionTypeEnum.kVerticalDimensionType)
                            countOK += 1
                        Catch : countFail += 1 : End Try
                    End If

                    '=====================================================
                    ' b1: Trái → Phải (phía trên) + đến mặt phải
                    '=====================================================
                    Dim topRow = cleanHoles.Where(Function(h) h.Y >= limitY).OrderBy(Function(h) h.X).ToList()
                    If topRow.Count > 0 AndAlso leftEdge IsNot Nothing Then
                        Dim prev = topRow.First()
                        ' cạnh trái → lỗ đầu
                        If (prev.X - minX) > TOL Then
                            Try
                                Dim tp = tg.CreatePoint2d((minX + prev.X) / 2, maxY + 3.2)
                                oSheet.DrawingDimensions.GeneralDimensions.AddLinear(tp,
                                    oSheet.CreateGeometryIntent(leftEdge),
                                    oSheet.CreateGeometryIntent(prev.Curve, PointIntentEnum.kCenterPointIntent),
                                    DimensionTypeEnum.kHorizontalDimensionType)
                                countOK += 1
                            Catch : countFail += 1 : End Try
                        End If
                        ' các lỗ giữa
                        For i = 1 To topRow.Count - 1
                            Dim h = topRow(i)
                            Dim dist = h.X - prev.X
                            If dist < TOL Then countZero += 1 : Continue For
                            Try
                                Dim tp = tg.CreatePoint2d((prev.X + h.X) / 2, maxY + 2.5)
                                oSheet.DrawingDimensions.GeneralDimensions.AddLinear(tp,
                                    oSheet.CreateGeometryIntent(prev.Curve, PointIntentEnum.kCenterPointIntent),
                                    oSheet.CreateGeometryIntent(h.Curve, PointIntentEnum.kCenterPointIntent),
                                    DimensionTypeEnum.kHorizontalDimensionType)
                                countOK += 1
                                prev = h
                            Catch : countFail += 1 : End Try
                        Next
                        ' lỗ cuối → mặt phải
                        If rightEdge IsNot Nothing AndAlso (maxX - prev.X) > TOL Then
                            Try
                                Dim tp = tg.CreatePoint2d((prev.X + maxX) / 2, maxY + 3.2)
                                oSheet.DrawingDimensions.GeneralDimensions.AddLinear(tp,
                                    oSheet.CreateGeometryIntent(prev.Curve, PointIntentEnum.kCenterPointIntent),
                                    oSheet.CreateGeometryIntent(rightEdge),
                                    DimensionTypeEnum.kHorizontalDimensionType)
                                countOK += 1
                            Catch : countFail += 1 : End Try
                        End If
                    End If

                    '=====================================================
                    ' b2: Trên → Dưới (bên trái) + đến mặt dưới
                    '=====================================================
                    Dim leftCol = cleanHoles.Where(Function(h) h.X <= limitX) _
                                            .OrderByDescending(Function(h) h.Y).ToList()
                    ' loại trùng Y
                    Dim uniqueLeft As New List(Of HoleInfo)
                    Dim lastY As Double = Double.MaxValue
                    For Each h In leftCol
                        If Math.Abs(h.Y - lastY) < TOL Then Continue For
                        uniqueLeft.Add(h)
                        lastY = h.Y
                    Next

                    If uniqueLeft.Count > 0 AndAlso topEdge IsNot Nothing Then
                        Dim prev = uniqueLeft.First()
                        If (maxY - prev.Y) > TOL Then
                            Try
                                Dim tp = tg.CreatePoint2d(minX - 3.2, (maxY + prev.Y) / 2)
                                oSheet.DrawingDimensions.GeneralDimensions.AddLinear(tp,
                                    oSheet.CreateGeometryIntent(topEdge),
                                    oSheet.CreateGeometryIntent(prev.Curve, PointIntentEnum.kCenterPointIntent),
                                    DimensionTypeEnum.kVerticalDimensionType)
                                countOK += 1
                            Catch : countFail += 1 : End Try
                        End If
                        For i = 1 To uniqueLeft.Count - 1
                            Dim h = uniqueLeft(i)
                            Dim dist = prev.Y - h.Y
                            If dist < TOL Then countZero += 1 : Continue For
                            Try
                                Dim tp = tg.CreatePoint2d(minX - 2.5, (prev.Y + h.Y) / 2)
                                oSheet.DrawingDimensions.GeneralDimensions.AddLinear(tp,
                                    oSheet.CreateGeometryIntent(prev.Curve, PointIntentEnum.kCenterPointIntent),
                                    oSheet.CreateGeometryIntent(h.Curve, PointIntentEnum.kCenterPointIntent),
                                    DimensionTypeEnum.kVerticalDimensionType)
                                countOK += 1
                                prev = h
                            Catch : countFail += 1 : End Try
                        Next
                        ' lỗ cuối → mặt dưới
                        If bottomEdge IsNot Nothing AndAlso (prev.Y - minY) > TOL Then
                            Try
                                Dim tp = tg.CreatePoint2d(minX - 3.2, (prev.Y + minY) / 2)
                                oSheet.DrawingDimensions.GeneralDimensions.AddLinear(tp,
                                    oSheet.CreateGeometryIntent(prev.Curve, PointIntentEnum.kCenterPointIntent),
                                    oSheet.CreateGeometryIntent(bottomEdge),
                                    DimensionTypeEnum.kVerticalDimensionType)
                                countOK += 1
                            Catch : countFail += 1 : End Try
                        End If
                    End If

                    '=====================================================
                    ' b3: Trái → Phải (phía dưới) + đến mặt phải
                    '=====================================================
                    Dim botRow = cleanHoles.Where(Function(h) h.Y < limitY).OrderBy(Function(h) h.X).ToList()
                    If botRow.Count > 0 AndAlso leftEdge IsNot Nothing Then
                        Dim prev = botRow.First()
                        If (prev.X - minX) > TOL Then
                            Try
                                Dim tp = tg.CreatePoint2d((minX + prev.X) / 2, minY - 3.2)
                                oSheet.DrawingDimensions.GeneralDimensions.AddLinear(tp,
                                    oSheet.CreateGeometryIntent(leftEdge),
                                    oSheet.CreateGeometryIntent(prev.Curve, PointIntentEnum.kCenterPointIntent),
                                    DimensionTypeEnum.kHorizontalDimensionType)
                                countOK += 1
                            Catch : countFail += 1 : End Try
                        End If
                        For i = 1 To botRow.Count - 1
                            Dim h = botRow(i)
                            Dim dist = h.X - prev.X
                            If dist < TOL Then countZero += 1 : Continue For
                            Try
                                Dim tp = tg.CreatePoint2d((prev.X + h.X) / 2, minY - 2.5)
                                oSheet.DrawingDimensions.GeneralDimensions.AddLinear(tp,
                                    oSheet.CreateGeometryIntent(prev.Curve, PointIntentEnum.kCenterPointIntent),
                                    oSheet.CreateGeometryIntent(h.Curve, PointIntentEnum.kCenterPointIntent),
                                    DimensionTypeEnum.kHorizontalDimensionType)
                                countOK += 1
                                prev = h
                            Catch : countFail += 1 : End Try
                        Next
                        If rightEdge IsNot Nothing AndAlso (maxX - prev.X) > TOL Then
                            Try
                                Dim tp = tg.CreatePoint2d((prev.X + maxX) / 2, minY - 3.2)
                                oSheet.DrawingDimensions.GeneralDimensions.AddLinear(tp,
                                    oSheet.CreateGeometryIntent(prev.Curve, PointIntentEnum.kCenterPointIntent),
                                    oSheet.CreateGeometryIntent(rightEdge),
                                    DimensionTypeEnum.kHorizontalDimensionType)
                                countOK += 1
                            Catch : countFail += 1 : End Try
                        End If
                    End If

                    '=====================================================
                    ' b4: Trên → Dưới (bên phải) + đến mặt dưới
                    '=====================================================
                    Dim rightCol = cleanHoles.Where(Function(h) h.X > limitX) _
                                             .OrderByDescending(Function(h) h.Y).ToList()
                    If rightCol.Count > 0 AndAlso topEdge IsNot Nothing Then
                        Dim prev = rightCol.First()
                        If (maxY - prev.Y) > TOL Then
                            Try
                                Dim tp = tg.CreatePoint2d(maxX + 3.2, (maxY + prev.Y) / 2)
                                oSheet.DrawingDimensions.GeneralDimensions.AddLinear(tp,
                                    oSheet.CreateGeometryIntent(topEdge),
                                    oSheet.CreateGeometryIntent(prev.Curve, PointIntentEnum.kCenterPointIntent),
                                    DimensionTypeEnum.kVerticalDimensionType)
                                countOK += 1
                            Catch : countFail += 1 : End Try
                        End If
                        For i = 1 To rightCol.Count - 1
                            Dim h = rightCol(i)
                            Dim dist = prev.Y - h.Y
                            If dist < TOL Then countZero += 1 : Continue For
                            Try
                                Dim tp = tg.CreatePoint2d(maxX + 2.5, (prev.Y + h.Y) / 2)
                                oSheet.DrawingDimensions.GeneralDimensions.AddLinear(tp,
                                    oSheet.CreateGeometryIntent(prev.Curve, PointIntentEnum.kCenterPointIntent),
                                    oSheet.CreateGeometryIntent(h.Curve, PointIntentEnum.kCenterPointIntent),
                                    DimensionTypeEnum.kVerticalDimensionType)
                                countOK += 1
                                prev = h
                            Catch : countFail += 1 : End Try
                        Next
                        If bottomEdge IsNot Nothing AndAlso (prev.Y - minY) > TOL Then
                            Try
                                Dim tp = tg.CreatePoint2d(maxX + 3.2, (prev.Y + minY) / 2)
                                oSheet.DrawingDimensions.GeneralDimensions.AddLinear(tp,
                                    oSheet.CreateGeometryIntent(prev.Curve, PointIntentEnum.kCenterPointIntent),
                                    oSheet.CreateGeometryIntent(bottomEdge),
                                    DimensionTypeEnum.kVerticalDimensionType)
                                countOK += 1
                            Catch : countFail += 1 : End Try
                        End If
                    End If

                Next

                oDrawDoc.Update()

                MessageBox.Show(
                    "Hoàn tất!" & vbCrLf & vbCrLf &
                    "Số dim: " & countOK & vbCrLf &
                    "Bỏ = 0: " & countZero & vbCrLf &
                    "Lỗi: " & countFail & vbCrLf & vbCrLf &
                    "• Đã thêm dim đến mặt cuối mỗi bước" & vbCrLf &
                    "• Lọc array bằng feature + mật độ",
                    "Dim Chain lỗ", MessageBoxButtons.OK, MessageBoxIcon.Information)

            Catch ex As Exception
                MessageBox.Show("Lỗi:" & vbCrLf & ex.Message, "Dim Chain lỗ", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try

        End Sub

        Public Class HoleInfo
            Public Curve As DrawingCurve
            Public Center As Point2d
            Public X As Double
            Public Y As Double
            Public Radius As Double
        End Class

    End Module

End Namespace