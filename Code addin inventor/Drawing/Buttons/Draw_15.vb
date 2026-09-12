Option Explicit On
Option Strict Off

Imports System.Windows.Forms
Imports Inventor
Imports System.Collections.Generic
Imports System.Linq

Namespace ToolInventor2020.Drawing.Buttons

    Public Module DimHoleDistance

        Public Sub OnExecute(ByVal Context As NameValueMap)

            Dim app As Inventor.Application = g_inventorApplication

            Try
                If app.ActiveDocument Is Nothing OrElse
                   app.ActiveDocument.DocumentType <> DocumentTypeEnum.kDrawingDocumentObject Then

                    MessageBox.Show("Vui lòng mở file Drawing (.idw)!", "Lỗi",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                Dim oDrawDoc As DrawingDocument = CType(app.ActiveDocument, DrawingDocument)
                Dim oSheet As Sheet = oDrawDoc.ActiveSheet
                Dim tg As TransientGeometry = app.TransientGeometry

                Dim totalHoles As Integer = 0
                Dim countOK As Integer = 0
                Dim countFail As Integer = 0
                Dim countZero As Integer = 0

                Const TOL As Double = 0.15
                Const OFFSET As Double = 3.0

                For Each oView As DrawingView In oSheet.DrawingViews

                    '=====================================================
                    ' 1. LẤY ĐƯỜNG TRÒN
                    '=====================================================
                    Dim holeList As New List(Of HoleInfo)

                    For Each oCurve As DrawingCurve In oView.DrawingCurves
                        Try
                            If oCurve.CurveType <> CurveTypeEnum.kCircleCurve Then Continue For

                            Dim center As Point2d = oCurve.CenterPoint
                            If center Is Nothing Then Continue For

                            Dim hi As New HoleInfo
                            hi.Curve = oCurve
                            hi.Center = center
                            hi.X = center.X
                            hi.Y = center.Y
                            holeList.Add(hi)
                        Catch
                        End Try
                    Next

                    If holeList.Count = 0 Then Continue For
                    totalHoles += holeList.Count

                    '=====================================================
                    ' 2. TÌM CẠNH TRÁI + CẠNH TRÊN
                    '=====================================================
                    Dim leftEdge As DrawingCurve = Nothing
                    Dim topEdge As DrawingCurve = Nothing
                    Dim minX As Double = Double.MaxValue
                    Dim maxY As Double = Double.MinValue

                    For Each oCurve As DrawingCurve In oView.DrawingCurves
                        Try
                            If oCurve.CurveType <> CurveTypeEnum.kLineCurve AndAlso
                               oCurve.CurveType <> CurveTypeEnum.kLineSegmentCurve Then Continue For

                            Dim p1 As Point2d = oCurve.StartPoint
                            Dim p2 As Point2d = oCurve.EndPoint

                            If Math.Abs(p1.X - p2.X) < 0.02 Then
                                If p1.X < minX Then
                                    minX = p1.X
                                    leftEdge = oCurve
                                End If
                            End If

                            If Math.Abs(p1.Y - p2.Y) < 0.02 Then
                                If p1.Y > maxY Then
                                    maxY = p1.Y
                                    topEdge = oCurve
                                End If
                            End If
                        Catch
                        End Try
                    Next

                    '=====================================================
                    ' 3. TÌM LỖ GỐC: TRÊN - TRÁI (gần cạnh trái + gần cạnh trên nhất)
                    '=====================================================
                    Dim originHole As HoleInfo = holeList _
                        .OrderBy(Function(h) h.X) _
                        .ThenByDescending(Function(h) h.Y) _
                        .First()

                    '=====================================================
                    ' 4. DIM NGANG: Cạnh trái → Lỗ trên-trái → các lỗ sang phải
                    '=====================================================
                    ' Lấy các lỗ theo thứ tự X tăng dần, ưu tiên Y cao (hàng trên)
                    Dim holesForX As List(Of HoleInfo) = holeList _
                        .OrderBy(Function(h) h.X) _
                        .ThenByDescending(Function(h) h.Y) _
                        .ToList()

                    ' 4.1 Cạnh trái → lỗ gốc (trên-trái)
                    If leftEdge IsNot Nothing Then
                        Try
                            If Math.Abs(originHole.X - minX) > TOL Then
                                Dim intent1 As GeometryIntent = oSheet.CreateGeometryIntent(leftEdge)
                                Dim intent2 As GeometryIntent = oSheet.CreateGeometryIntent(originHole.Curve, PointIntentEnum.kCenterPointIntent)

                                Dim textPos As Point2d = tg.CreatePoint2d((minX + originHole.X) / 2, maxY + OFFSET)

                                oSheet.DrawingDimensions.GeneralDimensions.AddLinear(
                                    textPos, intent1, intent2, DimensionTypeEnum.kHorizontalDimensionType)
                                countOK += 1
                            Else
                                countZero += 1
                            End If
                        Catch
                            countFail += 1
                        End Try
                    End If

                    ' 4.2 Từ lỗ gốc sang phải (chỉ lấy lỗ có X lớn hơn rõ)
                    Dim currentX As Double = originHole.X
                    Dim lastHole As HoleInfo = originHole

                    For Each h As HoleInfo In holesForX
                        If h Is originHole Then Continue For
                        If h.X - currentX < TOL Then
                            countZero += 1
                            Continue For
                        End If

                        Try
                            Dim intentA As GeometryIntent = oSheet.CreateGeometryIntent(lastHole.Curve, PointIntentEnum.kCenterPointIntent)
                            Dim intentB As GeometryIntent = oSheet.CreateGeometryIntent(h.Curve, PointIntentEnum.kCenterPointIntent)

                            Dim textPos As Point2d = tg.CreatePoint2d((lastHole.X + h.X) / 2, maxY + OFFSET)

                            oSheet.DrawingDimensions.GeneralDimensions.AddLinear(
                                textPos, intentA, intentB, DimensionTypeEnum.kHorizontalDimensionType)
                            countOK += 1

                            currentX = h.X
                            lastHole = h
                        Catch
                            countFail += 1
                        End Try
                    Next

                    '=====================================================
                    ' 5. DIM DỌC: Chỉ cột bên trái (từ lỗ trên-trái xuống)
                    '=====================================================
                    ' Lấy các lỗ gần cột trái (X gần với originHole.X)
                    Dim leftCol As List(Of HoleInfo) = holeList _
                        .Where(Function(h) Math.Abs(h.X - originHole.X) < 15) _
                        .OrderByDescending(Function(h) h.Y) _
                        .ToList()

                    ' 5.1 Cạnh trên → lỗ trên-trái
                    If topEdge IsNot Nothing Then
                        Try
                            If Math.Abs(maxY - originHole.Y) > TOL Then
                                Dim intent1 As GeometryIntent = oSheet.CreateGeometryIntent(topEdge)
                                Dim intent2 As GeometryIntent = oSheet.CreateGeometryIntent(originHole.Curve, PointIntentEnum.kCenterPointIntent)

                                Dim textPos As Point2d = tg.CreatePoint2d(minX - OFFSET, (maxY + originHole.Y) / 2)

                                oSheet.DrawingDimensions.GeneralDimensions.AddLinear(
                                    textPos, intent1, intent2, DimensionTypeEnum.kVerticalDimensionType)
                                countOK += 1
                            Else
                                countZero += 1
                            End If
                        Catch
                            countFail += 1
                        End Try
                    End If

                    ' 5.2 Từ trên xuống dưới trên cột trái
                    For i As Integer = 0 To leftCol.Count - 2
                        Try
                            Dim h1 As HoleInfo = leftCol(i)
                            Dim h2 As HoleInfo = leftCol(i + 1)

                            If Math.Abs(h1.Y - h2.Y) < TOL Then
                                countZero += 1
                                Continue For
                            End If

                            Dim intentA As GeometryIntent = oSheet.CreateGeometryIntent(h1.Curve, PointIntentEnum.kCenterPointIntent)
                            Dim intentB As GeometryIntent = oSheet.CreateGeometryIntent(h2.Curve, PointIntentEnum.kCenterPointIntent)

                            Dim textPos As Point2d = tg.CreatePoint2d(minX - OFFSET, (h1.Y + h2.Y) / 2)

                            oSheet.DrawingDimensions.GeneralDimensions.AddLinear(
                                textPos, intentA, intentB, DimensionTypeEnum.kVerticalDimensionType)
                            countOK += 1
                        Catch
                            countFail += 1
                        End Try
                    Next

                Next

                Try
                    oSheet.DrawingDimensions.Arrangea()
                Catch
                End Try

                oDrawDoc.Update()

                MessageBox.Show(
                    "Hoàn tất!" & vbCrLf & vbCrLf &
                    "Số đường tròn: " & totalHoles.ToString() & vbCrLf &
                    "Số kích thước: " & countOK.ToString() & vbCrLf &
                    "Đã bỏ = 0: " & countZero.ToString() & vbCrLf &
                    "Lỗi: " & countFail.ToString(),
                    "Dim Chain lỗ",
                    MessageBoxButtons.OK, MessageBoxIcon.Information)

            Catch ex As Exception
                MessageBox.Show("Lỗi:" & vbCrLf & ex.Message, "Dim Chain lỗ",
                                MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try

        End Sub

        Public Class HoleInfo
            Public Curve As DrawingCurve
            Public Center As Point2d
            Public X As Double
            Public Y As Double
        End Class

    End Module

End Namespace