Option Explicit On
Option Strict Off

Imports System.Windows.Forms
Imports Inventor
Imports System.Collections.Generic
Imports System.Linq

Namespace ToolInventor2020.Drawing.Buttons

    Public Module draw_15

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

                '=====================================================
                ' CHỌN VIEW
                '=====================================================
                Dim oView As DrawingView = Nothing

                Try
                    Dim oSelectSet As SelectSet = oDrawDoc.SelectSet
                    If oSelectSet.Count = 1 AndAlso TypeOf oSelectSet.Item(1) Is DrawingView Then
                        oView = CType(oSelectSet.Item(1), DrawingView)
                    End If
                Catch
                End Try

                If oView Is Nothing Then
                    MessageBox.Show("Hãy chọn 1 Drawing View trước khi chạy!", "Thông báo",
                                    MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Exit Sub
                End If

                '=====================================================
                ' THU THẬP LỖ (CẢI THIỆN)
                '=====================================================
                Dim holeList As New List(Of HoleInfo)
                Dim totalCurves As Integer = 0
                Dim circleFound As Integer = 0

                For Each oCurve As DrawingCurve In oView.DrawingCurves
                    totalCurves += 1

                    Try
                        ' Nhận cả Circle và Circular Arc
                        If oCurve.CurveType <> CurveTypeEnum.kCircleCurve AndAlso
                           oCurve.CurveType <> CurveTypeEnum.kCircularArcCurve Then
                            Continue For
                        End If

                        circleFound += 1

                        ' Lấy tâm
                        Dim center As Point2d = Nothing
                        Try
                            center = oCurve.CenterPoint
                        Catch
                            Continue For
                        End Try

                        If center Is Nothing Then Continue For

                        ' Kiểm tra thêm ModelGeometry (tăng độ chính xác)
                        Dim isHole As Boolean = True
                        Try
                            Dim modelGeom As Object = oCurve.ModelGeometry
                            If modelGeom IsNot Nothing AndAlso TypeOf modelGeom Is Edge Then
                                Dim ed As Edge = CType(modelGeom, Edge)
                                ' Có thể bỏ qua nếu muốn lọc chặt hơn
                            End If
                        Catch
                        End Try

                        If isHole Then
                            Dim hi As New HoleInfo
                            hi.Curve = oCurve
                            hi.Center = center
                            hi.X = center.X
                            hi.Y = center.Y

                            Try
                                hi.Radius = oCurve.Radius
                            Catch
                                hi.Radius = 0
                            End Try

                            holeList.Add(hi)
                        End If

                    Catch
                    End Try
                Next

                ' Thông báo debug nếu không tìm thấy
                If holeList.Count = 0 Then
                    MessageBox.Show(
                        "Không tìm thấy lỗ!" & vbCrLf & vbCrLf &
                        "Tổng DrawingCurves: " & totalCurves.ToString() & vbCrLf &
                        "Circle / Arc tìm được: " & circleFound.ToString() & vbCrLf & vbCrLf &
                        "Gợi ý:" & vbCrLf &
                        "1. Kiểm tra View có hiện lỗ không" & vbCrLf &
                        "2. Thử đổi sang View khác (Base View tốt hơn Section)" & vbCrLf &
                        "3. Đảm bảo lỗ là hình tròn hoàn chỉnh",
                        "Không tìm thấy lỗ",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Exit Sub
                End If

                '=====================================================
                ' THÔNG SỐ VIEW
                '=====================================================
                Dim viewLeft As Double = oView.Left
                Dim viewBottom As Double = oView.Bottom
                Dim viewWidth As Double = oView.Width
                Dim viewHeight As Double = oView.Height
                Dim viewTop As Double = viewBottom + viewHeight

                Dim countOK As Integer = 0
                Dim countFail As Integer = 0

                '=====================================================
                ' CHIỀU NGANG (TRÁI → PHẢI)
                '=====================================================
                Dim holesX As List(Of HoleInfo) = holeList.OrderBy(Function(h) h.X).ToList()

                ' Từ cạnh trái → lỗ đầu tiên
                Try
                    Dim firstHole As HoleInfo = holesX(0)

                    Dim leftEdgePoint As Point2d = tg.CreatePoint2d(viewLeft, firstHole.Y)
                    Dim intent1 As GeometryIntent = oSheet.CreateGeometryIntent(Nothing, leftEdgePoint)
                    Dim intent2 As GeometryIntent = oSheet.CreateGeometryIntent(firstHole.Curve, PointIntentEnum.kCircularLeftPointIntent)

                    Dim textPos As Point2d = tg.CreatePoint2d((viewLeft + firstHole.X) / 2, firstHole.Y + 1.5)

                    oSheet.DrawingDimensions.GeneralDimensions.AddLinear(
                        textPos, intent1, intent2, DimensionTypeEnum.kHorizontalDimensionType)

                    countOK += 1
                Catch
                    countFail += 1
                End Try

                ' Khoảng cách giữa các lỗ (ngang)
                For i As Integer = 0 To holesX.Count - 2
                    Try
                        Dim h1 As HoleInfo = holesX(i)
                        Dim h2 As HoleInfo = holesX(i + 1)

                        Dim intentA As GeometryIntent = oSheet.CreateGeometryIntent(h1.Curve, PointIntentEnum.kCircularRightPointIntent)
                        Dim intentB As GeometryIntent = oSheet.CreateGeometryIntent(h2.Curve, PointIntentEnum.kCircularLeftPointIntent)

                        Dim midX As Double = (h1.X + h2.X) / 2
                        Dim textY As Double = Math.Max(h1.Y, h2.Y) + 1.2

                        Dim textPos As Point2d = tg.CreatePoint2d(midX, textY)

                        oSheet.DrawingDimensions.GeneralDimensions.AddLinear(
                            textPos, intentA, intentB, DimensionTypeEnum.kHorizontalDimensionType)

                        countOK += 1
                    Catch
                        countFail += 1
                    End Try
                Next

                '=====================================================
                ' CHIỀU DỌC (TRÊN → DƯỚI)
                '=====================================================
                Dim holesY As List(Of HoleInfo) = holeList.OrderByDescending(Function(h) h.Y).ToList()

                ' Từ cạnh trên → lỗ cao nhất
                Try
                    Dim topHole As HoleInfo = holesY(0)

                    Dim topEdgePoint As Point2d = tg.CreatePoint2d(topHole.X, viewTop)
                    Dim intent1 As GeometryIntent = oSheet.CreateGeometryIntent(Nothing, topEdgePoint)
                    Dim intent2 As GeometryIntent = oSheet.CreateGeometryIntent(topHole.Curve, PointIntentEnum.kCircularTopPointIntent)

                    Dim textPos As Point2d = tg.CreatePoint2d(topHole.X + 1.5, (viewTop + topHole.Y) / 2)

                    oSheet.DrawingDimensions.GeneralDimensions.AddLinear(
                        textPos, intent1, intent2, DimensionTypeEnum.kVerticalDimensionType)

                    countOK += 1
                Catch
                    countFail += 1
                End Try

                ' Khoảng cách giữa các lỗ (dọc)
                For i As Integer = 0 To holesY.Count - 2
                    Try
                        Dim h1 As HoleInfo = holesY(i)
                        Dim h2 As HoleInfo = holesY(i + 1)

                        Dim intentA As GeometryIntent = oSheet.CreateGeometryIntent(h1.Curve, PointIntentEnum.kCircularBottomPointIntent)
                        Dim intentB As GeometryIntent = oSheet.CreateGeometryIntent(h2.Curve, PointIntentEnum.kCircularTopPointIntent)

                        Dim midY As Double = (h1.Y + h2.Y) / 2
                        Dim textX As Double = Math.Max(h1.X, h2.X) + 1.2

                        Dim textPos As Point2d = tg.CreatePoint2d(textX, midY)

                        oSheet.DrawingDimensions.GeneralDimensions.AddLinear(
                            textPos, intentA, intentB, DimensionTypeEnum.kVerticalDimensionType)

                        countOK += 1
                    Catch
                        countFail += 1
                    End Try
                Next

                oDrawDoc.Update()

                MessageBox.Show(
                    "Hoàn tất!" & vbCrLf & vbCrLf &
                    "Số lỗ tìm thấy: " & holeList.Count.ToString() & vbCrLf &
                    "Số kích thước đã đặt: " & countOK.ToString() & vbCrLf &
                    "Lỗi / bỏ qua: " & countFail.ToString(),
                    "Dim khoảng cách lỗ",
                    MessageBoxButtons.OK, MessageBoxIcon.Information)

            Catch ex As Exception
                MessageBox.Show("Lỗi:" & vbCrLf & ex.Message, "Dim khoảng cách lỗ",
                                MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try

        End Sub

        Public Class HoleInfo
            Public Curve As DrawingCurve
            Public Center As Point2d
            Public Radius As Double
            Public X As Double
            Public Y As Double
        End Class

    End Module

End Namespace