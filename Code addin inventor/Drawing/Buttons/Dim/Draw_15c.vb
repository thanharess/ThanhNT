Option Explicit On
Option Strict Off

Imports System.Windows.Forms
Imports Inventor
Imports System.Collections.Generic
Imports System.Linq

Namespace ToolInventor2020.Drawing.Buttons

    Public Module draw_15c

        Private Const TOL As Double = 0.03
        Private Const RATIO As Double = 0.5

        Public Sub OnExecute(ByVal Context As NameValueMap)

            Dim app As Inventor.Application = g_inventorApplication

            Try

                If app.ActiveDocument Is Nothing OrElse
                   app.ActiveDocument.DocumentType <> DocumentTypeEnum.kDrawingDocumentObject Then

                    MessageBox.Show(
                        "Vui lòng mở file Drawing (.idw)!",
                        "Lỗi",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error)

                    Exit Sub
                End If

                Dim oDrawDoc As DrawingDocument =
                    CType(app.ActiveDocument, DrawingDocument)

                Dim oSheet As Sheet = oDrawDoc.ActiveSheet
                Dim tg As TransientGeometry = app.TransientGeometry

                Dim countOK As Integer = 0
                Dim countFail As Integer = 0
                Dim countZero As Integer = 0
                Dim countAdd As Integer = 0

                For Each oView As DrawingView In oSheet.DrawingViews

                    '=====================================================
                    ' 1. LẤY TẤT CẢ LỖ
                    ' KHÔNG LỌC ARRAY
                    ' KHÔNG LỌC MẬT ĐỘ
                    ' KHÔNG XÉT ĐƯỜNG KÍNH
                    '=====================================================

                    Dim allHoles As New List(Of HoleInfo)

                    For Each oCurve As DrawingCurve In oView.DrawingCurves

                        Try

                            If oCurve.CurveType <>
                               CurveTypeEnum.kCircleCurve Then Continue For

                            Dim c As Point2d = oCurve.CenterPoint

                            If c Is Nothing Then Continue For

                            '---------------------------------------------
                            ' Chỉ loại nếu TÂM TRÙNG hoàn toàn với một lỗ
                            ' khác.
                            '
                            ' Không xét Radius.
                            '---------------------------------------------

                            Dim duplicated As Boolean = False

                            For Each oldHole As HoleInfo In allHoles

                                If Math.Abs(oldHole.X - c.X) <= TOL AndAlso
                                   Math.Abs(oldHole.Y - c.Y) <= TOL Then

                                    duplicated = True
                                    Exit For

                                End If

                            Next

                            If duplicated Then Continue For

                            Dim hi As New HoleInfo

                            hi.Curve = oCurve
                            hi.Center = c
                            hi.X = c.X
                            hi.Y = c.Y
                            hi.Dimmed = False

                            allHoles.Add(hi)

                        Catch
                        End Try

                    Next

                    If allHoles.Count = 0 Then Continue For


                    '=====================================================
                    ' 2. TÌM 4 CẠNH NGOÀI
                    '=====================================================

                    Dim leftEdge As DrawingCurve = Nothing
                    Dim rightEdge As DrawingCurve = Nothing
                    Dim topEdge As DrawingCurve = Nothing
                    Dim bottomEdge As DrawingCurve = Nothing

                    Dim minX As Double = Double.MaxValue
                    Dim maxX As Double = Double.MinValue

                    Dim minY As Double = Double.MaxValue
                    Dim maxY As Double = Double.MinValue


                    For Each oCurve As DrawingCurve In oView.DrawingCurves

                        Try

                            If oCurve.CurveType <>
                               CurveTypeEnum.kLineCurve AndAlso
                               oCurve.CurveType <>
                               CurveTypeEnum.kLineSegmentCurve Then Continue For

                            Dim p1 As Point2d = oCurve.StartPoint
                            Dim p2 As Point2d = oCurve.EndPoint

                            If p1 Is Nothing OrElse p2 Is Nothing Then Continue For


                            ' Cạnh đứng
                            If Math.Abs(p1.X - p2.X) < 0.03 Then

                                If p1.X < minX Then
                                    minX = p1.X
                                    leftEdge = oCurve
                                End If

                                If p1.X > maxX Then
                                    maxX = p1.X
                                    rightEdge = oCurve
                                End If

                            End If


                            ' Cạnh ngang
                            If Math.Abs(p1.Y - p2.Y) < 0.03 Then

                                If p1.Y > maxY Then
                                    maxY = p1.Y
                                    topEdge = oCurve
                                End If

                                If p1.Y < minY Then
                                    minY = p1.Y
                                    bottomEdge = oCurve
                                End If

                            End If

                        Catch
                        End Try

                    Next


                    If leftEdge Is Nothing OrElse
                       rightEdge Is Nothing OrElse
                       topEdge Is Nothing OrElse
                       bottomEdge Is Nothing Then

                        Continue For

                    End If


                    Dim viewW As Double = maxX - minX
                    Dim viewH As Double = maxY - minY

                    If viewW <= 0 OrElse viewH <= 0 Then Continue For


                    Dim limitX As Double =
                        minX + viewW * RATIO

                    Dim limitY As Double =
                        maxY - viewH * RATIO


                    '=====================================================
                    ' 3. DIM TỔNG
                    '=====================================================

                    Try

                        Dim tp As Point2d =
                            tg.CreatePoint2d(
                                (minX + maxX) / 2,
                                maxY + 5.5)

                        oSheet.DrawingDimensions.GeneralDimensions.AddLinear(
                            tp,
                            oSheet.CreateGeometryIntent(leftEdge),
                            oSheet.CreateGeometryIntent(rightEdge),
                            DimensionTypeEnum.kHorizontalDimensionType)

                        countOK += 1

                    Catch
                        countFail += 1
                    End Try


                    Try

                        Dim tp As Point2d =
                            tg.CreatePoint2d(
                                minX - 5.5,
                                (maxY + minY) / 2)

                        oSheet.DrawingDimensions.GeneralDimensions.AddLinear(
                            tp,
                            oSheet.CreateGeometryIntent(topEdge),
                            oSheet.CreateGeometryIntent(bottomEdge),
                            DimensionTypeEnum.kVerticalDimensionType)

                        countOK += 1

                    Catch
                        countFail += 1
                    End Try


                    '=====================================================
                    ' 4. CHUỖI TRÊN
                    '=====================================================

                    Dim topRow As List(Of HoleInfo) =
                        allHoles.Where(
                            Function(h) h.Y >= limitY).
                        OrderBy(
                            Function(h) h.X).
                        ToList()


                    If topRow.Count > 0 Then

                        Dim prev As HoleInfo = topRow.First()

                        ' Trái -> lỗ đầu
                        If Math.Abs(prev.X - minX) > TOL Then

                            If AddHorizontalDim(
                                oSheet,
                                tg,
                                leftEdge,
                                prev.Curve,
                                (minX + prev.X) / 2,
                                maxY + 3.2) Then

                                countOK += 1
                                prev.Dimmed = True

                            Else

                                countFail += 1

                            End If

                        Else

                            prev.Dimmed = True

                        End If


                        ' Lỗ -> lỗ
                        For i As Integer = 1 To topRow.Count - 1

                            Dim h As HoleInfo = topRow(i)

                            If Math.Abs(h.X - prev.X) <= TOL Then

                                ' Cùng tâm X -> không tạo dim 0
                                ' Nhưng vẫn đánh dấu lỗ đã được xử lý
                                h.Dimmed = True
                                countZero += 1

                                Continue For

                            End If


                            If AddHorizontalDim(
                                oSheet,
                                tg,
                                prev.Curve,
                                h.Curve,
                                (prev.X + h.X) / 2,
                                maxY + 2.5) Then

                                countOK += 1

                                prev.Dimmed = True
                                h.Dimmed = True

                            Else

                                countFail += 1

                            End If

                            prev = h

                        Next


                        ' Lỗ cuối -> phải
                        If Math.Abs(maxX - prev.X) > TOL Then

                            If AddHorizontalDim(
                                oSheet,
                                tg,
                                prev.Curve,
                                rightEdge,
                                (prev.X + maxX) / 2,
                                maxY + 3.2) Then

                                countOK += 1
                                prev.Dimmed = True

                            Else

                                countFail += 1

                            End If

                        End If

                    End If


                    '=====================================================
                    ' 5. CHUỖI TRÁI
                    '=====================================================

                    Dim leftCol As List(Of HoleInfo) =
                        allHoles.Where(
                            Function(h) h.X <= limitX).
                        OrderByDescending(
                            Function(h) h.Y).
                        ToList()


                    If leftCol.Count > 0 Then

                        Dim prev As HoleInfo = leftCol.First()

                        ' Trên -> lỗ đầu
                        If Math.Abs(maxY - prev.Y) > TOL Then

                            If AddVerticalDim(
                                oSheet,
                                tg,
                                topEdge,
                                prev.Curve,
                                minX - 3.2,
                                (maxY + prev.Y) / 2) Then

                                countOK += 1
                                prev.Dimmed = True

                            Else

                                countFail += 1

                            End If

                        Else

                            prev.Dimmed = True

                        End If


                        For i As Integer = 1 To leftCol.Count - 1

                            Dim h As HoleInfo = leftCol(i)

                            If Math.Abs(prev.Y - h.Y) <= TOL Then

                                h.Dimmed = True
                                countZero += 1

                                Continue For

                            End If


                            If AddVerticalDim(
                                oSheet,
                                tg,
                                prev.Curve,
                                h.Curve,
                                minX - 2.5,
                                (prev.Y + h.Y) / 2) Then

                                countOK += 1

                                prev.Dimmed = True
                                h.Dimmed = True

                            Else

                                countFail += 1

                            End If

                            prev = h

                        Next


                        If Math.Abs(prev.Y - minY) > TOL Then

                            If AddVerticalDim(
                                oSheet,
                                tg,
                                prev.Curve,
                                bottomEdge,
                                minX - 3.2,
                                (prev.Y + minY) / 2) Then

                                countOK += 1
                                prev.Dimmed = True

                            Else

                                countFail += 1

                            End If

                        End If

                    End If


                    '=====================================================
                    ' 6. CHUỖI DƯỚI
                    '=====================================================

                    Dim botRow As List(Of HoleInfo) =
                        allHoles.Where(
                            Function(h) h.Y < limitY).
                        OrderBy(
                            Function(h) h.X).
                        ToList()


                    If botRow.Count > 0 Then

                        Dim prev As HoleInfo = botRow.First()

                        If Math.Abs(prev.X - minX) > TOL Then

                            If AddHorizontalDim(
                                oSheet,
                                tg,
                                leftEdge,
                                prev.Curve,
                                (minX + prev.X) / 2,
                                minY - 3.2) Then

                                countOK += 1
                                prev.Dimmed = True

                            Else

                                countFail += 1

                            End If

                        Else

                            prev.Dimmed = True

                        End If


                        For i As Integer = 1 To botRow.Count - 1

                            Dim h As HoleInfo = botRow(i)

                            If Math.Abs(h.X - prev.X) <= TOL Then

                                h.Dimmed = True
                                countZero += 1

                                Continue For

                            End If


                            If AddHorizontalDim(
                                oSheet,
                                tg,
                                prev.Curve,
                                h.Curve,
                                (prev.X + h.X) / 2,
                                minY - 2.5) Then

                                countOK += 1

                                prev.Dimmed = True
                                h.Dimmed = True

                            Else

                                countFail += 1

                            End If

                            prev = h

                        Next


                        If Math.Abs(maxX - prev.X) > TOL Then

                            If AddHorizontalDim(
                                oSheet,
                                tg,
                                prev.Curve,
                                rightEdge,
                                (prev.X + maxX) / 2,
                                minY - 3.2) Then

                                countOK += 1
                                prev.Dimmed = True

                            Else

                                countFail += 1

                            End If

                        End If

                    End If


                    '=====================================================
                    ' 7. CHUỖI PHẢI
                    '=====================================================

                    Dim rightCol As List(Of HoleInfo) =
                        allHoles.Where(
                            Function(h) h.X > limitX).
                        OrderByDescending(
                            Function(h) h.Y).
                        ToList()


                    If rightCol.Count > 0 Then

                        Dim prev As HoleInfo = rightCol.First()

                        If Math.Abs(maxY - prev.Y) > TOL Then

                            If AddVerticalDim(
                                oSheet,
                                tg,
                                topEdge,
                                prev.Curve,
                                maxX + 3.2,
                                (maxY + prev.Y) / 2) Then

                                countOK += 1
                                prev.Dimmed = True

                            Else

                                countFail += 1

                            End If

                        Else

                            prev.Dimmed = True

                        End If


                        For i As Integer = 1 To rightCol.Count - 1

                            Dim h As HoleInfo = rightCol(i)

                            If Math.Abs(prev.Y - h.Y) <= TOL Then

                                h.Dimmed = True
                                countZero += 1

                                Continue For

                            End If


                            If AddVerticalDim(
                                oSheet,
                                tg,
                                prev.Curve,
                                h.Curve,
                                maxX + 2.5,
                                (prev.Y + h.Y) / 2) Then

                                countOK += 1

                                prev.Dimmed = True
                                h.Dimmed = True

                            Else

                                countFail += 1

                            End If

                            prev = h

                        Next


                        If Math.Abs(prev.Y - minY) > TOL Then

                            If AddVerticalDim(
                                oSheet,
                                tg,
                                prev.Curve,
                                bottomEdge,
                                maxX + 3.2,
                                (prev.Y + minY) / 2) Then

                                countOK += 1
                                prev.Dimmed = True

                            Else

                                countFail += 1

                            End If

                        End If

                    End If


                    '=====================================================
                    ' 8. QUÉT LẠI TOÀN BỘ LỖ
                    '
                    ' Đây là phần bổ sung mới.
                    '
                    ' Lỗ nào chưa được đánh dấu Dimmed thì tìm:
                    '   1. Lỗ cùng Y
                    '   2. Nếu không có -> lỗ cùng X
                    '   3. Nếu không có -> cạnh ngoài
                    '=====================================================

                    For Each hole As HoleInfo In allHoles

                        If hole.Dimmed Then Continue For


                        '=================================================
                        ' TÌM LỖ CÙNG ĐƯỜNG NGANG
                        '=================================================

                        Dim sameRow As New List(Of HoleInfo)

                        For Each other As HoleInfo In allHoles

                            If other Is hole Then Continue For

                            If Math.Abs(other.Y - hole.Y) <= TOL Then

                                sameRow.Add(other)

                            End If

                        Next


                        If sameRow.Count > 0 Then

                            Dim nearest As HoleInfo =
                                sameRow.OrderBy(
                                    Function(h) Math.Abs(h.X - hole.X)).
                                First()


                            If Math.Abs(nearest.X - hole.X) > TOL Then

                                Dim yPos As Double

                                If hole.Y >= limitY Then
                                    yPos = maxY + 2.5
                                Else
                                    yPos = minY - 2.5
                                End If


                                If AddHorizontalDim(
                                    oSheet,
                                    tg,
                                    hole.Curve,
                                    nearest.Curve,
                                    (hole.X + nearest.X) / 2,
                                    yPos) Then

                                    countOK += 1
                                    countAdd += 1
                                    hole.Dimmed = True

                                    Continue For

                                Else

                                    countFail += 1

                                End If

                            End If

                        End If


                        '=================================================
                        ' TÌM LỖ CÙNG ĐƯỜNG DỌC
                        '=================================================

                        Dim sameCol As New List(Of HoleInfo)

                        For Each other As HoleInfo In allHoles

                            If other Is hole Then Continue For

                            If Math.Abs(other.X - hole.X) <= TOL Then

                                sameCol.Add(other)

                            End If

                        Next


                        If sameCol.Count > 0 Then

                            Dim nearest As HoleInfo =
                                sameCol.OrderBy(
                                    Function(h) Math.Abs(h.Y - hole.Y)).
                                First()


                            If Math.Abs(nearest.Y - hole.Y) > TOL Then

                                Dim xPos As Double

                                If hole.X <= limitX Then
                                    xPos = minX - 2.5
                                Else
                                    xPos = maxX + 2.5
                                End If


                                If AddVerticalDim(
                                    oSheet,
                                    tg,
                                    hole.Curve,
                                    nearest.Curve,
                                    xPos,
                                    (hole.Y + nearest.Y) / 2) Then

                                    countOK += 1
                                    countAdd += 1
                                    hole.Dimmed = True

                                    Continue For

                                Else

                                    countFail += 1

                                End If

                            End If

                        End If


                        '=================================================
                        ' KHÔNG CÓ LỖ CÙNG HÀNG/CỘT
                        '
                        ' Dim trực tiếp đến cạnh gần nhất
                        '=================================================

                        If hole.X <= limitX Then

                            If AddHorizontalDim(
                                oSheet,
                                tg,
                                leftEdge,
                                hole.Curve,
                                (minX + hole.X) / 2,
                                minX - 3.2) Then

                                countOK += 1
                                countAdd += 1
                                hole.Dimmed = True

                            End If

                        Else

                            If AddHorizontalDim(
                                oSheet,
                                tg,
                                hole.Curve,
                                rightEdge,
                                (hole.X + maxX) / 2,
                                maxX + 3.2) Then

                                countOK += 1
                                countAdd += 1
                                hole.Dimmed = True

                            End If

                        End If

                    Next

                Next


                '=========================================================
                ' UPDATE
                '=========================================================

                oDrawDoc.Update()


                MessageBox.Show(
                    "Hoàn tất!" & vbCrLf & vbCrLf &
                    "Tổng số dim: " & countOK & vbCrLf &
                    "Dim bổ sung lượt 2: " & countAdd & vbCrLf &
                    "Bỏ dim = 0: " & countZero & vbCrLf &
                    "Lỗi: " & countFail & vbCrLf & vbCrLf &
                    "• Không lọc Circular Pattern" & vbCrLf &
                    "• Không lọc mật độ lỗ" & vbCrLf &
                    "• Không xét đường kính" & vbCrLf &
                    "• Lỗ sát nhau vẫn được giữ" & vbCrLf &
                    "• Đã quét lại để bổ sung lỗ thiếu",
                    "Dim Chain lỗ",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information)


            Catch ex As Exception

                MessageBox.Show(
                    "Lỗi:" & vbCrLf & ex.Message,
                    "Dim Chain lỗ",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error)

            End Try

        End Sub


        '=============================================================
        ' TẠO DIM NGANG
        '=============================================================

        Private Function AddHorizontalDim(
            ByVal oSheet As Sheet,
            ByVal tg As TransientGeometry,
            ByVal c1 As DrawingCurve,
            ByVal c2 As DrawingCurve,
            ByVal x As Double,
            ByVal y As Double) As Boolean

            Try

                Dim tp As Point2d =
                    tg.CreatePoint2d(x, y)

                oSheet.DrawingDimensions.GeneralDimensions.AddLinear(
                    tp,
                    oSheet.CreateGeometryIntent(
                        c1,
                        PointIntentEnum.kCenterPointIntent),
                    oSheet.CreateGeometryIntent(
                        c2,
                        PointIntentEnum.kCenterPointIntent),
                    DimensionTypeEnum.kHorizontalDimensionType)

                Return True

            Catch

                Return False

            End Try

        End Function


        '=============================================================
        ' TẠO DIM DỌC
        '=============================================================

        Private Function AddVerticalDim(
            ByVal oSheet As Sheet,
            ByVal tg As TransientGeometry,
            ByVal c1 As DrawingCurve,
            ByVal c2 As DrawingCurve,
            ByVal x As Double,
            ByVal y As Double) As Boolean

            Try

                Dim tp As Point2d =
                    tg.CreatePoint2d(x, y)

                oSheet.DrawingDimensions.GeneralDimensions.AddLinear(
                    tp,
                    oSheet.CreateGeometryIntent(
                        c1,
                        PointIntentEnum.kCenterPointIntent),
                    oSheet.CreateGeometryIntent(
                        c2,
                        PointIntentEnum.kCenterPointIntent),
                    DimensionTypeEnum.kVerticalDimensionType)

                Return True

            Catch

                Return False

            End Try

        End Function


        '=============================================================
        ' HOLE INFO
        '=============================================================

        Public Class HoleInfo

            Public Curve As DrawingCurve
            Public Center As Point2d

            Public X As Double
            Public Y As Double

            Public Dimmed As Boolean

        End Class

    End Module

End Namespace