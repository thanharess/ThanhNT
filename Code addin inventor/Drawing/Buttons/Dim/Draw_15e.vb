Option Explicit On
Option Strict Off

Imports System.Windows.Forms
Imports Inventor
Imports System.Collections.Generic
Imports System.Linq

Namespace ToolInventor2020.Drawing.Buttons

    Public Module draw_15e

        Private Const TOL As Double = 0.00
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
                Dim countAdd As Integer = 0
                Dim countZero As Integer = 0

                '=========================================================
                ' DUYỆT VIEW
                '=========================================================

                For Each oView As DrawingView In oSheet.DrawingViews

                    '=====================================================
                    ' 1. LẤY TẤT CẢ LỖ TRÒN (chỉ loại trùng TÂM)
                    '=====================================================

                    Dim allHoles As New List(Of HoleInfo)

                    For Each oCurve As DrawingCurve In oView.DrawingCurves

                        Try

                            If oCurve.CurveType <>
                               CurveTypeEnum.kCircleCurve Then Continue For

                            Dim c As Point2d = oCurve.CenterPoint

                            If c Is Nothing Then Continue For

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


                    Dim limitX As Double = minX + viewW * RATIO
                    Dim limitY As Double = maxY - viewH * RATIO


                    '=====================================================
                    ' DANH SÁCH TỌA ĐỘ ĐÃ DIM (ĐỂ CHỐNG TRÙNG)
                    '
                    ' usedX: các giá trị X đã tạo dim ngang
                    ' usedY: các giá trị Y đã tạo dim dọc
                    '
                    ' Tất cả nhóm dùng chung base (trái cho X, trên cho Y)
                    ' nên phải dedupe xuyên nhóm B1+B3 và B2+B4.
                    '=====================================================

                    Dim usedX As New List(Of Double)
                    Dim usedY As New List(Of Double)


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
                    ' 4. B1 - HÀNG TRÊN  (BASE = CẠNH TRÁI)
                    '=====================================================

                    Dim topRow As List(Of HoleInfo) =
                        allHoles.Where(
                            Function(h) h.Y >= limitY).
                        OrderBy(
                            Function(h) h.X).
                        ToList()


                    If topRow.Count > 0 Then

                        Dim result As Boolean =
                            AddBaselineChain(
                                oSheet,
                                app,
                                leftEdge,
                                topRow,
                                True,
                                maxY + 3.2,
                                usedX)

                        If result Then

                            For Each h As HoleInfo In topRow
                                h.Dimmed = True
                            Next

                        End If

                    End If


                    '=====================================================
                    ' 5. B2 - CỘT TRÁI  (BASE = CẠNH TRÊN)
                    '=====================================================

                    Dim leftCol As List(Of HoleInfo) =
                        allHoles.Where(
                            Function(h) h.X <= limitX).
                        OrderByDescending(
                            Function(h) h.Y).
                        ToList()


                    If leftCol.Count > 0 Then

                        Dim result As Boolean =
                            AddBaselineChain(
                                oSheet,
                                app,
                                topEdge,
                                leftCol,
                                False,
                                minX - 3.2,
                                usedY)

                        If result Then

                            For Each h As HoleInfo In leftCol
                                h.Dimmed = True
                            Next

                        End If

                    End If


                    '=====================================================
                    ' 6. B3 - HÀNG DƯỚI  (BASE = CẠNH TRÁI)
                    '=====================================================

                    Dim botRow As List(Of HoleInfo) =
                        allHoles.Where(
                            Function(h) h.Y < limitY).
                        OrderBy(
                            Function(h) h.X).
                        ToList()


                    If botRow.Count > 0 Then

                        Dim result As Boolean =
                            AddBaselineChain(
                                oSheet,
                                app,
                                leftEdge,
                                botRow,
                                True,
                                minY - 3.2,
                                usedX)

                        If result Then

                            For Each h As HoleInfo In botRow
                                h.Dimmed = True
                            Next

                        End If

                    End If


                    '=====================================================
                    ' 7. B4 - CỘT PHẢI  (BASE = CẠNH TRÊN)
                    '=====================================================

                    Dim rightCol As List(Of HoleInfo) =
                        allHoles.Where(
                            Function(h) h.X > limitX).
                        OrderByDescending(
                            Function(h) h.Y).
                        ToList()


                    If rightCol.Count > 0 Then

                        Dim result As Boolean =
                            AddBaselineChain(
                                oSheet,
                                app,
                                topEdge,
                                rightCol,
                                False,
                                maxX + 3.2,
                                usedY)

                        If result Then

                            For Each h As HoleInfo In rightCol
                                h.Dimmed = True
                            Next

                        End If

                    End If


                    '=====================================================
                    ' 8. QUÉT LẠI LỖ CHƯA ĐƯỢC DIM
                    ' (vẫn dedupe theo usedX / usedY)
                    '=====================================================

                    Dim missingHoles As List(Of HoleInfo) =
                        allHoles.Where(
                            Function(h) Not h.Dimmed).
                        ToList()


                    For Each hole As HoleInfo In missingHoles

                        Dim done As Boolean = False


                        '----- Thử ngang trước -----
                        Dim canX As Boolean =
                            Not ValueExists(usedX, hole.X)

                        If canX Then

                            If hole.X <= limitX Then

                                done =
                                    AddSingleBaseline(
                                        oSheet,
                                        app,
                                        leftEdge,
                                        hole.Curve,
                                        True,
                                        minX - 3.2)

                            Else

                                done =
                                    AddSingleBaseline(
                                        oSheet,
                                        app,
                                        leftEdge,
                                        hole.Curve,
                                        True,
                                        hole.Y + 2.5)

                            End If


                            If done Then

                                usedX.Add(hole.X)
                                hole.Dimmed = True
                                countOK += 1
                                countAdd += 1

                                Continue For

                            End If

                        End If


                        '----- Thử dọc -----
                        Dim canY As Boolean =
                            Not ValueExists(usedY, hole.Y)

                        If canY Then

                            If hole.X <= limitX Then

                                done =
                                    AddSingleBaseline(
                                        oSheet,
                                        app,
                                        topEdge,
                                        hole.Curve,
                                        False,
                                        minX - 3.2)

                            Else

                                done =
                                    AddSingleBaseline(
                                        oSheet,
                                        app,
                                        topEdge,
                                        hole.Curve,
                                        False,
                                        maxX + 3.2)

                            End If


                            If done Then

                                usedY.Add(hole.Y)
                                hole.Dimmed = True
                                countOK += 1
                                countAdd += 1

                            Else

                                countFail += 1

                            End If

                        End If

                    Next


                    '=====================================================
                    ' 9. QUÉT LẠI LẦN CUỐI
                    '=====================================================

                    For Each hole As HoleInfo In allHoles

                        If hole.Dimmed Then Continue For


                        Dim ok As Boolean = False


                        '----- Ngang -----
                        If Not ValueExists(usedX, hole.X) Then

                            ok =
                                AddSingleBaseline(
                                    oSheet,
                                    app,
                                    leftEdge,
                                    hole.Curve,
                                    True,
                                    minX - 4.0)

                            If ok Then
                                usedX.Add(hole.X)
                            End If

                        End If


                        '----- Dọc -----
                        If Not ok AndAlso Not ValueExists(usedY, hole.Y) Then

                            ok =
                                AddSingleBaseline(
                                    oSheet,
                                    app,
                                    topEdge,
                                    hole.Curve,
                                    False,
                                    maxX + 4.0)

                            If ok Then
                                usedY.Add(hole.Y)
                            End If

                        End If


                        If ok Then

                            hole.Dimmed = True
                            countOK += 1
                            countAdd += 1

                        End If

                    Next

                Next


                '=========================================================
                ' UPDATE
                '=========================================================

                oDrawDoc.Update()


                MessageBox.Show(
                    "Hoàn tất!" & vbCrLf & vbCrLf &
                    "Số dim tạo: " & countOK & vbCrLf &
                    "Dim bổ sung: " & countAdd & vbCrLf &
                    "Lỗi: " & countFail & vbCrLf &
                    "Bỏ = 0: " & countZero & vbCrLf & vbCrLf &
                    "• Dim tất cả lỗ" & vbCrLf &
                    "• Loại trùng TÂM và trùng KHOẢNG CÁCH" & vbCrLf &
                    "• Ngang: X trùng -> bỏ" & vbCrLf &
                    "• Dọc:  Y trùng -> bỏ",
                    "Dim Baseline lỗ",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information)


            Catch ex As Exception

                MessageBox.Show(
                    "Lỗi:" & vbCrLf & ex.Message,
                    "Dim Baseline lỗ",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error)

            End Try

        End Sub


        '=============================================================
        ' KIỂM TRA GIÁ TRỊ ĐÃ TỒN TẠI TRONG DANH SÁCH (sai số TOL)
        '=============================================================

        Private Function ValueExists(
            ByVal list As List(Of Double),
            ByVal value As Double) As Boolean

            If list Is Nothing Then Return False

            For Each v As Double In list

                If Math.Abs(v - value) <= TOL Then Return True

            Next

            Return False

        End Function


        '=============================================================
        ' TẠO CHUỖI DIM KIỂU BASELINE
        '
        ' - holeList  : danh sách lỗ trong nhóm
        ' - usedCoords: danh sách tọa độ đã dim trước đó
        '               (để chống trùng giữa các nhóm dùng cùng base)
        '
        ' Nếu tọa độ của lỗ đã tồn tại trong usedCoords -> bỏ qua
        ' (vì đã có 1 dim cùng giá trị, thêm nữa là trùng số)
        '=============================================================

        Private Function AddBaselineChain(
            ByVal oSheet As Sheet,
            ByVal app As Inventor.Application,
            ByVal baseCurve As DrawingCurve,
            ByVal holeList As List(Of HoleInfo),
            ByVal horizontal As Boolean,
            ByVal offset As Double,
            ByVal usedCoords As List(Of Double)) As Boolean

            Try

                If baseCurve Is Nothing Then Return False
                If holeList Is Nothing OrElse holeList.Count = 0 Then Return False
                If usedCoords Is Nothing Then usedCoords = New List(Of Double)


                Dim dimType As DimensionTypeEnum

                If horizontal Then
                    dimType = DimensionTypeEnum.kHorizontalDimensionType
                Else
                    dimType = DimensionTypeEnum.kVerticalDimensionType
                End If


                Dim baseIntent As GeometryIntent =
                    oSheet.CreateGeometryIntent(baseCurve)


                Dim okCount As Integer = 0


                For i As Integer = 0 To holeList.Count - 1

                    Dim hole As HoleInfo = holeList(i)

                    If hole Is Nothing Then Continue For
                    If hole.Curve Is Nothing Then Continue For


                    Dim coord As Double

                    If horizontal Then
                        coord = hole.X
                    Else
                        coord = hole.Y
                    End If


                    '----- Chống trùng: đã có dim cùng giá trị -----
                    If ValueExists(usedCoords, coord) Then Continue For


                    Try

                        Dim holeIntent As GeometryIntent =
                            oSheet.CreateGeometryIntent(
                                hole.Curve,
                                PointIntentEnum.kCenterPointIntent)


                        Dim placement As Point2d

                        If horizontal Then
                            placement = app.TransientGeometry.CreatePoint2d(
                                            hole.X, offset)
                        Else
                            placement = app.TransientGeometry.CreatePoint2d(
                                            offset, hole.Y)
                        End If


                        Dim newDim As GeneralDimension =
                            oSheet.DrawingDimensions.GeneralDimensions.AddLinear(
                                placement,
                                baseIntent,
                                holeIntent,
                                dimType)


                        If newDim IsNot Nothing Then

                            usedCoords.Add(coord)
                            okCount += 1

                        End If

                    Catch
                    End Try

                Next


                Return okCount > 0

            Catch

                Return False

            End Try

        End Function


        '=============================================================
        ' TẠO 1 DIM RIÊNG CHO LỖ BỊ THIẾU
        '=============================================================

        Private Function AddSingleBaseline(
            ByVal oSheet As Sheet,
            ByVal app As Inventor.Application,
            ByVal baseCurve As DrawingCurve,
            ByVal holeCurve As DrawingCurve,
            ByVal horizontal As Boolean,
            ByVal offset As Double) As Boolean

            Try

                If baseCurve Is Nothing Then Return False
                If holeCurve Is Nothing Then Return False


                Dim c As Point2d = holeCurve.CenterPoint

                If c Is Nothing Then Return False


                Dim baseIntent As GeometryIntent =
                    oSheet.CreateGeometryIntent(baseCurve)

                Dim holeIntent As GeometryIntent =
                    oSheet.CreateGeometryIntent(
                        holeCurve,
                        PointIntentEnum.kCenterPointIntent)


                Dim placement As Point2d

                If horizontal Then
                    placement = app.TransientGeometry.CreatePoint2d(
                                    c.X, offset)
                Else
                    placement = app.TransientGeometry.CreatePoint2d(
                                    offset, c.Y)
                End If


                Dim dimType As DimensionTypeEnum

                If horizontal Then
                    dimType = DimensionTypeEnum.kHorizontalDimensionType
                Else
                    dimType = DimensionTypeEnum.kVerticalDimensionType
                End If


                Dim newDim As GeneralDimension =
                    oSheet.DrawingDimensions.GeneralDimensions.AddLinear(
                        placement,
                        baseIntent,
                        holeIntent,
                        dimType)


                If newDim Is Nothing Then Return False


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