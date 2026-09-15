Option Explicit On
Option Strict Off
Imports System.Windows.Forms
Imports Inventor
Imports System.Collections.Generic

Namespace ToolInventor2020.Drawing.Buttons.Drawdim
    Public Module Delete_small_dims

        '=============================================================
        ' XÓA DIM CÓ GIÁ TRỊ NHỎ HƠN NGƯỠNG
        '=============================================================
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

                '===== NHẬP NGƯỠNG =====
                Dim input As String = InputBox(
                    "Nhập ngưỡng (mm)." & vbCrLf & vbCrLf &
                    "Xóa tất cả dim có giá trị NHỎ HƠN ngưỡng này." & vbCrLf &
                    "Ví dụ: nhập 5 để xóa các dim < 5mm." & vbCrLf & vbCrLf &
                    "Để trống = 0 (không xóa gì).",
                    "Xóa dim nhỏ",
                    "5")

                If input Is Nothing Then Exit Sub
                If input.Trim() = "" Then Exit Sub

                Dim thresholdMM As Double
                If Not Double.TryParse(input.Replace(",", ".").Trim(),
                                       Globalization.NumberStyles.Any,
                                       Globalization.CultureInfo.InvariantCulture,
                                       thresholdMM) Then
                    MessageBox.Show("Giá trị không hợp lệ.", "Lỗi",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                If thresholdMM <= 0 Then
                    MessageBox.Show("Ngưỡng phải > 0.", "Thông báo")
                    Exit Sub
                End If

                ' Đổi mm → cm (đơn vị nội bộ của Inventor API)
                Dim thresholdCM As Double = thresholdMM / 10.0

                '===== QUÉT VÀ XÓA =====
                Dim toDelete As New List(Of DrawingDimension)
                Dim nTotal As Integer = 0

                For Each oDim As DrawingDimension In oSheet.DrawingDimensions
                    Try
                        nTotal += 1

                        ' Chỉ xử lý dim dạng Linear
                        Dim linDim As LinearGeneralDimension = TryCast(oDim, LinearGeneralDimension)
                        If linDim Is Nothing Then Continue For

                        ' Lấy giá trị đo thực tế (đơn vị cm)
                        Dim valCM As Double = 0
                        Try
                            valCM = linDim.ModelValue
                        Catch
                            Try
                                valCM = linDim.Value
                            Catch
                                Continue For
                            End Try
                        End Try

                        ' Nếu < ngưỡng → đánh dấu xóa
                        If Math.Abs(valCM) < thresholdCM Then
                            toDelete.Add(oDim)
                        End If

                    Catch
                    End Try
                Next

                '===== XÓA =====
                Dim nDeleted As Integer = 0
                Dim nFail As Integer = 0

                For Each d As DrawingDimension In toDelete
                    Try
                        d.Delete()
                        nDeleted += 1
                    Catch
                        nFail += 1
                    End Try
                Next

                oDrawDoc.Update()

                MessageBox.Show(
                    "Hoàn tất!" & vbCrLf & vbCrLf &
                    "Tổng dim trên sheet: " & nTotal & vbCrLf &
                    "Đã xóa: " & nDeleted & vbCrLf &
                    "Lỗi: " & nFail & vbCrLf & vbCrLf &
                    "Ngưỡng: " & thresholdMM.ToString("0.0#") & " mm",
                    "Xóa dim nhỏ",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information)

            Catch ex As Exception
                MessageBox.Show("Lỗi:" & vbCrLf & ex.Message,
                                "Xóa dim nhỏ",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error)
            End Try

        End Sub

    End Module
End Namespace