Imports System
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Inventor


Namespace ThanhN.Drawing.Buttons
    Public Module Draw_5
        Public Sub OnExecute(ByVal Context As NameValueMap)

            ' Lấy ứng dụng Inventor đang chạy
            Dim invApp As Inventor.Application = System.Runtime.InteropServices.Marshal.GetActiveObject("Inventor.Application")

                ' Lấy tài liệu Drawing hiện tại
                Dim oDrawDoc As DrawingDocument = CType(invApp.ActiveDocument, DrawingDocument)

                ' Hộp thoại cho người dùng chọn
                Dim choice As String = InputBox("Nhập lựa chọn:" & vbCrLf &
                                        "1 = Reset Active Sheet" & vbCrLf &
                                        "2 = Reset All Sheets", "Reset Parts List")

            If choice = "1" Then
                ' Reset Active Sheet
                Dim oSheet As Sheet = oDrawDoc.ActiveSheet
                ResetPartsListOnSheet(oSheet)

            ElseIf choice = "2" Then
                ' Reset tất cả các Sheet
                For Each oSheet As Sheet In oDrawDoc.Sheets
                    ResetPartsListOnSheet(oSheet)
                Next

            Else
                MsgBox("Bạn chỉ được nhập 1 hoặc 2!", vbExclamation, "Lựa chọn không hợp lệ")
            End If
            MsgBox("Xong")
        End Sub

        Private Sub ResetPartsListOnSheet(oSheet As Sheet)
            If oSheet.PartsLists.Count > 0 Then
                For Each oPartsList As PartsList In oSheet.PartsLists
                    For Each oRow As PartsListRow In oPartsList.PartsListRows
                        For Each oColumn As PartsListColumn In oPartsList.PartsListColumns
                            oRow.Item(oColumn).Static = False
                        Next
                    Next
                Next
            End If

        End Sub

    End Module
        End Namespace