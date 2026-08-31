
Option Explicit On
Option Strict Off

Imports System
Imports System.Collections.Generic
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Windows.Forms
Imports Inventor

Namespace ThanhN.Assembly.Buttons.Frame

    Public Module Ass_Frame_2

        Private _invApp As Inventor.Application = Nothing

        '============================================================
        ' INVENTOR 2020
        '
        ' DISPLAY STATE SICK
        '============================================================

        Private Const SICK_DISPLAY_STATE As Integer = 46852

        '============================================================
        ' FRAME TREATMENT TYPE
        '============================================================

        Private Enum FrameTreatmentType

            Unknown = 0
            TrimExtend = 1
            Miter = 2
            Notch = 3
            LengthenShorten = 4
            EndCap = 5

        End Enum

        '============================================================
        ' SICK FRAME INFO
        '============================================================

        Private Class SickFrameInfo

            Public Property Node As Inventor.BrowserNode
            Public Property Label As String
            Public Property FullPath As String
            Public Property ToolTip As String
            Public Property NativeObject As Object
            Public Property TreatmentType As FrameTreatmentType
            Public Property DisplayState As Integer
            Public Property IsSick As Boolean

        End Class

        '============================================================
        ' DANH SÁCH
        '============================================================

        Private _sickTreatments As New List(Of SickFrameInfo)

        '============================================================
        ' GET INVENTOR APPLICATION
        '============================================================

        Private Function GetInventorApplication(
            ByVal Context As NameValueMap) As Inventor.Application

            Dim app As Inventor.Application = Nothing

            '--------------------------------------------------------
            ' CONTEXT
            '--------------------------------------------------------

            Try

                If Context IsNot Nothing Then

                    Try

                        app = DirectCast(
                            Context.Item("Application"),
                            Inventor.Application)

                    Catch
                    End Try

                End If

            Catch
            End Try

            If app IsNot Nothing Then
                Return app
            End If

            '--------------------------------------------------------
            ' FALLBACK
            '--------------------------------------------------------

            Try

                app = DirectCast(
                    Marshal.GetActiveObject(
                        "Inventor.Application"),
                    Inventor.Application)

            Catch ex As Exception

                MessageBox.Show(
                    "Không lấy được Inventor.Application." &
                    vbCrLf & vbCrLf &
                    ex.Message,
                    "Frame Generator",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error)

                Return Nothing

            End Try

            Return app

        End Function

        '============================================================
        ' MAIN
        '============================================================

        Public Sub OnExecute(
            ByVal Context As NameValueMap)

            Try

                _invApp =
                    GetInventorApplication(Context)

                If _invApp Is Nothing Then
                    Return
                End If

                Dim doc As Inventor.Document =
                    _invApp.ActiveDocument

                If doc Is Nothing Then

                    MessageBox.Show(
                        "Không có document đang mở.",
                        "Frame Generator",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning)

                    Return

                End If

                If doc.DocumentType <>
                   DocumentTypeEnum.kAssemblyDocumentObject Then

                    MessageBox.Show(
                        "Hãy chạy trong Assembly.",
                        "Frame Generator",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning)

                    Return

                End If

                Dim asmDoc As AssemblyDocument =
                    CType(doc, AssemblyDocument)

                '====================================================
                ' UPDATE ĐẦU
                '====================================================

                ForceUpdate(asmDoc)

                '====================================================
                ' SCAN
                '====================================================

                _sickTreatments.Clear()

                ScanFrameBrowser(doc)

                '====================================================
                ' KHÔNG CÓ TREATMENT
                '====================================================

                If _sickTreatments.Count = 0 Then

                    MessageBox.Show(
                        "Không tìm thấy Frame Treatment cần xử lý." &
                        vbCrLf & vbCrLf &
                        "Lưu ý:" &
                        vbCrLf &
                        "- End Cap trong Browser có tên: Sharp Corners" &
                        vbCrLf &
                        "- Lệnh tạo: Insert End Cap",
                        "Frame Generator",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information)

                    Return

                End If

                '====================================================
                ' THỐNG KÊ
                '====================================================

                Dim trimCount As Integer = 0
                Dim miterCount As Integer = 0
                Dim notchCount As Integer = 0
                Dim lengthenCount As Integer = 0
                Dim endCapCount As Integer = 0

                For Each item As SickFrameInfo In _sickTreatments

                    If item Is Nothing Then
                        Continue For
                    End If

                    Select Case item.TreatmentType

                        Case FrameTreatmentType.TrimExtend
                            trimCount += 1

                        Case FrameTreatmentType.Miter
                            miterCount += 1

                        Case FrameTreatmentType.Notch
                            notchCount += 1

                        Case FrameTreatmentType.LengthenShorten
                            lengthenCount += 1

                        Case FrameTreatmentType.EndCap
                            endCapCount += 1

                    End Select

                Next

                '====================================================
                ' LOG
                '====================================================

                Dim sb As New StringBuilder

                sb.AppendLine(
                    "========== FRAME GENERATOR ==========")

                sb.AppendLine()

                sb.AppendLine(
                    "FRAME TREATMENT ĐÃ PHÁT HIỆN:")

                sb.AppendLine()

                sb.AppendLine(
                    "Trim / Trim-Extend : " &
                    trimCount.ToString())

                sb.AppendLine(
                    "Miter / Mitre      : " &
                    miterCount.ToString())

                sb.AppendLine(
                    "Notch              : " &
                    notchCount.ToString())

                sb.AppendLine(
                    "Lengthen / Shorten : " &
                    lengthenCount.ToString())

                sb.AppendLine(
                    "End Cap / Sharp Corners : " &
                    endCapCount.ToString())

                sb.AppendLine()

                sb.AppendLine(
                    "--------------------------------------")

                For Each item As SickFrameInfo In _sickTreatments

                    If item Is Nothing Then
                        Continue For
                    End If

                    sb.AppendLine()

                    sb.AppendLine(
                        "TYPE: " &
                        GetTreatmentTypeName(
                            item.TreatmentType))

                    sb.AppendLine(
                        "NAME: " &
                        item.Label)

                    sb.AppendLine(
                        "PATH: " &
                        item.FullPath)

                    sb.AppendLine(
                        "DISPLAY STATE: " &
                        item.DisplayState.ToString())

                    sb.AppendLine(
                        "IS SICK: " &
                        item.IsSick.ToString())

                    If Not String.IsNullOrEmpty(
                        item.ToolTip) Then

                        sb.AppendLine(
                            "TOOLTIP: " &
                            item.ToolTip)

                    End If

                Next

                sb.AppendLine()

                sb.AppendLine(
                    "--------------------------------------")

                sb.AppendLine()

                sb.AppendLine(
                    "CÁCH XỬ LÝ:")

                sb.AppendLine()

                If trimCount > 0 Then

                    sb.AppendLine(
                        "Trim: DELETE node.")

                End If

                If miterCount > 0 Then

                    sb.AppendLine(
                        "Miter: DELETE node.")

                End If

                If notchCount > 0 Then

                    sb.AppendLine(
                        "Notch: DELETE node.")

                End If

                If lengthenCount > 0 Then

                    sb.AppendLine(
                        "Lengthen / Shorten: DELETE node.")

                End If

                If endCapCount > 0 Then

                    sb.AppendLine(
                        "Sharp Corners / End Cap: KHÔNG DELETE.")

                    sb.AppendLine(
                        "Sharp Corners / End Cap: CHỈ UPDATE / REBUILD.")

                End If

                sb.AppendLine()

                sb.AppendLine(
                    "Chỉ các Cut Treatment bị SICK mới được DELETE.")

                sb.AppendLine(
                    "Sharp Corners luôn được bảo vệ.")

                sb.AppendLine()

                sb.AppendLine(
                    "Bạn có muốn tiếp tục không?")

                Dim result As DialogResult =
                    MessageBox.Show(
                        sb.ToString(),
                        "FRAME GENERATOR - FIX",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Warning)

                If result <> DialogResult.Yes Then

                    MessageBox.Show(
                        "Đã hủy. Không thay đổi gì.",
                        "Frame Generator",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information)

                    Return

                End If

                '====================================================
                ' PROCESS
                '====================================================

                ProcessSickTreatments(asmDoc)

            Catch ex As Exception

                MessageBox.Show(
                    "LỖI:" &
                    vbCrLf & vbCrLf &
                    ex.Message &
                    vbCrLf & vbCrLf &
                    ex.StackTrace,
                    "Frame Generator Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error)

            End Try

        End Sub

        '============================================================
        ' FORCE UPDATE
        '============================================================

        Private Sub ForceUpdate(
            ByVal asmDoc As AssemblyDocument)

            If asmDoc Is Nothing Then
                Return
            End If

            Try
                asmDoc.Update2(True)
            Catch
            End Try

            Try
                asmDoc.Rebuild2()
            Catch
            End Try

            Try

                If _invApp IsNot Nothing Then
                    _invApp.ActiveView.Update()
                End If

            Catch
            End Try

        End Sub

        '============================================================
        ' SCAN FRAME BROWSER
        '============================================================

        Private Sub ScanFrameBrowser(
            ByVal doc As Inventor.Document)

            If doc Is Nothing Then
                Return
            End If

            Try

                Dim pane As Inventor.BrowserPane = Nothing

                Try

                    pane =
                        doc.BrowserPanes.Item("Model")

                Catch

                    Try

                        pane =
                            doc.BrowserPanes.ActivePane

                    Catch

                        pane = Nothing

                    End Try

                End Try

                If pane Is Nothing Then
                    Return
                End If

                If pane.TopNode Is Nothing Then
                    Return
                End If

                ScanBrowserNode(
                    pane.TopNode)

            Catch
            End Try

        End Sub

        '============================================================
        ' SCAN NODE
        '============================================================

        Private Sub ScanBrowserNode(
            ByVal node As Inventor.BrowserNode)

            If node Is Nothing Then
                Return
            End If

            Try

                Dim label As String = ""
                Dim fullPath As String = ""
                Dim tooltip As String = ""

                '----------------------------------------------------
                ' LABEL
                '----------------------------------------------------

                Try

                    label =
                        node.BrowserNodeDefinition.Label

                Catch

                    Try
                        label = node.FullPath
                    Catch
                        label = ""
                    End Try

                End Try

                '----------------------------------------------------
                ' PATH
                '----------------------------------------------------

                Try
                    fullPath = node.FullPath
                Catch
                    fullPath = label
                End Try

                '----------------------------------------------------
                ' TOOLTIP
                '----------------------------------------------------

                Try

                    tooltip =
                        node.BrowserNodeDefinition.
                        StateIconToolTipText

                Catch

                    tooltip = ""

                End Try

                '----------------------------------------------------
                ' DISPLAY STATE
                '----------------------------------------------------

                Dim state As Integer = -1

                Try

                    Dim ds As Object =
                        node.BrowserNodeDefinition.
                        DisplayState

                    state =
                        Convert.ToInt32(ds)

                Catch

                    state = -1

                End Try

                '----------------------------------------------------
                ' TREATMENT
                '----------------------------------------------------

                Dim treatmentType As FrameTreatmentType =
                    DetectTreatmentType(
                        label,
                        fullPath,
                        tooltip)

                '----------------------------------------------------
                ' END CAP / SHARP CORNERS
                '
                ' QUAN TRỌNG:
                '
                ' Sharp Corners là End Cap.
                '
                ' Không bắt buộc phải có SICK STATE
                ' mới nhận diện được End Cap.
                '----------------------------------------------------

                Dim isSharpCorners As Boolean =
                    IsSharpCornersNode(
                        label,
                        fullPath,
                        tooltip)

                '----------------------------------------------------
                ' SICK
                '----------------------------------------------------

                Dim isSick As Boolean = False

                If state =
                    SICK_DISPLAY_STATE Then

                    isSick = True

                End If

                '====================================================
                ' THÊM VÀO DANH SÁCH
                '
                ' CUT:
                '   chỉ thêm nếu SICK.
                '
                ' END CAP:
                '   Sharp Corners luôn được ghi nhận để debug/update.
                '====================================================

                If isSharpCorners Then

                    treatmentType =
                        FrameTreatmentType.EndCap

                End If

                If treatmentType =
                    FrameTreatmentType.EndCap Then

                    Dim info As New SickFrameInfo

                    info.Node = node
                    info.Label = label
                    info.FullPath = fullPath
                    info.ToolTip = tooltip
                    info.NativeObject = GetNativeObject(node)
                    info.TreatmentType =
                        FrameTreatmentType.EndCap
                    info.DisplayState = state
                    info.IsSick = isSick

                    If Not IsAlreadyAdded(node) Then
                        _sickTreatments.Add(info)
                    End If

                ElseIf isSick AndAlso
                       treatmentType <>
                       FrameTreatmentType.Unknown Then

                    Dim info As New SickFrameInfo

                    info.Node = node
                    info.Label = label
                    info.FullPath = fullPath
                    info.ToolTip = tooltip
                    info.NativeObject = GetNativeObject(node)
                    info.TreatmentType = treatmentType
                    info.DisplayState = state
                    info.IsSick = True

                    If Not IsAlreadyAdded(node) Then
                        _sickTreatments.Add(info)
                    End If

                End If

                '----------------------------------------------------
                ' CHILDREN
                '----------------------------------------------------

                Dim children As Object = Nothing

                Try

                    children =
                        node.BrowserNodes

                Catch

                    children = Nothing

                End Try

                If children IsNot Nothing Then

                    Try

                        For Each child As Object In children

                            Try

                                Dim childNode As Inventor.BrowserNode =
                                    DirectCast(
                                        child,
                                        Inventor.BrowserNode)

                                If childNode IsNot Nothing Then

                                    ScanBrowserNode(
                                        childNode)

                                End If

                            Catch
                            End Try

                        Next

                    Catch
                    End Try

                End If

            Catch
                ' Không để một node lỗi dừng toàn bộ scan.
            End Try

        End Sub

        '============================================================
        ' GET NATIVE OBJECT
        '============================================================

        Private Function GetNativeObject(
            ByVal node As Inventor.BrowserNode) As Object

            If node Is Nothing Then
                Return Nothing
            End If

            Try

                Return node.NativeObject

            Catch

                Return Nothing

            End Try

        End Function

        '============================================================
        ' SHARP CORNERS = END CAP
        '
        ' INVENTOR 2020
        '
        ' Browser:
        '     Sharp Corners
        '
        ' Command:
        '     Insert End Cap
        '============================================================

        Private Function IsSharpCornersNode(
            ByVal label As String,
            ByVal fullPath As String,
            ByVal tooltip As String) As Boolean

            Dim l As String =
                NormalizeText(label)

            Dim p As String =
                NormalizeText(fullPath)

            Dim t As String =
                NormalizeText(tooltip)

            '--------------------------------------------------------
            ' CHUẨN
            '--------------------------------------------------------

            If l.Contains("sharp corners") Then
                Return True
            End If

            If p.Contains("sharp corners") Then
                Return True
            End If

            If t.Contains("sharp corners") Then
                Return True
            End If

            '--------------------------------------------------------
            ' Một số trường hợp Inventor có thể bỏ khoảng trắng
            '--------------------------------------------------------

            If l.Contains("sharpcorners") Then
                Return True
            End If

            If p.Contains("sharpcorners") Then
                Return True
            End If

            If t.Contains("sharpcorners") Then
                Return True
            End If

            Return False

        End Function

        '============================================================
        ' NORMALIZE
        '============================================================

        Private Function NormalizeText(
            ByVal value As String) As String

            If value Is Nothing Then
                Return ""
            End If

            Return value.Trim().ToLowerInvariant()

        End Function

        '============================================================
        ' DETECT TREATMENT TYPE
        '============================================================

        Private Function DetectTreatmentType(
            ByVal label As String,
            ByVal fullPath As String,
            ByVal tooltip As String) As FrameTreatmentType

            Dim l As String =
                NormalizeText(label)

            Dim p As String =
                NormalizeText(fullPath)

            Dim t As String =
                NormalizeText(tooltip)

            '========================================================
            ' END CAP
            '
            ' PHẢI KIỂM TRA SHARP CORNERS TRƯỚC.
            '========================================================

            If IsSharpCornersNode(
                label,
                fullPath,
                tooltip) Then

                Return FrameTreatmentType.EndCap

            End If

            '========================================================
            ' TRIM
            '========================================================

            If l.Contains("trim") OrElse
               p.Contains("trim") OrElse
               t.Contains("trim") Then

                Return FrameTreatmentType.TrimExtend

            End If

            '========================================================
            ' MITER
            '========================================================

            If l.Contains("miter") OrElse
               l.Contains("mitre") OrElse
               p.Contains("miter") OrElse
               p.Contains("mitre") OrElse
               t.Contains("miter") OrElse
               t.Contains("mitre") Then

                Return FrameTreatmentType.Miter

            End If

            '========================================================
            ' NOTCH
            '========================================================

            If l.Contains("notch") OrElse
               p.Contains("notch") OrElse
               t.Contains("notch") Then

                Return FrameTreatmentType.Notch

            End If

            '========================================================
            ' LENGTHEN / SHORTEN
            '========================================================

            If l.Contains("lengthen") OrElse
               l.Contains("shorten") OrElse
               p.Contains("lengthen") OrElse
               p.Contains("shorten") OrElse
               t.Contains("lengthen") OrElse
               t.Contains("shorten") Then

                Return FrameTreatmentType.LengthenShorten

            End If

            '========================================================
            ' END CAP CŨ
            '
            ' Giữ thêm để phòng trường hợp một số Browser node
            ' dùng tên End Cap trực tiếp.
            '========================================================

            If l.Contains("end cap") OrElse
               l.Contains("endcap") OrElse
               p.Contains("end cap") OrElse
               p.Contains("endcap") OrElse
               t.Contains("end cap") OrElse
               t.Contains("endcap") Then

                Return FrameTreatmentType.EndCap

            End If

            Return FrameTreatmentType.Unknown

        End Function

        '============================================================
        ' TYPE NAME
        '============================================================

        Private Function GetTreatmentTypeName(
            ByVal treatmentType As FrameTreatmentType) As String

            Select Case treatmentType

                Case FrameTreatmentType.TrimExtend
                    Return "TRIM / TRIM-EXTEND"

                Case FrameTreatmentType.Miter
                    Return "MITER / MITRE"

                Case FrameTreatmentType.Notch
                    Return "NOTCH"

                Case FrameTreatmentType.LengthenShorten
                    Return "LENGTHEN / SHORTEN"

                Case FrameTreatmentType.EndCap
                    Return "END CAP / SHARP CORNERS"

                Case Else
                    Return "UNKNOWN"

            End Select

        End Function

        '============================================================
        ' CHECK DUPLICATE
        '============================================================

        Private Function IsAlreadyAdded(
            ByVal node As Inventor.BrowserNode) As Boolean

            If node Is Nothing Then
                Return False
            End If

            For Each item As SickFrameInfo In _sickTreatments

                If item Is Nothing Then
                    Continue For
                End If

                Try

                    If Object.ReferenceEquals(
                        item.Node,
                        node) Then

                        Return True

                    End If

                Catch
                End Try

            Next

            Return False

        End Function

        '============================================================
        ' PROCESS
        '============================================================

        Private Sub ProcessSickTreatments(
            ByVal asmDoc As AssemblyDocument)

            Dim deleteSuccess As Integer = 0
            Dim deleteFailed As Integer = 0
            Dim endCapUpdated As Integer = 0

            Dim resultLog As New StringBuilder

            resultLog.AppendLine(
                "========== FRAME GENERATOR FIX ==========")

            resultLog.AppendLine()

            '========================================================
            ' COPY LIST
            '========================================================

            Dim workList As New List(Of SickFrameInfo)

            For Each item As SickFrameInfo In _sickTreatments

                If item IsNot Nothing Then
                    workList.Add(item)
                End If

            Next

            '========================================================
            ' DELETE COMMAND
            '========================================================

            Dim deleteCmd As ControlDefinition = Nothing

            Try

                deleteCmd =
                    _invApp.CommandManager.
                    ControlDefinitions.Item("Delete")

            Catch ex As Exception

                resultLog.AppendLine(
                    "Không lấy được Delete command.")

                resultLog.AppendLine(
                    ex.Message)

            End Try

            '========================================================
            ' PROCESS
            '========================================================

            For i As Integer =
                workList.Count - 1 To 0 Step -1

                Dim item As SickFrameInfo =
                    workList(i)

                If item Is Nothing Then
                    Continue For
                End If

                resultLog.AppendLine(
                    "--------------------------------")

                resultLog.AppendLine(
                    "TYPE: " &
                    GetTreatmentTypeName(
                        item.TreatmentType))

                resultLog.AppendLine(
                    "NODE: " &
                    item.Label)

                resultLog.AppendLine(
                    "PATH: " &
                    item.FullPath)

                resultLog.AppendLine(
                    "STATE: " &
                    item.DisplayState.ToString())

                resultLog.AppendLine(
                    "SICK: " &
                    item.IsSick.ToString())

                '====================================================
                ' END CAP
                '
                ' SHARP CORNERS
                '
                ' TUYỆT ĐỐI KHÔNG DELETE.
                '====================================================

                If item.TreatmentType =
                    FrameTreatmentType.EndCap Then

                    resultLog.AppendLine(
                        "ACTION: KHÔNG DELETE")

                    resultLog.AppendLine(
                        "ACTION: UPDATE / REBUILD")

                    Try

                        asmDoc.Update2(True)

                        resultLog.AppendLine(
                            "Update2 = OK")

                    Catch ex As Exception

                        resultLog.AppendLine(
                            "Update2 lỗi: " &
                            ex.Message)

                    End Try

                    Try

                        asmDoc.Rebuild2()

                        resultLog.AppendLine(
                            "Rebuild2 = OK")

                    Catch ex As Exception

                        resultLog.AppendLine(
                            "Rebuild2 lỗi: " &
                            ex.Message)

                    End Try

                    Try

                        If _invApp IsNot Nothing Then

                            _invApp.ActiveView.Update()

                        End If

                    Catch
                    End Try

                    endCapUpdated += 1

                    resultLog.AppendLine(
                        "=> SHARP CORNERS / END CAP ĐÃ UPDATE")

                    Continue For

                End If

                '====================================================
                ' CUT TREATMENT
                '
                ' CHỈ DELETE NẾU THỰC SỰ SICK.
                '====================================================

                If Not item.IsSick Then

                    resultLog.AppendLine(
                        "=> KHÔNG SICK - BỎ QUA DELETE")

                    Continue For

                End If

                If item.TreatmentType <>
                   FrameTreatmentType.TrimExtend AndAlso
                   item.TreatmentType <>
                   FrameTreatmentType.Miter AndAlso
                   item.TreatmentType <>
                   FrameTreatmentType.Notch AndAlso
                   item.TreatmentType <>
                   FrameTreatmentType.LengthenShorten Then

                    Continue For

                End If

                '====================================================
                ' DELETE COMMAND
                '====================================================

                If deleteCmd Is Nothing Then

                    deleteFailed += 1

                    resultLog.AppendLine(
                        "=> FAIL: Delete command Nothing.")

                    Continue For

                End If

                '====================================================
                ' CLEAR SELECTION
                '====================================================

                Try
                    asmDoc.SelectSet.Clear()
                Catch
                End Try

                Try

                    Dim pane As BrowserPane =
                        asmDoc.BrowserPanes.Item("Model")

                    Try
                        pane.ClearSelection()
                    Catch
                    End Try

                Catch
                End Try

                '====================================================
                ' SELECT NODE
                '====================================================

                Dim selected As Boolean = False

                Try

                    item.Node.Select()

                    selected = True

                    resultLog.AppendLine(
                        "Node.Select = OK")

                Catch ex As Exception

                    resultLog.AppendLine(
                        "Node.Select lỗi: " &
                        ex.Message)

                End Try

                If Not selected Then

                    Try

                        item.Node.DoSelect()

                        selected = True

                        resultLog.AppendLine(
                            "Node.DoSelect = OK")

                    Catch ex As Exception

                        resultLog.AppendLine(
                            "Node.DoSelect lỗi: " &
                            ex.Message)

                    End Try

                End If

                If Not selected Then

                    deleteFailed += 1

                    resultLog.AppendLine(
                        "=> FAIL: Không select được node.")

                    Continue For

                End If

                '====================================================
                ' DELETE
                '====================================================

                Try

                    If Not deleteCmd.Enabled Then

                        deleteFailed += 1

                        resultLog.AppendLine(
                            "=> Delete command Disabled.")

                        Continue For

                    End If

                Catch
                End Try

                Try

                    deleteCmd.Execute()

                    resultLog.AppendLine(
                        "Delete.Execute = OK")

                Catch ex As Exception

                    deleteFailed += 1

                    resultLog.AppendLine(
                        "Delete lỗi: " &
                        ex.Message)

                    Continue For

                End Try

                '====================================================
                ' UPDATE
                '====================================================

                Try
                    asmDoc.Update2(True)
                Catch
                End Try

                Try

                    If _invApp IsNot Nothing Then
                        _invApp.ActiveView.Update()
                    End If

                Catch
                End Try

                '====================================================
                ' CHECK
                '====================================================

                Dim stillExists As Boolean =
                    BrowserNodeStillExists(
                        asmDoc,
                        item.FullPath)

                If stillExists Then

                    deleteFailed += 1

                    resultLog.AppendLine(
                        "=> DELETE CHƯA THÀNH CÔNG")

                Else

                    deleteSuccess += 1

                    resultLog.AppendLine(
                        "=> DELETE THÀNH CÔNG")

                End If

            Next

            '========================================================
            ' FINAL UPDATE
            '========================================================

            Try
                asmDoc.Update2(True)
            Catch
            End Try

            Try
                asmDoc.Rebuild2()
            Catch
            End Try

            Try

                If _invApp IsNot Nothing Then
                    _invApp.ActiveView.Update()
                End If

            Catch
            End Try

            '========================================================
            ' SCAN LẠI
            '========================================================

            _sickTreatments.Clear()

            ScanFrameBrowser(asmDoc)

            Dim remainTrim As Integer = 0
            Dim remainMiter As Integer = 0
            Dim remainNotch As Integer = 0
            Dim remainLengthen As Integer = 0
            Dim remainEndCap As Integer = 0

            For Each item As SickFrameInfo In _sickTreatments

                If item Is Nothing Then
                    Continue For
                End If

                Select Case item.TreatmentType

                    Case FrameTreatmentType.TrimExtend
                        remainTrim += 1

                    Case FrameTreatmentType.Miter
                        remainMiter += 1

                    Case FrameTreatmentType.Notch
                        remainNotch += 1

                    Case FrameTreatmentType.LengthenShorten
                        remainLengthen += 1

                    Case FrameTreatmentType.EndCap
                        remainEndCap += 1

                End Select

            Next

            '========================================================
            ' RESULT
            '========================================================

            resultLog.AppendLine()

            resultLog.AppendLine(
                "========================================")

            resultLog.AppendLine(
                "KẾT QUẢ")

            resultLog.AppendLine(
                "========================================")

            resultLog.AppendLine()

            resultLog.AppendLine(
                "Delete thành công : " &
                deleteSuccess.ToString())

            resultLog.AppendLine(
                "Delete thất bại   : " &
                deleteFailed.ToString())

            resultLog.AppendLine(
                "End Cap Update    : " &
                endCapUpdated.ToString())

            resultLog.AppendLine()

            resultLog.AppendLine(
                "SAU KHI XỬ LÝ:")

            resultLog.AppendLine(
                "Trim              : " &
                remainTrim.ToString())

            resultLog.AppendLine(
                "Miter              : " &
                remainMiter.ToString())

            resultLog.AppendLine(
                "Notch              : " &
                remainNotch.ToString())

            resultLog.AppendLine(
                "Lengthen/Shorten   : " &
                remainLengthen.ToString())

            resultLog.AppendLine(
                "Sharp Corners      : " &
                remainEndCap.ToString())

            resultLog.AppendLine()

            resultLog.AppendLine(
                "LƯU Ý: Sharp Corners = End Cap.")

            resultLog.AppendLine(
                "Sharp Corners KHÔNG BAO GIỜ DELETE.")

            '========================================================
            ' CLEAR
            '========================================================

            _sickTreatments.Clear()

            '========================================================
            ' SHOW
            '========================================================

            ShowDebugText(
                resultLog.ToString())

        End Sub

        '============================================================
        ' CHECK NODE STILL EXISTS
        '============================================================

        Private Function BrowserNodeStillExists(
            ByVal doc As Inventor.Document,
            ByVal targetPath As String) As Boolean

            If doc Is Nothing Then
                Return False
            End If

            If String.IsNullOrEmpty(targetPath) Then
                Return False
            End If

            Try

                Dim pane As BrowserPane =
                    doc.BrowserPanes.Item("Model")

                If pane Is Nothing Then
                    Return False
                End If

                Return SearchBrowserPath(
                    pane.TopNode,
                    targetPath)

            Catch

                Return False

            End Try

        End Function

        '============================================================
        ' SEARCH BROWSER PATH
        '============================================================

        Private Function SearchBrowserPath(
            ByVal node As Inventor.BrowserNode,
            ByVal targetPath As String) As Boolean

            If node Is Nothing Then
                Return False
            End If

            Try

                Dim currentPath As String = ""

                Try
                    currentPath = node.FullPath
                Catch
                End Try

                If String.Equals(
                    currentPath,
                    targetPath,
                    StringComparison.OrdinalIgnoreCase) Then

                    Return True

                End If

                Dim children As Object = Nothing

                Try
                    children = node.BrowserNodes
                Catch
                    children = Nothing
                End Try

                If children IsNot Nothing Then

                    Try

                        For Each child As Object In children

                            Try

                                Dim childNode As Inventor.BrowserNode =
                                    DirectCast(
                                        child,
                                        Inventor.BrowserNode)

                                If SearchBrowserPath(
                                    childNode,
                                    targetPath) Then

                                    Return True

                                End If

                            Catch
                            End Try

                        Next

                    Catch
                    End Try

                End If

            Catch
            End Try

            Return False

        End Function

        '============================================================
        ' ON DEBUG
        '
        ' DEBUG TOÀN BỘ BROWSER
        '
        ' ĐẶC BIỆT ĐÁNH DẤU:
        '
        ' Sharp Corners
        ' End Cap
        ' Sick
        '============================================================

        Public Sub OnDebug(
            ByVal Context As NameValueMap)

            Try

                _invApp =
                    GetInventorApplication(Context)

                If _invApp Is Nothing Then
                    Return
                End If

                Dim doc As Inventor.Document =
                    _invApp.ActiveDocument

                If doc Is Nothing Then

                    MessageBox.Show(
                        "Không có document đang mở.",
                        "Frame Debug",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning)

                    Return

                End If

                Dim sb As New StringBuilder

                sb.AppendLine(
                    "============================================================")

                sb.AppendLine(
                    " FRAME GENERATOR - SHARP CORNERS / END CAP DEBUG")

                sb.AppendLine(
                    " INVENTOR 2020")

                sb.AppendLine(
                    "============================================================")

                sb.AppendLine()

                sb.AppendLine(
                    "DOCUMENT: " &
                    doc.DisplayName)

                sb.AppendLine()

                sb.AppendLine(
                    "QUY ƯỚC:")

                sb.AppendLine(
                    "Sharp Corners = END CAP")

                sb.AppendLine(
                    "Command = Insert End Cap")

                sb.AppendLine()

                DebugBrowser(
                    doc,
                    sb)

                sb.AppendLine()

                sb.AppendLine(
                    "============================================================")

                sb.AppendLine(
                    " KẾT THÚC DEBUG")

                sb.AppendLine(
                    "============================================================")

                ShowDebugText(
                    sb.ToString())

            Catch ex As Exception

                MessageBox.Show(
                    ex.Message &
                    vbCrLf & vbCrLf &
                    ex.StackTrace,
                    "Frame Debug",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error)

            End Try

        End Sub

        '============================================================
        ' DEBUG BROWSER
        '============================================================

        Private Sub DebugBrowser(
            ByVal doc As Inventor.Document,
            ByVal sb As StringBuilder)

            Try

                Dim pane As Inventor.BrowserPane =
                    Nothing

                Try

                    pane =
                        doc.BrowserPanes.Item("Model")

                Catch

                    Try

                        pane =
                            doc.BrowserPanes.ActivePane

                    Catch

                        pane = Nothing

                    End Try

                End Try

                If pane Is Nothing Then

                    sb.AppendLine(
                        "Không lấy được Model Browser.")

                    Return

                End If

                If pane.TopNode Is Nothing Then

                    sb.AppendLine(
                        "TopNode = Nothing.")

                    Return

                End If

                sb.AppendLine(
                    "MODEL BROWSER: OK")

                sb.AppendLine()

                DebugBrowserNode(
                    pane.TopNode,
                    sb,
                    0)

            Catch ex As Exception

                sb.AppendLine(
                    "BROWSER ERROR: " &
                    ex.Message)

            End Try

        End Sub

        '============================================================
        ' DEBUG NODE
        '============================================================

        Private Sub DebugBrowserNode(
            ByVal node As Inventor.BrowserNode,
            ByVal sb As StringBuilder,
            ByVal level As Integer)

            If node Is Nothing Then
                Return
            End If

            Try

                Dim indent As String =
                    New String(
                        " "c,
                        level * 2)

                Dim label As String = ""
                Dim path As String = ""
                Dim tooltip As String = ""

                Dim state As Integer = -1

                Dim nativeType As String = ""

                '----------------------------------------------------
                ' LABEL
                '----------------------------------------------------

                Try

                    label =
                        node.BrowserNodeDefinition.Label

                Catch

                    Try
                        label = node.FullPath
                    Catch
                        label = "?"
                    End Try

                End Try

                '----------------------------------------------------
                ' PATH
                '----------------------------------------------------

                Try
                    path = node.FullPath
                Catch
                    path = ""
                End Try

                '----------------------------------------------------
                ' TOOLTIP
                '----------------------------------------------------

                Try

                    tooltip =
                        node.BrowserNodeDefinition.
                        StateIconToolTipText

                Catch

                    tooltip = ""

                End Try

                '----------------------------------------------------
                ' STATE
                '----------------------------------------------------

                Try

                    Dim ds As Object =
                        node.BrowserNodeDefinition.
                        DisplayState

                    state =
                        Convert.ToInt32(ds)

                Catch

                    state = -1

                End Try

                '----------------------------------------------------
                ' NATIVE TYPE
                '----------------------------------------------------

                Try

                    Dim obj As Object =
                        node.NativeObject

                    If obj IsNot Nothing Then

                        nativeType =
                            obj.GetType().FullName

                    End If

                Catch

                    nativeType = ""

                End Try

                '----------------------------------------------------
                ' DETECT
                '----------------------------------------------------

                Dim treatmentType As FrameTreatmentType =
                    DetectTreatmentType(
                        label,
                        path,
                        tooltip)

                Dim isSharpCorners As Boolean =
                    IsSharpCornersNode(
                        label,
                        path,
                        tooltip)

                Dim isSick As Boolean =
                    (state = SICK_DISPLAY_STATE)

                '====================================================
                ' CHỈ IN NODE QUAN TRỌNG
                '
                ' Sharp Corners
                ' End Cap
                ' Treatment
                ' Sick
                '====================================================

                Dim important As Boolean =
                    isSharpCorners OrElse
                    treatmentType <>
                    FrameTreatmentType.Unknown OrElse
                    isSick

                If important Then

                    sb.AppendLine()

                    sb.AppendLine(
                        indent &
                        "----------------------------------------")

                    sb.AppendLine(
                        indent &
                        "NAME: " &
                        label)

                    sb.AppendLine(
                        indent &
                        "TYPE: " &
                        GetTreatmentTypeName(
                            treatmentType))

                    sb.AppendLine(
                        indent &
                        "STATE: " &
                        state.ToString())

                    sb.AppendLine(
                        indent &
                        "SICK: " &
                        isSick.ToString())

                    sb.AppendLine(
                        indent &
                        "SHARP CORNERS: " &
                        isSharpCorners.ToString())

                    sb.AppendLine(
                        indent &
                        "TOOLTIP: " &
                        tooltip)

                    sb.AppendLine(
                        indent &
                        "NATIVE: " &
                        nativeType)

                    sb.AppendLine(
                        indent &
                        "PATH: " &
                        path)

                    If isSharpCorners Then

                        sb.AppendLine(
                            indent &
                            ">>> END CAP DETECTED <<<")

                        sb.AppendLine(
                            indent &
                            ">>> BROWSER NAME = SHARP CORNERS <<<")

                        sb.AppendLine(
                            indent &
                            ">>> COMMAND = INSERT END CAP <<<")

                    End If

                    If isSick Then

                        sb.AppendLine(
                            indent &
                            ">>> SICK NODE <<<")

                    End If

                End If

                '----------------------------------------------------
                ' CHILDREN
                '----------------------------------------------------

                Dim children As Object = Nothing

                Try
                    children = node.BrowserNodes
                Catch
                    children = Nothing
                End Try

                If children IsNot Nothing Then

                    Try

                        For Each child As Object In children

                            Try

                                Dim childNode As Inventor.BrowserNode =
                                    DirectCast(
                                        child,
                                        Inventor.BrowserNode)

                                If childNode IsNot Nothing Then

                                    DebugBrowserNode(
                                        childNode,
                                        sb,
                                        level + 1)

                                End If

                            Catch
                            End Try

                        Next

                    Catch
                    End Try

                End If

            Catch
            End Try

        End Sub

        '============================================================
        ' DEBUG FORM
        '============================================================

        Private Sub ShowDebugText(
            ByVal text As String)

            Dim f As New Windows.Forms.Form

            f.Text =
                "FRAME GENERATOR DEBUG - SHARP CORNERS / END CAP"

            f.Width = 1400
            f.Height = 850

            f.StartPosition =
                Windows.Forms.FormStartPosition.CenterScreen

            Dim tb As New Windows.Forms.TextBox

            tb.Multiline = True
            tb.ReadOnly = True

            tb.ScrollBars =
                Windows.Forms.ScrollBars.Both

            tb.WordWrap = False

            tb.Dock =
                Windows.Forms.DockStyle.Fill

            tb.Font =
                New System.Drawing.Font(
                    "Consolas",
                    9.0F)

            tb.Text = text

            f.Controls.Add(tb)

            f.ShowDialog()

        End Sub

    End Module

End Namespace

