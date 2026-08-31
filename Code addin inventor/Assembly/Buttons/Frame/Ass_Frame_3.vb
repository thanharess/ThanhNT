
Option Explicit On
Option Strict Off

Imports System
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Windows.Forms
Imports Inventor

Namespace ThanhN.Assembly.Buttons.Frame

    Public Module Ass_Frame_3

        Private _invApp As Inventor.Application = Nothing

        '============================================================
        ' MAIN BUTTON
        '============================================================

        Public Sub OnExecute(ByVal Context As NameValueMap)

            Try

                _invApp = GetInventorApplication(Context)

                If _invApp Is Nothing Then
                    Return
                End If

                Dim doc As Inventor.Document =
                    _invApp.ActiveDocument

                If doc Is Nothing Then

                    MessageBox.Show(
                        "Không có document đang mở.",
                        "Frame Generator Debug",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning)

                    Return

                End If

                Dim sb As New StringBuilder

                sb.AppendLine(
                    "============================================================")

                sb.AppendLine(
                    " FRAME GENERATOR BROWSER DEBUG - INVENTOR 2020")

                sb.AppendLine(
                    "============================================================")

                sb.AppendLine()

                sb.AppendLine(
                    "DOCUMENT: " &
                    doc.DisplayName)

                sb.AppendLine()

                sb.AppendLine(
                    "MODEL BROWSER: ")

                '----------------------------------------------------
                ' GET MODEL BROWSER
                '----------------------------------------------------

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

                    sb.AppendLine(
                        "KHÔNG LẤY ĐƯỢC MODEL BROWSER.")

                ElseIf pane.TopNode Is Nothing Then

                    sb.AppendLine(
                        "MODEL BROWSER TOP NODE = NOTHING.")

                Else

                    sb.AppendLine(
                        "OK")

                    sb.AppendLine()

                    sb.AppendLine(
                        "BẮT ĐẦU QUÉT TOÀN BỘ NODE...")

                    sb.AppendLine()

                    DebugBrowserNode(
                        pane.TopNode,
                        sb,
                        0)

                End If

                sb.AppendLine()

                sb.AppendLine(
                    "============================================================")

                sb.AppendLine(
                    " KẾT THÚC SCAN")

                sb.AppendLine(
                    "============================================================")

                ShowDebugText(
                    sb.ToString())

            Catch ex As Exception

                MessageBox.Show(
                    "DEBUG ERROR:" &
                    vbCrLf &
                    vbCrLf &
                    ex.Message,
                    "Frame Generator Debug",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error)

            End Try

        End Sub


        '============================================================
        ' GET INVENTOR
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

                        app =
                            DirectCast(
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
            ' ACTIVE OBJECT
            '--------------------------------------------------------

            Try

                app =
                    DirectCast(
                        Marshal.GetActiveObject(
                            "Inventor.Application"),
                        Inventor.Application)

            Catch ex As Exception

                MessageBox.Show(
                    "Không lấy được Inventor.Application." &
                    vbCrLf &
                    vbCrLf &
                    ex.Message,
                    "Frame Debug",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error)

                Return Nothing

            End Try

            Return app

        End Function


        '============================================================
        ' DEBUG BROWSER NODE
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

                '----------------------------------------------------
                ' VARIABLES
                '----------------------------------------------------

                Dim label As String = ""
                Dim path As String = ""
                Dim tooltip As String = ""

                Dim nativeObj As Object = Nothing

                Dim nativeName As String = ""
                Dim nativeType As String = ""

                Dim definitionType As String = ""

                Dim health As String = ""
                Dim displayState As String = ""

                Dim childCount As Integer = 0

                '----------------------------------------------------
                ' LABEL
                '----------------------------------------------------

                Try

                    label =
                        node.BrowserNodeDefinition.Label

                Catch

                    Try

                        label =
                            node.FullPath

                    Catch

                        label = ""

                    End Try

                End Try

                '----------------------------------------------------
                ' PATH
                '----------------------------------------------------

                Try

                    path =
                        node.FullPath

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
                ' DEFINITION TYPE
                '----------------------------------------------------

                Try

                    If node.BrowserNodeDefinition IsNot Nothing Then

                        definitionType =
                            node.BrowserNodeDefinition.
                            GetType().
                            FullName

                    End If

                Catch

                    definitionType = ""

                End Try

                '----------------------------------------------------
                ' NATIVE OBJECT
                '----------------------------------------------------

                Try

                    nativeObj =
                        node.NativeObject

                Catch

                    nativeObj = Nothing

                End Try

                If nativeObj IsNot Nothing Then

                    Try

                        nativeType =
                            nativeObj.GetType().FullName

                    Catch

                        nativeType = ""

                    End Try

                    Try

                        nativeName =
                            Convert.ToString(
                                CallByName(
                                    nativeObj,
                                    "Name",
                                    CallType.Get))

                    Catch

                        nativeName = ""

                    End Try

                End If

                '----------------------------------------------------
                ' HEALTH
                '----------------------------------------------------

                Try

                    If nativeObj IsNot Nothing Then

                        Dim h As Object = Nothing

                        Try

                            h =
                                CallByName(
                                    nativeObj,
                                    "HealthStatus",
                                    CallType.Get)

                        Catch

                            h = Nothing

                        End Try

                        If h IsNot Nothing Then

                            health =
                                Convert.ToString(h)

                        End If

                    End If

                Catch

                    health = ""

                End Try

                '----------------------------------------------------
                ' DISPLAY STATE
                '----------------------------------------------------

                Try

                    Dim ds As Object =
                        node.BrowserNodeDefinition.
                        DisplayState

                    displayState =
                        Convert.ToString(ds)

                Catch

                    displayState = ""

                End Try

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

                        childCount =
                            Convert.ToInt32(
                                children.Count)

                    Catch

                        childCount = 0

                    End Try

                End If

                '----------------------------------------------------
                ' KEYWORD
                '----------------------------------------------------

                Dim allText As String =
                    (
                        label & " " &
                        path & " " &
                        tooltip & " " &
                        nativeName & " " &
                        nativeType & " " &
                        definitionType
                    ).ToLowerInvariant()

                Dim isEndCap As Boolean = False

                If allText.Contains("end cap") Then
                    isEndCap = True
                End If

                If allText.Contains("endcap") Then
                    isEndCap = True
                End If

                If allText.Contains("end treatment") Then
                    isEndCap = True
                End If

                If allText.Contains("endtreatment") Then
                    isEndCap = True
                End If

                If allText.Contains("cap") Then
                    isEndCap = True
                End If

                '----------------------------------------------------
                ' PRINT NODE
                '----------------------------------------------------

                sb.AppendLine(
                    indent &
                    "================================================")

                sb.AppendLine(
                    indent &
                    "NODE")

                sb.AppendLine(
                    indent &
                    "LABEL: " &
                    label)

                sb.AppendLine(
                    indent &
                    "PATH: " &
                    path)

                sb.AppendLine(
                    indent &
                    "NATIVE TYPE: " &
                    nativeType)

                sb.AppendLine(
                    indent &
                    "NATIVE NAME: " &
                    nativeName)

                sb.AppendLine(
                    indent &
                    "HEALTH: " &
                    health)

                sb.AppendLine(
                    indent &
                    "DISPLAY STATE: " &
                    displayState)

                sb.AppendLine(
                    indent &
                    "BROWSER NODE TYPE: " &
                    definitionType)

                sb.AppendLine(
                    indent &
                    "TOOLTIP: " &
                    tooltip)

                sb.AppendLine(
                    indent &
                    "CHILD COUNT: " &
                    childCount.ToString())

                '----------------------------------------------------
                ' END CAP CANDIDATE
                '----------------------------------------------------

                If isEndCap Then

                    sb.AppendLine()

                    sb.AppendLine(
                        indent &
                        "*************** END CAP CANDIDATE ***************")

                    sb.AppendLine(
                        indent &
                        "END CAP DETECTED = TRUE")

                    sb.AppendLine(
                        indent &
                        "**************************************************")

                End If

                '----------------------------------------------------
                ' PROPERTY DEBUG
                ' CHỈ KHI CÓ NATIVE OBJECT
                '----------------------------------------------------

                If nativeObj IsNot Nothing Then

                    DebugProperties(
                        nativeObj,
                        sb,
                        level)

                End If

                sb.AppendLine()

                '----------------------------------------------------
                ' CHILDREN
                '----------------------------------------------------

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

                            Catch exChild As Exception

                                sb.AppendLine(
                                    indent &
                                    "CHILD ERROR: " &
                                    exChild.Message)

                            End Try

                        Next

                    Catch exChildren As Exception

                        sb.AppendLine(
                            indent &
                            "CHILD COLLECTION ERROR: " &
                            exChildren.Message)

                    End Try

                End If

            Catch ex As Exception

                sb.AppendLine(
                    New String(
                        " "c,
                        level * 2) &
                    "NODE ERROR: " &
                    ex.Message)

            End Try

        End Sub


        '============================================================
        ' DEBUG PROPERTIES
        '
        ' QUAN TRỌNG:
        ' Không dùng:
        '
        ' For Each p As Property
        '
        ' Dùng Object để tương thích Inventor 2020
        '============================================================

        Private Sub DebugProperties(
            ByVal nativeObj As Object,
            ByVal sb As StringBuilder,
            ByVal level As Integer)

            If nativeObj Is Nothing Then
                Return
            End If

            Dim indent As String =
                New String(
                    " "c,
                    (level + 1) * 2)

            Try

                Dim propSets As Object = Nothing

                Try

                    propSets =
                        CallByName(
                            nativeObj,
                            "PropertySets",
                            CallType.Get)

                Catch

                    propSets = Nothing

                End Try

                If propSets Is Nothing Then
                    Return
                End If

                Dim setCount As Integer = 0

                Try

                    setCount =
                        Convert.ToInt32(
                            CallByName(
                                propSets,
                                "Count",
                                CallType.Get))

                Catch

                    setCount = 0

                End Try

                For setIndex As Integer =
                    1 To setCount

                    Dim propSet As Object = Nothing

                    Try

                        propSet =
                            CallByName(
                                propSets,
                                "Item",
                                CallType.Method,
                                setIndex)

                    Catch

                        propSet = Nothing

                    End Try

                    If propSet Is Nothing Then
                        Continue For
                    End If

                    Dim setName As String = ""

                    Try

                        setName =
                            Convert.ToString(
                                CallByName(
                                    propSet,
                                    "Name",
                                    CallType.Get))

                    Catch

                        setName = ""

                    End Try

                    '------------------------------------------------
                    ' CHỈ QUAN TÂM PROPERTY SET LIÊN QUAN
                    '------------------------------------------------

                    Dim setLower As String =
                        setName.ToLowerInvariant()

                    If setLower.Contains("user defined") OrElse
                       setLower.Contains("frame") OrElse
                       setLower.Contains("inventor") Then

                        sb.AppendLine(
                            indent &
                            "PROPERTY SET: " &
                            setName)

                        Dim propCount As Integer = 0

                        Try

                            propCount =
                                Convert.ToInt32(
                                    CallByName(
                                        propSet,
                                        "Count",
                                        CallType.Get))

                        Catch

                            propCount = 0

                        End Try

                        For propIndex As Integer =
                            1 To propCount

                            Dim p As Object = Nothing

                            Try

                                p =
                                    CallByName(
                                        propSet,
                                        "Item",
                                        CallType.Method,
                                        propIndex)

                            Catch

                                p = Nothing

                            End Try

                            If p Is Nothing Then
                                Continue For
                            End If

                            Dim pName As String = ""
                            Dim pValue As String = ""

                            Try

                                pName =
                                    Convert.ToString(
                                        CallByName(
                                            p,
                                            "Name",
                                            CallType.Get))

                            Catch

                                pName = ""

                            End Try

                            Try

                                pValue =
                                    Convert.ToString(
                                        CallByName(
                                            p,
                                            "Value",
                                            CallType.Get))

                            Catch

                                pValue = ""

                            End Try

                            sb.AppendLine(
                                indent &
                                "  PROPERTY: " &
                                pName &
                                " = " &
                                pValue)

                            '----------------------------------------
                            ' CUTDETAIL
                            '----------------------------------------

                            If pName.ToUpperInvariant().
                                Contains("CUTDETAIL") Then

                                sb.AppendLine(
                                    indent &
                                    "  >>> CUTDETAIL FOUND <<<")

                                sb.AppendLine(
                                    indent &
                                    "  NAME: " &
                                    pName)

                                sb.AppendLine(
                                    indent &
                                    "  VALUE: " &
                                    pValue)

                            End If

                        Next

                    End If

                Next

            Catch ex As Exception

                sb.AppendLine(
                    indent &
                    "PROPERTY ERROR: " &
                    ex.Message)

            End Try

        End Sub


        '============================================================
        ' DEBUG FORM
        '============================================================

        Private Sub ShowDebugText(
            ByVal text As String)

            Dim f As New Windows.Forms.Form

            f.Text =
                "FRAME GENERATOR - END CAP DEBUG"

            f.Width = 1500

            f.Height = 900

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

            tb.Text =
                text

            f.Controls.Add(tb)

            f.ShowDialog()

        End Sub

    End Module

End Namespace
