Option Explicit On
Option Strict Off

Imports Inventor
Imports System
Imports System.Drawing
Imports System.Windows.Forms

Namespace ThanhN.Drawing.Buttons

    Public Module Draw_7

        '=========================================================
        ' BUTTON EXECUTE
        '=========================================================
        Public Sub OnExecute(ByVal Context As NameValueMap)

            Dim invApp As Inventor.Application = Nothing

            Try

                invApp = g_inventorApplication

                '=================================================
                ' KIỂM TRA INVENTOR
                '=================================================
                If invApp Is Nothing Then
                    MessageBox.Show(
                        "Không tìm thấy Inventor Application.",
                        "Sheet Navigator",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error)

                    Exit Sub
                End If


                '=================================================
                ' KIỂM TRA DOCUMENT
                '=================================================
                If invApp.ActiveDocument Is Nothing Then

                    MessageBox.Show(
                        "Không có tài liệu đang mở.",
                        "Sheet Navigator",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error)

                    Exit Sub

                End If


                If invApp.ActiveDocument.DocumentType <>
                   Inventor.DocumentTypeEnum.kDrawingDocumentObject Then

                    MessageBox.Show(
                        "Chỉ sử dụng chức năng này trong Drawing.",
                        "Sheet Navigator",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning)

                    Exit Sub

                End If


                '=================================================
                ' TẠO FORM
                '=================================================
                Dim frm As New SheetNavigatorForm(invApp)

                frm.Show()


            Catch ex As Exception

                MessageBox.Show(
                    "Lỗi Sheet Navigator:" &
                    vbCrLf &
                    ex.Message,
                    "Sheet Navigator",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error)

            End Try

        End Sub

    End Module


    '=============================================================
    ' SHEET NAVIGATOR FORM
    ' INVENTOR 2020
    '=============================================================
    Public Class SheetNavigatorForm

        Inherits Form


        '=========================================================
        ' INVENTOR
        '=========================================================
        Private invApp As Inventor.Application

        Private oDrawDoc As Inventor.DrawingDocument


        '=========================================================
        ' CONTROLS
        '=========================================================
        Private lblInfo As System.Windows.Forms.Label

        Private txtPage As System.Windows.Forms.TextBox


        Private btnFirst As System.Windows.Forms.Button
        Private btnPrev As System.Windows.Forms.Button
        Private btnNext As System.Windows.Forms.Button
        Private btnLast As System.Windows.Forms.Button
        Private btnGo As System.Windows.Forms.Button
        Private btnClose As System.Windows.Forms.Button


        Private chkLeft As System.Windows.Forms.CheckBox
        Private chkRight As System.Windows.Forms.CheckBox


        '=========================================================
        ' DRAG FORM
        '=========================================================
        Private isDragging As Boolean = False

        Private dragCursorPoint As System.Drawing.Point

        Private dragFormPoint As System.Drawing.Point


        '=========================================================
        ' CONSTRUCTOR
        '=========================================================
        Public Sub New(
            ByVal app As Inventor.Application)

            MyBase.New()


            '=====================================================
            ' INVENTOR APPLICATION
            '=====================================================
            invApp = app


            '=====================================================
            ' DRAWING DOCUMENT
            '=====================================================
            If invApp Is Nothing Then
                Throw New Exception(
                    "Inventor Application không hợp lệ.")
            End If


            If invApp.ActiveDocument Is Nothing Then
                Throw New Exception(
                    "Không có document đang mở.")
            End If


            If invApp.ActiveDocument.DocumentType <>
               Inventor.DocumentTypeEnum.kDrawingDocumentObject Then

                Throw New Exception(
                    "Document hiện tại không phải Drawing.")

            End If


            oDrawDoc =
                CType(
                    invApp.ActiveDocument,
                    Inventor.DrawingDocument)


            '=====================================================
            ' FORM
            '=====================================================
            Me.Text = ""

            Me.Width = 250

            Me.Height = 95


            Me.StartPosition =
                FormStartPosition.CenterScreen


            ' Không header
            Me.FormBorderStyle =
                FormBorderStyle.None


            ' Luôn hiện
            Me.TopMost = True


            ' Không hiện taskbar
            Me.ShowInTaskbar = False


            ' Trong suốt nhẹ
            Me.Opacity = 0.7


            ' Màu nền tối
            Me.BackColor =
                System.Drawing.Color.FromArgb(
                    45,
                    45,
                    48)


            '=====================================================
            ' INFO
            '=====================================================
            lblInfo =
                New System.Windows.Forms.Label()


            lblInfo.Left = 10

            lblInfo.Top = 8

            lblInfo.Width = 150

            lblInfo.Height = 18


            lblInfo.ForeColor =
                System.Drawing.Color.White


            lblInfo.Font =
                New System.Drawing.Font(
                    "Arial",
                    8)


            Me.Controls.Add(lblInfo)


            '=====================================================
            ' CLOSE BUTTON
            '=====================================================
            btnClose =
                New System.Windows.Forms.Button()


            btnClose.Text = "X"


            btnClose.Left = 215

            btnClose.Top = 5


            btnClose.Width = 25

            btnClose.Height = 20


            btnClose.BackColor =
                System.Drawing.Color.FromArgb(
                    90,
                    40,
                    40)


            btnClose.ForeColor =
                System.Drawing.Color.White


            btnClose.FlatStyle =
                FlatStyle.Flat


            Me.Controls.Add(btnClose)


            '=====================================================
            ' FIRST
            '=====================================================
            btnFirst =
                New System.Windows.Forms.Button()


            btnFirst.Text = "|<"


            btnFirst.Left = 10

            btnFirst.Top = 30


            btnFirst.Width = 40

            btnFirst.Height = 25


            '=====================================================
            ' PREVIOUS
            '=====================================================
            btnPrev =
                New System.Windows.Forms.Button()


            btnPrev.Text = "<"


            btnPrev.Left = 55

            btnPrev.Top = 30


            btnPrev.Width = 35

            btnPrev.Height = 25


            '=====================================================
            ' NEXT
            '=====================================================
            btnNext =
                New System.Windows.Forms.Button()


            btnNext.Text = ">"


            btnNext.Left = 95

            btnNext.Top = 30


            btnNext.Width = 35

            btnNext.Height = 25


            '=====================================================
            ' LAST
            '=====================================================
            btnLast =
                New System.Windows.Forms.Button()


            btnLast.Text = ">|"


            btnLast.Left = 135

            btnLast.Top = 30


            btnLast.Width = 40

            btnLast.Height = 25


            Me.Controls.Add(btnFirst)

            Me.Controls.Add(btnPrev)

            Me.Controls.Add(btnNext)

            Me.Controls.Add(btnLast)


            '=====================================================
            ' PAGE INPUT
            '=====================================================
            txtPage =
                New System.Windows.Forms.TextBox()


            txtPage.Left = 10

            txtPage.Top = 60


            txtPage.Width = 45

            txtPage.Height = 20


            Me.Controls.Add(txtPage)


            '=====================================================
            ' GO BUTTON
            '=====================================================
            btnGo =
                New System.Windows.Forms.Button()


            btnGo.Text = "Chuyển Sheet"


            btnGo.Left = 60
            btnGo.Top = 58
            btnGo.Width = 100
            btnGo.Height = 22


            Me.Controls.Add(btnGo)


            '=====================================================
            ' LEFT
            '=====================================================
            chkLeft =
                New System.Windows.Forms.CheckBox()


            chkLeft.Text = "L"


            chkLeft.Left = 180

            chkLeft.Top = 60


            chkLeft.Width = 35


            chkLeft.ForeColor =
                System.Drawing.Color.White


            Me.Controls.Add(chkLeft)


            '=====================================================
            ' RIGHT
            '=====================================================
            chkRight =
                New System.Windows.Forms.CheckBox()


            chkRight.Text = "R"


            chkRight.Left = 215

            chkRight.Top = 60


            chkRight.Width = 35


            chkRight.ForeColor =
                System.Drawing.Color.White


            Me.Controls.Add(chkRight)


            '=====================================================
            ' STYLE BUTTON
            '=====================================================
            Dim ctrl As Control


            For Each ctrl In Me.Controls

                If TypeOf ctrl Is
                   System.Windows.Forms.Button Then


                    ctrl.BackColor =
                        System.Drawing.Color.FromArgb(
                            63,
                            63,
                            70)


                    ctrl.ForeColor =
                        System.Drawing.Color.White


                    CType(
                        ctrl,
                        Button).FlatStyle =
                            FlatStyle.Flat

                End If

            Next


            '=====================================================
            ' EVENTS
            '=====================================================
            AddHandler btnFirst.Click,
                AddressOf btnFirst_Click


            AddHandler btnPrev.Click,
                AddressOf btnPrev_Click


            AddHandler btnNext.Click,
                AddressOf btnNext_Click


            AddHandler btnLast.Click,
                AddressOf btnLast_Click


            AddHandler btnGo.Click,
                AddressOf btnGo_Click


            AddHandler btnClose.Click,
                AddressOf btnClose_Click


            AddHandler chkLeft.CheckedChanged,
                AddressOf chkLeft_CheckedChanged


            AddHandler chkRight.CheckedChanged,
                AddressOf chkRight_CheckedChanged


            AddHandler Me.Move,
                AddressOf Form_Move


            '=====================================================
            ' DRAG FORM
            '=====================================================
            AddHandler Me.MouseDown,
                AddressOf Form_MouseDown


            AddHandler Me.MouseMove,
                AddressOf Form_MouseMove


            AddHandler Me.MouseUp,
                AddressOf Form_MouseUp


            AddHandler lblInfo.MouseDown,
                AddressOf Form_MouseDown


            AddHandler lblInfo.MouseMove,
                AddressOf Form_MouseMove


            AddHandler lblInfo.MouseUp,
                AddressOf Form_MouseUp


            '=====================================================
            ' UPDATE
            '=====================================================
            UpdateInfo()

        End Sub


        '=========================================================
        ' CURRENT SHEET
        '=========================================================
        Private Function GetCurrentSheetIndex() As Integer

            Try

                Dim i As Integer


                For i = 1 To oDrawDoc.Sheets.Count

                    If oDrawDoc.Sheets.Item(i).Name =
                       oDrawDoc.ActiveSheet.Name Then

                        Return i

                    End If

                Next


            Catch
            End Try


            Return 1

        End Function


        '=========================================================
        ' UPDATE INFO
        '=========================================================
        Private Sub UpdateInfo()

            Try

                Dim currentIndex As Integer =
                    GetCurrentSheetIndex()


                lblInfo.Text =
                    "Sheet " &
                    currentIndex.ToString() &
                    "/" &
                    oDrawDoc.Sheets.Count.ToString()


            Catch
            End Try

        End Sub


        '=========================================================
        ' GO TO SHEET
        '=========================================================
        Private Sub GoToSheet(
            ByVal index As Integer)

            Try

                If oDrawDoc Is Nothing Then
                    Exit Sub
                End If


                If index < 1 Then
                    index = 1
                End If


                If index > oDrawDoc.Sheets.Count Then
                    index = oDrawDoc.Sheets.Count
                End If


                oDrawDoc.Sheets.Item(index).Activate()


                UpdateInfo()


            Catch ex As Exception

                MessageBox.Show(
                    ex.Message,
                    "Sheet Navigator",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning)

            End Try

        End Sub


        '=========================================================
        ' KEEP POSITION
        '=========================================================
        Private Sub Form_Move(
            ByVal sender As Object,
            ByVal e As EventArgs)

            Try

                If chkLeft.Checked Then

                    Me.Left = 900

                    Me.Top = 900

                End If


                If chkRight.Checked Then

                    Me.Left = 2800

                    Me.Top = 900

                End If

            Catch
            End Try

        End Sub


        '=========================================================
        ' DOCK LEFT
        '=========================================================
        Private Sub chkLeft_CheckedChanged(
    ByVal sender As Object,
    ByVal e As EventArgs)

            Try

                If chkLeft.Checked Then

                    chkRight.Checked = False

                    Dim wa As Rectangle =
                Screen.PrimaryScreen.WorkingArea

                    '=============================================
                    ' Căn giữa theo chiều dọc
                    ' Sát về phía trái màn hình
                    '=============================================
                    Me.Left = wa.Left + 10

                    Me.Top =
                wa.Top +
                ((wa.Height - Me.Height) \ 2)

                End If

            Catch
            End Try

        End Sub


        '=========================================================
        ' DOCK RIGHT
        '=========================================================
        Private Sub chkRight_CheckedChanged(
    ByVal sender As Object,
    ByVal e As EventArgs)

            Try

                If chkRight.Checked Then

                    chkLeft.Checked = False

                    Dim wa As Rectangle =
                Screen.PrimaryScreen.WorkingArea

                    Dim halfWidth As Integer =
                wa.Width \ 2

                    '=============================================
                    ' GIỮA NỬA PHẢI + PHÍA DƯỚI MÀN HÌNH
                    '=============================================
                    Me.Left =
                wa.Left +
                halfWidth +
                ((halfWidth - Me.Width) \ 2)

                    Me.Top =
                wa.Bottom -
                Me.Height -
                10

                End If

            Catch
            End Try

        End Sub


        '=========================================================
        ' DRAG FORM - MOUSE DOWN
        '=========================================================
        Private Sub Form_MouseDown(
            ByVal sender As Object,
            ByVal e As MouseEventArgs)

            Try

                If chkLeft.Checked OrElse
                   chkRight.Checked Then

                    Exit Sub

                End If


                If e.Button =
                   MouseButtons.Left Then


                    isDragging = True


                    dragCursorPoint =
                        Cursor.Position


                    dragFormPoint =
                        Me.Location

                End If

            Catch
            End Try

        End Sub


        '=========================================================
        ' DRAG FORM - MOUSE MOVE
        '=========================================================
        Private Sub Form_MouseMove(
            ByVal sender As Object,
            ByVal e As MouseEventArgs)

            Try

                If isDragging Then


                    Dim diff As System.Drawing.Point


                    diff =
                        System.Drawing.Point.Subtract(
                            Cursor.Position,
                            New Size(
                                dragCursorPoint))


                    Me.Location =
                        System.Drawing.Point.Add(
                            dragFormPoint,
                            New Size(diff))

                End If

            Catch
            End Try

        End Sub


        '=========================================================
        ' DRAG FORM - MOUSE UP
        '=========================================================
        Private Sub Form_MouseUp(
            ByVal sender As Object,
            ByVal e As MouseEventArgs)

            isDragging = False

        End Sub


        '=========================================================
        ' CLOSE
        '=========================================================
        Private Sub btnClose_Click(
            ByVal sender As Object,
            ByVal e As EventArgs)

            Try

                Me.Close()

                Me.Dispose()

            Catch
            End Try

        End Sub


        '=========================================================
        ' FIRST
        '=========================================================
        Private Sub btnFirst_Click(
            ByVal sender As Object,
            ByVal e As EventArgs)

            GoToSheet(1)

        End Sub


        '=========================================================
        ' PREVIOUS
        '=========================================================
        Private Sub btnPrev_Click(
            ByVal sender As Object,
            ByVal e As EventArgs)

            Dim currentIndex As Integer =
                GetCurrentSheetIndex()


            GoToSheet(
                currentIndex - 1)

        End Sub


        '=========================================================
        ' NEXT
        '=========================================================
        Private Sub btnNext_Click(
            ByVal sender As Object,
            ByVal e As EventArgs)

            Dim currentIndex As Integer =
                GetCurrentSheetIndex()


            GoToSheet(
                currentIndex + 1)

        End Sub


        '=========================================================
        ' LAST
        '=========================================================
        Private Sub btnLast_Click(
            ByVal sender As Object,
            ByVal e As EventArgs)

            GoToSheet(
                oDrawDoc.Sheets.Count)

        End Sub


        '=========================================================
        ' GO
        '=========================================================
        Private Sub btnGo_Click(
            ByVal sender As Object,
            ByVal e As EventArgs)

            Try

                Dim pageNumber As Integer


                If Integer.TryParse(
                    txtPage.Text,
                    pageNumber) Then


                    GoToSheet(
                        pageNumber)


                Else

                    MessageBox.Show(
                        "Nhập số Sheet hợp lệ.",
                        "Sheet Navigator",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning)

                End If


            Catch ex As Exception

                MessageBox.Show(
                    ex.Message,
                    "Sheet Navigator",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning)

            End Try

        End Sub

        Private Sub InitializeComponent()
            Me.SuspendLayout()
            '
            'SheetNavigatorForm
            '
            Me.ClientSize = New System.Drawing.Size(284, 261)
            Me.Name = "SheetNavigatorForm"
            Me.ResumeLayout(False)

        End Sub

        Private Sub SheetNavigatorForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        End Sub
    End Class

End Namespace