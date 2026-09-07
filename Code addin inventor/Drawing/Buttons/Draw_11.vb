Option Explicit On
Option Strict Off

Imports Inventor
Imports System
Imports System.Windows.Forms
Imports SD = System.Drawing          ' Alias tránh xung đột

Namespace ToolInventor2020.Drawing.Buttons

    Public Module Draw_11

        Public Sub OnExecute(ByVal Context As NameValueMap)
            Dim invApp As Inventor.Application = Nothing
            Try
                invApp = g_inventorApplication

                If invApp Is Nothing Then
                    MessageBox.Show("Không tìm thấy Inventor Application.", "Sheet Navigator", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                If invApp.ActiveDocument Is Nothing Then
                    MessageBox.Show("Không có tài liệu đang mở.", "Sheet Navigator", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                If invApp.ActiveDocument.DocumentType <> DocumentTypeEnum.kDrawingDocumentObject Then
                    MessageBox.Show("Chỉ sử dụng trong Drawing.", "Sheet Navigator", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Exit Sub
                End If

                Dim frm As New SheetNavigatorForma(invApp)
                frm.Show()

            Catch ex As Exception
                MessageBox.Show("Lỗi Sheet Navigator:" & vbCrLf & ex.Message, "Sheet Navigator", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

    End Module

    '=============================================================
    ' SHEET NAVIGATOR FORM - LIGHT + CLEAN (Inventor 2020)
    '=============================================================
    Public Class SheetNavigatorForma
        Inherits Form

        Private m_invApp As Inventor.Application
        Private m_drawDoc As DrawingDocument

        '----- Thay toàn bộ phần khai báo control bằng đoạn này -----

        Private lblInfo As System.Windows.Forms.Label
        Private txtPage As System.Windows.Forms.TextBox          ' ← sửa chỗ này
        Private btnFirst As System.Windows.Forms.Button
        Private btnPrev As System.Windows.Forms.Button
        Private btnNext As System.Windows.Forms.Button
        Private btnLast As System.Windows.Forms.Button
        Private btnGo As System.Windows.Forms.Button
        Private btnClose As System.Windows.Forms.Button
        Private chkLeft As System.Windows.Forms.CheckBox
        Private chkRight As System.Windows.Forms.CheckBox

        Private isDragging As Boolean = False
        Private dragCursorPoint As SD.Point
        Private dragFormPoint As SD.Point

        Public Sub New(ByVal app As Inventor.Application)
            MyBase.New()

            m_invApp = app

            If m_invApp Is Nothing OrElse m_invApp.ActiveDocument Is Nothing Then
                Throw New Exception("Inventor / Document không hợp lệ.")
            End If

            If m_invApp.ActiveDocument.DocumentType <> DocumentTypeEnum.kDrawingDocumentObject Then
                Throw New Exception("Document không phải Drawing.")
            End If

            m_drawDoc = CType(m_invApp.ActiveDocument, DrawingDocument)

            '----- Form -----
            Me.Text = "Sheet Nav"
            Me.FormBorderStyle = FormBorderStyle.FixedToolWindow
            Me.StartPosition = FormStartPosition.Manual
            Me.Size = New SD.Size(260, 120)
            Me.ShowInTaskbar = False
            Me.BackColor = SD.Color.FromArgb(45, 45, 48)
            Me.TopMost = True

            '----- Label -----
            lblInfo = New Label()
            lblInfo.Left = 8
            lblInfo.Top = 6
            lblInfo.Width = 150
            lblInfo.Height = 18
            lblInfo.ForeColor = SD.Color.White
            lblInfo.Font = New SD.Font("Segoe UI", 8)
            Me.Controls.Add(lblInfo)

            '----- Close -----
            btnClose = New Button()
            btnClose.Text = "X"
            btnClose.Left = 160
            btnClose.Top = 6
            btnClose.Width = 24
            btnClose.Height = 20
            btnClose.FlatStyle = FlatStyle.Flat
            btnClose.BackColor = SD.Color.FromArgb(90, 40, 40)
            btnClose.ForeColor = SD.Color.White
            Me.Controls.Add(btnClose)

            '----- Navigation buttons -----
            btnFirst = CreateBtn("|<", 8, 28, 38)
            btnPrev = CreateBtn("<", 50, 28, 34)
            btnNext = CreateBtn(">", 88, 28, 34)
            btnLast = CreateBtn(">|", 126, 28, 38)

            '----- Page + Go -----
            txtPage = New System.Windows.Forms.TextBox()
            txtPage.Left = 8
            txtPage.Top = 60
            txtPage.Width = 42
            txtPage.Height = 20
            Me.Controls.Add(txtPage)
            Me.Controls.Add(txtPage)

            btnGo = CreateBtn("Go", 54, 58, 70)
            btnGo.Height = 22

            '----- Dock -----
            chkLeft = New System.Windows.Forms.CheckBox()
            chkLeft.Text = "L"
            chkLeft.Left = 140
            chkLeft.Top = 60
            chkLeft.Width = 35
            chkLeft.ForeColor = SD.Color.White
            chkLeft.BackColor = SD.Color.Transparent
            Me.Controls.Add(chkLeft)

            chkRight = New System.Windows.Forms.CheckBox()
            chkRight.Text = "R"
            chkRight.Left = 180
            chkRight.Top = 60
            chkRight.Width = 35
            chkRight.ForeColor = SD.Color.White
            chkRight.BackColor = SD.Color.Transparent
            Me.Controls.Add(chkRight)

            ' Style buttons
            For Each c As Control In Me.Controls
                If TypeOf c Is Button AndAlso Not Object.ReferenceEquals(c, btnClose) Then
                    c.BackColor = SD.Color.FromArgb(63, 63, 70)
                    c.ForeColor = SD.Color.White
                    CType(c, Button).FlatStyle = FlatStyle.Flat
                End If
            Next

            ' Events
            AddHandler btnFirst.Click, AddressOf btnFirst_Click
            AddHandler btnPrev.Click, AddressOf btnPrev_Click
            AddHandler btnNext.Click, AddressOf btnNext_Click
            AddHandler btnLast.Click, AddressOf btnLast_Click
            AddHandler btnGo.Click, AddressOf btnGo_Click
            AddHandler btnClose.Click, AddressOf btnClose_Click
            AddHandler chkLeft.CheckedChanged, AddressOf chkLeft_CheckedChanged
            AddHandler chkRight.CheckedChanged, AddressOf chkRight_CheckedChanged

            AddHandler Me.MouseDown, AddressOf Form_MouseDown
            AddHandler Me.MouseMove, AddressOf Form_MouseMove
            AddHandler Me.MouseUp, AddressOf Form_MouseUp
            AddHandler lblInfo.MouseDown, AddressOf Form_MouseDown
            AddHandler lblInfo.MouseMove, AddressOf Form_MouseMove
            AddHandler lblInfo.MouseUp, AddressOf Form_MouseUp

            UpdateInfo()
            Me.Location = New SD.Point(Screen.PrimaryScreen.WorkingArea.Right - Me.Width - 20, 80)
        End Sub

        Private Function CreateBtn(txt As String, x As Integer, y As Integer, w As Integer) As Button
            Dim b As New Button()
            b.Text = txt
            b.Left = x
            b.Top = y
            b.Width = w
            b.Height = 24
            Me.Controls.Add(b)
            Return b
        End Function

        Private Function GetCurrentSheetIndex() As Integer
            Try
                Dim name As String = m_drawDoc.ActiveSheet.Name
                For i As Integer = 1 To m_drawDoc.Sheets.Count
                    If m_drawDoc.Sheets.Item(i).Name = name Then Return i
                Next
            Catch
            End Try
            Return 1
        End Function

        Private Sub UpdateInfo()
            Try
                Dim cur As Integer = GetCurrentSheetIndex()
                lblInfo.Text = "Sheet " & cur.ToString() & " / " & m_drawDoc.Sheets.Count.ToString()
            Catch
            End Try
        End Sub

        Private Sub GoToSheet(index As Integer)
            Try
                If m_drawDoc Is Nothing Then Exit Sub
                If index < 1 Then index = 1
                If index > m_drawDoc.Sheets.Count Then index = m_drawDoc.Sheets.Count
                m_drawDoc.Sheets.Item(index).Activate()
                UpdateInfo()
            Catch
            End Try
        End Sub

        Private Sub chkLeft_CheckedChanged(sender As Object, e As EventArgs)
            If chkLeft.Checked Then
                chkRight.Checked = False
                Dim wa = Screen.PrimaryScreen.WorkingArea
                Me.Left = wa.Left + 8
                Me.Top = wa.Top + (wa.Height - Me.Height) \ 2
            End If
        End Sub

        Private Sub chkRight_CheckedChanged(sender As Object, e As EventArgs)
            If chkRight.Checked Then
                chkLeft.Checked = False
                Dim wa = Screen.PrimaryScreen.WorkingArea
                Me.Left = wa.Right - Me.Width - 8
                Me.Top = wa.Bottom - Me.Height - 12
            End If
        End Sub

        Private Sub Form_MouseDown(sender As Object, e As MouseEventArgs)
            If chkLeft.Checked OrElse chkRight.Checked Then Exit Sub
            If e.Button = MouseButtons.Left Then
                isDragging = True
                dragCursorPoint = Cursor.Position
                dragFormPoint = Me.Location
            End If
        End Sub

        Private Sub Form_MouseMove(sender As Object, e As MouseEventArgs)
            If isDragging Then
                Dim diff As SD.Point = SD.Point.Subtract(Cursor.Position, New SD.Size(dragCursorPoint))
                Me.Location = SD.Point.Add(dragFormPoint, New SD.Size(diff))
            End If
        End Sub

        Private Sub Form_MouseUp(sender As Object, e As MouseEventArgs)
            isDragging = False
        End Sub

        Private Sub btnClose_Click(sender As Object, e As EventArgs)
            Me.Close()
            Me.Dispose()
        End Sub

        Private Sub btnFirst_Click(sender As Object, e As EventArgs)
            GoToSheet(1)
        End Sub

        Private Sub btnPrev_Click(sender As Object, e As EventArgs)
            GoToSheet(GetCurrentSheetIndex() - 1)
        End Sub

        Private Sub btnNext_Click(sender As Object, e As EventArgs)
            GoToSheet(GetCurrentSheetIndex() + 1)
        End Sub

        Private Sub btnLast_Click(sender As Object, e As EventArgs)
            GoToSheet(m_drawDoc.Sheets.Count)
        End Sub

        Private Sub btnGo_Click(sender As Object, e As EventArgs)
            Dim page As Integer
            If Integer.TryParse(txtPage.Text, page) Then
                GoToSheet(page)
            Else
                MessageBox.Show("Nhập số Sheet hợp lệ.", "Sheet Navigator", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
        End Sub

        Protected Overrides Sub OnFormClosed(e As FormClosedEventArgs)
            MyBase.OnFormClosed(e)
            m_drawDoc = Nothing
            m_invApp = Nothing
        End Sub

        Private Sub InitializeComponent()
            Me.SuspendLayout()
            '
            'SheetNavigatorForma
            '
            Me.ClientSize = New System.Drawing.Size(284, 261)
            Me.Name = "SheetNavigatorForma"
            Me.ResumeLayout(False)

        End Sub

        Private Sub SheetNavigatorForma_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        End Sub
    End Class

End Namespace