Imports Inventor

Namespace ToolInventor2020
    Public Class PartButtons
        Private Shared Function LoadIconFromPath(path As String) As stdole.IPictureDisp
            Try
                If String.IsNullOrEmpty(path) Then Return Nothing
                If Not System.IO.File.Exists(path) Then Return Nothing
                Using bmp As New System.Drawing.Bitmap(path)
                    Dim clone As New System.Drawing.Bitmap(bmp)
                    Try
                        Return PictureDispConverter.ToIPictureDisp(clone)
                    Finally
                        clone.Dispose()
                    End Try
                End Using
            Catch
                Return Nothing
            End Try
        End Function

        Public Shared Sub Register(controlDefs As Inventor.ControlDefinitions, addInClientID As String, buttonsList As System.Collections.Generic.List(Of ButtonDefinition), largeIcon As stdole.IPictureDisp, smallIcon As stdole.IPictureDisp)
            ' Multi-line description/tooltips so you can edit them easily
            ' Per-button icon file paths (update these paths later to change icons per button)
            Dim iconsFolder As String = "C:\Users\thanh\source\repos\ThanhN\Code addin inventor\Images\Button\Part"
            Dim btn1LargePath As String = System.IO.Path.Combine(iconsFolder, "o3.bmp")
            Dim btn1SmallPath As String = System.IO.Path.Combine(iconsFolder, "o1.bmp")
            Dim btn2LargePath As String = btn1LargePath
            Dim btn2SmallPath As String = btn1SmallPath
            Dim btn3LargePath As String = btn1LargePath
            Dim btn3SmallPath As String = btn1SmallPath
            Dim btn4LargePath As String = btn1LargePath
            Dim btn4SmallPath As String = btn1SmallPath
            Dim btn5LargePath As String = btn1LargePath
            Dim btn5SmallPath As String = btn1SmallPath
            Dim btn6LargePath As String = btn1LargePath
            Dim btn6SmallPath As String = btn1SmallPath
            Dim btn7LargePath As String = btn1LargePath
            Dim btn7SmallPath As String = btn1SmallPath
            Dim btn8LargePath As String = btn1LargePath
            Dim btn8SmallPath As String = btn1SmallPath
            Dim btn9LargePath As String = btn1LargePath
            Dim btn9SmallPath As String = btn1SmallPath
            Dim btn10LargePath As String = btn1LargePath
            Dim btn10SmallPath As String = btn1SmallPath
            Dim btn11LargePath As String = btn1LargePath
            Dim btn11SmallPath As String = btn1SmallPath

            ' Load per-button icons (fallback to provided largeIcon/smallIcon when file missing)
            Dim btn1LargeIcon As stdole.IPictureDisp = If(System.IO.File.Exists(btn1LargePath), LoadIconFromPath(btn1LargePath), largeIcon)
            Dim btn1SmallIcon As stdole.IPictureDisp = If(System.IO.File.Exists(btn1SmallPath), LoadIconFromPath(btn1SmallPath), smallIcon)
            Dim btn2LargeIcon As stdole.IPictureDisp = If(System.IO.File.Exists(btn2LargePath), LoadIconFromPath(btn2LargePath), largeIcon)
            Dim btn2SmallIcon As stdole.IPictureDisp = If(System.IO.File.Exists(btn2SmallPath), LoadIconFromPath(btn2SmallPath), smallIcon)
            Dim btn3LargeIcon As stdole.IPictureDisp = If(System.IO.File.Exists(btn3LargePath), LoadIconFromPath(btn3LargePath), largeIcon)
            Dim btn3SmallIcon As stdole.IPictureDisp = If(System.IO.File.Exists(btn3SmallPath), LoadIconFromPath(btn3SmallPath), smallIcon)
            Dim btn4LargeIcon As stdole.IPictureDisp = If(System.IO.File.Exists(btn4LargePath), LoadIconFromPath(btn4LargePath), largeIcon)
            Dim btn4SmallIcon As stdole.IPictureDisp = If(System.IO.File.Exists(btn4SmallPath), LoadIconFromPath(btn4SmallPath), smallIcon)
            Dim btn5LargeIcon As stdole.IPictureDisp = If(System.IO.File.Exists(btn5LargePath), LoadIconFromPath(btn5LargePath), largeIcon)
            Dim btn5SmallIcon As stdole.IPictureDisp = If(System.IO.File.Exists(btn5SmallPath), LoadIconFromPath(btn5SmallPath), smallIcon)
            Dim btn6LargeIcon As stdole.IPictureDisp = If(System.IO.File.Exists(btn6LargePath), LoadIconFromPath(btn6LargePath), largeIcon)
            Dim btn6SmallIcon As stdole.IPictureDisp = If(System.IO.File.Exists(btn6SmallPath), LoadIconFromPath(btn6SmallPath), smallIcon)
            Dim btn7LargeIcon As stdole.IPictureDisp = If(System.IO.File.Exists(btn7LargePath), LoadIconFromPath(btn7LargePath), largeIcon)
            Dim btn7SmallIcon As stdole.IPictureDisp = If(System.IO.File.Exists(btn7SmallPath), LoadIconFromPath(btn7SmallPath), smallIcon)
            Dim btn8LargeIcon As stdole.IPictureDisp = If(System.IO.File.Exists(btn8LargePath), LoadIconFromPath(btn8LargePath), largeIcon)
            Dim btn8SmallIcon As stdole.IPictureDisp = If(System.IO.File.Exists(btn8SmallPath), LoadIconFromPath(btn8SmallPath), smallIcon)
            Dim btn9LargeIcon As stdole.IPictureDisp = If(System.IO.File.Exists(btn9LargePath), LoadIconFromPath(btn9LargePath), largeIcon)
            Dim btn9SmallIcon As stdole.IPictureDisp = If(System.IO.File.Exists(btn9SmallPath), LoadIconFromPath(btn9SmallPath), smallIcon)
            Dim btn10LargeIcon As stdole.IPictureDisp = If(System.IO.File.Exists(btn10LargePath), LoadIconFromPath(btn10LargePath), largeIcon)
            Dim btn10SmallIcon As stdole.IPictureDisp = If(System.IO.File.Exists(btn10SmallPath), LoadIconFromPath(btn10SmallPath), smallIcon)
            Dim btn11LargeIcon As stdole.IPictureDisp = If(System.IO.File.Exists(btn11LargePath), LoadIconFromPath(btn11LargePath), largeIcon)
            Dim btn11SmallIcon As stdole.IPictureDisp = If(System.IO.File.Exists(btn11SmallPath), LoadIconFromPath(btn11SmallPath), smallIcon)
            Dim partBtn1 As ButtonDefinition = controlDefs.AddButtonDefinition("Thêm dung sai vào dim sketch", "ToolInventor2020_Part_Btn1", CommandTypesEnum.kShapeEditCmdType, addInClientID,
                                                                    "Thêm dung sai vào dim sketch" & vbCrLf & "Mô tả: Thêm dung sai (tolerance) vào dimension trong sketch." & vbCrLf & "Ghi chú: chỉnh theo nhu cầu.",
                                                                     "Tooltip: Thêm dung sai vào dim sketch (edit)", btn1SmallIcon, btn1LargeIcon)
            AddHandler partBtn1.OnExecute, AddressOf Part.Buttons.solid.Part_Solid_1.OnExecute
            buttonsList.Add(partBtn1)

            Dim partBtn2 As ButtonDefinition = controlDefs.AddButtonDefinition("Xóa fix all sketch", "ToolInventor2020_Part_Btn2", CommandTypesEnum.kShapeEditCmdType, addInClientID,
                                                                               "Xóa fix all sketch" & vbCrLf & "Mô tả: Xóa trạng thái 'fix' cho tất cả sketch entities.",
                                                                               "Tooltip: Xóa fix all sketch (edit)", btn2SmallIcon, btn2LargeIcon)
            AddHandler partBtn2.OnExecute, AddressOf Part.Buttons.solid.Part_Solid_2.OnExecute
            buttonsList.Add(partBtn2)

            Dim partBtn3 As ButtonDefinition = controlDefs.AddButtonDefinition("Tạo coil line 3d", "ToolInventor2020_Part_Btn3", CommandTypesEnum.kShapeEditCmdType, addInClientID,
                                                                               "Tạo coil line 3d" & vbCrLf & "Mô tả: Tạo đường coil 3D từ sketch hiện tại.",
                                                                               "Tooltip: Tạo coil line 3d (edit)", btn3SmallIcon, btn3LargeIcon)
            AddHandler partBtn3.OnExecute, AddressOf Part.Buttons.solid.Part_Solid_3.OnExecute
            buttonsList.Add(partBtn3)

            Dim partBtn4 As ButtonDefinition = controlDefs.AddButtonDefinition("Thay tên body", "ToolInventor2020_Part_Btn4", CommandTypesEnum.kShapeEditCmdType, addInClientID,
                                                                               "Thay tên body" & vbCrLf & "Mô tả: Đổi tên body theo quy tắc đặt tên của bạn.",
                                                                               "Tooltip: Thay tên body (edit)", btn4SmallIcon, btn4LargeIcon)
            AddHandler partBtn4.OnExecute, AddressOf Part.Buttons.solid.Part_Solid_4.OnExecute
            buttonsList.Add(partBtn4)

            Dim partBtn5 As ButtonDefinition = controlDefs.AddButtonDefinition("Export từng body sang các định dạng khác", "ToolInventor2020_Part_Btn5", CommandTypesEnum.kShapeEditCmdType, addInClientID,
                                                                               "Export từng body sang các định dạng khác" & vbCrLf & "Mô tả: Xuất từng body ra file STEP/IGES/...",
                                                                               "Tooltip: Export từng body (edit)", btn5SmallIcon, btn5LargeIcon)
            AddHandler partBtn5.OnExecute, AddressOf Part.Buttons.solid.Part_Solid_5.OnExecute
            buttonsList.Add(partBtn5)

            Dim partBtn6 As ButtonDefinition = controlDefs.AddButtonDefinition("Thay tên body part theo thứ tự", "ToolInventor2020_Part_Btn6", CommandTypesEnum.kShapeEditCmdType, addInClientID,
                                                                               "Thay tên body part theo thứ tự" & vbCrLf & "Mô tả: Đổi tên body theo thứ tự trong danh sách.",
                                                                               "Tooltip: Thay tên body theo thứ tự (edit)", btn6SmallIcon, btn6LargeIcon)
            AddHandler partBtn6.OnExecute, AddressOf Part.Buttons.solid.Part_Solid_6.OnExecute
            buttonsList.Add(partBtn6)

            Dim partBtn7 As ButtonDefinition = controlDefs.AddButtonDefinition("Part Action 7", "ToolInventor2020_Part_Btn7", CommandTypesEnum.kShapeEditCmdType, addInClientID,
                                                                               "Part Action 7" & vbCrLf & "Mô tả: (chỉnh sửa) Action 7 cho Part.",
                                                                               "Tooltip: Part Action 7 (edit)", btn7SmallIcon, btn7LargeIcon)
            AddHandler partBtn7.OnExecute, AddressOf Part.Buttons.Button7.OnExecute
            buttonsList.Add(partBtn7)

            Dim partBtn8 As ButtonDefinition = controlDefs.AddButtonDefinition("Part Action 8", "ToolInventor2020_Part_Btn8", CommandTypesEnum.kShapeEditCmdType, addInClientID,
                                                                               "Part Action 8" & vbCrLf & "Mô tả: (chỉnh sửa) Action 8 cho Part.",
                                                                               "Tooltip: Part Action 8 (edit)", btn8SmallIcon, btn8LargeIcon)
            AddHandler partBtn8.OnExecute, AddressOf Part.Buttons.Button8.OnExecute
            buttonsList.Add(partBtn8)

            Dim partBtn9 As ButtonDefinition = controlDefs.AddButtonDefinition("Part Action 9", "ToolInventor2020_Part_Btn9", CommandTypesEnum.kShapeEditCmdType, addInClientID,
                                                                               "Part Action 9" & vbCrLf & "Mô tả: (chỉnh sửa) Action 9 cho Part.",
                                                                               "Tooltip: Part Action 9 (edit)", btn9SmallIcon, btn9LargeIcon)
            AddHandler partBtn9.OnExecute, AddressOf Part.Buttons.Button9.OnExecute
            buttonsList.Add(partBtn9)

            Dim partBtn10 As ButtonDefinition = controlDefs.AddButtonDefinition("Part Action 10", "ToolInventor2020_Part_Btn10", CommandTypesEnum.kShapeEditCmdType, addInClientID,
                                                                                 "Part Action 10" & vbCrLf & "Mô tả: (chỉnh sửa) Action 10 cho Part.",
                                                                                 "Tooltip: Part Action 10 (edit)", btn10SmallIcon, btn10LargeIcon)
            AddHandler partBtn10.OnExecute, AddressOf Part.Buttons.Button10.OnExecute
            buttonsList.Add(partBtn10)

            Dim partBtn11 As ButtonDefinition = controlDefs.AddButtonDefinition("Part Action 11", "ToolInventor2020_Part_Btn11", CommandTypesEnum.kShapeEditCmdType, addInClientID,
                                                                                "Part Action 11" & vbCrLf & "Mô tả: (chỉnh sửa) Action 11 cho Part.",
                                                                                "Tooltip: Part Action 11 (edit)", largeIcon, smallIcon)
            AddHandler partBtn11.OnExecute, AddressOf Part.Buttons.Button11.OnExecute
            buttonsList.Add(partBtn11)

            Dim partBtn12 As ButtonDefinition = controlDefs.AddButtonDefinition("Part Action 12", "ToolInventor2020_Part_Btn12", CommandTypesEnum.kShapeEditCmdType, addInClientID,
                                                                                "Part Action 12" & vbCrLf & "Mô tả: (chỉnh sửa) Action 12 cho Part.",
                                                                                "Tooltip: Part Action 12 (edit)", largeIcon, smallIcon)
            AddHandler partBtn12.OnExecute, AddressOf Part.Buttons.Button12.OnExecute
            buttonsList.Add(partBtn12)

            Dim partBtn13 As ButtonDefinition = controlDefs.AddButtonDefinition("Part Action 13", "ToolInventor2020_Part_Btn13", CommandTypesEnum.kShapeEditCmdType, addInClientID,
                                                                                "Part Action 13" & vbCrLf & "Mô tả: (chỉnh sửa) Action 13 cho Part.",
                                                                                "Tooltip: Part Action 13 (edit)", largeIcon, smallIcon)
            AddHandler partBtn13.OnExecute, AddressOf Part.Buttons.Button13.OnExecute
            buttonsList.Add(partBtn13)

            Dim partBtn14 As ButtonDefinition = controlDefs.AddButtonDefinition("Part Action 14", "ToolInventor2020_Part_Btn14", CommandTypesEnum.kShapeEditCmdType, addInClientID,
                                                                                "Part Action 14" & vbCrLf & "Mô tả: (chỉnh sửa) Action 14 cho Part.",
                                                                                "Tooltip: Part Action 14 (edit)", largeIcon, smallIcon)
            AddHandler partBtn14.OnExecute, AddressOf Part.Buttons.Button14.OnExecute
            buttonsList.Add(partBtn14)

            Dim partBtn15 As ButtonDefinition = controlDefs.AddButtonDefinition("Part Action 15", "ToolInventor2020_Part_Btn15", CommandTypesEnum.kShapeEditCmdType, addInClientID,
                                                                                "Part Action 15" & vbCrLf & "Mô tả: (chỉnh sửa) Action 15 cho Part.",
                                                                                "Tooltip: Part Action 15 (edit)", largeIcon, smallIcon)
            AddHandler partBtn15.OnExecute, AddressOf Part.Buttons.Button15.OnExecute
            buttonsList.Add(partBtn15)

        End Sub
    End Class
End Namespace
