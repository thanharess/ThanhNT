Imports Inventor

Namespace ThanhN
    Public Class PartButtons
        Public Shared Sub Register(controlDefs As Inventor.ControlDefinitions, addInClientID As String, buttonsList As System.Collections.Generic.List(Of ButtonDefinition))
            ' Multi-line description/tooltips so you can edit them easily
            Dim descs As String() = New String() {
                "Thêm dung sai vào dim sketch" & vbCrLf & "Mô tả: Thêm dung sai (tolerance) vào dimension trong sketch." & vbCrLf & "Ghi chú: chỉnh theo nhu cầu.",
                "Xóa fix all sketch" & vbCrLf & "Mô tả: Xóa trạng thái 'fix' cho tất cả sketch entities.",
                "Tạo coil line 3d" & vbCrLf & "Mô tả: Tạo đường coil 3D từ sketch hiện tại.",
                "Thay tên body" & vbCrLf & "Mô tả: Đổi tên body theo quy tắc đặt tên của bạn.",
                "Export từng body sang các định dạng khác" & vbCrLf & "Mô tả: Xuất từng body ra file STEP/IGES/...",
                "Thay tên body part the thứ tự" & vbCrLf & "Mô tả: Đổi tên body theo thứ tự trong danh sách.",
                "Part Action 7" & vbCrLf & "Mô tả: (chỉnh sửa) Action 7 cho Part.",
                "Part Action 8" & vbCrLf & "Mô tả: (chỉnh sửa) Action 8 cho Part.",
                "Part Action 9" & vbCrLf & "Mô tả: (chỉnh sửa) Action 9 cho Part.",
                "Part Action 10" & vbCrLf & "Mô tả: (chỉnh sửa) Action 10 cho Part.",
                "Part Action 11" & vbCrLf & "Mô tả: (chỉnh sửa) Action 11 cho Part.",
                "Part Action 12" & vbCrLf & "Mô tả: (chỉnh sửa) Action 12 cho Part.",
                "Part Action 13" & vbCrLf & "Mô tả: (chỉnh sửa) Action 13 cho Part.",
                "Part Action 14" & vbCrLf & "Mô tả: (chỉnh sửa) Action 14 cho Part.",
                "Part Action 15" & vbCrLf & "Mô tả: (chỉnh sửa) Action 15 cho Part."}
            Dim tips As String() = New String() {
                "Tooltip: Thêm dung sai vào dim sketch (edit)",
                "Tooltip: Xóa fix all sketch (edit)",
                "Tooltip: Tạo coil line 3d (edit)",
                "Tooltip: Thay tên body (edit)",
                "Tooltip: Export từng body (edit)",
                "Tooltip: Thay tên body theo thứ tự (edit)",
                "Tooltip: Part Action 7 (edit)",
                "Tooltip: Part Action 8 (edit)",
                "Tooltip: Part Action 9 (edit)",
                "Tooltip: Part Action 10 (edit)",
                "Tooltip: Part Action 11 (edit)",
                "Tooltip: Part Action 12 (edit)",
                "Tooltip: Part Action 13 (edit)",
                "Tooltip: Part Action 14 (edit)",
                "Tooltip: Part Action 15 (edit)"}

            ' Create Part buttons explicitly (no loop) so each button can have distinct implementation
            Dim partBtn1 As ButtonDefinition = controlDefs.AddButtonDefinition("Thêm dung sai vào dim sketch", "ThanhN_Part_Btn1", CommandTypesEnum.kShapeEditCmdType, addInClientID, descs(0), tips(0))
            AddHandler partBtn1.OnExecute, AddressOf Part.Buttons.Button1.OnExecute
            buttonsList.Add(partBtn1)

            Dim partBtn2 As ButtonDefinition = controlDefs.AddButtonDefinition("Xóa fix all sketch", "ThanhN_Part_Btn2", CommandTypesEnum.kShapeEditCmdType, addInClientID, descs(1), tips(1))
            AddHandler partBtn2.OnExecute, AddressOf Part.Buttons.Button2.OnExecute
            buttonsList.Add(partBtn2)

            Dim partBtn3 As ButtonDefinition = controlDefs.AddButtonDefinition("Tạo coil line 3d", "ThanhN_Part_Btn3", CommandTypesEnum.kShapeEditCmdType, addInClientID, descs(2), tips(2))
            AddHandler partBtn3.OnExecute, AddressOf Part.Buttons.Button3.OnExecute
            buttonsList.Add(partBtn3)

            Dim partBtn4 As ButtonDefinition = controlDefs.AddButtonDefinition("Thay tên body", "ThanhN_Part_Btn4", CommandTypesEnum.kShapeEditCmdType, addInClientID, descs(3), tips(3))
            AddHandler partBtn4.OnExecute, AddressOf Part.Buttons.Button4.OnExecute
            buttonsList.Add(partBtn4)

            Dim partBtn5 As ButtonDefinition = controlDefs.AddButtonDefinition("Export từng body sang các định dạng khác", "ThanhN_Part_Btn5", CommandTypesEnum.kShapeEditCmdType, addInClientID, descs(4), tips(4))
            AddHandler partBtn5.OnExecute, AddressOf Part.Buttons.Button5.OnExecute
            buttonsList.Add(partBtn5)

            Dim partBtn6 As ButtonDefinition = controlDefs.AddButtonDefinition("Thay tên body part the thứ tự", "ThanhN_Part_Btn6", CommandTypesEnum.kShapeEditCmdType, addInClientID, descs(5), tips(5))
            AddHandler partBtn6.OnExecute, AddressOf Part.Buttons.Button6.OnExecute
            buttonsList.Add(partBtn6)

            Dim partBtn7 As ButtonDefinition = controlDefs.AddButtonDefinition("Part Action 7", "ThanhN_Part_Btn7", CommandTypesEnum.kShapeEditCmdType, addInClientID, descs(6), tips(6))
            AddHandler partBtn7.OnExecute, AddressOf Part.Buttons.Button7.OnExecute
            buttonsList.Add(partBtn7)

            Dim partBtn8 As ButtonDefinition = controlDefs.AddButtonDefinition("Part Action 8", "ThanhN_Part_Btn8", CommandTypesEnum.kShapeEditCmdType, addInClientID, descs(7), tips(7))
            AddHandler partBtn8.OnExecute, AddressOf Part.Buttons.Button8.OnExecute
            buttonsList.Add(partBtn8)

            Dim partBtn9 As ButtonDefinition = controlDefs.AddButtonDefinition("Part Action 9", "ThanhN_Part_Btn9", CommandTypesEnum.kShapeEditCmdType, addInClientID, descs(8), tips(8))
            AddHandler partBtn9.OnExecute, AddressOf Part.Buttons.Button9.OnExecute
            buttonsList.Add(partBtn9)

            Dim partBtn10 As ButtonDefinition = controlDefs.AddButtonDefinition("Part Action 10", "ThanhN_Part_Btn10", CommandTypesEnum.kShapeEditCmdType, addInClientID, descs(9), tips(9))
            AddHandler partBtn10.OnExecute, AddressOf Part.Buttons.Button10.OnExecute
            buttonsList.Add(partBtn10)

            Dim partBtn11 As ButtonDefinition = controlDefs.AddButtonDefinition("Part Action 11", "ThanhN_Part_Btn11", CommandTypesEnum.kShapeEditCmdType, addInClientID, descs(10), tips(10))
            AddHandler partBtn11.OnExecute, AddressOf Part.Buttons.Button11.OnExecute
            buttonsList.Add(partBtn11)

            Dim partBtn12 As ButtonDefinition = controlDefs.AddButtonDefinition("Part Action 12", "ThanhN_Part_Btn12", CommandTypesEnum.kShapeEditCmdType, addInClientID, descs(11), tips(11))
            AddHandler partBtn12.OnExecute, AddressOf Part.Buttons.Button12.OnExecute
            buttonsList.Add(partBtn12)

            Dim partBtn13 As ButtonDefinition = controlDefs.AddButtonDefinition("Part Action 13", "ThanhN_Part_Btn13", CommandTypesEnum.kShapeEditCmdType, addInClientID, descs(12), tips(12))
            AddHandler partBtn13.OnExecute, AddressOf Part.Buttons.Button13.OnExecute
            buttonsList.Add(partBtn13)

            Dim partBtn14 As ButtonDefinition = controlDefs.AddButtonDefinition("Part Action 14", "ThanhN_Part_Btn14", CommandTypesEnum.kShapeEditCmdType, addInClientID, descs(13), tips(13))
            AddHandler partBtn14.OnExecute, AddressOf Part.Buttons.Button14.OnExecute
            buttonsList.Add(partBtn14)

            Dim partBtn15 As ButtonDefinition = controlDefs.AddButtonDefinition("Part Action 15", "ThanhN_Part_Btn15", CommandTypesEnum.kShapeEditCmdType, addInClientID, descs(14), tips(14))
            AddHandler partBtn15.OnExecute, AddressOf Part.Buttons.Button15.OnExecute
            buttonsList.Add(partBtn15)
        End Sub
    End Class
End Namespace
