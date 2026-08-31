Imports System.Diagnostics.Contracts
Imports Inventor
Imports ThanhN.ThanhN.Assembly.Buttons

Namespace ThanhN
    Public Class AssemblyButtons
        Public Shared Sub Register(controlDefs As Inventor.ControlDefinitions, addInClientID As String, buttonsList As System.Collections.Generic.List(Of ButtonDefinition))


            ' Create Assembly buttons explicitly (no loop) so each button can have distinct implementation
            Dim assemblyBtn1 As ButtonDefinition = controlDefs.AddButtonDefinition("Suppress & Unsuppress contrain all", "ThanhN_Assembly_Btn1", CommandTypesEnum.kShapeEditCmdType, addInClientID,
                                                                                   Nothing,
                                                                                   Nothing)
            AddHandler assemblyBtn1.OnExecute, AddressOf Assembly.Buttons.caclenhlapghep.Ass_1.OnExecute
            buttonsList.Add(assemblyBtn1)

            Dim assemblyBtn2 As ButtonDefinition = controlDefs.AddButtonDefinition("Auto contrain Keep position", "ThanhN_Assembly_Btn2", CommandTypesEnum.kShapeEditCmdType, addInClientID,
                                                                                   Nothing,
                                                                                   "code giữ nguyên vị trí các cum & gán contrain tự động")
            AddHandler assemblyBtn2.OnExecute, AddressOf Assembly.Buttons.caclenhlapghep.Ass_2.OnExecute
            buttonsList.Add(assemblyBtn2)

            Dim assemblyBtn3 As ButtonDefinition = controlDefs.AddButtonDefinition("Contrain về gốc tọa dộ 2 chi tiết", "ThanhN_Assembly_Btn3", CommandTypesEnum.kShapeEditCmdType, addInClientID,
                                                                                   Nothing, "Contrain cụm chi tiết hoặc part về gốc tọa độ của nhau")
            AddHandler assemblyBtn3.OnExecute, AddressOf Assembly.Buttons.caclenhlapghep.Ass_3.OnExecute
            buttonsList.Add(assemblyBtn3)

            Dim assemblyBtn4 As ButtonDefinition = controlDefs.AddButtonDefinition("Contrain all to select", "ThanhN_Assembly_Btn4", CommandTypesEnum.kShapeEditCmdType, addInClientID,
                                                                                   Nothing,
                                                                                   "Contrain cụm chi tiết & part tất cả về gốc tọa độ của chi tiết hoặc cụm chi tiết được chọn")
            AddHandler assemblyBtn4.OnExecute, AddressOf Assembly.Buttons.caclenhlapghep.Ass_4.OnExecute
            buttonsList.Add(assemblyBtn4)

            Dim assemblyBtn5 As ButtonDefinition = controlDefs.AddButtonDefinition("Xóa Contrain & ground,.. ", "ThanhN_Assembly_Btn5", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing)
            AddHandler assemblyBtn5.OnExecute, AddressOf Assembly.Buttons.caclenhlapghep.Ass_5.OnExecute
            buttonsList.Add(assemblyBtn5)

            Dim assemblyBtn6 As ButtonDefinition = controlDefs.AddButtonDefinition("Covert to sheetmetal", "ThanhN_Assembly_Btn6", CommandTypesEnum.kShapeEditCmdType, addInClientID,
                                                                                   Nothing,
                                                                                   "Chuyển part thành sheet metal")
            AddHandler assemblyBtn6.OnExecute, AddressOf Assembly.Buttons.part.Ass_Part_1.OnExecute
            buttonsList.Add(assemblyBtn6)

            Dim assemblyBtn7 As ButtonDefinition = controlDefs.AddButtonDefinition("Đổi đơn vị part & Cụm lắp", "ThanhN_Assembly_Btn7", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing)
            AddHandler assemblyBtn7.OnExecute, AddressOf Assembly.Buttons.part.Ass_Part_2.OnExecute
            buttonsList.Add(assemblyBtn7)

            Dim assemblyBtn8 As ButtonDefinition = controlDefs.AddButtonDefinition("Generic part to steel", "ThanhN_Assembly_Btn8", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing,
                                                                                   "Thay đổi vật liệu part từ generic sang steel")
            AddHandler assemblyBtn8.OnExecute, AddressOf Assembly.Buttons.part.Ass_Part_3.OnExecute
            buttonsList.Add(assemblyBtn8)

            Dim assemblyBtn9 As ButtonDefinition = controlDefs.AddButtonDefinition("Trải ALL Sheetmetal", "ThanhN_Assembly_Btn9", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing)
            AddHandler assemblyBtn9.OnExecute, AddressOf Assembly.Buttons.part.Ass_Part_4.OnExecute
            buttonsList.Add(assemblyBtn9)

            Dim assemblyBtn10 As ButtonDefinition = controlDefs.AddButtonDefinition("Tat ALL Adaptive cum LG", "ThanhN_Assembly_Btn10", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing)
            AddHandler assemblyBtn10.OnExecute, AddressOf Assembly.Buttons.caclenhlapghep.Ass_6.OnExecute
            buttonsList.Add(assemblyBtn10)

            Dim assemblyBtn11 As ButtonDefinition = controlDefs.AddButtonDefinition("An all plane part", "ThanhN_Assembly_Btn11", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing)
            AddHandler assemblyBtn11.OnExecute, AddressOf Assembly.Buttons.caclenhlapghep.Ass_7.OnExecute
            buttonsList.Add(assemblyBtn11)

            Dim assemblyBtn12 As ButtonDefinition = controlDefs.AddButtonDefinition("Xoa mau ghi de len part", "ThanhN_Assembly_Btn12", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing)
            AddHandler assemblyBtn12.OnExecute, AddressOf Assembly.Buttons.caclenhlapghep.Ass_8.OnExecute
            buttonsList.Add(assemblyBtn12)

            Dim assemblyBtn13 As ButtonDefinition = controlDefs.AddButtonDefinition("Import step to part", "ThanhN_Assembly_Btn13", CommandTypesEnum.kShapeEditCmdType, addInClientID,
                                                                                    Nothing,
                                                                                    "Import all file to part tự lưu, xóa liên kết lưu file tự động")
            AddHandler assemblyBtn13.OnExecute, AddressOf Assembly.Buttons.Import_step_to_part.OnExecute
            buttonsList.Add(assemblyBtn13)

            Dim assemblyBtn14 As ButtonDefinition = controlDefs.AddButtonDefinition("Save coppy to replace part", "ThanhN_Assembly_Btn14", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing)
            AddHandler assemblyBtn14.OnExecute, AddressOf Assembly.Buttons.caclenhlapghep.Ass_Frame_1.OnExecute
            buttonsList.Add(assemblyBtn14)

            Dim assemblyBtn15 As ButtonDefinition = controlDefs.AddButtonDefinition("Design Assitan", "ThanhN_Assembly_Btn15", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing)
            AddHandler assemblyBtn15.OnExecute, AddressOf Assembly.Buttons.caclenhlapghep.Ass_10.OnExecute
            buttonsList.Add(assemblyBtn15)

            Dim assemblyBtn16 As ButtonDefinition = controlDefs.AddButtonDefinition("Btach PL trong Ass top LVer", "ThanhN_Assembly_Btn16", CommandTypesEnum.kShapeEditCmdType, addInClientID,
                                                                                    Nothing,
                                                                                    "Tạo, ghép các chi tiết shetmetal để Phuc cụ bóc tách top lever")
            AddHandler assemblyBtn16.OnExecute, AddressOf Assembly.Buttons.caclenhlapghep.Ass_11.OnExecute
            buttonsList.Add(assemblyBtn16)

            Dim assemblyBtn17 As ButtonDefinition = controlDefs.AddButtonDefinition("Lọc các PL trong ASS ALL", "ThanhN_Assembly_Btn17", CommandTypesEnum.kShapeEditCmdType, addInClientID,
                                                                                    Nothing, "Tạo, ghép shetmetal to assembly all lever - lọc đếm phân loại ko ộng dồn ko trùng partnumber")
            AddHandler assemblyBtn17.OnExecute, AddressOf Assembly.Buttons.caclenhlapghep.Ass_12.OnExecute
            buttonsList.Add(assemblyBtn17)

            Dim assemblyBtn18 As ButtonDefinition = controlDefs.AddButtonDefinition("Tập hợp PL trong ASS ALL", "ThanhN_Assembly_Btn18", CommandTypesEnum.kShapeEditCmdType, addInClientID,
                                                                                    Nothing, "Tạo, ghép shetmetal to assembly all lever lấy tất cả các tấm kể cả trung tên partnumber")
            AddHandler assemblyBtn18.OnExecute, AddressOf Assembly.Buttons.caclenhlapghep.Ass_13.OnExecute
            buttonsList.Add(assemblyBtn18)

            Dim assemblyBtn19 As ButtonDefinition = controlDefs.AddButtonDefinition("Ghép shetmetal,mua,thư viện trong ASS ALL", "ThanhN_Assembly_Btn19", CommandTypesEnum.kShapeEditCmdType, addInClientID,
                                                                                   Nothing, "Tạo, ghép shetmetal,mua,thư viện to assembly all lever lấy tổng partnumber")
            AddHandler assemblyBtn19.OnExecute, AddressOf Assembly.Buttons.caclenhlapghep.ass_14.OnExecute
            buttonsList.Add(assemblyBtn19)

            Dim assemblyBtn20 As ButtonDefinition = controlDefs.AddButtonDefinition("Lọc các loại tấm xuất hiện trong ASS ALL", "ThanhN_Assembly_Btn20", CommandTypesEnum.kShapeEditCmdType, addInClientID,
                                                                                   Nothing, "")
            AddHandler assemblyBtn20.OnExecute, AddressOf Assembly.Buttons.caclenhlapghep.ass_15.OnExecute
            buttonsList.Add(assemblyBtn20)

            Dim assemblyBtn21 As ButtonDefinition = controlDefs.AddButtonDefinition("UPDATE DESIGN STANDARD", "ThanhN_Assembly_Btn21", CommandTypesEnum.kShapeEditCmdType, addInClientID,
                                                                                   Nothing, "Up date cho các tool tinh toán tiêu chuẩn ví dụ như buloong, key,...")
            AddHandler assemblyBtn21.OnExecute, AddressOf Assembly.Buttons.caclenhlapghep.ass_16.OnExecute
            buttonsList.Add(assemblyBtn21)

            Dim assemblyBtn22 As ButtonDefinition = controlDefs.AddButtonDefinition("Thông số part", "ThanhN_Assembly_Btn22", CommandTypesEnum.kShapeEditCmdType, addInClientID,
                                                                                   Nothing, "Up date cho các tool tinh toán tiêu chuẩn ví dụ như buloong, key,...")
            AddHandler assemblyBtn22.OnExecute, AddressOf Assembly.Buttons.caclenhlapghep.ass_17.OnExecute
            buttonsList.Add(assemblyBtn22)

            Dim assemblyBtn23 As ButtonDefinition = controlDefs.AddButtonDefinition("Auto drawing v8", "ThanhN_Assembly_Btn23", CommandTypesEnum.kShapeEditCmdType, addInClientID,
                                                                                 Nothing, "Auto drawing cho ALL lever.")
            AddHandler assemblyBtn23.OnExecute, AddressOf Assembly.Buttons.AutoCreateDrawing.AutoDrawingV8.OnExecute
            buttonsList.Add(assemblyBtn23)

            Dim assemblyBtn24 As ButtonDefinition = controlDefs.AddButtonDefinition("Auto drawing ASS", "ThanhN_Assembly_Btn24", CommandTypesEnum.kShapeEditCmdType, addInClientID,
                                                                            Nothing, "Auto drawing cho Top lever chỉ áp dụng cho các Assembly.")
            AddHandler assemblyBtn24.OnExecute, AddressOf Assembly.Buttons.AutoCreateDrawing.AutoDrawingASSTopLV.OnExecute
            buttonsList.Add(assemblyBtn24)


            Dim assemblyBtn25 As ButtonDefinition = controlDefs.AddButtonDefinition("Auto drawing ASS", "ThanhN_Assembly_Btn25", CommandTypesEnum.kShapeEditCmdType, addInClientID,
                                                                            Nothing, "Auto drawing cho Top lever chỉ áp dụng cho các Assembly.")
            AddHandler assemblyBtn25.OnExecute, AddressOf Assembly.Buttons.AutoCreateDrawing.AutoDrawingASSpartTopLV.OnExecute
            buttonsList.Add(assemblyBtn25)

            Dim assemblyBtn26 As ButtonDefinition = controlDefs.AddButtonDefinition("Xem lỗi cắt Frame", "ThanhN_Assembly_Btn26", CommandTypesEnum.kShapeEditCmdType, addInClientID,
                                                                          Nothing, "Auto drawing cho Top lever chỉ áp dụng cho các Assembly.")
            AddHandler assemblyBtn26.OnExecute, AddressOf Assembly.Buttons.Frame.Ass_Frame_1.OnExecute
            buttonsList.Add(assemblyBtn26)

            Dim assemblyBtn27 As ButtonDefinition = controlDefs.AddButtonDefinition("Xem lỗi cắt Frame2", "ThanhN_Assembly_Btn27", CommandTypesEnum.kShapeEditCmdType, addInClientID,
                                                                          Nothing, "Auto drawing cho Top lever chỉ áp dụng cho các Assembly.")
            AddHandler assemblyBtn27.OnExecute, AddressOf Assembly.Buttons.Frame.Ass_Frame_2.OnExecute
            buttonsList.Add(assemblyBtn27)

            Dim assemblyBtn28 As ButtonDefinition = controlDefs.AddButtonDefinition("Xem lỗi cắt Frame3", "ThanhN_Assembly_Btn28", CommandTypesEnum.kShapeEditCmdType, addInClientID,
                                                                          Nothing, "Auto drawing cho Top lever chỉ áp dụng cho các Assembly.")
            AddHandler assemblyBtn28.OnExecute, AddressOf Assembly.Buttons.Frame.Ass_Frame_3.OnExecute
            buttonsList.Add(assemblyBtn28)
        End Sub
    End Class
End Namespace
