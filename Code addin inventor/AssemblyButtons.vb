Imports System.Diagnostics.Contracts
Imports Inventor

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

            Dim assemblyBtn10 As ButtonDefinition = controlDefs.AddButtonDefinition("Assembly Action 10", "ThanhN_Assembly_Btn10", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing)
            '  AddHandler assemblyBtn10.OnExecute, AddressOf Assembly.Buttons.caclenhlapghep.Button10.OnExecute
            buttonsList.Add(assemblyBtn10)

            Dim assemblyBtn11 As ButtonDefinition = controlDefs.AddButtonDefinition("Assembly Action 11", "ThanhN_Assembly_Btn11", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing)
            'AddHandler assemblyBtn11.OnExecute, AddressOf Assembly.Buttons.caclenhlapghep.Button11.OnExecute
            buttonsList.Add(assemblyBtn11)

            Dim assemblyBtn12 As ButtonDefinition = controlDefs.AddButtonDefinition("Assembly Action 12", "ThanhN_Assembly_Btn12", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing)
            ' AddHandler assemblyBtn12.OnExecute, AddressOf Assembly.Buttons.caclenhlapghep.Button12.OnExecute
            buttonsList.Add(assemblyBtn12)

            Dim assemblyBtn13 As ButtonDefinition = controlDefs.AddButtonDefinition("Assembly Action 13", "ThanhN_Assembly_Btn13", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing)
            ' AddHandler assemblyBtn13.OnExecute, AddressOf Assembly.Buttons.caclenhlapghep.Button13.OnExecute
            buttonsList.Add(assemblyBtn13)

            Dim assemblyBtn14 As ButtonDefinition = controlDefs.AddButtonDefinition("Assembly Action 14", "ThanhN_Assembly_Btn14", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing)
            ' AddHandler assemblyBtn14.OnExecute, AddressOf Assembly.Buttons.caclenhlapghep.Button14.OnExecute
            buttonsList.Add(assemblyBtn14)

            Dim assemblyBtn15 As ButtonDefinition = controlDefs.AddButtonDefinition("Assembly Action 15", "ThanhN_Assembly_Btn15", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing)
            ' AddHandler assemblyBtn15.OnExecute, AddressOf Assembly.Buttons.caclenhlapghep.Button15.OnExecute
            buttonsList.Add(assemblyBtn15)
        End Sub
    End Class
End Namespace
