Imports Inventor

Namespace ToolInventor2020
    Public Class Assembly3Buttons
        Public Shared Sub Register(controlDefs As Inventor.ControlDefinitions, addInClientID As String, buttonsList As System.Collections.Generic.List(Of ButtonDefinition))

            ' Create Assembly buttons explicitly (no loop) so each button can have distinct implementation
            Dim assemblyBtnb1 As ButtonDefinition = controlDefs.AddButtonDefinition("Ghi tên theo chuỗi Top lever Partnumber", "ToolInventor2020_Assembly_Btnb1", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing)
            AddHandler assemblyBtnb1.OnExecute, AddressOf Assembly3.Buttons.Button1.OnExecute
            buttonsList.Add(assemblyBtnb1)

            Dim assemblyBtnb2 As ButtonDefinition = controlDefs.AddButtonDefinition("Ghi tên theo chuỗi Top lever Stocknumber", "ToolInventor2020_Assembly_Btnb2", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing)
            AddHandler assemblyBtnb2.OnExecute, AddressOf Assembly3.Buttons.Button2.OnExecute
            buttonsList.Add(assemblyBtnb2)

            Dim assemblyBtnb3 As ButtonDefinition = controlDefs.AddButtonDefinition("Đánh STT cho item1 Top lever", "ToolInventor2020_Assembly_Btnb3", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing)
            AddHandler assemblyBtnb3.OnExecute, AddressOf Assembly3.Buttons.Button3.OnExecute
            buttonsList.Add(assemblyBtnb3)

            Dim assemblyBtnb4 As ButtonDefinition = controlDefs.AddButtonDefinition("Đánh STT cho cụm xếp VT partnumber Top lever", "ToolInventor2020_Assembly_Btnb4", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing)
            AddHandler assemblyBtnb4.OnExecute, AddressOf Assembly3.Buttons.Button4.OnExecute
            buttonsList.Add(assemblyBtnb4)

            Dim assemblyBtnb5 As ButtonDefinition = controlDefs.AddButtonDefinition("Đánh STT, item1 VT Buy Top lever về cuối", "ToolInventor2020_Assembly_Btnb5", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing)
            AddHandler assemblyBtnb5.OnExecute, AddressOf Assembly3.Buttons.Button5.OnExecute
            buttonsList.Add(assemblyBtnb5)

            Dim assemblyBtnb6 As ButtonDefinition = controlDefs.AddButtonDefinition("Đánh STT top lever", "ToolInventor2020_Assembly_Btnb6", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing)
            AddHandler assemblyBtnb6.OnExecute, AddressOf Assembly3.Buttons.Button6.OnExecute
            buttonsList.Add(assemblyBtnb6)

            Dim assemblyBtnb7 As ButtonDefinition = controlDefs.AddButtonDefinition("Ghi tên PLxx sheetmetal vào trong PartNB", "ToolInventor2020_Assembly_Btnb7", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing)
            AddHandler assemblyBtnb7.OnExecute, AddressOf Assembly3.Buttons.Button7.OnExecute
            buttonsList.Add(assemblyBtnb7)

            Dim assemblyBtnb8 As ButtonDefinition = controlDefs.AddButtonDefinition("Ghi tên PLxx sheetmetal vào trong StockNB", "ToolInventor2020_Assembly_Btnb8", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, "Code này sẽ thay tên stocknumber 
theo kích thước bao sau khi trải ")
            AddHandler assemblyBtnb8.OnExecute, AddressOf Assembly3.Buttons.Button8.OnExecute
            buttonsList.Add(assemblyBtnb8)

            Dim assemblyBtnb9 As ButtonDefinition = controlDefs.AddButtonDefinition("Điền thông tin chiều dày tấm vào PL BOM", "ToolInventor2020_Assembly_Btnb9", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, "
Điền thông tin chiều dày theo thickness. chỉ áp dụng với cá sheetmetal! = t1,t2,t3,t4,...")
            AddHandler assemblyBtnb9.OnExecute, AddressOf Assembly3.Buttons.Button9.OnExecute
            buttonsList.Add(assemblyBtnb9)

            Dim assemblyBtnb10 As ButtonDefinition = controlDefs.AddButtonDefinition("STT từ Item sang item1 chỉ dánh part top lever", "ToolInventor2020_Assembly_Btnb10", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, "Code này chỉ coppy SST sang item1 cho part
toplever")
            AddHandler assemblyBtnb10.OnExecute, AddressOf Assembly3.Buttons.Button10.OnExecute
            buttonsList.Add(assemblyBtnb10)

            Dim assemblyBtnb11 As ButtonDefinition = controlDefs.AddButtonDefinition("Item, Item Qty part sang Item1, SL Part all lever", "ToolInventor2020_Assembly_Btnb11", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, "Coppy Item, Item Qty part all leveer sang Item1, SL Part all lever
tối đa 3 lever & không copy vào cụm lắp. chỉ copy part")
            AddHandler assemblyBtnb11.OnExecute, AddressOf Assembly3.Buttons.Button11.OnExecute
            buttonsList.Add(assemblyBtnb11)

            Dim assemblyBtnb12 As ButtonDefinition = controlDefs.AddButtonDefinition("Item, Item Qty part sang Item1, SL Part, PL all lever", "ToolInventor2020_Assembly_Btnb12", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, "Coppy Item, Item Qty part all leveer sang Item1, SL Part, thickness PL sheetmetal all lever
tối đa 3 lever & không copy vào cụm lắp. chỉ copy part")
            AddHandler assemblyBtnb12.OnExecute, AddressOf Assembly3.Buttons.Button12.OnExecute
            buttonsList.Add(assemblyBtnb12)

            Dim assemblyBtnb13 As ButtonDefinition = controlDefs.AddButtonDefinition("Item, Item Qty part sang Item1, SL Part, PL all lever", "ToolInventor2020_Assembly_Btnb13", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing)
            AddHandler assemblyBtnb13.OnExecute, AddressOf Assembly3.Buttons.Button13.OnExecute
            buttonsList.Add(assemblyBtnb13)

            Dim assemblyBtnb14 As ButtonDefinition = controlDefs.AddButtonDefinition("Assembly Action 14", "ToolInventor2020_Assembly_Btnb14", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing)
            AddHandler assemblyBtnb14.OnExecute, AddressOf Assembly3.Buttons.Button14.OnExecute
            buttonsList.Add(assemblyBtnb14)

            Dim assemblyBtnb15 As ButtonDefinition = controlDefs.AddButtonDefinition("11111111111111111111", "ToolInventor2020_Assembly_Btnb15", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing)
            AddHandler assemblyBtnb15.OnExecute, AddressOf Assembly3.Buttons.Button15.OnExecute
            buttonsList.Add(assemblyBtnb15)

            ' --- Temporary test buttons in case pulldown is missing ---
            Dim assemblyTest1 As ButtonDefinition = controlDefs.AddButtonDefinition("Test Action 1 (Assembly3)", "ToolInventor2020_Assembly_Test1", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, "Temporary test button")
            AddHandler assemblyTest1.OnExecute, AddressOf Assembly3.Buttons.Button1.OnExecute
            buttonsList.Add(assemblyTest1)

            Dim assemblyTest2 As ButtonDefinition = controlDefs.AddButtonDefinition("Test Action 2 (Assembly3)", "ToolInventor2020_Assembly_Test2", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, "Temporary test button")
            AddHandler assemblyTest2.OnExecute, AddressOf Assembly3.Buttons.Button2.OnExecute
            buttonsList.Add(assemblyTest2)

        End Sub
    End Class
End Namespace
