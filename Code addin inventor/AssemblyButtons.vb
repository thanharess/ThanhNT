Imports Inventor

Namespace ThanhN
    Public Class AssemblyButtons
        Public Shared Sub Register(controlDefs As Inventor.ControlDefinitions, addInClientID As String, buttonsList As System.Collections.Generic.List(Of ButtonDefinition))


            ' Create Assembly buttons explicitly (no loop) so each button can have distinct implementation
            Dim assemblyBtn1 As ButtonDefinition = controlDefs.AddButtonDefinition("Assembly Action 1", "ThanhN_Assembly_Btn1", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing)
            AddHandler assemblyBtn1.OnExecute, AddressOf Assembly.Buttons.Button1.OnExecute
            buttonsList.Add(assemblyBtn1)

            Dim assemblyBtn2 As ButtonDefinition = controlDefs.AddButtonDefinition("Assembly Action 2", "ThanhN_Assembly_Btn2", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing)
            AddHandler assemblyBtn2.OnExecute, AddressOf Assembly.Buttons.Button2.OnExecute
            buttonsList.Add(assemblyBtn2)

            Dim assemblyBtn3 As ButtonDefinition = controlDefs.AddButtonDefinition("Assembly Action 3", "ThanhN_Assembly_Btn3", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing)
            AddHandler assemblyBtn3.OnExecute, AddressOf Assembly.Buttons.Button3.OnExecute
            buttonsList.Add(assemblyBtn3)

            Dim assemblyBtn4 As ButtonDefinition = controlDefs.AddButtonDefinition("Assembly Action 4", "ThanhN_Assembly_Btn4", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing)
            AddHandler assemblyBtn4.OnExecute, AddressOf Assembly.Buttons.Button4.OnExecute
            buttonsList.Add(assemblyBtn4)

            Dim assemblyBtn5 As ButtonDefinition = controlDefs.AddButtonDefinition("Assembly Action 5", "ThanhN_Assembly_Btn5", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing)
            AddHandler assemblyBtn5.OnExecute, AddressOf Assembly.Buttons.Button5.OnExecute
            buttonsList.Add(assemblyBtn5)

            Dim assemblyBtn6 As ButtonDefinition = controlDefs.AddButtonDefinition("Assembly Action 6", "ThanhN_Assembly_Btn6", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing)
            AddHandler assemblyBtn6.OnExecute, AddressOf Assembly.Buttons.Button6.OnExecute
            buttonsList.Add(assemblyBtn6)

            Dim assemblyBtn7 As ButtonDefinition = controlDefs.AddButtonDefinition("Assembly Action 7", "ThanhN_Assembly_Btn7", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing)
            AddHandler assemblyBtn7.OnExecute, AddressOf Assembly.Buttons.Button7.OnExecute
            buttonsList.Add(assemblyBtn7)

            Dim assemblyBtn8 As ButtonDefinition = controlDefs.AddButtonDefinition("Assembly Action 8", "ThanhN_Assembly_Btn8", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing)
            AddHandler assemblyBtn8.OnExecute, AddressOf Assembly.Buttons.Button8.OnExecute
            buttonsList.Add(assemblyBtn8)

            Dim assemblyBtn9 As ButtonDefinition = controlDefs.AddButtonDefinition("Assembly Action 9", "ThanhN_Assembly_Btn9", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing)
            AddHandler assemblyBtn9.OnExecute, AddressOf Assembly.Buttons.Button9.OnExecute
            buttonsList.Add(assemblyBtn9)

            Dim assemblyBtn10 As ButtonDefinition = controlDefs.AddButtonDefinition("Assembly Action 10", "ThanhN_Assembly_Btn10", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing)
            AddHandler assemblyBtn10.OnExecute, AddressOf Assembly.Buttons.Button10.OnExecute
            buttonsList.Add(assemblyBtn10)

            Dim assemblyBtn11 As ButtonDefinition = controlDefs.AddButtonDefinition("Assembly Action 11", "ThanhN_Assembly_Btn11", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing)
            AddHandler assemblyBtn11.OnExecute, AddressOf Assembly.Buttons.Button11.OnExecute
            buttonsList.Add(assemblyBtn11)

            Dim assemblyBtn12 As ButtonDefinition = controlDefs.AddButtonDefinition("Assembly Action 12", "ThanhN_Assembly_Btn12", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing)
            AddHandler assemblyBtn12.OnExecute, AddressOf Assembly.Buttons.Button12.OnExecute
            buttonsList.Add(assemblyBtn12)

            Dim assemblyBtn13 As ButtonDefinition = controlDefs.AddButtonDefinition("Assembly Action 13", "ThanhN_Assembly_Btn13", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing)
            AddHandler assemblyBtn13.OnExecute, AddressOf Assembly.Buttons.Button13.OnExecute
            buttonsList.Add(assemblyBtn13)

            Dim assemblyBtn14 As ButtonDefinition = controlDefs.AddButtonDefinition("Assembly Action 14", "ThanhN_Assembly_Btn14", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing)
            AddHandler assemblyBtn14.OnExecute, AddressOf Assembly.Buttons.Button14.OnExecute
            buttonsList.Add(assemblyBtn14)

            Dim assemblyBtn15 As ButtonDefinition = controlDefs.AddButtonDefinition("Assembly Action 15", "ThanhN_Assembly_Btn15", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing)
            AddHandler assemblyBtn15.OnExecute, AddressOf Assembly.Buttons.Button15.OnExecute
            buttonsList.Add(assemblyBtn15)
        End Sub
    End Class
End Namespace
