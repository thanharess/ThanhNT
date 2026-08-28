Imports Inventor

Namespace ThanhN
    Public Class DrawingButtons
        Public Shared Sub Register(controlDefs As Inventor.ControlDefinitions, addInClientID As String, buttonsList As System.Collections.Generic.List(Of ButtonDefinition))

            ' Create Drawing buttons explicitly (no loop) so each button can have distinct implementation
            Dim drawingBtn1 As ButtonDefinition = controlDefs.AddButtonDefinition("Drawing Action 1", "ThanhN_Drawing_Btn1", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing)
            AddHandler drawingBtn1.OnExecute, AddressOf Drawing.Buttons.Button1.OnExecute
            buttonsList.Add(drawingBtn1)

            Dim DrawingBtn2 As ButtonDefinition = controlDefs.AddButtonDefinition("Drawing Action 2", "ThanhN_Drawing_Btn2", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing)
            AddHandler DrawingBtn2.OnExecute, AddressOf Drawing.Buttons.Button2.OnExecute
            buttonsList.Add(DrawingBtn2)

            Dim DrawingBtn3 As ButtonDefinition = controlDefs.AddButtonDefinition("Drawing Action 3", "ThanhN_Drawing_Btn3", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing)
            AddHandler DrawingBtn3.OnExecute, AddressOf Drawing.Buttons.Button3.OnExecute
            buttonsList.Add(DrawingBtn3)

            Dim DrawingBtn4 As ButtonDefinition = controlDefs.AddButtonDefinition("Drawing Action 4", "ThanhN_Drawing_Btn4", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing)
            AddHandler DrawingBtn4.OnExecute, AddressOf Drawing.Buttons.Button4.OnExecute
            buttonsList.Add(DrawingBtn4)

            Dim DrawingBtn5 As ButtonDefinition = controlDefs.AddButtonDefinition("Drawing Action 5", "ThanhN_Drawing_Btn5", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing)
            AddHandler DrawingBtn5.OnExecute, AddressOf Drawing.Buttons.Button5.OnExecute
            buttonsList.Add(DrawingBtn5)

            Dim DrawingBtn6 As ButtonDefinition = controlDefs.AddButtonDefinition("Drawing Action 6", "ThanhN_Drawing_Btn6", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing)
            AddHandler DrawingBtn6.OnExecute, AddressOf Drawing.Buttons.Button6.OnExecute
            buttonsList.Add(DrawingBtn6)

            Dim DrawingBtn7 As ButtonDefinition = controlDefs.AddButtonDefinition("Drawing Action 7", "ThanhN_Drawing_Btn7", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing)
            AddHandler DrawingBtn7.OnExecute, AddressOf Drawing.Buttons.Button7.OnExecute
            buttonsList.Add(DrawingBtn7)

            Dim DrawingBtn8 As ButtonDefinition = controlDefs.AddButtonDefinition("Drawing Action 8", "ThanhN_Drawing_Btn8", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing)
            AddHandler DrawingBtn8.OnExecute, AddressOf Drawing.Buttons.Button8.OnExecute
            buttonsList.Add(DrawingBtn8)

            Dim DrawingBtn9 As ButtonDefinition = controlDefs.AddButtonDefinition("Drawing Action 9", "ThanhN_Drawing_Btn9", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing)
            AddHandler DrawingBtn9.OnExecute, AddressOf Drawing.Buttons.Button9.OnExecute
            buttonsList.Add(DrawingBtn9)

            Dim DrawingBtn10 As ButtonDefinition = controlDefs.AddButtonDefinition("Drawing Action 10", "ThanhN_Drawing_Btn10", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing)
            AddHandler DrawingBtn10.OnExecute, AddressOf Drawing.Buttons.Button10.OnExecute
            buttonsList.Add(DrawingBtn10)

            Dim DrawingBtn11 As ButtonDefinition = controlDefs.AddButtonDefinition("Drawing Action 11", "ThanhN_Drawing_Btn11", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing)
            AddHandler DrawingBtn11.OnExecute, AddressOf Drawing.Buttons.Button11.OnExecute
            buttonsList.Add(DrawingBtn11)

            Dim DrawingBtn12 As ButtonDefinition = controlDefs.AddButtonDefinition("Drawing Action 12", "ThanhN_Drawing_Btn12", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing)
            AddHandler DrawingBtn12.OnExecute, AddressOf Drawing.Buttons.Button12.OnExecute
            buttonsList.Add(DrawingBtn12)

            Dim DrawingBtn13 As ButtonDefinition = controlDefs.AddButtonDefinition("Drawing Action 13", "ThanhN_Drawing_Btn13", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing)
            AddHandler DrawingBtn13.OnExecute, AddressOf Drawing.Buttons.Button13.OnExecute
            buttonsList.Add(DrawingBtn13)

            Dim DrawingBtn14 As ButtonDefinition = controlDefs.AddButtonDefinition("Drawing Action 14", "ThanhN_Drawing_Btn14", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing)
            AddHandler DrawingBtn14.OnExecute, AddressOf Drawing.Buttons.Button14.OnExecute
            buttonsList.Add(DrawingBtn14)

            Dim DrawingBtn15 As ButtonDefinition = controlDefs.AddButtonDefinition("Drawing Action 15", "ThanhN_Drawing_Btn15", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing)
            AddHandler DrawingBtn15.OnExecute, AddressOf Drawing.Buttons.Button15.OnExecute
            buttonsList.Add(DrawingBtn15)

        End Sub
    End Class
End Namespace
