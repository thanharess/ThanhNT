Imports Inventor

Namespace ToolInventor2020
    Public Class Assembly2Buttons
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

        Public Shared Sub Register(controlDefs As Inventor.ControlDefinitions, addInClientID As String, buttonsList As System.Collections.Generic.List(Of ButtonDefinition))

            ' Load shared icons for Assembly2 buttons
            Dim iconsFolder As String = "C:\Users\thanh\source\repos\ThanhN\Code addin inventor\Images\Button\Assembly2"
            Dim asmLargePath As String = System.IO.Path.Combine(iconsFolder, "a2_large.bmp")
            Dim asmSmallPath As String = System.IO.Path.Combine(iconsFolder, "a2_small.bmp")
            Dim smallIcon As stdole.IPictureDisp = Nothing
            Dim largeIcon As stdole.IPictureDisp = Nothing
            Try
                If System.IO.File.Exists(asmLargePath) Then largeIcon = LoadIconFromPath(asmLargePath)
                If System.IO.File.Exists(asmSmallPath) Then smallIcon = LoadIconFromPath(asmSmallPath) Else smallIcon = largeIcon
            Catch
                smallIcon = Nothing
                largeIcon = Nothing
            End Try

            ' Create Assembly buttons explicitly (no loop) so each button can have distinct implementation
            Dim assemblyBtna1 As ButtonDefinition = controlDefs.AddButtonDefinition("Thay tên ,STT BOM", "ToolInventor2020_Assembly_Btna1", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing, smallIcon, largeIcon)
            AddHandler assemblyBtna1.OnExecute, AddressOf Assembly2.Buttons.BOMcode.Ass_Bom_1.OnExecute
            buttonsList.Add(assemblyBtna1)

            Dim assemblyBtna2 As ButtonDefinition = controlDefs.AddButtonDefinition("Item1 Buy Top lever về cuối", "ToolInventor2020_Assembly_Btna2", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing, smallIcon, largeIcon)
            AddHandler assemblyBtna2.OnExecute, AddressOf Assembly2.Buttons.BOMcode.Ass_Bom_2.OnExecute
            buttonsList.Add(assemblyBtna2)

            Dim assemblyBtna3 As ButtonDefinition = controlDefs.AddButtonDefinition("Ghi item1,t,SL", "ToolInventor2020_Assembly_Btna3", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing, smallIcon, largeIcon)
            AddHandler assemblyBtna3.OnExecute, AddressOf Assembly2.Buttons.BOMcode.Ass_Bom_3.OnExecute
            buttonsList.Add(assemblyBtna3)

            Dim assemblyBtna4 As ButtonDefinition = controlDefs.AddButtonDefinition("Ghi PL part,stocknumber", "ToolInventor2020_Assembly_Btna4", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing, smallIcon, largeIcon)
            AddHandler assemblyBtna4.OnExecute, AddressOf Assembly2.Buttons.BOMcode.Ass_Bom_4.OnExecute
            buttonsList.Add(assemblyBtna4)

            Dim assemblyBtna5 As ButtonDefinition = controlDefs.AddButtonDefinition("2", "ToolInventor2020_Assembly_Btna5", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing, smallIcon, largeIcon)
            AddHandler assemblyBtna5.OnExecute, AddressOf Assembly2.Buttons.BOMcode.Ass_Bom_5.OnExecute
            buttonsList.Add(assemblyBtna5)

            Dim assemblyBtna6 As ButtonDefinition = controlDefs.AddButtonDefinition("3", "ToolInventor2020_Assembly_Btna6", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing, smallIcon, largeIcon)
            AddHandler assemblyBtna6.OnExecute, AddressOf Assembly2.Buttons.BOMcode.Ass_Bom_6.OnExecute
            buttonsList.Add(assemblyBtna6)

            Dim assemblyBtna7 As ButtonDefinition = controlDefs.AddButtonDefinition("4", "ToolInventor2020_Assembly_Btna7", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing, smallIcon, largeIcon)
            AddHandler assemblyBtna7.OnExecute, AddressOf Assembly2.Buttons.BOMcode.Ass_Bom_7.OnExecute
            buttonsList.Add(assemblyBtna7)

            Dim assemblyBtna8 As ButtonDefinition = controlDefs.AddButtonDefinition("5", "ToolInventor2020_Assembly_Btna8", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, "Code này sẽ thay tên stocknumber 
theo kích thước bao sau khi trải ")
            AddHandler assemblyBtna8.OnExecute, AddressOf Assembly2.Buttons.BOMcode.Ass_Bom_8.OnExecute
            buttonsList.Add(assemblyBtna8)

            Dim assemblyBtna9 As ButtonDefinition = controlDefs.AddButtonDefinition("9", "ToolInventor2020_Assembly_Btna9", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, "
Điền thông tin chiều dày theo thickness. chỉ áp dụng với cá sheetmetal! = t1,t2,t3,t4,...", smallIcon, largeIcon)
            AddHandler assemblyBtna9.OnExecute, AddressOf Assembly2.Buttons.BOMcode.Ass_Bom_9.OnExecute
            buttonsList.Add(assemblyBtna9)

            Dim assemblyBtna10 As ButtonDefinition = controlDefs.AddButtonDefinition("6", "ToolInventor2020_Assembly_Btna10", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, "Code này chỉ coppy SST sang item1 cho part
toplever")
            AddHandler assemblyBtna10.OnExecute, AddressOf Assembly2.Buttons.BOMcode.Ass_Bom_10.OnExecute
            buttonsList.Add(assemblyBtna10)

            Dim assemblyBtna11 As ButtonDefinition = controlDefs.AddButtonDefinition("7", "ToolInventor2020_Assembly_Btna11", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, "Coppy Item, Item Qty part all leveer sang Item1, SL Part all lever
tối đa 3 lever & không copy vào cụm lắp. chỉ copy part")
            AddHandler assemblyBtna11.OnExecute, AddressOf Assembly2.Buttons.BOMcode.Ass_Bom_11.OnExecute
            buttonsList.Add(assemblyBtna11)

            Dim assemblyBtna12 As ButtonDefinition = controlDefs.AddButtonDefinition("8", "ToolInventor2020_Assembly_Btna12", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, "Coppy Item, Item Qty part all leveer sang Item1, SL Part, thickness PL sheetmetal all lever
tối đa 3 lever & không copy vào cụm lắp. chỉ copy part")
            AddHandler assemblyBtna12.OnExecute, AddressOf Assembly2.Buttons.BOMcode.Ass_Bom_12.OnExecute
            buttonsList.Add(assemblyBtna12)

            Dim assemblyBtna13 As ButtonDefinition = controlDefs.AddButtonDefinition("9", "ToolInventor2020_Assembly_Btna13", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing, smallIcon, largeIcon)
            AddHandler assemblyBtna13.OnExecute, AddressOf Assembly2.Buttons.BOMcode.Ass_Bom_13.OnExecute
            buttonsList.Add(assemblyBtna13)

            Dim assemblyBtna14 As ButtonDefinition = controlDefs.AddButtonDefinition("Assembly Action 14", "ToolInventor2020_Assembly_Btna14", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing, smallIcon, largeIcon)
            AddHandler assemblyBtna14.OnExecute, AddressOf Assembly2.Buttons.BOMcode.Ass_Bom_14.OnExecute
            buttonsList.Add(assemblyBtna14)

            Dim assemblyBtna15 As ButtonDefinition = controlDefs.AddButtonDefinition("Assembly Action 15", "ToolInventor2020_Assembly_Btna15", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing, smallIcon, largeIcon)
            AddHandler assemblyBtna15.OnExecute, AddressOf Assembly2.Buttons.BOMcode.Ass_Bom_15.OnExecute
            buttonsList.Add(assemblyBtna15)

            Dim assemblyBtna16 As ButtonDefinition = controlDefs.AddButtonDefinition("Assembly Action 16", "ThanhN_Assembly_Btna16", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing, smallIcon, largeIcon)
            'AddHandler assemblyBtna16.OnExecute, AddressOf Assembly2.Buttons.BOMcode.Ass_Bom_16.OnExecute
            buttonsList.Add(assemblyBtna16)

            Dim assemblyBtna17 As ButtonDefinition = controlDefs.AddButtonDefinition("Assembly Action 17", "ThanhN_Assembly_Btna17", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing, smallIcon, largeIcon)
            'AddHandler assemblyBtna17.OnExecute, AddressOf Assembly2.Buttons.BOMcode.Ass_Bom_17.OnExecute
            buttonsList.Add(assemblyBtna17)

        End Sub
    End Class
End Namespace
