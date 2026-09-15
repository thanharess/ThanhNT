Imports Inventor

Namespace ToolInventor2020
    Public Class ToolButtonDrawing
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

            Dim assemblyFolder4 As String = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)
            Dim configured As String = Nothing
            Try
                configured = My.Settings.ImageFolder
            Catch
                configured = Nothing
            End Try

            Dim iconsFolder As String = Nothing
            If Not String.IsNullOrWhiteSpace(configured) AndAlso System.IO.Directory.Exists(configured) Then
                iconsFolder = System.IO.Path.Combine(configured, "Drawing")
            Else
                iconsFolder = System.IO.Path.Combine(assemblyFolder4, "Code addin inventor", "Images", "Drawing")
            End If

#Region "Prepare icon paths"
            Dim ToolDrawLargePath1 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim ToolDrawSmallPath1 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim ToolDrawLargePath2 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim ToolDrawSmallPath2 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim ToolDrawLargePath3 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim ToolDrawSmallPath3 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim ToolDrawLargePath4 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim ToolDrawSmallPath4 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim ToolDrawLargePath5 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim ToolDrawSmallPath5 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim ToolDrawLargePath6 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim ToolDrawSmallPath6 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim ToolDrawLargePath7 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim ToolDrawSmallPath7 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim ToolDrawLargePath8 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim ToolDrawSmallPath8 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim ToolDrawLargePath9 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim ToolDrawSmallPath9 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim ToolDrawLargePath10 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim ToolDrawSmallPath10 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim ToolDrawLargePath11 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim ToolDrawSmallPath11 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim ToolDrawLargePath12 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim ToolDrawSmallPath12 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim ToolDrawLargePath13 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim ToolDrawSmallPath13 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim ToolDrawLargePath14 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim ToolDrawSmallPath14 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim ToolDrawLargePath15 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim ToolDrawSmallPath15 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim ToolDrawLargePath16 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim ToolDrawSmallPath16 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim ToolDrawLargePath17 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim ToolDrawSmallPath17 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim ToolDrawLargePath18 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim ToolDrawSmallPath18 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim ToolDrawLargePath19 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim ToolDrawSmallPath19 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim ToolDrawLargePath20 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim ToolDrawSmallPath20 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim ToolDrawLargePath21 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim ToolDrawSmallPath21 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim ToolDrawLargePath22 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim ToolDrawSmallPath22 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim ToolDrawLargePath23 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim ToolDrawSmallPath23 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim ToolDrawLargePath24 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim ToolDrawSmallPath24 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim ToolDrawLargePath25 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim ToolDrawSmallPath25 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim ToolDrawLargePath26 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim ToolDrawSmallPath26 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim ToolDrawLargePath27 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")
            Dim ToolDrawSmallPath27 As String = System.IO.Path.Combine(iconsFolder, "i34 1.bmp")
            Dim ToolDrawLargePath28 As String = System.IO.Path.Combine(iconsFolder, "i34.bmp")

            ' Prepare IPictureDisp icons (use provided largeIcon/smallIcon as fallback)
            Dim ToolDrawLargeIcon1 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawLargePath1), LoadIconFromPath(ToolDrawLargePath1), largeIcon)
            Dim ToolDrawSmallIcon1 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawSmallPath1), LoadIconFromPath(ToolDrawSmallPath1), smallIcon)
            Dim ToolDrawLargeIcon2 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawLargePath2), LoadIconFromPath(ToolDrawLargePath2), largeIcon)
            Dim ToolDrawSmallIcon2 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawSmallPath2), LoadIconFromPath(ToolDrawSmallPath2), smallIcon)
            Dim ToolDrawLargeIcon3 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawLargePath3), LoadIconFromPath(ToolDrawLargePath3), largeIcon)
            Dim ToolDrawSmallIcon3 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawSmallPath3), LoadIconFromPath(ToolDrawSmallPath3), smallIcon)
            Dim ToolDrawLargeIcon4 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawLargePath4), LoadIconFromPath(ToolDrawLargePath4), largeIcon)
            Dim ToolDrawSmallIcon4 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawSmallPath4), LoadIconFromPath(ToolDrawSmallPath4), smallIcon)
            Dim ToolDrawLargeIcon5 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawLargePath5), LoadIconFromPath(ToolDrawLargePath5), largeIcon)
            Dim ToolDrawSmallIcon5 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawSmallPath5), LoadIconFromPath(ToolDrawSmallPath5), smallIcon)
            Dim ToolDrawLargeIcon6 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawLargePath6), LoadIconFromPath(ToolDrawLargePath6), largeIcon)
            Dim ToolDrawSmallIcon6 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawSmallPath6), LoadIconFromPath(ToolDrawSmallPath6), smallIcon)
            Dim ToolDrawLargeIcon7 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawLargePath7), LoadIconFromPath(ToolDrawLargePath7), largeIcon)
            Dim ToolDrawSmallIcon7 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawSmallPath7), LoadIconFromPath(ToolDrawSmallPath7), smallIcon)
            Dim ToolDrawLargeIcon8 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawLargePath8), LoadIconFromPath(ToolDrawLargePath8), largeIcon)
            Dim ToolDrawSmallIcon8 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawSmallPath8), LoadIconFromPath(ToolDrawSmallPath8), smallIcon)
            Dim ToolDrawLargeIcon9 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawLargePath9), LoadIconFromPath(ToolDrawLargePath9), largeIcon)
            Dim ToolDrawSmallIcon9 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawSmallPath9), LoadIconFromPath(ToolDrawSmallPath9), smallIcon)
            Dim ToolDrawLargeIcon10 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawLargePath10), LoadIconFromPath(ToolDrawLargePath10), largeIcon)
            Dim ToolDrawSmallIcon10 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawSmallPath10), LoadIconFromPath(ToolDrawSmallPath10), smallIcon)
            Dim ToolDrawLargeIcon11 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawLargePath11), LoadIconFromPath(ToolDrawLargePath11), largeIcon)
            Dim ToolDrawSmallIcon11 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawSmallPath11), LoadIconFromPath(ToolDrawSmallPath11), smallIcon)
            Dim ToolDrawLargeIcon12 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawLargePath12), LoadIconFromPath(ToolDrawLargePath12), largeIcon)
            Dim ToolDrawSmallIcon12 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawSmallPath12), LoadIconFromPath(ToolDrawSmallPath12), smallIcon)
            Dim ToolDrawLargeIcon13 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawLargePath13), LoadIconFromPath(ToolDrawLargePath13), largeIcon)
            Dim ToolDrawSmallIcon13 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawSmallPath13), LoadIconFromPath(ToolDrawSmallPath13), smallIcon)
            Dim ToolDrawLargeIcon14 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawLargePath14), LoadIconFromPath(ToolDrawLargePath14), largeIcon)
            Dim ToolDrawSmallIcon14 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawSmallPath14), LoadIconFromPath(ToolDrawSmallPath14), smallIcon)
            Dim ToolDrawLargeIcon15 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawLargePath15), LoadIconFromPath(ToolDrawLargePath15), largeIcon)
            Dim ToolDrawSmallIcon15 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawSmallPath15), LoadIconFromPath(ToolDrawSmallPath15), smallIcon)
            Dim ToolDrawLargeIcon16 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawLargePath16), LoadIconFromPath(ToolDrawLargePath16), largeIcon)
            Dim ToolDrawSmallIcon16 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawSmallPath16), LoadIconFromPath(ToolDrawSmallPath16), smallIcon)
            Dim ToolDrawLargeIcon17 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawLargePath17), LoadIconFromPath(ToolDrawLargePath17), largeIcon)
            Dim ToolDrawSmallIcon17 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawSmallPath17), LoadIconFromPath(ToolDrawSmallPath17), smallIcon)
            Dim ToolDrawLargeIcon18 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawLargePath18), LoadIconFromPath(ToolDrawLargePath18), largeIcon)
            Dim ToolDrawSmallIcon18 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawSmallPath18), LoadIconFromPath(ToolDrawSmallPath18), smallIcon)
            Dim ToolDrawLargeIcon19 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawLargePath19), LoadIconFromPath(ToolDrawLargePath19), largeIcon)
            Dim ToolDrawSmallIcon19 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawSmallPath19), LoadIconFromPath(ToolDrawSmallPath19), smallIcon)
            Dim ToolDrawLargeIcon20 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawLargePath20), LoadIconFromPath(ToolDrawLargePath20), largeIcon)
            Dim ToolDrawSmallIcon20 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawSmallPath20), LoadIconFromPath(ToolDrawSmallPath20), smallIcon)
            Dim ToolDrawLargeIcon21 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawLargePath21), LoadIconFromPath(ToolDrawLargePath21), largeIcon)
            Dim ToolDrawSmallIcon21 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawSmallPath21), LoadIconFromPath(ToolDrawSmallPath21), smallIcon)
            Dim ToolDrawLargeIcon22 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawLargePath22), LoadIconFromPath(ToolDrawLargePath22), largeIcon)
            Dim ToolDrawSmallIcon22 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawSmallPath22), LoadIconFromPath(ToolDrawSmallPath22), smallIcon)
            Dim ToolDrawLargeIcon23 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawLargePath23), LoadIconFromPath(ToolDrawLargePath23), largeIcon)
            Dim ToolDrawSmallIcon23 As stdole.IPictureDisp = If(System.IO.File.Exists(ToolDrawSmallPath23), LoadIconFromPath(ToolDrawSmallPath23), smallIcon)
#End Region


            ' Create Drawing buttons explicitly (no loop) so each button can have distinct implementation
            Dim TooldrawingBtn1 As ButtonDefinition = controlDefs.AddButtonDefinition("Sửa số thập phân dim", "ToolInventor2020_Drawing_Btn1", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing, ToolDrawSmallIcon1, ToolDrawLargeIcon1)
            AddHandler TooldrawingBtn1.OnExecute, AddressOf Drawing.Buttons.Draw_1.OnExecute
            buttonsList.Add(TooldrawingBtn1)

            Dim TooldrawingBtn2 As ButtonDefinition = controlDefs.AddButtonDefinition("Auto Dim", "ToolInventor2020_Drawing_Btn2", CommandTypesEnum.kShapeEditCmdType, addInClientID,
                                                                                   Nothing, "Auto dim kích thước line, lỗ, lỗ ren", ToolDrawSmallIcon2, ToolDrawLargeIcon2)
            AddHandler TooldrawingBtn2.OnExecute, AddressOf Drawing.Buttons.Drawdim.draw_2.OnExecute
            buttonsList.Add(TooldrawingBtn2)

            Dim TooldrawingBtn3 As ButtonDefinition = controlDefs.AddButtonDefinition("Xóa, căn chỉnh dim", "ToolInventor2020_Drawing_Btn3", CommandTypesEnum.kShapeEditCmdType, addInClientID,
                                                                                   Nothing,
                                                                                   "Xóa, căn chỉnh dim bằng arange", ToolDrawSmallIcon3, ToolDrawLargeIcon3)
            AddHandler TooldrawingBtn3.OnExecute, AddressOf Drawing.Buttons.Drawdim.draw_3.OnExecute
            buttonsList.Add(TooldrawingBtn3)

            Dim TooldrawingBtn4 As ButtonDefinition = controlDefs.AddButtonDefinition("Tìm Dim bị edit", "ToolInventor2020_Drawing_Btn4", CommandTypesEnum.kShapeEditCmdType, addInClientID,
                                                                                   Nothing, "Kiểm tra kim bị edit rồi hiện màu đỏ những dim bị chỉnh sửa", ToolDrawSmallIcon4, ToolDrawLargeIcon4)
            AddHandler TooldrawingBtn4.OnExecute, AddressOf Drawing.Buttons.Draw_4.OnExecute
            buttonsList.Add(TooldrawingBtn4)

            Dim TooldrawingBtn5 As ButtonDefinition = controlDefs.AddButtonDefinition("Reset part list", "ToolInventor2020_Drawing_Btn5", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing, ToolDrawSmallIcon5, ToolDrawLargeIcon5)
            AddHandler TooldrawingBtn5.OnExecute, AddressOf Drawing.Buttons.Draw_5.OnExecute
            buttonsList.Add(TooldrawingBtn5)

            Dim TooldrawingBtn6 As ButtonDefinition = controlDefs.AddButtonDefinition("Ghi Partlist ENG,VIE", "ToolInventor2020_Drawing_Btn6", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing,
                                                                                  "1,2. Update Drawing Views trong sheet hoặc all sheet thay vì ấn bằng tay từng sheet 1" & vbLf &
                                                                                   "3. Ghi mã chi tiết vào category trong partlist. yêu cầu partlist phải có sẵn ô này thì mới được!", ToolDrawSmallIcon6, ToolDrawLargeIcon6)
            AddHandler TooldrawingBtn6.OnExecute, AddressOf Drawing.Buttons.DrawPartList.draw_6.OnExecute
            buttonsList.Add(TooldrawingBtn6)

            Dim TooldrawingBtn7 As ButtonDefinition = controlDefs.AddButtonDefinition("Nút chuyển Sheet", "ToolInventor2020_Drawing_Btn7", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing, ToolDrawSmallIcon7, ToolDrawLargeIcon7)
            AddHandler TooldrawingBtn7.OnExecute, AddressOf Drawing.Buttons.Draw_7.OnExecute
            buttonsList.Add(TooldrawingBtn7)

            Dim TooldrawingBtn8 As ButtonDefinition = controlDefs.AddButtonDefinition("Đổi scale view", "ToolInventor2020_Drawing_Btn8", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing, ToolDrawSmallIcon8, ToolDrawLargeIcon8)
            AddHandler TooldrawingBtn8.OnExecute, AddressOf Drawing.Buttons.DrawView.Draw_8.OnExecute
            buttonsList.Add(TooldrawingBtn8)

            Dim TooldrawingBtn9 As ButtonDefinition = controlDefs.AddButtonDefinition("Xóa Part,ASS trùng lặp", "ToolInventor2020_Drawing_Btn9", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing, ToolDrawSmallIcon9, ToolDrawLargeIcon9)
            AddHandler TooldrawingBtn9.OnExecute, AddressOf Drawing.Buttons.Draw_9.OnExecute
            buttonsList.Add(TooldrawingBtn9)

            Dim TooldrawingBtn10 As ButtonDefinition = controlDefs.AddButtonDefinition("Xóa centermark, centerline", "ToolInventor2020_Drawing_Btn10", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing, ToolDrawSmallIcon10, ToolDrawLargeIcon10)
            AddHandler TooldrawingBtn10.OnExecute, AddressOf Drawing.Buttons.Draw_10.OnExecute
            buttonsList.Add(TooldrawingBtn10)
            '
            Dim TooldrawingBtn11 As ButtonDefinition = controlDefs.AddButtonDefinition("Update Views", "ToolInventor2020_Drawing_Btn11", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing,
                                                                                   "Update Drawing Views trong sheet hoặc all sheet thay vì ấn bằng tay từng sheet 1", ToolDrawSmallIcon11, ToolDrawLargeIcon11)
            AddHandler TooldrawingBtn11.OnExecute, AddressOf Drawing.Buttons.Vi_tri_file.OnExecute
            buttonsList.Add(TooldrawingBtn11)

            Dim TooldrawingBtn12 As ButtonDefinition = controlDefs.AddButtonDefinition("Auto Dim hole ko ar2", "ToolInventor2020_Drawing_Btn12", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing,
                                                                                   "Ghi mã chi tiết vào category trong partlist. yêu cầu partlist phải có sẵn ô này thì mới được!", ToolDrawSmallIcon12, ToolDrawLargeIcon12)
            ' AddHandler TooldrawingBtn12.OnExecute, AddressOf Drawing.Buttons.Drawdim.Draw_dim_base_line_b.OnExecute
            ' buttonsList.Add(TooldrawingBtn12)

            Dim TooldrawingBtn13 As ButtonDefinition = controlDefs.AddButtonDefinition("Thay chữ cái view sheet", "ToolInventor2020_Drawing_Btn13", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing,
                                                                                   "Đổi tên view hiển thị trong sheet", ToolDrawSmallIcon13, ToolDrawLargeIcon13)
            AddHandler TooldrawingBtn13.OnExecute, AddressOf Drawing.Buttons.draw_13.OnExecute
            buttonsList.Add(TooldrawingBtn13)

            Dim TooldrawingBtn14 As ButtonDefinition = controlDefs.AddButtonDefinition("Đổi scale Hatch", "ToolInventor2020_Drawing_Btn14", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing,
                                                                                   "Chưa ok", ToolDrawSmallIcon14, ToolDrawLargeIcon14)
            AddHandler TooldrawingBtn14.OnExecute, AddressOf Drawing.Buttons.draw_14.OnExecute
            buttonsList.Add(TooldrawingBtn14)

            Dim TooldrawingBtn15 As ButtonDefinition = controlDefs.AddButtonDefinition("Delete", "ToolInventor2020_Drawing_Btn15", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing,
                                                                                   "Xóa các thứ liên quan dến text. ko hỗ trợ xóa weld cho bản 2024 trở xuống", ToolDrawSmallIcon15, ToolDrawLargeIcon15)
            AddHandler TooldrawingBtn15.OnExecute, AddressOf Drawing.Buttons.Drawdelete.Draw_delete.OnExecute

            buttonsList.Add(TooldrawingBtn15)


            Dim TooldrawingBtn17 As ButtonDefinition = controlDefs.AddButtonDefinition("chain", "ToolInventor2020_Drawing_Btn17", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing, ToolDrawSmallIcon17, ToolDrawLargeIcon17)
            ' AddHandler TooldrawingBtn17.OnExecute, AddressOf Drawing.Buttons.Drawdim.Draw_dim_chain_line_b.OnExecute

            ' buttonsList.Add(TooldrawingBtn17)
            Dim TooldrawingBtn18 As ButtonDefinition = controlDefs.AddButtonDefinition("Auto Dim hole ko ard2", "ToolInventor2020_Drawing_Btn18", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing, Nothing, ToolDrawSmallIcon18, ToolDrawLargeIcon18)
            ' AddHandler TooldrawingBtn18.OnExecute, AddressOf Drawing.Buttons.draw_15e.OnExecute

            ' buttonsList.Add(TooldrawingBtn18)





            ''''''''' tool ngoài

            Dim TooldrawingBtn16 As ButtonDefinition = controlDefs.AddButtonDefinition("Mở nơi lưu File", "ToolInventor2020_Drawing_Btn16", CommandTypesEnum.kShapeEditCmdType, addInClientID, Nothing,
                                                                                   "Mở vị trí lưu file", ToolDrawSmallIcon16, ToolDrawLargeIcon16)
            AddHandler TooldrawingBtn16.OnExecute, AddressOf Toolngoai.Vitrifile.Vitrifile

            buttonsList.Add(TooldrawingBtn16)



        End Sub
    End Class
End Namespace
