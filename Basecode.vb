Imports System.Runtime.InteropServices
Imports Inventor
Imports Microsoft.Win32

Namespace ThanhN
    <ProgIdAttribute("ThanhN.StandardAddInServer"),
    GuidAttribute("27785725-854b-490a-ac86-ab9dad7f3cc5")>
    Public Class Basecode
        Implements Inventor.ApplicationAddInServer

        ' AddIn client ID must be a managed String. Define it explicitly to avoid
        ' passing a COM object (System.__ComObject) where a String is expected.
        Private Const AddInClientID As String = "27785725-854b-490a-ac86-ab9dad7f3cc5"

        Private WithEvents m_uiEvents As UserInterfaceEvents
        Private m_partButtons As New System.Collections.Generic.List(Of ButtonDefinition)
        Private m_assemblyButtons As New System.Collections.Generic.List(Of ButtonDefinition)
        Private m_drawingButtons As New System.Collections.Generic.List(Of ButtonDefinition)

#Region "ApplicationAddInServer Members"

        ' This method is called by Inventor when it loads the AddIn. The AddInSiteObject provides access  
        ' to the Inventor Application object. The FirstTime flag indicates if the AddIn is loaded for
        ' the first time. However, with the introduction of the ribbon this argument is always true.
        Public Sub Activate(ByVal addInSiteObject As Inventor.ApplicationAddInSite, ByVal firstTime As Boolean) Implements Inventor.ApplicationAddInServer.Activate
            ' Initialize AddIn members.
            g_inventorApplication = addInSiteObject.Application

            ' Connect to the user-interface events to handle a ribbon reset.
            m_uiEvents = g_inventorApplication.UserInterfaceManager.UserInterfaceEvents

            ' Create button definitions for Part, Assembly and Drawing: 15 buttons each.
            Dim controlDefs As Inventor.ControlDefinitions = g_inventorApplication.CommandManager.ControlDefinitions

            ' No icons: use text-only buttons to avoid bitmap/COM issues
            Dim largeIcon As stdole.IPictureDisp = Nothing
            Dim smallIcon As stdole.IPictureDisp = Nothing

            ' Create Part buttons explicitly (no loop) so each button can have distinct implementation
            Dim partBtn1 As ButtonDefinition = controlDefs.AddButtonDefinition("Thêm dung sai vào dim sketch", "ThanhN_Part_Btn1", CommandTypesEnum.kShapeEditCmdType, AddInClientID, Nothing, Nothing)
            AddHandler partBtn1.OnExecute, AddressOf Part.Buttons.Button1.OnExecute
            m_partButtons.Add(partBtn1)

            Dim partBtn2 As ButtonDefinition = controlDefs.AddButtonDefinition("Xóa fix all sketch", "ThanhN_Part_Btn2", CommandTypesEnum.kShapeEditCmdType, AddInClientID, Nothing, Nothing)
            AddHandler partBtn2.OnExecute, AddressOf Part.Buttons.Button2.OnExecute
            m_partButtons.Add(partBtn2)

            Dim partBtn3 As ButtonDefinition = controlDefs.AddButtonDefinition("Tạo coil line 3d", "ThanhN_Part_Btn3", CommandTypesEnum.kShapeEditCmdType, AddInClientID, Nothing, Nothing)
            AddHandler partBtn3.OnExecute, AddressOf Part.Buttons.Button3.OnExecute
            m_partButtons.Add(partBtn3)

            Dim partBtn4 As ButtonDefinition = controlDefs.AddButtonDefinition("Thay tên body", "ThanhN_Part_Btn4", CommandTypesEnum.kShapeEditCmdType, AddInClientID, Nothing, Nothing)
            AddHandler partBtn4.OnExecute, AddressOf Part.Buttons.Button4.OnExecute
            m_partButtons.Add(partBtn4)

            Dim partBtn5 As ButtonDefinition = controlDefs.AddButtonDefinition("Part Action 5", "ThanhN_Part_Btn5", CommandTypesEnum.kShapeEditCmdType, AddInClientID, Nothing, Nothing)
            AddHandler partBtn5.OnExecute, AddressOf Part.Buttons.Button5.OnExecute
            m_partButtons.Add(partBtn5)

            Dim partBtn6 As ButtonDefinition = controlDefs.AddButtonDefinition("Part Action 6", "ThanhN_Part_Btn6", CommandTypesEnum.kShapeEditCmdType, AddInClientID, Nothing, Nothing)
            AddHandler partBtn6.OnExecute, AddressOf Part.Buttons.Button6.OnExecute
            m_partButtons.Add(partBtn6)

            Dim partBtn7 As ButtonDefinition = controlDefs.AddButtonDefinition("Part Action 7", "ThanhN_Part_Btn7", CommandTypesEnum.kShapeEditCmdType, AddInClientID, Nothing, Nothing)
            AddHandler partBtn7.OnExecute, AddressOf Part.Buttons.Button7.OnExecute
            m_partButtons.Add(partBtn7)

            Dim partBtn8 As ButtonDefinition = controlDefs.AddButtonDefinition("Part Action 8", "ThanhN_Part_Btn8", CommandTypesEnum.kShapeEditCmdType, AddInClientID, Nothing, Nothing)
            AddHandler partBtn8.OnExecute, AddressOf Part.Buttons.Button8.OnExecute
            m_partButtons.Add(partBtn8)

            Dim partBtn9 As ButtonDefinition = controlDefs.AddButtonDefinition("Part Action 9", "ThanhN_Part_Btn9", CommandTypesEnum.kShapeEditCmdType, AddInClientID, Nothing, Nothing)
            AddHandler partBtn9.OnExecute, AddressOf Part.Buttons.Button9.OnExecute
            m_partButtons.Add(partBtn9)

            Dim partBtn10 As ButtonDefinition = controlDefs.AddButtonDefinition("Part Action 10", "ThanhN_Part_Btn10", CommandTypesEnum.kShapeEditCmdType, AddInClientID, Nothing, Nothing)
            AddHandler partBtn10.OnExecute, AddressOf Part.Buttons.Button10.OnExecute
            m_partButtons.Add(partBtn10)

            Dim partBtn11 As ButtonDefinition = controlDefs.AddButtonDefinition("Part Action 11", "ThanhN_Part_Btn11", CommandTypesEnum.kShapeEditCmdType, AddInClientID, Nothing, Nothing)
            AddHandler partBtn11.OnExecute, AddressOf Part.Buttons.Button11.OnExecute
            m_partButtons.Add(partBtn11)

            Dim partBtn12 As ButtonDefinition = controlDefs.AddButtonDefinition("Part Action 12", "ThanhN_Part_Btn12", CommandTypesEnum.kShapeEditCmdType, AddInClientID, Nothing, Nothing)
            AddHandler partBtn12.OnExecute, AddressOf Part.Buttons.Button12.OnExecute
            m_partButtons.Add(partBtn12)

            Dim partBtn13 As ButtonDefinition = controlDefs.AddButtonDefinition("Part Action 13", "ThanhN_Part_Btn13", CommandTypesEnum.kShapeEditCmdType, AddInClientID, Nothing, Nothing)
            AddHandler partBtn13.OnExecute, AddressOf Part.Buttons.Button13.OnExecute
            m_partButtons.Add(partBtn13)

            Dim partBtn14 As ButtonDefinition = controlDefs.AddButtonDefinition("Part Action 14", "ThanhN_Part_Btn14", CommandTypesEnum.kShapeEditCmdType, AddInClientID, Nothing, Nothing)
            AddHandler partBtn14.OnExecute, AddressOf Part.Buttons.Button14.OnExecute
            m_partButtons.Add(partBtn14)

            Dim partBtn15 As ButtonDefinition = controlDefs.AddButtonDefinition("Part Action 15", "ThanhN_Part_Btn15", CommandTypesEnum.kShapeEditCmdType, AddInClientID, Nothing, Nothing)
            AddHandler partBtn15.OnExecute, AddressOf Part.Buttons.Button15.OnExecute
            m_partButtons.Add(partBtn15)

            ' Create Assembly buttons explicitly (no loop) so each button can have distinct implementation
            Dim assemblyBtn1 As ButtonDefinition = controlDefs.AddButtonDefinition("Assembly Action 1", "ThanhN_Assembly_Btn1", CommandTypesEnum.kShapeEditCmdType, AddInClientID, Nothing, Nothing)
            AddHandler assemblyBtn1.OnExecute, AddressOf Assembly.Buttons.Button1.OnExecute
            m_assemblyButtons.Add(assemblyBtn1)

            Dim assemblyBtn2 As ButtonDefinition = controlDefs.AddButtonDefinition("Assembly Action 2", "ThanhN_Assembly_Btn2", CommandTypesEnum.kShapeEditCmdType, AddInClientID, Nothing, Nothing)
            AddHandler assemblyBtn2.OnExecute, AddressOf Assembly.Buttons.Button2.OnExecute
            m_assemblyButtons.Add(assemblyBtn2)

            Dim assemblyBtn3 As ButtonDefinition = controlDefs.AddButtonDefinition("Assembly Action 3", "ThanhN_Assembly_Btn3", CommandTypesEnum.kShapeEditCmdType, AddInClientID, Nothing, Nothing)
            AddHandler assemblyBtn3.OnExecute, AddressOf Assembly.Buttons.Button3.OnExecute
            m_assemblyButtons.Add(assemblyBtn3)

            Dim assemblyBtn4 As ButtonDefinition = controlDefs.AddButtonDefinition("Assembly Action 4", "ThanhN_Assembly_Btn4", CommandTypesEnum.kShapeEditCmdType, AddInClientID, Nothing, Nothing)
            AddHandler assemblyBtn4.OnExecute, AddressOf Assembly.Buttons.Button4.OnExecute
            m_assemblyButtons.Add(assemblyBtn4)

            Dim assemblyBtn5 As ButtonDefinition = controlDefs.AddButtonDefinition("Assembly Action 5", "ThanhN_Assembly_Btn5", CommandTypesEnum.kShapeEditCmdType, AddInClientID, Nothing, Nothing)
            AddHandler assemblyBtn5.OnExecute, AddressOf Assembly.Buttons.Button5.OnExecute
            m_assemblyButtons.Add(assemblyBtn5)

            Dim assemblyBtn6 As ButtonDefinition = controlDefs.AddButtonDefinition("Assembly Action 6", "ThanhN_Assembly_Btn6", CommandTypesEnum.kShapeEditCmdType, AddInClientID, Nothing, Nothing)
            AddHandler assemblyBtn6.OnExecute, AddressOf Assembly.Buttons.Button6.OnExecute
            m_assemblyButtons.Add(assemblyBtn6)

            Dim assemblyBtn7 As ButtonDefinition = controlDefs.AddButtonDefinition("Assembly Action 7", "ThanhN_Assembly_Btn7", CommandTypesEnum.kShapeEditCmdType, AddInClientID, Nothing, Nothing)
            AddHandler assemblyBtn7.OnExecute, AddressOf Assembly.Buttons.Button7.OnExecute
            m_assemblyButtons.Add(assemblyBtn7)

            Dim assemblyBtn8 As ButtonDefinition = controlDefs.AddButtonDefinition("Assembly Action 8", "ThanhN_Assembly_Btn8", CommandTypesEnum.kShapeEditCmdType, AddInClientID, Nothing, Nothing)
            AddHandler assemblyBtn8.OnExecute, AddressOf Assembly.Buttons.Button8.OnExecute
            m_assemblyButtons.Add(assemblyBtn8)

            Dim assemblyBtn9 As ButtonDefinition = controlDefs.AddButtonDefinition("Assembly Action 9", "ThanhN_Assembly_Btn9", CommandTypesEnum.kShapeEditCmdType, AddInClientID, Nothing, Nothing)
            AddHandler assemblyBtn9.OnExecute, AddressOf Assembly.Buttons.Button9.OnExecute
            m_assemblyButtons.Add(assemblyBtn9)

            Dim assemblyBtn10 As ButtonDefinition = controlDefs.AddButtonDefinition("Assembly Action 10", "ThanhN_Assembly_Btn10", CommandTypesEnum.kShapeEditCmdType, AddInClientID, Nothing, Nothing)
            AddHandler assemblyBtn10.OnExecute, AddressOf Assembly.Buttons.Button10.OnExecute
            m_assemblyButtons.Add(assemblyBtn10)

            Dim assemblyBtn11 As ButtonDefinition = controlDefs.AddButtonDefinition("Assembly Action 11", "ThanhN_Assembly_Btn11", CommandTypesEnum.kShapeEditCmdType, AddInClientID, Nothing, Nothing)
            AddHandler assemblyBtn11.OnExecute, AddressOf Assembly.Buttons.Button11.OnExecute
            m_assemblyButtons.Add(assemblyBtn11)

            Dim assemblyBtn12 As ButtonDefinition = controlDefs.AddButtonDefinition("Assembly Action 12", "ThanhN_Assembly_Btn12", CommandTypesEnum.kShapeEditCmdType, AddInClientID, Nothing, Nothing)
            AddHandler assemblyBtn12.OnExecute, AddressOf Assembly.Buttons.Button12.OnExecute
            m_assemblyButtons.Add(assemblyBtn12)

            Dim assemblyBtn13 As ButtonDefinition = controlDefs.AddButtonDefinition("Assembly Action 13", "ThanhN_Assembly_Btn13", CommandTypesEnum.kShapeEditCmdType, AddInClientID, Nothing, Nothing)
            AddHandler assemblyBtn13.OnExecute, AddressOf Assembly.Buttons.Button13.OnExecute
            m_assemblyButtons.Add(assemblyBtn13)

            Dim assemblyBtn14 As ButtonDefinition = controlDefs.AddButtonDefinition("Assembly Action 14", "ThanhN_Assembly_Btn14", CommandTypesEnum.kShapeEditCmdType, AddInClientID, Nothing, Nothing)
            AddHandler assemblyBtn14.OnExecute, AddressOf Assembly.Buttons.Button14.OnExecute
            m_assemblyButtons.Add(assemblyBtn14)

            Dim assemblyBtn15 As ButtonDefinition = controlDefs.AddButtonDefinition("Assembly Action 15", "ThanhN_Assembly_Btn15", CommandTypesEnum.kShapeEditCmdType, AddInClientID, Nothing, Nothing)
            AddHandler assemblyBtn15.OnExecute, AddressOf Assembly.Buttons.Button15.OnExecute
            m_assemblyButtons.Add(assemblyBtn15)


            ' Create Drawing buttons explicitly for full customization
            Dim drawingBtn1 As ButtonDefinition = controlDefs.AddButtonDefinition("Drawing Action 1", "ThanhN_Drawing_Btn1", CommandTypesEnum.kShapeEditCmdType, AddInClientID, Nothing, Nothing)
            AddHandler drawingBtn1.OnExecute, Sub(Context As NameValueMap) DrawingButton_OnExecute(1, Context)
            m_drawingButtons.Add(drawingBtn1)
            Dim drawingBtn2 As ButtonDefinition = controlDefs.AddButtonDefinition("Drawing Action 2", "ThanhN_Drawing_Btn2", CommandTypesEnum.kShapeEditCmdType, AddInClientID, Nothing, Nothing)
            AddHandler drawingBtn2.OnExecute, Sub(Context As NameValueMap) DrawingButton_OnExecute(2, Context)
            m_drawingButtons.Add(drawingBtn2)
            Dim drawingBtn3 As ButtonDefinition = controlDefs.AddButtonDefinition("Drawing Action 3", "ThanhN_Drawing_Btn3", CommandTypesEnum.kShapeEditCmdType, AddInClientID, Nothing, Nothing)
            AddHandler drawingBtn3.OnExecute, Sub(Context As NameValueMap) DrawingButton_OnExecute(3, Context)
            m_drawingButtons.Add(drawingBtn3)
            Dim drawingBtn4 As ButtonDefinition = controlDefs.AddButtonDefinition("Drawing Action 4", "ThanhN_Drawing_Btn4", CommandTypesEnum.kShapeEditCmdType, AddInClientID, Nothing, Nothing)
            AddHandler drawingBtn4.OnExecute, Sub(Context As NameValueMap) DrawingButton_OnExecute(4, Context)
            m_drawingButtons.Add(drawingBtn4)
            Dim drawingBtn5 As ButtonDefinition = controlDefs.AddButtonDefinition("Drawing Action 5", "ThanhN_Drawing_Btn5", CommandTypesEnum.kShapeEditCmdType, AddInClientID, Nothing, Nothing)
            AddHandler drawingBtn5.OnExecute, Sub(Context As NameValueMap) DrawingButton_OnExecute(5, Context)
            m_drawingButtons.Add(drawingBtn5)
            Dim drawingBtn6 As ButtonDefinition = controlDefs.AddButtonDefinition("Drawing Action 6", "ThanhN_Drawing_Btn6", CommandTypesEnum.kShapeEditCmdType, AddInClientID, Nothing, Nothing)
            AddHandler drawingBtn6.OnExecute, Sub(Context As NameValueMap) DrawingButton_OnExecute(6, Context)
            m_drawingButtons.Add(drawingBtn6)
            Dim drawingBtn7 As ButtonDefinition = controlDefs.AddButtonDefinition("Drawing Action 7", "ThanhN_Drawing_Btn7", CommandTypesEnum.kShapeEditCmdType, AddInClientID, Nothing, Nothing)
            AddHandler drawingBtn7.OnExecute, Sub(Context As NameValueMap) DrawingButton_OnExecute(7, Context)
            m_drawingButtons.Add(drawingBtn7)
            Dim drawingBtn8 As ButtonDefinition = controlDefs.AddButtonDefinition("Drawing Action 8", "ThanhN_Drawing_Btn8", CommandTypesEnum.kShapeEditCmdType, AddInClientID, Nothing, Nothing)
            AddHandler drawingBtn8.OnExecute, Sub(Context As NameValueMap) DrawingButton_OnExecute(8, Context)
            m_drawingButtons.Add(drawingBtn8)
            Dim drawingBtn9 As ButtonDefinition = controlDefs.AddButtonDefinition("Drawing Action 9", "ThanhN_Drawing_Btn9", CommandTypesEnum.kShapeEditCmdType, AddInClientID, Nothing, Nothing)
            AddHandler drawingBtn9.OnExecute, Sub(Context As NameValueMap) DrawingButton_OnExecute(9, Context)
            m_drawingButtons.Add(drawingBtn9)
            Dim drawingBtn10 As ButtonDefinition = controlDefs.AddButtonDefinition("Drawing Action 10", "ThanhN_Drawing_Btn10", CommandTypesEnum.kShapeEditCmdType, AddInClientID, Nothing, Nothing)
            AddHandler drawingBtn10.OnExecute, Sub(Context As NameValueMap) DrawingButton_OnExecute(10, Context)
            m_drawingButtons.Add(drawingBtn10)
            Dim drawingBtn11 As ButtonDefinition = controlDefs.AddButtonDefinition("Drawing Action 11", "ThanhN_Drawing_Btn11", CommandTypesEnum.kShapeEditCmdType, AddInClientID, Nothing, Nothing)
            AddHandler drawingBtn11.OnExecute, Sub(Context As NameValueMap) DrawingButton_OnExecute(11, Context)
            m_drawingButtons.Add(drawingBtn11)
            Dim drawingBtn12 As ButtonDefinition = controlDefs.AddButtonDefinition("Drawing Action 12", "ThanhN_Drawing_Btn12", CommandTypesEnum.kShapeEditCmdType, AddInClientID, Nothing, Nothing)
            AddHandler drawingBtn12.OnExecute, Sub(Context As NameValueMap) DrawingButton_OnExecute(12, Context)
            m_drawingButtons.Add(drawingBtn12)
            Dim drawingBtn13 As ButtonDefinition = controlDefs.AddButtonDefinition("Drawing Action 13", "ThanhN_Drawing_Btn13", CommandTypesEnum.kShapeEditCmdType, AddInClientID, Nothing, Nothing)
            AddHandler drawingBtn13.OnExecute, Sub(Context As NameValueMap) DrawingButton_OnExecute(13, Context)
            m_drawingButtons.Add(drawingBtn13)
            Dim drawingBtn14 As ButtonDefinition = controlDefs.AddButtonDefinition("Drawing Action 14", "ThanhN_Drawing_Btn14", CommandTypesEnum.kShapeEditCmdType, AddInClientID, Nothing, Nothing)
            AddHandler drawingBtn14.OnExecute, Sub(Context As NameValueMap) DrawingButton_OnExecute(14, Context)
            m_drawingButtons.Add(drawingBtn14)
            Dim drawingBtn15 As ButtonDefinition = controlDefs.AddButtonDefinition("Drawing Action 15", "ThanhN_Drawing_Btn15", CommandTypesEnum.kShapeEditCmdType, AddInClientID, Nothing, Nothing)
            AddHandler drawingBtn15.OnExecute, Sub(Context As NameValueMap) DrawingButton_OnExecute(15, Context)
            m_drawingButtons.Add(drawingBtn15)

            ' Add to the user interface, if it's the first time.
            If firstTime Then
                AddToUserInterface()
            End If
        End Sub

        ' This method is called by Inventor when the AddIn is unloaded. The AddIn will be
        ' unloaded either manually by the user or when the Inventor session is terminated.
        Public Sub Deactivate() Implements Inventor.ApplicationAddInServer.Deactivate

            ' TODO:  Add ApplicationAddInServer.Deactivate implementation

            ' Release objects.
            m_uiEvents = Nothing
            g_inventorApplication = Nothing

            System.GC.Collect()
            System.GC.WaitForPendingFinalizers()
        End Sub

        ' This property is provided to allow the AddIn to expose an API of its own to other 
        ' programs. Typically, this  would be done by implementing the AddIn's API
        ' interface in a class and returning that class object through this property.
        Public ReadOnly Property Automation() As Object Implements Inventor.ApplicationAddInServer.Automation
            Get
                Return Nothing
            End Get
        End Property

        ' Note:this method is now obsolete, you should use the 
        ' ControlDefinition functionality for implementing commands.
        Public Sub ExecuteCommand(ByVal commandID As Integer) Implements Inventor.ApplicationAddInServer.ExecuteCommand
        End Sub

#End Region

#Region "User interface definition"
        ' Sub where the user-interface creation is done.  This is called when
        ' the add-in loaded and also if the user interface is reset.
        Private Sub AddToUserInterface()
            ' This is where you'll add code to add buttons to the ribbon.

            '** Create a separate custom tab (not inside the Tools tab) for Assembly and Drawing.

            ' Helper lambda to add a tab/panel/buttons to a ribbon.
            Dim AddTabPanelButtons = Sub(ribbonName As String, tabDisplayName As String, tabInternalName As String, panelDisplayName As String, panelInternalName As String, buttons As System.Collections.Generic.List(Of ButtonDefinition))
                                         Try
                                             Dim ribbonObj As Ribbon = g_inventorApplication.UserInterfaceManager.Ribbons.Item(ribbonName)

                                             Dim customTab As RibbonTab = Nothing
                                             Try
                                                 customTab = ribbonObj.RibbonTabs.Add(tabDisplayName, tabInternalName, AddInClientID)
                                             Catch
                                                 ' Tab likely exists - try to get it
                                                 Try
                                                     customTab = ribbonObj.RibbonTabs.Item(tabInternalName)
                                                 Catch
                                                     customTab = Nothing
                                                 End Try
                                             End Try

                                             If customTab IsNot Nothing Then
                                                 Dim customPanel As RibbonPanel = Nothing
                                                 Try
                                                     customPanel = customTab.RibbonPanels.Add(panelDisplayName, panelInternalName, AddInClientID)
                                                 Catch
                                                     Try
                                                         customPanel = customTab.RibbonPanels.Item(panelInternalName)
                                                     Catch
                                                         customPanel = Nothing
                                                     End Try
                                                 End Try

                                                 If customPanel IsNot Nothing Then
                                                     If buttons IsNot Nothing Then
                                                         For Each bd As ButtonDefinition In buttons
                                                             customPanel.CommandControls.AddButton(bd)
                                                         Next
                                                     End If
                                                 End If
                                             End If
                                         Catch
                                             ' Ignore failures for missing ribbons or other issues
                                         End Try
                                     End Sub

            ' Add to Assembly ribbon as a separate tab.
            AddTabPanelButtons("Assembly", "ThanhN", "ThanhN_AssemblyTab", "Main", "ThanhN_AssemblyPanel", m_assemblyButtons)

            ' Add to Drawing ribbon as a separate tab.
            AddTabPanelButtons("Drawing", "ThanhN", "ThanhN_DrawingTab", "Main", "ThanhN_DrawingPanel", m_drawingButtons)

            ' Add to Part ribbon as a separate tab.
            AddTabPanelButtons("Part", "ThanhN", "ThanhN_PartTab", "Main", "ThanhN_PartPanel", m_partButtons)
        End Sub

        Private Sub m_uiEvents_OnResetRibbonInterface(Context As NameValueMap) Handles m_uiEvents.OnResetRibbonInterface
            ' The ribbon was reset, so add back the add-ins user-interface.
            AddToUserInterface()
        End Sub

        ' Handlers for per-environment buttons (route to new modules)
        Private Sub PartButton_OnExecute(actionIndex As Integer, Context As NameValueMap)
            Try
                Select Case actionIndex
                    Case 1
                        Part.Buttons.Button1.OnExecute(Context)
                    Case 2
                        Part.Buttons.Button2.OnExecute(Context)
                    Case 3
                        Part.Buttons.Button3.OnExecute(Context)
                    Case 4
                        Part.Buttons.Button4.OnExecute(Context)
                    Case 5
                        Part.Buttons.Button5.OnExecute(Context)
                    Case 6
                        Part.Buttons.Button6.OnExecute(Context)
                    Case 7
                        Part.Buttons.Button7.OnExecute(Context)
                    Case 8
                        Part.Buttons.Button8.OnExecute(Context)
                    Case 9
                        Part.Buttons.Button9.OnExecute(Context)
                    Case 10
                        Part.Buttons.Button10.OnExecute(Context)
                    Case 11
                        Part.Buttons.Button11.OnExecute(Context)
                    Case 12
                        Part.Buttons.Button12.OnExecute(Context)
                    Case 13
                        Part.Buttons.Button13.OnExecute(Context)
                    Case 14
                        Part.Buttons.Button14.OnExecute(Context)
                    Case 15
                        Part.Buttons.Button15.OnExecute(Context)
                End Select
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End Sub

        Private Sub AssemblyButton_OnExecute(actionIndex As Integer, Context As NameValueMap)
            Try
                Select Case actionIndex
                    Case 1
                        Assembly.Buttons.Button1.OnExecute(Context)
                    Case 2
                        Assembly.Buttons.Button2.OnExecute(Context)
                    Case 3
                        Assembly.Buttons.Button3.OnExecute(Context)
                    Case 4
                        Assembly.Buttons.Button4.OnExecute(Context)
                    Case 5
                        Assembly.Buttons.Button5.OnExecute(Context)
                    Case 6
                        Assembly.Buttons.Button6.OnExecute(Context)
                    Case 7
                        Assembly.Buttons.Button7.OnExecute(Context)
                    Case 8
                        Assembly.Buttons.Button8.OnExecute(Context)
                    Case 9
                        Assembly.Buttons.Button9.OnExecute(Context)
                    Case 10
                        Assembly.Buttons.Button10.OnExecute(Context)
                    Case 11
                        Assembly.Buttons.Button11.OnExecute(Context)
                    Case 12
                        Assembly.Buttons.Button12.OnExecute(Context)
                    Case 13
                        Assembly.Buttons.Button13.OnExecute(Context)
                    Case 14
                        Assembly.Buttons.Button14.OnExecute(Context)
                    Case 15
                        Assembly.Buttons.Button15.OnExecute(Context)
                End Select
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End Sub

        Private Sub DrawingButton_OnExecute(actionIndex As Integer, Context As NameValueMap)
            Try
                Select Case actionIndex
                    Case 1
                        Drawing.Buttons.Button1.OnExecute(Context)
                    Case 2
                        Drawing.Buttons.Button2.OnExecute(Context)
                    Case 3
                        Drawing.Buttons.Button3.OnExecute(Context)
                    Case 4
                        Drawing.Buttons.Button4.OnExecute(Context)
                    Case 5
                        Drawing.Buttons.Button5.OnExecute(Context)
                    Case 6
                        Drawing.Buttons.Button6.OnExecute(Context)
                    Case 7
                        Drawing.Buttons.Button7.OnExecute(Context)
                    Case 8
                        Drawing.Buttons.Button8.OnExecute(Context)
                    Case 9
                        Drawing.Buttons.Button9.OnExecute(Context)
                    Case 10
                        Drawing.Buttons.Button10.OnExecute(Context)
                    Case 11
                        Drawing.Buttons.Button11.OnExecute(Context)
                    Case 12
                        Drawing.Buttons.Button12.OnExecute(Context)
                    Case 13
                        Drawing.Buttons.Button13.OnExecute(Context)
                    Case 14
                        Drawing.Buttons.Button14.OnExecute(Context)
                    Case 15
                        Drawing.Buttons.Button15.OnExecute(Context)
                End Select
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End Sub
#End Region

    End Class
End Namespace


Public Module Globals
    ' Inventor application object.
    Public g_inventorApplication As Inventor.Application

#Region "Function to get the add-in client ID."
    ' This function uses reflection to get the GuidAttribute associated with the add-in.
    Public Function AddInClientID() As String
        Dim guid As String = ""
        Try
            Dim t As Type = GetType(ThanhN.Basecode)
            Dim customAttributes() As Object = t.GetCustomAttributes(GetType(GuidAttribute), False)
            Dim guidAttribute As GuidAttribute = CType(customAttributes(0), GuidAttribute)
            guid = "{" + guidAttribute.Value.ToString() + "}"
        Catch
        End Try

        Return guid
    End Function
#End Region

#Region "hWnd Wrapper Class"
    ' This class is used to wrap a Win32 hWnd as a .Net IWind32Window class.
    ' This is primarily used for parenting a dialog to the Inventor window.
    '
    ' For example:
    ' myForm.Show(New WindowWrapper(g_inventorApplication.MainFrameHWND))
    '
    Public Class WindowWrapper
        Implements System.Windows.Forms.IWin32Window
        Public Sub New(ByVal handle As IntPtr)
            _hwnd = handle
        End Sub

        Public ReadOnly Property Handle() As IntPtr _
          Implements System.Windows.Forms.IWin32Window.Handle
            Get
                Return _hwnd
            End Get
        End Property

        Private _hwnd As IntPtr
    End Class
#End Region

#Region "Image Converter"
    ' Class used to convert bitmaps and icons from their .Net native types into
    ' an IPictureDisp object which is what the Inventor API requires. A typical
    ' usage is shown below where MyIcon is a bitmap or icon that's available
    ' as a resource of the project.
    '
    ' Dim smallIcon As stdole.IPictureDisp = PictureDispConverter.ToIPictureDisp(My.Resources.MyIcon)

    Public NotInheritable Class PictureDispConverter
        <DllImport("OleAut32.dll", EntryPoint:="OleCreatePictureIndirect", ExactSpelling:=True, PreserveSig:=False)> _
        Private Shared Function OleCreatePictureIndirect( _
            <MarshalAs(UnmanagedType.AsAny)> ByVal picdesc As Object, _
            ByRef iid As Guid, _
            <MarshalAs(UnmanagedType.Bool)> ByVal fOwn As Boolean) As stdole.IPictureDisp
        End Function

        Shared iPictureDispGuid As Guid = GetType(stdole.IPictureDisp).GUID

        Private NotInheritable Class PICTDESC
            Private Sub New()
            End Sub

            'Picture Types
            Public Const PICTYPE_BITMAP As Short = 1
            Public Const PICTYPE_ICON As Short = 3

            <StructLayout(LayoutKind.Sequential)> _
            Public Class Icon
                Friend cbSizeOfStruct As Integer = Marshal.SizeOf(GetType(PICTDESC.Icon))
                Friend picType As Integer = PICTDESC.PICTYPE_ICON
                Friend hicon As IntPtr = IntPtr.Zero
                Friend unused1 As Integer
                Friend unused2 As Integer

                Friend Sub New(ByVal icon As System.Drawing.Icon)
                    Me.hicon = icon.ToBitmap().GetHicon()
                End Sub
            End Class

            <StructLayout(LayoutKind.Sequential)> _
            Public Class Bitmap
                Friend cbSizeOfStruct As Integer = Marshal.SizeOf(GetType(PICTDESC.Bitmap))
                Friend picType As Integer = PICTDESC.PICTYPE_BITMAP
                Friend hbitmap As IntPtr = IntPtr.Zero
                Friend hpal As IntPtr = IntPtr.Zero
                Friend unused As Integer

                Friend Sub New(ByVal bitmap As System.Drawing.Bitmap)
                    Me.hbitmap = bitmap.GetHbitmap()
                End Sub
            End Class
        End Class

        Public Shared Function ToIPictureDisp(ByVal icon As System.Drawing.Icon) As stdole.IPictureDisp
            Dim pictIcon As New PICTDESC.Icon(icon)
            Return OleCreatePictureIndirect(pictIcon, iPictureDispGuid, True)
        End Function

        Public Shared Function ToIPictureDisp(ByVal bmp As System.Drawing.Bitmap) As stdole.IPictureDisp
            Dim pictBmp As New PICTDESC.Bitmap(bmp)
            Return OleCreatePictureIndirect(pictBmp, iPictureDispGuid, True)
        End Function
    End Class
#End Region

End Module
