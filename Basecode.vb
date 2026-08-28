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

        Private WithEvents m_uievents As UserInterfaceEvents
        Private m_partButtons As New System.Collections.Generic.List(Of ButtonDefinition)
        Private m_assemblyButtons As New System.Collections.Generic.List(Of ButtonDefinition)
        Private m_drawingButtons As New System.Collections.Generic.List(Of ButtonDefinition)
        Private m_assembly2Buttons As New System.Collections.Generic.List(Of ButtonDefinition)

#Region "ApplicationAddInServer Members"

        ' This method is called by Inventor when it loads the AddIn. The AddInSiteObject provides access  
        ' to the Inventor Application object. The FirstTime flag indicates if the AddIn is loaded for
        ' the first time. However, with the introduction of the ribbon this argument is always true.
        Public Sub Activate(ByVal addInSiteObject As Inventor.ApplicationAddInSite, ByVal firstTime As Boolean) Implements Inventor.ApplicationAddInServer.Activate
            ' Initialize AddIn members.
            g_inventorApplication = addInSiteObject.Application

            ' Connect to the user-interface events to handle a ribbon reset.
            M_uievents = g_inventorApplication.UserInterfaceManager.UserInterfaceEvents

            ' Create button definitions for Part, Assembly and Drawing: 15 buttons each.
            Dim controlDefs As Inventor.ControlDefinitions = g_inventorApplication.CommandManager.ControlDefinitions

            ' No icons: use text-only buttons to avoid bitmap/COM issues
            Dim largeIcon As stdole.IPictureDisp = Nothing
            Dim smallIcon As stdole.IPictureDisp = Nothing
            ' Create Part buttons via helper class
            PartButtons.Register(controlDefs, AddInClientID, m_partButtons)

            ' Create Assembly buttons via helper class
            AssemblyButtons.Register(controlDefs, AddInClientID, m_assemblyButtons)


            ' Create Drawing buttons via helper class
            DrawingButtons.Register(controlDefs, AddInClientID, m_drawingButtons)

            ' Create Assembly buttons via helper class
            Assembly2Buttons.Register(controlDefs, AddInClientID, m_assembly2Buttons)


            ' Ensure the user interface is created when the add-in activates.
            ' Call AddToUserInterface unconditionally so the Assembly tab/panel is created
            ' even when firstTime is False (the ribbon can be reset at other times).
            AddToUserInterface()
        End Sub

        ' This method is called by Inventor when the AddIn is unloaded. The AddIn will be
        ' unloaded either manually by the user or when the Inventor session is terminated.
        Public Sub Deactivate() Implements Inventor.ApplicationAddInServer.Deactivate

            ' TODO:  Add ApplicationAddInServer.Deactivate implementation

            ' Release objects.
            m_uievents = Nothing
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

            ' Add to Assembly ribbon as a separate tab.
            AddTabPanelButtons("Assembly", "BOM ADDIN", "ThanhN_AssemblyTab2", "Main2", "ThanhN_AssemblyPanel2", m_assembly2Buttons)
        End Sub

        Private Sub muievents_onresetribbonInterface(Context As NameValueMap) Handles m_uievents.OnResetRibbonInterface
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
                        Part.Buttons.Button1.OnExecute(Context)
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

        ' Button dispatcher methods removed: handlers wired directly to specific Button classes

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
        <DllImport("OleAut32.dll", EntryPoint:="OleCreatePictureIndirect", ExactSpelling:=True, PreserveSig:=False)>
        Private Shared Function OleCreatePictureIndirect(
            <MarshalAs(UnmanagedType.AsAny)> ByVal picdesc As Object,
            ByRef iid As Guid,
            <MarshalAs(UnmanagedType.Bool)> ByVal fOwn As Boolean) As stdole.IPictureDisp
        End Function

        Shared iPictureDispGuid As Guid = GetType(stdole.IPictureDisp).GUID

        Private NotInheritable Class PICTDESC
            Private Sub New()
            End Sub

            'Picture Types
            Public Const PICTYPE_BITMAP As Short = 1
            Public Const PICTYPE_ICON As Short = 3

            <StructLayout(LayoutKind.Sequential)>
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

            <StructLayout(LayoutKind.Sequential)>
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
