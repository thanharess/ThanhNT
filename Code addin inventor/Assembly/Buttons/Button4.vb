Imports Inventor

Namespace ThanhN.Assembly.Buttons
    Public Module Button4
        Public Sub OnExecute(ByVal Context As NameValueMap)
            Try
                g_inventorApplication.UserInterfaceManager.UserInteractionManager.PostStatus("Executed Assembly Action 4")
            Catch ex As Exception
                Try
                    g_inventorApplication.UserInterfaceManager.UserInteractionManager.PostStatus("Error in Assembly Action 4: " & ex.Message)
                Catch
                End Try
            End Try
        End Sub
    End Module
End Namespace
