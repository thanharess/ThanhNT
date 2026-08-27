Imports Inventor

Namespace ThanhN.Part.Buttons
    Public Module Button7
        Public Sub OnExecute(ByVal Context As NameValueMap)
            Try
                g_inventorApplication.UserInterfaceManager.UserInteractionManager.PostStatus("Executed Part Action 7")
            Catch ex As Exception
                Try
                    g_inventorApplication.UserInterfaceManager.UserInteractionManager.PostStatus("Error in Part Action 7: " & ex.Message)
                Catch
                End Try
            End Try
        End Sub
    End Module
End Namespace
