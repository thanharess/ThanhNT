Imports Inventor

Namespace ThanhN.Drawing.Buttons
    Public Module Draw_13
        Public Sub OnExecute(ByVal Context As NameValueMap)
            Try
                g_inventorApplication.UserInterfaceManager.UserInteractionManager.PostStatus("Executed Drawing Action 13")
            Catch ex As Exception
                Try
                    g_inventorApplication.UserInterfaceManager.UserInteractionManager.PostStatus("Error in Drawing Action 13: " & ex.Message)
                Catch
                End Try
            End Try
        End Sub
    End Module
End Namespace
