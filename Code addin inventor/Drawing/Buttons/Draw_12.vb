Imports Inventor

Namespace ThanhN.Drawing.Buttons
    Public Module Draw_12
        Public Sub OnExecute(ByVal Context As NameValueMap)
            Try
                g_inventorApplication.UserInterfaceManager.UserInteractionManager.PostStatus("Executed Drawing Action 12")
            Catch ex As Exception
                Try
                    g_inventorApplication.UserInterfaceManager.UserInteractionManager.PostStatus("Error in Drawing Action 12: " & ex.Message)
                Catch
                End Try
            End Try
        End Sub
    End Module
End Namespace
