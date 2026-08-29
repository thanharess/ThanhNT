Imports Inventor

Namespace ThanhN.Drawing.Buttons
    Public Module Draw_15
        Public Sub OnExecute(ByVal Context As NameValueMap)
            Try
                g_inventorApplication.UserInterfaceManager.UserInteractionManager.PostStatus("Executed Drawing Action 15")
            Catch ex As Exception
                Try
                    g_inventorApplication.UserInterfaceManager.UserInteractionManager.PostStatus("Error in Drawing Action 15: " & ex.Message)
                Catch
                End Try
            End Try
        End Sub
    End Module
End Namespace
