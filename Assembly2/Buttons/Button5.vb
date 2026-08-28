Imports Inventor
Imports System.Windows.Forms
Imports Microsoft.VisualBasic
Imports System.Collections.Generic

Namespace ThanhN.Assembly2.Buttons

    Public Module Button5
        Public Sub OnExecute(ByVal Context As NameValueMap)
            Try
            Catch ex As Exception
                Try
                    g_inventorApplication.UserInterfaceManager.UserInteractionManager.PostStatus("Error in Part Action 5: " & ex.Message)
                Catch
                End Try
            End Try
        End Sub
    End Module
End Namespace
