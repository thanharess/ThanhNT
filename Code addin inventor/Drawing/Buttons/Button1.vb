Imports Inventor

Namespace ThanhN.Drawing.Buttons
    Public Module Button1
        Public Sub OnExecute(ByVal Context As NameValueMap)
            Try
                ' Action1 for Drawing: ask user for tolerance value, then select drawing dimensions and set symmetric tolerance
                Dim input As String = Microsoft.VisualBasic.Interaction.InputBox("Enter symmetric tolerance value (e.g. 0.05)", "Tolerance value", "0.05")
                If String.IsNullOrWhiteSpace(input) Then
                    g_inventorApplication.UserInterfaceManager.UserInteractionManager.PostStatus("Drawing Action 1 cancelled by user")
                    Exit Sub
                End If

                Dim tol As Double
                If Not Double.TryParse(input, tol) Then
                    g_inventorApplication.UserInterfaceManager.UserInteractionManager.PostStatus("Invalid tolerance value: " & input)
                    Exit Sub
                End If

                Dim oMsg As String = "Select drawing dimensions to apply default tolerance (Press Esc to continue)"
                While True
                    ' SelectionFilterEnum does not have a drawing-dimension filter in all Inventor versions.
                    ' Use no filter and validate the picked object at runtime.
                    Dim picked As Object = Nothing
                    picked = g_inventorApplication.CommandManager.Pick(Nothing, oMsg)
                    If picked Is Nothing Then Exit While

                    Dim oDimension As DimensionConstraint = TryCast(picked, DimensionConstraint)
                    If oDimension Is Nothing Then
                        ' Not a dimension; inform user and continue selection
                        g_inventorApplication.UserInterfaceManager.UserInteractionManager.PostStatus("Selected object is not a dimension. Please select a drawing dimension.")
                        Continue While
                    End If

                    oDimension.Parameter.Tolerance.SetToSymmetric(tol)
                End While

                g_inventorApplication.UserInterfaceManager.UserInteractionManager.PostStatus("Drawing Action 1 completed")
            Catch ex As Exception
                Try
                    g_inventorApplication.UserInterfaceManager.UserInteractionManager.PostStatus("Error in Drawing Action 1: " & ex.Message)
                Catch
                End Try
            End Try
        End Sub
    End Module
End Namespace
