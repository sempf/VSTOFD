Imports 
Public Class ThisAddIn

    Private Sub ThisAddIn_Startup(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Startup
        Dim currentProject As MSProject.Project = Application.ActiveProject
        For Each currentTask As MSProject.Task In currentProject.Tasks
            If currentTask.LateStart Then
                'just get the latest one
                Dim currentResource As MSProject.Resource = currentTask.Resources(0)
                Dim theirEmail As String = currentResource.EMailAddress
                Dim outgoingMail As New System.Net.Mail.SmtpClient
                outgoingMail.Send("bill@pointweb.net", _
                              theirEmail, _
                              "This task is late", _
                              "This taks is in the collection of late tasks.  Do something!")
            End If
        Next
    End Sub

    Private Sub ThisAddIn_Shutdown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shutdown

    End Sub


End Class
