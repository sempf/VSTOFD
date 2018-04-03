Public Class ThisAddIn

    Private Sub ThisAddIn_Startup(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Startup
        For Each currentFilter As MSProject.Filter In Application.ActiveProject.TaskFilters
            currentFilter.Delete()
        Next
        'Apply your formatting here!
        Application.ActiveProject.TaskFilters("Critical").Apply()



    End Sub

    Private Sub ThisAddIn_Shutdown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shutdown

    End Sub

    Function WorkingDays(ByVal selectedResource As MSProject.Resource, ByVal year As Integer) As Integer
        Dim selectedYear As MSProject.Year = selectedResource.Calendar.Years(year)
        Dim result As Integer = 0
        For Each countMonth As MSProject.Month In selectedYear.Months
            For Each countDay As MSProject.Day In countMonth.Days
                If countDay.Working Then
                    result = result + 1
                End If
            Next
        Next
    End Function

    Function CalcOvertime(ByVal selectedResource As MSProject.Resource) As Integer
        Dim result As Double = 0.0
        For Each currentAssignment As MSProject.Assignment In selectedResource.Assignments
            result = result + CDbl(currentAssignment.OvertimeCost)
        Next
    End Function
End Class
