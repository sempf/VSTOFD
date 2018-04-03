Public Class ThisAddIn

    Private Sub ThisAddIn_Startup(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Startup
        Dim currentDocument As Visio.Document
        currentDocument = Application.ActiveDocument

        Dim currentPage As Visio.Page
        currentPage = Application.ActivePage

        Dim circleShape As Visio.Shape
        circleShape = Application.ActivePage.DrawOval(100, 100, 200, 200)

        'Get an object to hold the shape
        Dim neatShape As Visio.Shape
        'Instead of using a Draw method that will give us a shape
        'we go get a copy of it.
        neatShape = currentDocument.Masters.ItemU("PC")
        'USe the Drop method to place the shape.
        currentPage.Drop(neatShape, 4, 4)

        Dim squareShape As Visio.Shape
        squareShape = Application.ActivePage.DrawRectangle(600, 100, 620, 300)
        With squareShape
            .TextStyle = "Basic"
            .LineStyle = "TextOnly"
            .FillStyle = "TextOnly"
        End With

        Dim titleChars As Visio.Characters
        titleChars = squareShape.Characters

        With titleChars
            'set the text
            .Text = "This is the title"
            'Set the font to 18
            .CharProps(CShort(Visio.VisCellIndices.visCharacterSize)) = CShort(18)
        End With

        Dim labelChars As Visio.Characters
        labelChars = neatShape.Characters
        'Just for fun, we will use this computers name.
        labelChars.Text = My.Computer.Name

        'currentDocument.DataRecordsets.Add("", "", 0)
        'currentDocument.SaveAs("A new name.vsd")

        Dim theSaveAsWeb As Visio.

    End Sub

    Private Sub ThisAddIn_Shutdown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shutdown

    End Sub

End Class

