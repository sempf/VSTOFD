'Gives us access to the Presentation
Imports Microsoft.Office.Interop.PowerPoint
'This is where the Global objects are
Imports Microsoft.Office.Core
'We need this for the Dictionary of images
Imports System.Collections.Generic

Public Class ImagePicker


    Private Shared ImageLibrary As New Dictionary(Of String, String)


    Private Sub ImagePicker_Startup(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Startup
        LoadImageLibrary()
    End Sub

    Private Sub ImagePicker_Shutdown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shutdown
        ImageLibrary = Nothing
    End Sub
    Public Shared Sub addSlideFour()
        Dim active As DocumentWindow = Globals.ImagePicker.Application.ActiveWindow
        If active.Selection.Type = PpSelectionType.ppSelectionText Then
            active.Selection.TextRange.Font.Bold = MsoTriState.msoTrue
        End If

    End Sub
    Private Shared Sub LoadImageLibrary()
        'Go find the images directory and get a list of names
        Dim imageDir As New System.IO.DirectoryInfo("C:\ProductImages")
        'Load them into a class level generic list
        For Each imageFile As System.IO.FileInfo In imageDir.GetFiles
            ImageLibrary.Add(imageFile.Name.Substring(0, imageFile.Name.IndexOf(".")), imageFile.FullName)
        Next
    End Sub

    Public Shared Sub PickImages()
        Dim active As Presentation = Globals.ImagePicker.Application.ActivePresentation
        'For each slide in the presentation
        For Each slideToCheck As Slide In active.Slides
            'Check the title for the product name
            If ImageLibrary.ContainsKey(slideToCheck.Shapes.Title.TextFrame.TextRange.Text) Then
                'If we get it, drop a pictre of hte image on the page
                slideToCheck.Shapes.AddPicture(ImageLibrary.Item( _
                                               slideToCheck.Shapes.Title.TextFrame.TextRange.Text), _
                                               MsoTriState.msoFalse, MsoTriState.msoTrue, 20, 100)
            End If
        Next
    End Sub

End Class
Public Class myInspector
    Implements IDocumentInspector

    Public Sub Fix(ByVal Doc As Object, ByVal Hwnd As Integer, ByRef Status As Microsoft.Office.Core.MsoDocInspectorStatus, ByRef Result As String) Implements Microsoft.Office.Core.IDocumentInspector.Fix

    End Sub

    Public Sub GetInfo(ByRef Name As String, ByRef Desc As String) Implements Microsoft.Office.Core.IDocumentInspector.GetInfo

    End Sub

    Public Sub Inspect(ByVal Doc As Object, ByRef Status As Microsoft.Office.Core.MsoDocInspectorStatus, ByRef Result As String, ByRef Action As String) Implements Microsoft.Office.Core.IDocumentInspector.Inspect

    End Sub
End Class