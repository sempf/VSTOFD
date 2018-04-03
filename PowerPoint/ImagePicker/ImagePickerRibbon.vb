Imports Microsoft.Office.Tools.Ribbon
Imports ImagePicker.ImagePicker
'Gives us access to the Presentation
Imports Microsoft.Office.Interop.PowerPoint
'This is where the Global objects are
Imports Microsoft.Office.Core

Public Class ImagePickerRibbon

    Private Sub ImagePickerRibbon_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles Button1.Click
        ImagePicker.PickImages()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles Button2.Click
        Dim active As Presentation = Globals.ImagePicker.Application.ActivePresentation
        Dim custom As CustomLayout = active.SlideMaster.CustomLayouts.Item(PpSlideLayout.ppLayoutClipartAndText)
        active.Slides.AddSlide(4, custom)
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles Button3.Click
        addSlideFour()
    End Sub
End Class
