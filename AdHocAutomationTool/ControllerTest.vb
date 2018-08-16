Imports System.Text
Imports AdHocAutomationTool.Controller
Imports AdHocAutomationTool.Macro
Imports AdHocAutomationTool.FileUtility
Imports NUnit.Framework
Imports System.IO
Imports Microsoft.Office.Interop.Word

<TestFixture()> Public Class ControllerTest

    <Test()> Public Sub RunSelectedOneMacroOnFilesTest()
        ' All other controller subs and functions seemed like they will just be calls to other methods that are already tested.

        Dim template As New FileInfo("C:\Users\Brian\Google Drive\HCC\Summer 2018\Capstone\Test\TemplateWithBoldMacro.dotm")
        AddMacrosFromTemplate(template)

        Dim macro As New Macro("BoldTheFile", template)
        SelectMacro(macro)

        Dim file As New FileInfo("C:\Users\Brian\Google Drive\HCC\Summer 2018\Capstone\Test\File.docm")
        SelectedFiles.Add(file)

        RunSelectedMacrosOnFiles()

        Dim word As New Application
        Dim doc As Document

        doc = word.Documents.Open("C:\Users\Brian\Google Drive\HCC\Summer 2018\Capstone\Test\File.docm")

        Assert.IsTrue(CBool(doc.Range.Bold))

        doc.Close(WdSaveOptions.wdDoNotSaveChanges)
        word.Quit()
    End Sub

    <Test()> Public Sub RunSelectedTwoMacrosOnFilesTest()
        ' All other controller subs and functions seemed like they will just be calls to other methods that are already tested.

        Dim boldTemplate As New FileInfo("C:\Users\Brian\Google Drive\HCC\Summer 2018\Capstone\Test\TemplateWithBoldMacro.dotm")
        AddMacrosFromTemplate(boldTemplate)

        Dim italicTemplate As New FileInfo("C:\Users\Brian\Google Drive\HCC\Summer 2018\Capstone\Test\TemplateWithItalicMacro.dotm")
        AddMacrosFromTemplate(italicTemplate)

        Dim boldMacro As New Macro("BoldTheFile", boldTemplate)
        SelectMacro(boldMacro)

        Dim italicMacro As New Macro("ItalicTheFile", italicTemplate)
        SelectMacro(italicMacro)

        Dim file As New FileInfo("C:\Users\Brian\Google Drive\HCC\Summer 2018\Capstone\Test\File.docm")
        SelectedFiles.Add(file)

        RunSelectedMacrosOnFiles()

        Dim word As New Application
        Dim doc As Document

        doc = word.Documents.Open("C:\Users\Brian\Google Drive\HCC\Summer 2018\Capstone\Test\File.docm")

        Assert.IsTrue(CBool(doc.Range.Bold))

        doc.Close(WdSaveOptions.wdDoNotSaveChanges)
        word.Quit()
    End Sub
End Class