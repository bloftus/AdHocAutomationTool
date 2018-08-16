Imports AdHocAutomationTool.Macro
Imports NUnit.Framework
Imports System.IO

<TestFixture()> Public Class MacroTest

    <Test()> Public Sub NewTest()

        Dim tempTemplate As New FileInfo("C:\Users\Brian\Google Drive\HCC\Summer 2018\Capstone\Test\TemplateWithBoldMacro.dotm")
        Dim newMacro As New Macro("Test", tempTemplate)

        Assert.NotNull(newMacro)

    End Sub

    <Test()> Public Sub AddTemplateAndMacrosTest()

        Dim tempTemplate As New FileInfo("C:\Users\Brian\Google Drive\HCC\Summer 2018\Capstone\Test\TemplateWithBoldMacro.dotm")
        Dim macros As Collection

        macros = AddTemplateAndMacros(tempTemplate)

        Assert.AreEqual(1, macros.Count)

        macros.Clear()

        Dim invalidTemplate As New FileInfo("C:\Users\Brian\Google Drive\HCC\Summer 2018\Capstone\Test\InvalidTemplate.dotm")

        macros = AddTemplateAndMacros(invalidTemplate)

        Assert.AreEqual(0, macros.Count)

        macros.Clear()

        Dim noMacrosTemplate As New FileInfo("C:\Users\Brian\Google Drive\HCC\Summer 2018\Capstone\Test\NoMacros.dotm")

        macros = AddTemplateAndMacros(noMacrosTemplate)

        Assert.AreEqual(0, macros.Count)

    End Sub

    <Test()> Public Sub RemoveTemplateAndMacrosTest()

        Dim template As New FileInfo("C:\Users\Brian\Google Drive\HCC\Summer 2018\Capstone\Test\TemplateWithBoldMacro.dotm")
        Dim testMacro As New Macro("HereIsAMacro", template)
        AddMacro(testMacro)

        RemoveTemplateAndMacros(template)

        Dim macros As Collection
        macros = AllMacros

        Assert.AreEqual(0, macros.Count)

        Dim notFoundTemplate As New FileInfo("C:\Users\Brian\Google Drive\HCC\Summer 2018\Capstone\Test\NotFoundTemplate.dotm")
        AddMacro(testMacro)

        RemoveTemplateAndMacros(notFoundTemplate)

        macros = AllMacros

        Assert.AreEqual(1, macros.Count)

    End Sub

End Class