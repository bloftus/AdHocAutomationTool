Imports System.IO
Imports NUnit.Framework
Imports AdHocAutomationTool.FileUtility

<TestFixture()> Public Class FileUtilityTest

    <Test()> Public Sub GetFilesInFolderTest()
        Dim files As Collection

        Dim invalidPath As String = "C:\Users\Brian\Google Drive\HCC\Summer 2018\Capstone\Test\Invalid"
        files = GetFilesInFolder(invalidPath)
        Assert.AreEqual(0, files.Count)

        files.Clear()

        Dim filePath As String = "C:\Users\Brian\Google Drive\HCC\Summer 2018\Capstone\Test\File.docm"
        files = GetFilesInFolder(filePath)
        Assert.AreEqual(0, files.Count)

        files.Clear()

        Dim folderPathOneFile As String = "C:\Users\Brian\Google Drive\HCC\Summer 2018\Capstone\Test\Has One File"
        files = GetFilesInFolder(folderPathOneFile)
        Assert.AreEqual(1, files.Count)

        files.Clear()

        Dim folderPathNestedFiles As String = "C:\Users\Brian\Google Drive\HCC\Summer 2018\Capstone\Test\Has Nested Files"
        files = GetFilesInFolder(folderPathNestedFiles)
        Assert.AreEqual(2, files.Count)

        files.Clear()

    End Sub

    <Test()> Public Sub WriteListOfFilesTest()

        Dim files As New Collection
        Dim testFile As New FileInfo("C:\Users\Brian\Google Drive\HCC\Summer 2018\Capstone\Test\File.docm")
        files.Add(testFile)

        WriteListOfFiles(files, Application.LocalUserAppDataPath + "\files.txt")

        Assert.That(File.Exists(Application.LocalUserAppDataPath + "\files.txt"))

        Dim fileContent As String
        fileContent = My.Computer.FileSystem.ReadAllText(Application.LocalUserAppDataPath + "\files.txt")

        Assert.AreEqual("C:\Users\Brian\Google Drive\HCC\Summer 2018\Capstone\Test\File.docm" + vbCrLf, fileContent)

    End Sub

    <Test()> Public Sub ImportListOfFilesTest()
        'invalid path
        Dim invalidFilePath As String = "C:\Users\Brian\Google Drive\HCC\Summer 2018\Capstone\Test\invalid.txt"
        Dim files As Collection

        files = ImportListOfFiles(invalidFilePath)

        Assert.AreEqual(0, files.Count)

        files.Clear()

        ' text file that contains the path to one file.
        Dim textFilePath As String = "C:\Users\Brian\Google Drive\HCC\Summer 2018\Capstone\Test\files.txt"
        ' textFilePath = Replace(textFilePath, "\", "\\")
        files = ImportListOfFiles(textFilePath)

        Assert.AreEqual(1, files.Count)

        files.Clear()

    End Sub

End Class