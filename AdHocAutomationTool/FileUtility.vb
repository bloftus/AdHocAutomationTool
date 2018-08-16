Option Explicit On

Imports System.IO

Public Class FileUtility

    Private Shared _selectedFiles As New Collection
    Private Shared _templates As New Collection

    Public Shared Property SelectedFiles As Collection
        Get
            Return _selectedFiles
        End Get
        Set(value As Collection)
            _selectedFiles = value
        End Set
    End Property

    Public Shared Property Templates As Collection
        Get
            Return _templates
        End Get
        Set(value As Collection)
            _templates = value
        End Set
    End Property

    ''' <summary>
    ''' If the user selects a folder this function will return the files in it.
    ''' </summary>
    ''' <returns>A collection of FileInfo objects</returns>
    Public Shared Function GetFilesInFolder(folderPath As String) As Collection
        Dim result As New Collection

        If Directory.Exists(folderPath) Then
            Dim filepaths As String()
            filepaths = Directory.GetFiles(folderPath, "*.docm", SearchOption.AllDirectories)

            Dim filepath As String
            Dim file As FileInfo
            For Each filepath In filepaths

                file = New FileInfo(filepath)
                result.Add(file)

            Next
        End If

        Return result
    End Function

    ''' <summary>
    ''' Writes the filepaths into a text file so they can be loaded in a later session.
    ''' </summary>
    Public Shared Sub WriteListOfFiles(files As Collection, outputFile As String)

        Dim textFile As StreamWriter
        textFile = My.Computer.FileSystem.OpenTextFileWriter(outputFile, False)

        Dim myFile As FileInfo
        For Each myFile In files
            textFile.WriteLine(myFile.FullName)
        Next

        textFile.Close()

    End Sub

    ''' <summary>
    ''' Reads a text file of filepaths and makes fileinfo objects.
    ''' </summary>
    ''' <returns>A collection of fileinfo objects</returns>
    Public Shared Function ImportListOfFiles(txtFilePath As String) As Collection
        Dim result As New Collection

        If File.Exists(txtFilePath) Then
            Dim textFile As StreamReader
            textFile = My.Computer.FileSystem.OpenTextFileReader(txtFilePath)

            Dim line As String
            Dim myFile As FileInfo
            Do While textFile.Peek <> -1
                line = textFile.ReadLine
                If File.Exists(line) Then
                    myFile = New FileInfo(line)
                    result.Add(myFile)
                End If
            Loop
            textFile.Close()
        End If

        Return result
    End Function

    ''' <summary>
    ''' Returns the index of a file in a collection or -1 if not found
    ''' </summary>
    ''' <param name="fileToFind"></param>
    ''' <param name="files"></param>
    ''' <returns>index of file or -1 if not found</returns>
    Public Shared Function GetFileIndex(fileToFind As FileInfo, files As Collection) As Integer

        For i = 1 To files.Count
            Dim tempFile As FileInfo = files.Item(i)
            If tempFile.FullName = fileToFind.FullName Then
                Return i
                Exit For
            End If
        Next

        Return -1
    End Function

    ''' <summary>
    ''' Wraps GetFileIndex to use it to check if a File is in the collection or not.
    ''' Makes the code more clear when a boolean can be used.
    ''' </summary>
    ''' <param name="fileToLoad"></param>
    ''' <param name="files"></param>
    ''' <returns>True if file is already found, false if not found</returns>
    Public Shared Function IsFileInCollection(fileToLoad As FileInfo, files As Collection) As Boolean
        Dim result As Boolean = True

        Dim foundIndex As Integer
        foundIndex = GetFileIndex(fileToLoad, files)

        If foundIndex = -1 Then
            result = False
        End If

        Return result
    End Function

End Class
