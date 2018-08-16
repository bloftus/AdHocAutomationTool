Option Explicit On

Imports System.IO
Imports AdHocAutomationTool.FileUtility
Imports AdHocAutomationTool.Macro
Imports Microsoft.Office.Interop.Word
Imports AdHocAutomationTool.Logging

Public Class Controller

    ''' <summary>
    ''' Moves a macro from all macros to selected macros.
    ''' </summary>
    ''' <param name="macroToAdd"></param>
    Public Shared Sub MoveMacroToSelected(macroToAdd As Macro)
        SelectMacro(macroToAdd)
    End Sub

    ''' <summary>
    ''' Moves a macro from selected macros back to all macros
    ''' </summary>
    ''' <param name="macroToRemove"></param>
    Public Shared Sub MoveMacroFromSelected(macroToRemove As Macro)
        DeSelectMacro(macroToRemove)
    End Sub

    ''' <summary>
    ''' When the user selects a folder the files in it are loaded into a collection.
    ''' </summary>
    ''' <param name="folderPath"></param>
    Public Shared Sub SelectFolder(folderPath As String)
        Dim files As Collection
        files = GetFilesInFolder(folderPath)

        Dim file As FileInfo

        For Each file In files
            Dim fileLoaded As Boolean
            fileLoaded = IsFileInCollection(file, SelectedFiles)

            If fileLoaded = False Then
                SelectedFiles.Add(file)
            End If
        Next

    End Sub

    ''' <summary>
    ''' When the user selects files those Files are added to a collection.
    ''' </summary>
    ''' <param name="filepaths"></param>
    Public Shared Sub SelectFiles(filepaths As Collection)
        Dim filepath As String
        Dim tempfile As FileInfo

        For Each filepath In filepaths
            If filepath Like "*.docm" Then
                tempfile = New FileInfo(filepath)
                Dim fileLoaded As Boolean
                fileLoaded = IsFileInCollection(tempfile, SelectedFiles)

                If fileLoaded = False Then
                    SelectedFiles.Add(tempfile)
                End If
            End If
        Next
    End Sub

    ''' <summary>
    ''' When the user selects a template file it's macros are loaded into a collection.
    ''' </summary>
    ''' <param name="template"></param>
    Public Shared Sub AddMacrosFromTemplate(template As FileInfo)
        Dim macros As New Collection
        Dim templates As Collection = FileUtility.Templates

        Dim templateLoaded As Boolean
        templateLoaded = IsFileInCollection(template, templates)

        If templateLoaded = False Then
            macros = AddTemplateAndMacros(template)

            Dim macro As Macro
            For Each macro In macros
                AllMacros.Add(macro)
            Next
        End If

    End Sub

    ''' <summary>
    ''' The user can also decide to remove a template and all of its associated macros will no longer show up.
    ''' </summary>
    ''' <param name="template"></param>
    Public Shared Sub RemoveTemplateAndMacros(template As FileInfo)
        Macro.RemoveTemplateAndMacros(template)
    End Sub

    ''' <summary>
    ''' The main function of the application. This will run the selected macros on the selected files.
    ''' </summary>
    Public Shared Sub RunSelectedMacrosOnFiles()
        Dim word As New Application
        Dim document As Document
        Dim file As FileInfo
        Dim originalTemplate As FileInfo
        Dim macro As Macro

        For Each file In SelectedFiles
            document = word.Documents.Open(file.FullName)

            originalTemplate = New FileInfo(document.AttachedTemplate.FullName)
            For Each macro In SelectedMacros
                document.AttachedTemplate = macro.Template.FullName
                word.Run(macro.Name)
            Next

            document.AttachedTemplate = originalTemplate.FullName
            document.Close(WdSaveOptions.wdSaveChanges)
        Next

        word.Quit()
        WriteLog()
    End Sub

    ''' <summary>
    ''' Presents a copy of the Selected Files collection
    ''' </summary>
    ''' <returns>A copy of the Selected Files collection</returns>
    Public Shared Function GetFiles() As Collection
        Return FileUtility.SelectedFiles
    End Function

    ''' <summary>
    ''' Presents a copy of the Templates collection
    ''' </summary>
    ''' <returns>A copy of the Templates collection</returns>
    Public Shared Function GetTemplates() As Collection
        Return FileUtility.Templates
    End Function

    ''' <summary>
    ''' Presents a copy of the All Macros Collection
    ''' </summary>
    ''' <returns>A cop of the All Macros Collection</returns>
    Public Shared Function GetAllMacros() As Collection
        Return Macro.AllMacros
    End Function

    ''' <summary>
    ''' Presents a copy of the Selected Macros Collection
    ''' </summary>
    ''' <returns>A copy of the Selected Macros Collection</returns>
    Public Shared Function GetSelectedMacros() As Collection
        Return Macro.SelectedMacros
    End Function

    ''' <summary>
    ''' Writes the Selected Files to a text file so they will be loaded the next time the user runs the 
    ''' application.
    ''' </summary>
    Public Shared Sub WriteSelectedFiles()
        Dim outputTextFile As String
        outputTextFile = My.Computer.FileSystem.SpecialDirectories.CurrentUserApplicationData + "\SelectedFiles.txt"
        WriteListOfFiles(SelectedFiles, outputTextFile)
    End Sub

    ''' <summary>
    ''' If there is a text file containing the files from a previous session they will be loaded so the user
    ''' can continue where they left off.
    ''' </summary>
    Public Shared Sub ImportSelectedFiles()
        Dim inputTextFile As String
        inputTextFile = My.Computer.FileSystem.SpecialDirectories.CurrentUserApplicationData + "\SelectedFiles.txt"

        Dim files As Collection
        files = ImportListOfFiles(inputTextFile)
        SelectedFiles = files
    End Sub

    ''' <summary>
    ''' Writes the templates to a text file so they can be loaded the next time the user runs the application.
    ''' </summary>
    Public Shared Sub WriteTemplates()
        Dim outputTextFile As String
        outputTextFile = My.Computer.FileSystem.SpecialDirectories.CurrentUserApplicationData + "\Templates.txt"
        WriteListOfFiles(FileUtility.Templates, outputTextFile)
    End Sub

    ''' <summary>
    ''' If there is a text files containin the templates from a previous session they will be loaded so he user
    ''' can continue where they left off.
    ''' </summary>
    Public Shared Sub ImportTemplates()
        Dim inputTextFile As String
        inputTextFile = My.Computer.FileSystem.SpecialDirectories.CurrentUserApplicationData + "\Templates.txt"

        Dim templates As Collection
        templates = ImportListOfFiles(inputTextFile)

        Dim template As FileInfo
        For Each template In templates
            Dim macros As Collection
            macros = AddTemplateAndMacros(template)
            Dim macro As Macro
            For Each macro In macros
                AllMacros.Add(macro)
            Next
        Next
    End Sub

    ''' <summary>
    ''' Allows the user to remove any files that they do not want to run macros on.
    ''' </summary>
    ''' <param name="filepaths"></param>
    Public Shared Sub RemoveFiles(filepaths As Collection)
        Dim filepath As String

        For Each filepath In filepaths
            For i = SelectedFiles.Count To 1 Step -1
                If SelectedFiles.Item(i).Fullname = filepath Then
                    SelectedFiles.Remove(i)
                End If
            Next
        Next
    End Sub

    ''' <summary>
    ''' Allows the user to rearrange their selected macros by moving one up in the list.
    ''' </summary>
    ''' <param name="macroToMove"></param>
    Public Shared Sub MoveMacroHigherInList(macroToMove As Macro)
        MoveMacroUp(macroToMove)
    End Sub

    ''' <summary>
    ''' Allows the user to rearrange their selected macros by moving one down in the list.
    ''' </summary>
    ''' <param name="macroToMove"></param>
    Public Shared Sub MoveMacroLowerInList(macroToMove As Macro)
        MoveMacroDown(macroToMove)
    End Sub
End Class
