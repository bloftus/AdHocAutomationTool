Option Explicit On

Imports System.IO
Imports Microsoft.Office.Interop.Word
Imports AdHocAutomationTool.FileUtility
Imports Microsoft.Vbe.Interop

Public Class Macro

    Private Shared _allMacros As New Collection
    Private Shared _selectedMacros As New Collection

    Private _name As String
    Private _template As FileInfo

    Public Property Name As String
        Get
            Return _name
        End Get
        Set(value As String)
            _name = value
        End Set
    End Property

    Public Property Template As FileInfo
        Get
            Return _template
        End Get
        Set(value As FileInfo)
            _template = value
        End Set
    End Property

    Public Shared Property AllMacros As Collection
        Get
            Return _allMacros
        End Get
        Set(value As Collection)
            _allMacros = value
        End Set
    End Property

    Public Shared Property SelectedMacros As Collection
        Get
            Return _selectedMacros
        End Get
        Set(value As Collection)
            _selectedMacros = value
        End Set
    End Property

    Public Sub New(name As String, template As FileInfo)
        Me.Name = name
        Me.Template = template
    End Sub

    ''' <summary>
    ''' Gets the macros out of one file.
    ''' </summary>
    ''' <param name="template"></param>
    ''' <returns></returns>
    Shared Function AddTemplateAndMacros(template As FileInfo) As Collection
        Dim result As New Collection

        If template.Exists() Then
            FileUtility.Templates.Add(template)

            Dim word As New Microsoft.Office.Interop.Word.Application
            Dim doc As Document
            doc = word.Documents.Open(template.FullName)

            Dim myComponent As VBComponent
            For Each myComponent In doc.VBProject.VBComponents
                If myComponent.Type = vbext_ComponentType.vbext_ct_StdModule Then
                    Dim line As String = ""
                    Dim lineNumber As Integer = 1
                    Dim macroName As String = ""
                    Dim macro As Macro
                    Do While lineNumber < myComponent.CodeModule.CountOfLines
                        line = myComponent.CodeModule.Lines(lineNumber, 1)
                        If IsSubDeclaration(line) Then
                            macroName = myComponent.CodeModule.ProcOfLine(lineNumber, vbext_ProcKind.vbext_pk_Proc)
                            macro = New Macro(macroName, template)
                            result.Add(macro)
                            lineNumber += myComponent.CodeModule.ProcCountLines(macroName, vbext_ProcKind.vbext_pk_Proc)
                        End If
                        lineNumber += 1
                    Loop
                End If
            Next

            doc.Close(WdSaveOptions.wdDoNotSaveChanges)
            word.Quit()
        End If

        Return result
    End Function

    ''' <summary>
    ''' Looks for the word Sub to see if the line is a sub declaration.
    ''' </summary>
    ''' <param name="line">Line of text most likely from a code module</param>
    ''' <returns>True if Sub declaration</returns>
    Private Shared Function IsSubDeclaration(line As String) As Boolean
        Dim result As Boolean = False

        If line Like "*Sub *()" Then
            result = True
        End If

        Return result
    End Function

    ''' <summary>
    ''' Gets the macros out of one file.
    ''' </summary>
    ''' <param name="template"></param>
    Shared Sub RemoveTemplateAndMacros(template As FileInfo)

        Dim foundIndex As Integer = -1
        foundIndex = GetFileIndex(template, FileUtility.Templates)

        If foundIndex <> -1 Then
            FileUtility.Templates.Remove(foundIndex)
            Dim tempMacro As Macro
            For i = 1 To AllMacros.Count
                tempMacro = AllMacros.Item(i)
                If tempMacro.Template.FullName = template.FullName Then
                    AllMacros.Remove(i)
                Else
                    i += 1
                End If
            Next
            For i = 1 To SelectedMacros.Count
                tempMacro = SelectedMacros.Item(i)
                If tempMacro.Template.FullName = template.FullName Then
                    SelectedMacros.Remove(i)
                Else
                    i += 1
                End If
            Next
        End If

    End Sub

    ''' <summary>
    ''' Add a macro to the all macros collection.
    ''' </summary>
    ''' <param name="macroToAdd"></param>
    Public Shared Sub AddMacro(macroToAdd As Macro)
        AllMacros.Add(macroToAdd)
    End Sub

    ''' <summary>
    ''' Firsts finds the index of the macroToRemove and if found removes it.
    ''' </summary>
    ''' <param name="macroToRemove"></param>
    Public Shared Sub RemoveMacro(macroToRemove As Macro)
        Dim foundIndex As Integer = -1
        foundIndex = GetMacroIndex(macroToRemove, AllMacros)

        If foundIndex <> -1 Then
            AllMacros.Remove(foundIndex)
        End If
    End Sub

    ''' <summary>
    ''' Move a macro from the all macros collection to the selected macros collection.
    ''' </summary>
    ''' <param name="macroToSelect"></param>
    Public Shared Sub SelectMacro(macroToSelect As Macro)

        RemoveMacro(macroToSelect)
        SelectedMacros.Add(macroToSelect)

    End Sub

    ''' <summary>
    ''' Move a macro from the selected macros collection to the all macros collection.
    ''' </summary>
    ''' <param name="macroToDeSelect"></param>
    Public Shared Sub DeSelectMacro(macroToDeSelect As Macro)

        Dim foundIndex As Integer = -1
        foundIndex = GetMacroIndex(macroToDeSelect, SelectedMacros)
        If foundIndex <> -1 Then
            SelectedMacros.Remove(foundIndex)
        End If

        AddMacro(macroToDeSelect)

    End Sub

    ''' <summary>
    ''' Reorder the macros by moving the macro passed in up one position.
    ''' </summary>
    ''' <param name="macroToMove"></param>
    Public Shared Sub MoveMacroUp(macroToMove As Macro)
        Dim currentIndex As Integer
        currentIndex = GetMacroIndex(macroToMove, SelectedMacros)

        If currentIndex > 1 Then
            SelectedMacros.Remove(currentIndex)
            SelectedMacros.Add(macroToMove,, currentIndex - 1)
        End If
    End Sub

    ''' <summary>
    ''' Reorder the macros by moving the macro passed in down one position
    ''' </summary>
    ''' <param name="macroToMove"></param>
    Public Shared Sub MoveMacroDown(macroToMove As Macro)
        Dim currentIndex As Integer
        currentIndex = GetMacroIndex(macroToMove, SelectedMacros)

        If currentIndex > -1 Then
            SelectedMacros.Remove(currentIndex)
            SelectedMacros.Add(macroToMove,,, currentIndex - 1)
        End If
    End Sub

    ''' <summary>
    ''' Returns the index of the macro. If the macro isn't found in the collection -1 is returned
    ''' </summary>
    ''' <param name="macroToFind"></param>
    ''' <returns>The index or -1 if not found</returns>
    Private Shared Function GetMacroIndex(macroToFind As Macro, macros As Collection) As Integer

        For i = 1 To macros.Count
            Dim tempMacro As Macro = macros.Item(i)
            If tempMacro.Name = macroToFind.Name Then
                If tempMacro.Template.FullName = macroToFind.Template.FullName Then
                    Return i
                    Exit For
                End If
            End If
        Next

        Return -1

    End Function

End Class
