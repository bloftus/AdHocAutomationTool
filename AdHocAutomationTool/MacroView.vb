Imports AdHocAutomationTool.Controller
Imports System.IO

Public Class MacroView

    ''' <summary>
    ''' Allows the user to select a single template.
    ''' It also refreshes the listboxes on the screen.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub btnAddTemplate_Click(sender As Object, e As EventArgs) Handles btnAddTemplate.Click
        Dim fileSelector As New OpenFileDialog()
        fileSelector.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        fileSelector.Multiselect = True

        If fileSelector.ShowDialog = System.Windows.Forms.DialogResult.OK Then
            Dim filepath As String
            Dim file As FileInfo
            For Each filepath In fileSelector.FileNames
                If filepath Like "*.dotm" Or filepath Like "*.dot" Then
                    file = New FileInfo(filepath)
                    AddMacrosFromTemplate(file)
                End If
            Next
        End If

        LoadTemplateListBox()
        LoadAllMacroListBox()

    End Sub

    ''' <summary>
    ''' Loads the list boxes.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub MacroView_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadAllMacroListBox()
        LoadSelectedMacroListBox()
        LoadTemplateListBox()
    End Sub

    ''' <summary>
    ''' Allows the user to remove a template and its associated macros.
    ''' It also refreshs the list boxes.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub btnRemoveTemplate_Click(sender As Object, e As EventArgs) Handles btnRemoveTemplate.Click
        RemoveTemplateAndMacros(lstbSelectedTemplates.SelectedItem)
        LoadAllMacroListBox()
        LoadSelectedMacroListBox()
        LoadTemplateListBox()
    End Sub

    ''' <summary>
    ''' Refreshes the listbox that displays the templates
    ''' </summary>
    Private Sub LoadTemplateListBox()
        lstbSelectedTemplates.Items.Clear()

        Dim templates As New Collection
        templates = GetTemplates()

        Dim template As FileInfo
        For Each template In templates
            lstbSelectedTemplates.Items.Add(template)
        Next
    End Sub

    ''' <summary>
    ''' Refreshes the listbox that displays the macros that haven't been selected.
    ''' </summary>
    Private Sub LoadAllMacroListBox()
        lstbAllMacros.Items.Clear()

        Dim allMacros As New Collection
        allMacros = GetAllMacros()

        Dim macro As Macro
        For Each macro In allMacros
            lstbAllMacros.Items.Add(macro.Template.FullName + "\" + macro.Name)
        Next
    End Sub

    ''' <summary>
    ''' Refreshes the listbox that displays the selected macros.
    ''' </summary>
    Private Sub LoadSelectedMacroListBox()
        lstbSelectedMacros.Items.Clear()

        Dim selectedMacros As New Collection
        selectedMacros = GetSelectedMacros()

        Dim macro As Macro
        For Each macro In GetSelectedMacros()
            lstbSelectedMacros.Items.Add(macro.Template.FullName + "\" + macro.Name)
        Next
    End Sub

    ''' <summary>
    ''' Allows the user to select a macro
    ''' It also refreshes the list boxes.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub btnAddSelected_Click(sender As Object, e As EventArgs) Handles btnAddSelected.Click
        If Not lstbAllMacros.SelectedItem = Nothing Then
            Dim macroToAdd As Macro
            macroToAdd = PathToMacro(lstbAllMacros.SelectedItem)

            MoveMacroToSelected(macroToAdd)

            LoadAllMacroListBox()
            LoadSelectedMacroListBox()
        End If
    End Sub

    ''' <summary>
    ''' Allows the user to deselect a macro
    ''' It also refreshes the list boxes.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub btnRemoveSelected_Click(sender As Object, e As EventArgs) Handles btnRemoveSelected.Click
        If Not lstbSelectedMacros.SelectedItem = Nothing Then
            Dim macroToRemove As Macro
            macroToRemove = PathToMacro(lstbSelectedMacros.SelectedItem)

            MoveMacroFromSelected(macroToRemove)

            LoadAllMacroListBox()
            LoadSelectedMacroListBox()
        End If
    End Sub

    ''' <summary>
    ''' Takes a path like FilePath\MacroName and returns a Macro object
    ''' 
    ''' In case there may be macros with the same name in more than one template they must be displayed
    ''' as a path to disambiguate them. This Function is part of the View as it only relates to how
    ''' the macros are displayed and wouldn't be relevant if the GUI was changed.
    ''' </summary>
    ''' <param name="macroPath"></param>
    ''' <returns>Macro object</returns>
    Private Function PathToMacro(macroPath As String) As Macro
        Dim returnMacro As Macro
        Dim macroName As String
        Dim template As FileInfo

        Dim slashPosition As Integer
        slashPosition = InStrRev(macroPath, "\")

        macroName = Strings.Right(macroPath, Len(macroPath) - slashPosition)
        template = New FileInfo(Strings.Left(macroPath, slashPosition - 1))

        returnMacro = New Macro(macroName, template)

        Return returnMacro
    End Function

    Private Sub btnMoveUp_Click(sender As Object, e As EventArgs) Handles btnMoveUp.Click
        If Not lstbSelectedMacros.SelectedItem = Nothing Then
            Dim macroToMove As Macro
            macroToMove = PathToMacro(lstbSelectedMacros.SelectedItem)
            MoveMacroHigherInList(macroToMove)
        End If

        LoadSelectedMacroListBox()
    End Sub

    Private Sub btnMoveDown_Click(sender As Object, e As EventArgs) Handles btnMoveDown.Click
        If Not lstbSelectedMacros.SelectedItem = Nothing Then
            Dim macroToMove As Macro
            macroToMove = PathToMacro(lstbSelectedMacros.SelectedItem)
            MoveMacroLowerInList(macroToMove)
        End If

        LoadSelectedMacroListBox()
    End Sub
End Class