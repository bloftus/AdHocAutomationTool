Imports AdHocAutomationTool.Controller
Imports System.IO

Public Class MainView
    Private Sub btnSelectFiles_Click(sender As Object, e As EventArgs) Handles btnSelectFiles.Click
        Dim fileSelector As New OpenFileDialog()
        fileSelector.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        fileSelector.Multiselect = True
        If fileSelector.ShowDialog = System.Windows.Forms.DialogResult.OK Then
            Dim filepath As String
            Dim files As New Collection

            For Each filepath In fileSelector.FileNames
                files.Add(filepath)
            Next
            SelectFiles(files)
        End If
        LoadFileListBox()
    End Sub

    Private Sub btnSelectMacros_Click(sender As Object, e As EventArgs) Handles btnSelectMacros.Click
        MacroView.Show()
    End Sub

    Private Sub btnSelectFolder_Click(sender As Object, e As EventArgs) Handles btnSelectFolder.Click
        Dim folderSelector As New FolderBrowserDialog
        folderSelector.SelectedPath = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        If folderSelector.ShowDialog = System.Windows.Forms.DialogResult.OK Then
            SelectFolder(folderSelector.SelectedPath)
        End If
        LoadFileListBox()
    End Sub

    Private Sub btnRun_Click(sender As Object, e As EventArgs) Handles btnRun.Click
        RunSelectedMacrosOnFiles()
        lblStatus.Text = "Done!"
    End Sub

    Private Sub LoadFileListBox()
        lstbFiles.Items.Clear()

        Dim files As New Collection
        files = GetFiles()

        Dim file As FileInfo
        For Each file In files
            lstbFiles.Items.Add(file.FullName)
        Next
    End Sub

    Private Sub MainView_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ImportSelectedFiles()
        ImportTemplates()
        LoadFileListBox()
        Me.Show()
    End Sub

    Private Sub MainView_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        WriteSelectedFiles()
        WriteTemplates()
    End Sub

    Private Sub BtnRemoveSelected_Click(sender As Object, e As EventArgs) Handles BtnRemoveSelected.Click
        Dim filepaths As New Collection
        Dim filepath As String

        For Each filepath In lstbFiles.SelectedItems
            filepaths.Add(filepath)
        Next

        RemoveFiles(filepaths)

        LoadFileListBox()
    End Sub

End Class
