Imports System.IO
Imports AdHocAutomationTool.FileUtility
Imports AdHocAutomationTool.Macro

Public Class Logging

    ''' <summary>
    ''' Writes a log show who ran macros on files and when.
    ''' Its not that useful because it doesn't say what macros were run on how many files but extended logs like
    ''' that can come in a future iteration.
    ''' </summary>
    Public Shared Sub WriteLog()

        Dim outputFile As String
        outputFile = My.Computer.FileSystem.SpecialDirectories.CurrentUserApplicationData + "\AdHocLog.txt"
        Dim textFile As StreamWriter
        textFile = My.Computer.FileSystem.OpenTextFileWriter(outputFile, True)

        textFile.WriteLine(Now + "|" + Environment.UserName + "|Ran " + SelectedMacros.Count.ToString + " macros on " + SelectedFiles.Count.ToString + " files.")

        textFile.Close()
    End Sub

End Class
