<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MainView
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.btnSelectFiles = New System.Windows.Forms.Button()
        Me.btnSelectMacros = New System.Windows.Forms.Button()
        Me.btnRun = New System.Windows.Forms.Button()
        Me.btnSelectFolder = New System.Windows.Forms.Button()
        Me.lstbFiles = New System.Windows.Forms.ListBox()
        Me.BtnRemoveSelected = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lblStatus = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'btnSelectFiles
        '
        Me.btnSelectFiles.Location = New System.Drawing.Point(28, 110)
        Me.btnSelectFiles.Name = "btnSelectFiles"
        Me.btnSelectFiles.Size = New System.Drawing.Size(102, 23)
        Me.btnSelectFiles.TabIndex = 2
        Me.btnSelectFiles.Text = "Select &Files"
        Me.btnSelectFiles.UseVisualStyleBackColor = True
        '
        'btnSelectMacros
        '
        Me.btnSelectMacros.Location = New System.Drawing.Point(28, 188)
        Me.btnSelectMacros.Name = "btnSelectMacros"
        Me.btnSelectMacros.Size = New System.Drawing.Size(102, 23)
        Me.btnSelectMacros.TabIndex = 3
        Me.btnSelectMacros.Text = "Select &Macros"
        Me.btnSelectMacros.UseVisualStyleBackColor = True
        '
        'btnRun
        '
        Me.btnRun.Location = New System.Drawing.Point(42, 227)
        Me.btnRun.Name = "btnRun"
        Me.btnRun.Size = New System.Drawing.Size(75, 23)
        Me.btnRun.TabIndex = 4
        Me.btnRun.Text = "Run!"
        Me.btnRun.UseVisualStyleBackColor = True
        '
        'btnSelectFolder
        '
        Me.btnSelectFolder.Location = New System.Drawing.Point(28, 149)
        Me.btnSelectFolder.Name = "btnSelectFolder"
        Me.btnSelectFolder.Size = New System.Drawing.Size(102, 23)
        Me.btnSelectFolder.TabIndex = 5
        Me.btnSelectFolder.Text = "Select F&older"
        Me.btnSelectFolder.UseVisualStyleBackColor = True
        '
        'lstbFiles
        '
        Me.lstbFiles.FormattingEnabled = True
        Me.lstbFiles.Location = New System.Drawing.Point(153, 38)
        Me.lstbFiles.Name = "lstbFiles"
        Me.lstbFiles.SelectionMode = System.Windows.Forms.SelectionMode.MultiSimple
        Me.lstbFiles.Size = New System.Drawing.Size(608, 290)
        Me.lstbFiles.TabIndex = 6
        '
        'BtnRemoveSelected
        '
        Me.BtnRemoveSelected.Location = New System.Drawing.Point(392, 334)
        Me.BtnRemoveSelected.Name = "BtnRemoveSelected"
        Me.BtnRemoveSelected.Size = New System.Drawing.Size(130, 23)
        Me.BtnRemoveSelected.TabIndex = 7
        Me.BtnRemoveSelected.Text = "Remove Selected Files"
        Me.BtnRemoveSelected.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(414, 19)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(87, 13)
        Me.Label1.TabIndex = 8
        Me.Label1.Text = "Selected Files"
        '
        'lblStatus
        '
        Me.lblStatus.AutoSize = True
        Me.lblStatus.Location = New System.Drawing.Point(12, 339)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(0, 13)
        Me.lblStatus.TabIndex = 9
        '
        'MainView
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(773, 362)
        Me.Controls.Add(Me.lblStatus)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.BtnRemoveSelected)
        Me.Controls.Add(Me.lstbFiles)
        Me.Controls.Add(Me.btnSelectFolder)
        Me.Controls.Add(Me.btnRun)
        Me.Controls.Add(Me.btnSelectMacros)
        Me.Controls.Add(Me.btnSelectFiles)
        Me.Name = "MainView"
        Me.Text = " Ad Hoc Automation"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnSelectFiles As Button
    Friend WithEvents btnSelectMacros As Button
    Friend WithEvents btnRun As Button
    Friend WithEvents btnSelectFolder As Button
    Friend WithEvents lstbFiles As ListBox
    Friend WithEvents BtnRemoveSelected As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents lblStatus As Label
End Class
