<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MacroView
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
        Me.lstbSelectedMacros = New System.Windows.Forms.ListBox()
        Me.lstbAllMacros = New System.Windows.Forms.ListBox()
        Me.btnAddSelected = New System.Windows.Forms.Button()
        Me.btnRemoveSelected = New System.Windows.Forms.Button()
        Me.lstbSelectedTemplates = New System.Windows.Forms.ListBox()
        Me.btnAddTemplate = New System.Windows.Forms.Button()
        Me.btnRemoveTemplate = New System.Windows.Forms.Button()
        Me.btnMoveUp = New System.Windows.Forms.Button()
        Me.btnMoveDown = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'lstbSelectedMacros
        '
        Me.lstbSelectedMacros.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.3!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstbSelectedMacros.FormattingEnabled = True
        Me.lstbSelectedMacros.HorizontalExtent = 600
        Me.lstbSelectedMacros.HorizontalScrollbar = True
        Me.lstbSelectedMacros.Location = New System.Drawing.Point(844, 36)
        Me.lstbSelectedMacros.Name = "lstbSelectedMacros"
        Me.lstbSelectedMacros.Size = New System.Drawing.Size(383, 511)
        Me.lstbSelectedMacros.TabIndex = 0
        '
        'lstbAllMacros
        '
        Me.lstbAllMacros.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.3!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstbAllMacros.FormattingEnabled = True
        Me.lstbAllMacros.HorizontalExtent = 600
        Me.lstbAllMacros.HorizontalScrollbar = True
        Me.lstbAllMacros.Location = New System.Drawing.Point(377, 36)
        Me.lstbAllMacros.Name = "lstbAllMacros"
        Me.lstbAllMacros.Size = New System.Drawing.Size(383, 511)
        Me.lstbAllMacros.TabIndex = 1
        '
        'btnAddSelected
        '
        Me.btnAddSelected.Location = New System.Drawing.Point(766, 246)
        Me.btnAddSelected.Name = "btnAddSelected"
        Me.btnAddSelected.Size = New System.Drawing.Size(75, 25)
        Me.btnAddSelected.TabIndex = 2
        Me.btnAddSelected.Text = "Add >>"
        Me.btnAddSelected.UseVisualStyleBackColor = True
        '
        'btnRemoveSelected
        '
        Me.btnRemoveSelected.Location = New System.Drawing.Point(766, 281)
        Me.btnRemoveSelected.Name = "btnRemoveSelected"
        Me.btnRemoveSelected.Size = New System.Drawing.Size(75, 25)
        Me.btnRemoveSelected.TabIndex = 3
        Me.btnRemoveSelected.Text = "<< Remove"
        Me.btnRemoveSelected.UseVisualStyleBackColor = True
        '
        'lstbSelectedTemplates
        '
        Me.lstbSelectedTemplates.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.3!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstbSelectedTemplates.FormattingEnabled = True
        Me.lstbSelectedTemplates.HorizontalExtent = 600
        Me.lstbSelectedTemplates.HorizontalScrollbar = True
        Me.lstbSelectedTemplates.Location = New System.Drawing.Point(12, 101)
        Me.lstbSelectedTemplates.Name = "lstbSelectedTemplates"
        Me.lstbSelectedTemplates.Size = New System.Drawing.Size(359, 446)
        Me.lstbSelectedTemplates.TabIndex = 4
        '
        'btnAddTemplate
        '
        Me.btnAddTemplate.Location = New System.Drawing.Point(139, 10)
        Me.btnAddTemplate.Name = "btnAddTemplate"
        Me.btnAddTemplate.Size = New System.Drawing.Size(105, 25)
        Me.btnAddTemplate.TabIndex = 5
        Me.btnAddTemplate.Text = "Add Template"
        Me.btnAddTemplate.UseVisualStyleBackColor = True
        '
        'btnRemoveTemplate
        '
        Me.btnRemoveTemplate.Location = New System.Drawing.Point(139, 41)
        Me.btnRemoveTemplate.Name = "btnRemoveTemplate"
        Me.btnRemoveTemplate.Size = New System.Drawing.Size(105, 25)
        Me.btnRemoveTemplate.TabIndex = 6
        Me.btnRemoveTemplate.Text = "Remove Template"
        Me.btnRemoveTemplate.UseVisualStyleBackColor = True
        '
        'btnMoveUp
        '
        Me.btnMoveUp.Font = New System.Drawing.Font("Microsoft Sans Serif", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnMoveUp.Location = New System.Drawing.Point(1236, 229)
        Me.btnMoveUp.Margin = New System.Windows.Forms.Padding(0)
        Me.btnMoveUp.Name = "btnMoveUp"
        Me.btnMoveUp.Padding = New System.Windows.Forms.Padding(5, 0, 0, 0)
        Me.btnMoveUp.Size = New System.Drawing.Size(27, 42)
        Me.btnMoveUp.TabIndex = 7
        Me.btnMoveUp.Text = "⮙"
        Me.btnMoveUp.UseVisualStyleBackColor = True
        '
        'btnMoveDown
        '
        Me.btnMoveDown.Font = New System.Drawing.Font("Microsoft Sans Serif", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnMoveDown.Location = New System.Drawing.Point(1236, 281)
        Me.btnMoveDown.Name = "btnMoveDown"
        Me.btnMoveDown.Padding = New System.Windows.Forms.Padding(5, 0, 0, 0)
        Me.btnMoveDown.Size = New System.Drawing.Size(27, 42)
        Me.btnMoveDown.TabIndex = 8
        Me.btnMoveDown.Text = "⮛"
        Me.btnMoveDown.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(535, 17)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(66, 13)
        Me.Label1.TabIndex = 9
        Me.Label1.Text = "All Macros"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(984, 17)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(102, 13)
        Me.Label2.TabIndex = 10
        Me.Label2.Text = "Selected Macros"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(159, 82)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(65, 13)
        Me.Label3.TabIndex = 11
        Me.Label3.Text = "Templates"
        '
        'MacroView
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1272, 553)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnMoveDown)
        Me.Controls.Add(Me.btnMoveUp)
        Me.Controls.Add(Me.btnRemoveTemplate)
        Me.Controls.Add(Me.btnAddTemplate)
        Me.Controls.Add(Me.lstbSelectedTemplates)
        Me.Controls.Add(Me.btnRemoveSelected)
        Me.Controls.Add(Me.btnAddSelected)
        Me.Controls.Add(Me.lstbAllMacros)
        Me.Controls.Add(Me.lstbSelectedMacros)
        Me.Name = "MacroView"
        Me.Text = "MacroView"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lstbSelectedMacros As ListBox
    Friend WithEvents lstbAllMacros As ListBox
    Friend WithEvents btnAddSelected As Button
    Friend WithEvents btnRemoveSelected As Button
    Friend WithEvents lstbSelectedTemplates As ListBox
    Friend WithEvents btnAddTemplate As Button
    Friend WithEvents btnRemoveTemplate As Button
    Friend WithEvents btnMoveUp As Button
    Friend WithEvents btnMoveDown As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Label3 As Label
End Class
