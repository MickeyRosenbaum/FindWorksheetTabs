<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.btnFindTabs = New System.Windows.Forms.Button()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.btnWriteToNotepad = New System.Windows.Forms.Button()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnFindTabs
        '
        Me.btnFindTabs.Location = New System.Drawing.Point(315, 24)
        Me.btnFindTabs.Name = "btnFindTabs"
        Me.btnFindTabs.Size = New System.Drawing.Size(171, 61)
        Me.btnFindTabs.TabIndex = 0
        Me.btnFindTabs.Text = "Select Excel File"
        Me.btnFindTabs.UseVisualStyleBackColor = True
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'DataGridView1
        '
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Location = New System.Drawing.Point(178, 111)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.RowTemplate.Height = 24
        Me.DataGridView1.Size = New System.Drawing.Size(444, 156)
        Me.DataGridView1.TabIndex = 1
        '
        'btnWriteToNotepad
        '
        Me.btnWriteToNotepad.Location = New System.Drawing.Point(339, 297)
        Me.btnWriteToNotepad.Name = "btnWriteToNotepad"
        Me.btnWriteToNotepad.Size = New System.Drawing.Size(123, 42)
        Me.btnWriteToNotepad.TabIndex = 2
        Me.btnWriteToNotepad.Text = "Write To Notepad"
        Me.btnWriteToNotepad.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 351)
        Me.Controls.Add(Me.btnWriteToNotepad)
        Me.Controls.Add(Me.DataGridView1)
        Me.Controls.Add(Me.btnFindTabs)
        Me.Name = "Form1"
        Me.Text = "Copy List of Tabs to Notepad"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents btnFindTabs As Button
    Friend WithEvents OpenFileDialog1 As OpenFileDialog
    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents btnWriteToNotepad As Button
End Class
