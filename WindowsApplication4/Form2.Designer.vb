<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form2
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form2))
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
        Me.txtFileDirectory = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.cmdDisplay = New System.Windows.Forms.Button()
        Me.cmdCheck = New System.Windows.Forms.Button()
        Me.cmdExcel = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtAccountno = New System.Windows.Forms.TextBox()
        Me.ProgBar = New System.Windows.Forms.ProgressBar()
        CType(Me.DataGridView1,System.ComponentModel.ISupportInitialize).BeginInit
        Me.SuspendLayout
        '
        'DataGridView1
        '
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Location = New System.Drawing.Point(12, 91)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.Size = New System.Drawing.Size(838, 248)
        Me.DataGridView1.TabIndex = 0
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(818, 20)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(30, 23)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "...."
        Me.Button1.UseVisualStyleBackColor = true
        '
        'txtFileDirectory
        '
        Me.txtFileDirectory.Enabled = false
        Me.txtFileDirectory.Location = New System.Drawing.Point(12, 23)
        Me.txtFileDirectory.Name = "txtFileDirectory"
        Me.txtFileDirectory.Size = New System.Drawing.Size(800, 20)
        Me.txtFileDirectory.TabIndex = 2
        Me.txtFileDirectory.Text = "C:\Users\User\Desktop\Customerx.csv"
        '
        'Label1
        '
        Me.Label1.AutoSize = true
        Me.Label1.Location = New System.Drawing.Point(12, 7)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(68, 13)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "File Directory"
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'cmdDisplay
        '
        Me.cmdDisplay.Location = New System.Drawing.Point(773, 49)
        Me.cmdDisplay.Name = "cmdDisplay"
        Me.cmdDisplay.Size = New System.Drawing.Size(75, 23)
        Me.cmdDisplay.TabIndex = 4
        Me.cmdDisplay.Text = "Display"
        Me.cmdDisplay.UseVisualStyleBackColor = true
        '
        'cmdCheck
        '
        Me.cmdCheck.Location = New System.Drawing.Point(692, 49)
        Me.cmdCheck.Name = "cmdCheck"
        Me.cmdCheck.Size = New System.Drawing.Size(75, 23)
        Me.cmdCheck.TabIndex = 5
        Me.cmdCheck.Text = "Check"
        Me.cmdCheck.UseVisualStyleBackColor = true
        '
        'cmdExcel
        '
        Me.cmdExcel.Location = New System.Drawing.Point(611, 49)
        Me.cmdExcel.Name = "cmdExcel"
        Me.cmdExcel.Size = New System.Drawing.Size(75, 23)
        Me.cmdExcel.TabIndex = 6
        Me.cmdExcel.Text = "Excel"
        Me.cmdExcel.UseVisualStyleBackColor = true
        '
        'Label2
        '
        Me.Label2.AutoSize = true
        Me.Label2.Location = New System.Drawing.Point(12, 54)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(39, 13)
        Me.Label2.TabIndex = 7
        Me.Label2.Text = "Label2"
        '
        'txtAccountno
        '
        Me.txtAccountno.Location = New System.Drawing.Point(57, 49)
        Me.txtAccountno.MaxLength = 32999999
        Me.txtAccountno.Name = "txtAccountno"
        Me.txtAccountno.Size = New System.Drawing.Size(100, 20)
        Me.txtAccountno.TabIndex = 8
        '
        'ProgBar
        '
        Me.ProgBar.Location = New System.Drawing.Point(15, 75)
        Me.ProgBar.Name = "ProgBar"
        Me.ProgBar.Size = New System.Drawing.Size(836, 10)
        Me.ProgBar.TabIndex = 9
        '
        'Form2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6!, 13!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(860, 351)
        Me.Controls.Add(Me.ProgBar)
        Me.Controls.Add(Me.txtAccountno)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.cmdExcel)
        Me.Controls.Add(Me.cmdCheck)
        Me.Controls.Add(Me.cmdDisplay)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtFileDirectory)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.DataGridView1)
        Me.Icon = CType(resources.GetObject("$this.Icon"),System.Drawing.Icon)
        Me.Name = "Form2"
        Me.Text = "Points Checker"
        CType(Me.DataGridView1,System.ComponentModel.ISupportInitialize).EndInit
        Me.ResumeLayout(false)
        Me.PerformLayout

End Sub
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents txtFileDirectory As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents cmdDisplay As System.Windows.Forms.Button
    Friend WithEvents cmdCheck As System.Windows.Forms.Button
    Friend WithEvents cmdExcel As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtAccountno As System.Windows.Forms.TextBox
    Friend WithEvents ProgBar As System.Windows.Forms.ProgressBar
End Class
