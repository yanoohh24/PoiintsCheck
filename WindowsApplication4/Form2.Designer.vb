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
        Me.txtMemberDate = New System.Windows.Forms.TextBox()
        Me.cmdTop100 = New System.Windows.Forms.Button()
        Me.cmdTop1000 = New System.Windows.Forms.Button()
        Me.cmdTop10000 = New System.Windows.Forms.Button()
        Me.txtTop = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.cmdNegDB = New System.Windows.Forms.Button()
        Me.cmdNegPOS = New System.Windows.Forms.Button()
        Me.ListView1 = New System.Windows.Forms.ListView()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.StatCount = New System.Windows.Forms.ToolStripStatusLabel()
        Me.StatProgress = New System.Windows.Forms.ToolStripStatusLabel()
        CType(Me.DataGridView1,System.ComponentModel.ISupportInitialize).BeginInit
        Me.StatusStrip1.SuspendLayout
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
        Me.txtAccountno.Location = New System.Drawing.Point(642, 91)
        Me.txtAccountno.MaxLength = 32999999
        Me.txtAccountno.Name = "txtAccountno"
        Me.txtAccountno.Size = New System.Drawing.Size(100, 20)
        Me.txtAccountno.TabIndex = 8
        Me.txtAccountno.Visible = false
        '
        'ProgBar
        '
        Me.ProgBar.Location = New System.Drawing.Point(15, 75)
        Me.ProgBar.Name = "ProgBar"
        Me.ProgBar.Size = New System.Drawing.Size(836, 10)
        Me.ProgBar.TabIndex = 9
        '
        'txtMemberDate
        '
        Me.txtMemberDate.Location = New System.Drawing.Point(748, 91)
        Me.txtMemberDate.MaxLength = 32999999
        Me.txtMemberDate.Name = "txtMemberDate"
        Me.txtMemberDate.Size = New System.Drawing.Size(100, 20)
        Me.txtMemberDate.TabIndex = 10
        Me.txtMemberDate.Visible = false
        '
        'cmdTop100
        '
        Me.cmdTop100.Location = New System.Drawing.Point(15, 49)
        Me.cmdTop100.Name = "cmdTop100"
        Me.cmdTop100.Size = New System.Drawing.Size(65, 23)
        Me.cmdTop100.TabIndex = 11
        Me.cmdTop100.Text = "TOP 100"
        Me.cmdTop100.UseVisualStyleBackColor = true
        '
        'cmdTop1000
        '
        Me.cmdTop1000.Location = New System.Drawing.Point(86, 49)
        Me.cmdTop1000.Name = "cmdTop1000"
        Me.cmdTop1000.Size = New System.Drawing.Size(65, 23)
        Me.cmdTop1000.TabIndex = 12
        Me.cmdTop1000.Text = "TOP 1000"
        Me.cmdTop1000.UseVisualStyleBackColor = true
        '
        'cmdTop10000
        '
        Me.cmdTop10000.Location = New System.Drawing.Point(157, 49)
        Me.cmdTop10000.Name = "cmdTop10000"
        Me.cmdTop10000.Size = New System.Drawing.Size(73, 23)
        Me.cmdTop10000.TabIndex = 13
        Me.cmdTop10000.Text = "TOP 10000"
        Me.cmdTop10000.UseVisualStyleBackColor = true
        '
        'txtTop
        '
        Me.txtTop.Location = New System.Drawing.Point(280, 49)
        Me.txtTop.MaxLength = 32999999
        Me.txtTop.Name = "txtTop"
        Me.txtTop.Size = New System.Drawing.Size(100, 20)
        Me.txtTop.TabIndex = 14
        '
        'Label3
        '
        Me.Label3.AutoSize = true
        Me.Label3.Location = New System.Drawing.Point(236, 52)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(38, 13)
        Me.Label3.TabIndex = 15
        Me.Label3.Text = "Enter :"
        '
        'Label4
        '
        Me.Label4.AutoSize = true
        Me.Label4.Location = New System.Drawing.Point(386, 52)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(56, 13)
        Me.Label4.TabIndex = 16
        Me.Label4.Text = "Negative :"
        '
        'cmdNegDB
        '
        Me.cmdNegDB.Location = New System.Drawing.Point(501, 49)
        Me.cmdNegDB.Name = "cmdNegDB"
        Me.cmdNegDB.Size = New System.Drawing.Size(47, 23)
        Me.cmdNegDB.TabIndex = 17
        Me.cmdNegDB.Text = "DB"
        Me.cmdNegDB.UseVisualStyleBackColor = true
        '
        'cmdNegPOS
        '
        Me.cmdNegPOS.Location = New System.Drawing.Point(448, 49)
        Me.cmdNegPOS.Name = "cmdNegPOS"
        Me.cmdNegPOS.Size = New System.Drawing.Size(47, 23)
        Me.cmdNegPOS.TabIndex = 18
        Me.cmdNegPOS.Text = "POS"
        Me.cmdNegPOS.UseVisualStyleBackColor = true
        '
        'ListView1
        '
        Me.ListView1.GridLines = true
        Me.ListView1.Location = New System.Drawing.Point(12, 204)
        Me.ListView1.Name = "ListView1"
        Me.ListView1.Size = New System.Drawing.Size(836, 135)
        Me.ListView1.TabIndex = 19
        Me.ListView1.UseCompatibleStateImageBehavior = false
        Me.ListView1.View = System.Windows.Forms.View.Details
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.StatCount, Me.StatProgress})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 345)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(860, 22)
        Me.StatusStrip1.TabIndex = 20
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'StatCount
        '
        Me.StatCount.Name = "StatCount"
        Me.StatCount.Size = New System.Drawing.Size(46, 17)
        Me.StatCount.Text = "Count :"
        '
        'StatProgress
        '
        Me.StatProgress.Name = "StatProgress"
        Me.StatProgress.Size = New System.Drawing.Size(61, 17)
        Me.StatProgress.Text = "Progress : "
        '
        'Form2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6!, 13!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(860, 367)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.DataGridView1)
        Me.Controls.Add(Me.ListView1)
        Me.Controls.Add(Me.cmdNegPOS)
        Me.Controls.Add(Me.cmdNegDB)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.txtTop)
        Me.Controls.Add(Me.cmdTop10000)
        Me.Controls.Add(Me.cmdTop1000)
        Me.Controls.Add(Me.cmdTop100)
        Me.Controls.Add(Me.txtMemberDate)
        Me.Controls.Add(Me.ProgBar)
        Me.Controls.Add(Me.txtAccountno)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.cmdExcel)
        Me.Controls.Add(Me.cmdCheck)
        Me.Controls.Add(Me.cmdDisplay)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtFileDirectory)
        Me.Controls.Add(Me.Button1)
        Me.Icon = CType(resources.GetObject("$this.Icon"),System.Drawing.Icon)
        Me.Name = "Form2"
        Me.Text = "Points Checker"
        CType(Me.DataGridView1,System.ComponentModel.ISupportInitialize).EndInit
        Me.StatusStrip1.ResumeLayout(false)
        Me.StatusStrip1.PerformLayout
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
    Friend WithEvents txtMemberDate As System.Windows.Forms.TextBox
    Friend WithEvents cmdTop100 As System.Windows.Forms.Button
    Friend WithEvents cmdTop1000 As System.Windows.Forms.Button
    Friend WithEvents cmdTop10000 As System.Windows.Forms.Button
    Friend WithEvents txtTop As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cmdNegDB As System.Windows.Forms.Button
    Friend WithEvents cmdNegPOS As System.Windows.Forms.Button
    Friend WithEvents ListView1 As System.Windows.Forms.ListView
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatCount As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatProgress As System.Windows.Forms.ToolStripStatusLabel
End Class
