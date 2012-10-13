<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmTwitter
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmTwitter))
        Me.cmdClose = New System.Windows.Forms.Button()
        Me.cmdSend = New System.Windows.Forms.Button()
        Me.lblMessage = New System.Windows.Forms.Label()
        Me.txtMessage = New System.Windows.Forms.TextBox()
        Me.lblCaracteres = New System.Windows.Forms.Label()
        Me.lblResultado = New System.Windows.Forms.Label()
        Me.cmdExcel = New System.Windows.Forms.Button()
        Me.lblProcesing = New System.Windows.Forms.Label()
        Me.cmbImportaOptions = New System.Windows.Forms.ComboBox()
        Me.lblOptions = New System.Windows.Forms.Label()
        Me.txtSeach = New System.Windows.Forms.TextBox()
        Me.lblSearch = New System.Windows.Forms.Label()
        Me.lblTweets = New System.Windows.Forms.Label()
        Me.txtTweets = New System.Windows.Forms.TextBox()
        Me.lblInReplayToID = New System.Windows.Forms.Label()
        Me.txtInReplayToID = New System.Windows.Forms.TextBox()
        Me.lblAPI = New System.Windows.Forms.Label()
        Me.lblVersion = New System.Windows.Forms.Label()
        Me.btnAdd = New System.Windows.Forms.Button()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ToolStripStatusLabel2 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.lblUserName = New System.Windows.Forms.Label()
        Me.cmbUserName = New System.Windows.Forms.ComboBox()
        Me.cmdAddUser = New System.Windows.Forms.Button()
        Me.cmdDeleteUser = New System.Windows.Forms.Button()
        Me.StatusStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdClose
        '
        Me.cmdClose.Location = New System.Drawing.Point(495, 382)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.Size = New System.Drawing.Size(75, 39)
        Me.cmdClose.TabIndex = 0
        Me.cmdClose.Text = "&Close"
        Me.cmdClose.UseVisualStyleBackColor = True
        '
        'cmdSend
        '
        Me.cmdSend.Enabled = False
        Me.cmdSend.Location = New System.Drawing.Point(495, 161)
        Me.cmdSend.Name = "cmdSend"
        Me.cmdSend.Size = New System.Drawing.Size(75, 39)
        Me.cmdSend.TabIndex = 1
        Me.cmdSend.Text = "&Send"
        Me.cmdSend.UseVisualStyleBackColor = True
        '
        'lblMessage
        '
        Me.lblMessage.AutoSize = True
        Me.lblMessage.Location = New System.Drawing.Point(30, 83)
        Me.lblMessage.Name = "lblMessage"
        Me.lblMessage.Size = New System.Drawing.Size(50, 13)
        Me.lblMessage.TabIndex = 15
        Me.lblMessage.Text = "Message"
        '
        'txtMessage
        '
        Me.txtMessage.Location = New System.Drawing.Point(107, 54)
        Me.txtMessage.Multiline = True
        Me.txtMessage.Name = "txtMessage"
        Me.txtMessage.Size = New System.Drawing.Size(463, 95)
        Me.txtMessage.TabIndex = 14
        '
        'lblCaracteres
        '
        Me.lblCaracteres.AutoSize = True
        Me.lblCaracteres.Location = New System.Drawing.Point(337, 161)
        Me.lblCaracteres.Name = "lblCaracteres"
        Me.lblCaracteres.Size = New System.Drawing.Size(58, 13)
        Me.lblCaracteres.TabIndex = 16
        Me.lblCaracteres.Text = "140-0=140"
        '
        'lblResultado
        '
        Me.lblResultado.AutoSize = True
        Me.lblResultado.Location = New System.Drawing.Point(33, 200)
        Me.lblResultado.MaximumSize = New System.Drawing.Size(500, 0)
        Me.lblResultado.MinimumSize = New System.Drawing.Size(5, 50)
        Me.lblResultado.Name = "lblResultado"
        Me.lblResultado.Size = New System.Drawing.Size(5, 50)
        Me.lblResultado.TabIndex = 18
        '
        'cmdExcel
        '
        Me.cmdExcel.Enabled = False
        Me.cmdExcel.Location = New System.Drawing.Point(495, 331)
        Me.cmdExcel.Name = "cmdExcel"
        Me.cmdExcel.Size = New System.Drawing.Size(75, 40)
        Me.cmdExcel.TabIndex = 19
        Me.cmdExcel.Text = "Export to Excel"
        Me.cmdExcel.UseVisualStyleBackColor = True
        '
        'lblProcesing
        '
        Me.lblProcesing.AutoSize = True
        Me.lblProcesing.Location = New System.Drawing.Point(103, 381)
        Me.lblProcesing.MaximumSize = New System.Drawing.Size(150, 0)
        Me.lblProcesing.MinimumSize = New System.Drawing.Size(150, 0)
        Me.lblProcesing.Name = "lblProcesing"
        Me.lblProcesing.Size = New System.Drawing.Size(150, 13)
        Me.lblProcesing.TabIndex = 20
        Me.lblProcesing.Text = "..."
        '
        'cmbImportaOptions
        '
        Me.cmbImportaOptions.Enabled = False
        Me.cmbImportaOptions.FormattingEnabled = True
        Me.cmbImportaOptions.Location = New System.Drawing.Point(107, 301)
        Me.cmbImportaOptions.Name = "cmbImportaOptions"
        Me.cmbImportaOptions.Size = New System.Drawing.Size(285, 21)
        Me.cmbImportaOptions.TabIndex = 21
        Me.cmbImportaOptions.Text = "Search"
        '
        'lblOptions
        '
        Me.lblOptions.AutoSize = True
        Me.lblOptions.Location = New System.Drawing.Point(30, 301)
        Me.lblOptions.Name = "lblOptions"
        Me.lblOptions.Size = New System.Drawing.Size(43, 13)
        Me.lblOptions.TabIndex = 23
        Me.lblOptions.Text = "Options"
        '
        'txtSeach
        '
        Me.txtSeach.Enabled = False
        Me.txtSeach.Location = New System.Drawing.Point(107, 328)
        Me.txtSeach.Name = "txtSeach"
        Me.txtSeach.Size = New System.Drawing.Size(285, 20)
        Me.txtSeach.TabIndex = 22
        '
        'lblSearch
        '
        Me.lblSearch.AutoSize = True
        Me.lblSearch.Location = New System.Drawing.Point(30, 331)
        Me.lblSearch.Name = "lblSearch"
        Me.lblSearch.Size = New System.Drawing.Size(41, 13)
        Me.lblSearch.TabIndex = 24
        Me.lblSearch.Text = "Search"
        '
        'lblTweets
        '
        Me.lblTweets.AutoSize = True
        Me.lblTweets.Location = New System.Drawing.Point(30, 357)
        Me.lblTweets.Name = "lblTweets"
        Me.lblTweets.Size = New System.Drawing.Size(42, 13)
        Me.lblTweets.TabIndex = 26
        Me.lblTweets.Text = "Tweets"
        '
        'txtTweets
        '
        Me.txtTweets.Enabled = False
        Me.txtTweets.Location = New System.Drawing.Point(107, 354)
        Me.txtTweets.Name = "txtTweets"
        Me.txtTweets.Size = New System.Drawing.Size(285, 20)
        Me.txtTweets.TabIndex = 25
        Me.txtTweets.Text = "3200"
        '
        'lblInReplayToID
        '
        Me.lblInReplayToID.AutoSize = True
        Me.lblInReplayToID.Location = New System.Drawing.Point(33, 161)
        Me.lblInReplayToID.Name = "lblInReplayToID"
        Me.lblInReplayToID.Size = New System.Drawing.Size(68, 13)
        Me.lblInReplayToID.TabIndex = 28
        Me.lblInReplayToID.Text = "In Replay To"
        '
        'txtInReplayToID
        '
        Me.txtInReplayToID.Enabled = False
        Me.txtInReplayToID.Location = New System.Drawing.Point(107, 158)
        Me.txtInReplayToID.Name = "txtInReplayToID"
        Me.txtInReplayToID.Size = New System.Drawing.Size(220, 20)
        Me.txtInReplayToID.TabIndex = 27
        '
        'lblAPI
        '
        Me.lblAPI.AutoSize = True
        Me.lblAPI.Location = New System.Drawing.Point(104, 408)
        Me.lblAPI.MaximumSize = New System.Drawing.Size(150, 0)
        Me.lblAPI.MinimumSize = New System.Drawing.Size(150, 0)
        Me.lblAPI.Name = "lblAPI"
        Me.lblAPI.Size = New System.Drawing.Size(150, 13)
        Me.lblAPI.TabIndex = 29
        Me.lblAPI.Text = "API ()"
        '
        'lblVersion
        '
        Me.lblVersion.AutoSize = True
        Me.lblVersion.Location = New System.Drawing.Point(365, 381)
        Me.lblVersion.MaximumSize = New System.Drawing.Size(50, 0)
        Me.lblVersion.MinimumSize = New System.Drawing.Size(50, 0)
        Me.lblVersion.Name = "lblVersion"
        Me.lblVersion.Size = New System.Drawing.Size(50, 13)
        Me.lblVersion.TabIndex = 30
        Me.lblVersion.Text = "0.1.4"
        '
        'btnAdd
        '
        Me.btnAdd.Enabled = False
        Me.btnAdd.Location = New System.Drawing.Point(495, 283)
        Me.btnAdd.Name = "btnAdd"
        Me.btnAdd.Size = New System.Drawing.Size(75, 39)
        Me.btnAdd.TabIndex = 31
        Me.btnAdd.Text = "Schedule Tweets"
        Me.btnAdd.UseVisualStyleBackColor = True
        '
        'Timer1
        '
        Me.Timer1.Interval = 1000
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1, Me.ToolStripStatusLabel2})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 436)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(582, 22)
        Me.StatusStrip1.TabIndex = 33
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(34, 17)
        Me.ToolStripStatusLabel1.Text = "00:00"
        '
        'ToolStripStatusLabel2
        '
        Me.ToolStripStatusLabel2.Name = "ToolStripStatusLabel2"
        Me.ToolStripStatusLabel2.Size = New System.Drawing.Size(39, 17)
        Me.ToolStripStatusLabel2.Text = "Status"
        '
        'lblUserName
        '
        Me.lblUserName.AutoSize = True
        Me.lblUserName.Location = New System.Drawing.Point(37, 26)
        Me.lblUserName.Name = "lblUserName"
        Me.lblUserName.Size = New System.Drawing.Size(60, 13)
        Me.lblUserName.TabIndex = 35
        Me.lblUserName.Text = "User Name"
        '
        'cmbUserName
        '
        Me.cmbUserName.FormattingEnabled = True
        Me.cmbUserName.Location = New System.Drawing.Point(107, 23)
        Me.cmbUserName.Name = "cmbUserName"
        Me.cmbUserName.Size = New System.Drawing.Size(285, 21)
        Me.cmbUserName.TabIndex = 34
        '
        'cmdAddUser
        '
        Me.cmdAddUser.Location = New System.Drawing.Point(414, 14)
        Me.cmdAddUser.Name = "cmdAddUser"
        Me.cmdAddUser.Size = New System.Drawing.Size(75, 37)
        Me.cmdAddUser.TabIndex = 36
        Me.cmdAddUser.Text = "&Add User"
        Me.cmdAddUser.UseVisualStyleBackColor = True
        '
        'cmdDeleteUser
        '
        Me.cmdDeleteUser.Location = New System.Drawing.Point(495, 14)
        Me.cmdDeleteUser.Name = "cmdDeleteUser"
        Me.cmdDeleteUser.Size = New System.Drawing.Size(75, 37)
        Me.cmdDeleteUser.TabIndex = 37
        Me.cmdDeleteUser.Text = "&Delete"
        Me.cmdDeleteUser.UseVisualStyleBackColor = True
        '
        'frmTwitter
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(582, 458)
        Me.Controls.Add(Me.cmdDeleteUser)
        Me.Controls.Add(Me.cmdAddUser)
        Me.Controls.Add(Me.lblUserName)
        Me.Controls.Add(Me.cmbUserName)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.btnAdd)
        Me.Controls.Add(Me.lblVersion)
        Me.Controls.Add(Me.lblAPI)
        Me.Controls.Add(Me.lblInReplayToID)
        Me.Controls.Add(Me.txtInReplayToID)
        Me.Controls.Add(Me.lblTweets)
        Me.Controls.Add(Me.txtTweets)
        Me.Controls.Add(Me.lblSearch)
        Me.Controls.Add(Me.lblOptions)
        Me.Controls.Add(Me.txtSeach)
        Me.Controls.Add(Me.cmbImportaOptions)
        Me.Controls.Add(Me.lblProcesing)
        Me.Controls.Add(Me.cmdExcel)
        Me.Controls.Add(Me.lblResultado)
        Me.Controls.Add(Me.lblCaracteres)
        Me.Controls.Add(Me.lblMessage)
        Me.Controls.Add(Me.txtMessage)
        Me.Controls.Add(Me.cmdSend)
        Me.Controls.Add(Me.cmdClose)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "frmTwitter"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "ExcelTweet"
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cmdClose As System.Windows.Forms.Button
    Friend WithEvents cmdSend As System.Windows.Forms.Button
    Friend WithEvents lblMessage As System.Windows.Forms.Label
    Friend WithEvents txtMessage As System.Windows.Forms.TextBox
    Friend WithEvents lblCaracteres As System.Windows.Forms.Label
    Friend WithEvents lblResultado As System.Windows.Forms.Label
    Friend WithEvents cmdExcel As System.Windows.Forms.Button
    Friend WithEvents lblProcesing As System.Windows.Forms.Label
    Friend WithEvents cmbImportaOptions As System.Windows.Forms.ComboBox
    Friend WithEvents lblOptions As System.Windows.Forms.Label
    Friend WithEvents txtSeach As System.Windows.Forms.TextBox
    Friend WithEvents lblSearch As System.Windows.Forms.Label
    Friend WithEvents lblTweets As System.Windows.Forms.Label
    Friend WithEvents txtTweets As System.Windows.Forms.TextBox
    Friend WithEvents lblInReplayToID As System.Windows.Forms.Label
    Friend WithEvents txtInReplayToID As System.Windows.Forms.TextBox
    Friend WithEvents lblAPI As System.Windows.Forms.Label
    Friend WithEvents lblVersion As System.Windows.Forms.Label
    Friend WithEvents btnAdd As System.Windows.Forms.Button
    Public WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents ToolStripStatusLabel2 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents lblUserName As System.Windows.Forms.Label
    Friend WithEvents cmbUserName As System.Windows.Forms.ComboBox
    Friend WithEvents cmdAddUser As System.Windows.Forms.Button
    Friend WithEvents cmdDeleteUser As System.Windows.Forms.Button

End Class
