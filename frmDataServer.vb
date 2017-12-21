Public Class frmDataServer
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents cmdOK As System.Windows.Forms.Button
    Friend WithEvents txtUserName As System.Windows.Forms.TextBox
    Friend WithEvents lblUserName As System.Windows.Forms.Label
    Friend WithEvents lblPassword As System.Windows.Forms.Label
    Friend WithEvents txtPassword As System.Windows.Forms.TextBox
    Friend WithEvents lblDatabase As System.Windows.Forms.Label
    Friend WithEvents txtDatabase As System.Windows.Forms.TextBox
    Friend WithEvents lblServer As System.Windows.Forms.Label
    Friend WithEvents txtServer As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmDataServer))
        Me.cmdCancel = New System.Windows.Forms.Button
        Me.cmdOK = New System.Windows.Forms.Button
        Me.txtUserName = New System.Windows.Forms.TextBox
        Me.lblUserName = New System.Windows.Forms.Label
        Me.lblPassword = New System.Windows.Forms.Label
        Me.txtPassword = New System.Windows.Forms.TextBox
        Me.lblDatabase = New System.Windows.Forms.Label
        Me.txtDatabase = New System.Windows.Forms.TextBox
        Me.lblServer = New System.Windows.Forms.Label
        Me.txtServer = New System.Windows.Forms.TextBox
        Me.SuspendLayout()
        '
        'cmdCancel
        '
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Location = New System.Drawing.Point(75, 225)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.TabIndex = 6
        Me.cmdCancel.Text = "Cancel"
        '
        'cmdOK
        '
        Me.cmdOK.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.cmdOK.Location = New System.Drawing.Point(75, 200)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.TabIndex = 5
        Me.cmdOK.Text = "OK"
        '
        'txtUserName
        '
        Me.txtUserName.Location = New System.Drawing.Point(25, 25)
        Me.txtUserName.Name = "txtUserName"
        Me.txtUserName.Size = New System.Drawing.Size(175, 20)
        Me.txtUserName.TabIndex = 1
        Me.txtUserName.Text = ""
        '
        'lblUserName
        '
        Me.lblUserName.Location = New System.Drawing.Point(25, 0)
        Me.lblUserName.Name = "lblUserName"
        Me.lblUserName.Size = New System.Drawing.Size(175, 23)
        Me.lblUserName.TabIndex = 26
        Me.lblUserName.Text = "UserName"
        Me.lblUserName.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'lblPassword
        '
        Me.lblPassword.Location = New System.Drawing.Point(25, 50)
        Me.lblPassword.Name = "lblPassword"
        Me.lblPassword.Size = New System.Drawing.Size(175, 23)
        Me.lblPassword.TabIndex = 28
        Me.lblPassword.Text = "Password"
        Me.lblPassword.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'txtPassword
        '
        Me.txtPassword.Location = New System.Drawing.Point(25, 75)
        Me.txtPassword.Name = "txtPassword"
        Me.txtPassword.PasswordChar = Microsoft.VisualBasic.ChrW(42)
        Me.txtPassword.Size = New System.Drawing.Size(175, 20)
        Me.txtPassword.TabIndex = 2
        Me.txtPassword.Text = ""
        '
        'lblDatabase
        '
        Me.lblDatabase.Location = New System.Drawing.Point(25, 100)
        Me.lblDatabase.Name = "lblDatabase"
        Me.lblDatabase.Size = New System.Drawing.Size(175, 23)
        Me.lblDatabase.TabIndex = 30
        Me.lblDatabase.Text = "Database (CADM)"
        Me.lblDatabase.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'txtDatabase
        '
        Me.txtDatabase.Location = New System.Drawing.Point(25, 125)
        Me.txtDatabase.Name = "txtDatabase"
        Me.txtDatabase.Size = New System.Drawing.Size(175, 20)
        Me.txtDatabase.TabIndex = 3
        Me.txtDatabase.Text = ""
        '
        'lblServer
        '
        Me.lblServer.Location = New System.Drawing.Point(25, 150)
        Me.lblServer.Name = "lblServer"
        Me.lblServer.Size = New System.Drawing.Size(175, 23)
        Me.lblServer.TabIndex = 32
        Me.lblServer.Text = "Server (www.silverbulletinc.com)"
        Me.lblServer.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'txtServer
        '
        Me.txtServer.Location = New System.Drawing.Point(25, 175)
        Me.txtServer.Name = "txtServer"
        Me.txtServer.Size = New System.Drawing.Size(175, 20)
        Me.txtServer.TabIndex = 4
        Me.txtServer.Text = ""
        '
        'frmDataServer
        '
        Me.AcceptButton = Me.cmdOK
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.CancelButton = Me.cmdCancel
        Me.ClientSize = New System.Drawing.Size(217, 266)
        Me.Controls.Add(Me.lblServer)
        Me.Controls.Add(Me.txtServer)
        Me.Controls.Add(Me.lblDatabase)
        Me.Controls.Add(Me.txtDatabase)
        Me.Controls.Add(Me.lblPassword)
        Me.Controls.Add(Me.txtPassword)
        Me.Controls.Add(Me.lblUserName)
        Me.Controls.Add(Me.txtUserName)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdOK)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmDataServer"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "SQL Server"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub frmDataServer_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim strDatabaseConnection As String
        'Dim intStart As Integer
        'Dim intEnd As Integer
        strDatabaseConnection = frmDataTree.myConnectString
        Try
            Me.txtUserName.Text = frmDataTree.strUserName
            Me.txtPassword.Text = frmDataTree.strPassword
            Me.txtDatabase.Text = frmDataTree.strDatabase
            Me.txtServer.Text = frmDataTree.strServer
            'If strDatabaseConnection Like "Provider=SQLOLEDB*" Then
            '    'UserName
            '    intStart = InStr(strDatabaseConnection, "User ID=") + 8
            '    intEnd = InStr(intStart, strDatabaseConnection, ";") - 0
            '    Me.txtUserName.Text = Mid(strDatabaseConnection, intStart, intEnd - intStart)
            '    'Password
            '    Me.txtPassword.Text = ""
            '    'Database
            '    intStart = InStr(strDatabaseConnection, "Initial Catalog=") + 16
            '    intEnd = InStr(intStart, strDatabaseConnection, ";") - 0
            '    Me.txtDatabase.Text = Mid(strDatabaseConnection, intStart, intEnd - intStart)
            '    'Server
            '    intStart = InStr(strDatabaseConnection, "Data Source=") + 12
            '    intEnd = InStr(intStart, strDatabaseConnection, ";") - 0
            '    Me.txtServer.Text = Mid(strDatabaseConnection, intStart, intEnd - intStart)
            'Else
            '    Me.txtUserName.Text = ""
            '    Me.txtPassword.Text = ""
            '    Me.txtDatabase.Text = ""
            '    Me.txtServer.Text = ""
            'End If
        Catch ex As Exception

        End Try
        If Me.txtUserName.Text = "" Then
            Me.txtUserName.Focus()
        ElseIf Me.txtPassword.Text = "" Then
            Me.txtPassword.Focus()
        ElseIf Me.txtDatabase.Text = "" Then
            Me.txtDatabase.Focus()
        ElseIf Me.txtServer.Text = "" Then
            Me.txtServer.Focus()
        End If

    End Sub

    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click
        frmDataTree.strUserName = Me.txtUserName.Text
        frmDataTree.strPassword = Me.txtPassword.Text
        frmDataTree.strDatabase = Me.txtDatabase.Text
        frmDataTree.strServer = Me.txtServer.Text

    End Sub

End Class
