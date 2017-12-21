Public Class frmDataGrid
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
    Friend WithEvents cmdMoveUp As System.Windows.Forms.Button
    Friend WithEvents cmdMoveDown As System.Windows.Forms.Button
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents cmdOK As System.Windows.Forms.Button
    Friend WithEvents lstColumns As System.Windows.Forms.CheckedListBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmDataGrid))
        Me.cmdMoveUp = New System.Windows.Forms.Button
        Me.cmdMoveDown = New System.Windows.Forms.Button
        Me.cmdCancel = New System.Windows.Forms.Button
        Me.cmdOK = New System.Windows.Forms.Button
        Me.lstColumns = New System.Windows.Forms.CheckedListBox
        Me.SuspendLayout()
        '
        'cmdMoveUp
        '
        Me.cmdMoveUp.Location = New System.Drawing.Point(250, 25)
        Me.cmdMoveUp.Name = "cmdMoveUp"
        Me.cmdMoveUp.TabIndex = 1
        Me.cmdMoveUp.Text = "Move Up"
        '
        'cmdMoveDown
        '
        Me.cmdMoveDown.Location = New System.Drawing.Point(250, 50)
        Me.cmdMoveDown.Name = "cmdMoveDown"
        Me.cmdMoveDown.TabIndex = 2
        Me.cmdMoveDown.Text = "Move Down"
        '
        'cmdCancel
        '
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Location = New System.Drawing.Point(250, 275)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.TabIndex = 24
        Me.cmdCancel.Text = "Cancel"
        '
        'cmdOK
        '
        Me.cmdOK.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.cmdOK.Location = New System.Drawing.Point(250, 250)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.TabIndex = 23
        Me.cmdOK.Text = "OK"
        '
        'lstColumns
        '
        Me.lstColumns.Location = New System.Drawing.Point(25, 25)
        Me.lstColumns.Name = "lstColumns"
        Me.lstColumns.Size = New System.Drawing.Size(200, 274)
        Me.lstColumns.TabIndex = 26
        '
        'frmDataGrid
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(342, 316)
        Me.Controls.Add(Me.lstColumns)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdOK)
        Me.Controls.Add(Me.cmdMoveDown)
        Me.Controls.Add(Me.cmdMoveUp)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmDataGrid"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Column Order"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub frmDataGrid_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim intCounter As Integer
        'Create Column List Box
        For intCounter = 0 To frmDataTree.myColumnList.Items.Count - 1
            Me.lstColumns.Items.Add( _
                frmDataTree.myColumnList.Items(intCounter), _
                frmDataTree.myColumnList.CheckedItems.Contains(frmDataTree.myColumnList.Items(intCounter)))
        Next
        'Set Data Tree Shared Column List
        frmDataTree.myColumnList = Me.lstColumns

    End Sub

    Private Sub cmdMoveUp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdMoveUp.Click
        Dim mySelectedIndex As Integer
        Dim myCheckState As CheckState
        Dim myItem As New Object
        'Get Selected Index
        mySelectedIndex = Me.lstColumns.SelectedIndex
        'Get Checked State
        If Me.lstColumns.CheckedItems.Contains(Me.lstColumns.SelectedItem) Then
            myCheckState = CheckState.Checked
        Else
            myCheckState = CheckState.Unchecked
        End If
        'Get Selected Item
        myItem = Me.lstColumns.SelectedItem
        'Move Selected Item
        If mySelectedIndex > 0 Then
            Me.lstColumns.Items.Remove(Me.lstColumns.SelectedItem)
            Me.lstColumns.Items.Insert(mySelectedIndex - 1, myItem)
            Me.lstColumns.SelectedIndex = mySelectedIndex - 1
            Me.lstColumns.SetItemCheckState(mySelectedIndex - 1, myCheckState)
        End If

    End Sub

    Private Sub cmdMoveDown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdMoveDown.Click
        Dim mySelectedIndex As Integer
        Dim myCheckState As CheckState
        Dim myItem As New Object
        'Get Selected Index
        mySelectedIndex = Me.lstColumns.SelectedIndex
        'Get Checked State
        If Me.lstColumns.CheckedItems.Contains(Me.lstColumns.SelectedItem) Then
            myCheckState = CheckState.Checked
        Else
            myCheckState = CheckState.Unchecked
        End If
        'Get Selected Item
        myItem = Me.lstColumns.SelectedItem
        'Move Selected Item
        If mySelectedIndex < Me.lstColumns.Items.Count - 1 Then
            Me.lstColumns.Items.Remove(Me.lstColumns.SelectedItem)
            Me.lstColumns.Items.Insert(mySelectedIndex + 1, myItem)
            Me.lstColumns.SelectedIndex = mySelectedIndex + 1
            Me.lstColumns.SetItemCheckState(mySelectedIndex + 1, myCheckState)
        End If

    End Sub

End Class
