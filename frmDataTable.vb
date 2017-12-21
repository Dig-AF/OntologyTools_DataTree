Public Class frmDatabase
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
    Friend WithEvents dbConnection As System.Data.OleDb.OleDbConnection
    Friend WithEvents cboChildTable As System.Windows.Forms.ComboBox
    Friend WithEvents lstChildField As System.Windows.Forms.ListBox
    Friend WithEvents cboParentTable As System.Windows.Forms.ComboBox
    Friend WithEvents lstParentField As System.Windows.Forms.ListBox
    Friend WithEvents txtChildID As System.Windows.Forms.TextBox
    Friend WithEvents txtChildName As System.Windows.Forms.TextBox
    Friend WithEvents txtChildDescription As System.Windows.Forms.TextBox
    Friend WithEvents txtParentChildID As System.Windows.Forms.TextBox
    Friend WithEvents txtParentParentID As System.Windows.Forms.TextBox
    Friend WithEvents cmdAddChildID As System.Windows.Forms.Button
    Friend WithEvents cmdAddChildName As System.Windows.Forms.Button
    Friend WithEvents cmdAddChildDescription As System.Windows.Forms.Button
    Friend WithEvents cmdAddParentChildID As System.Windows.Forms.Button
    Friend WithEvents cmdAddParentParentID As System.Windows.Forms.Button
    Friend WithEvents tblTable As System.Data.OleDb.OleDbDataAdapter
    Friend WithEvents lblChildID As System.Windows.Forms.Label
    Friend WithEvents lblChildName As System.Windows.Forms.Label
    Friend WithEvents lblChildDescription As System.Windows.Forms.Label
    Friend WithEvents lblParentChildID As System.Windows.Forms.Label
    Friend WithEvents lblParentParentID As System.Windows.Forms.Label
    Friend WithEvents lblChildTable As System.Windows.Forms.Label
    Friend WithEvents lblParentTable As System.Windows.Forms.Label
    Friend WithEvents cmdOK As System.Windows.Forms.Button
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents sqlSelect As System.Data.OleDb.OleDbCommand
    Friend WithEvents lblCriteriaEquals As System.Windows.Forms.Label
    Friend WithEvents txtCriteriaValue As System.Windows.Forms.TextBox
    Friend WithEvents cmdAddCriteria As System.Windows.Forms.Button
    Friend WithEvents lblCriteria As System.Windows.Forms.Label
    Friend WithEvents txtCriteria As System.Windows.Forms.TextBox
    Friend WithEvents lblCriteriaValue As System.Windows.Forms.Label
    Friend WithEvents cmdClearCriteria As System.Windows.Forms.Button
    Friend WithEvents lblParentAutoNumber As System.Windows.Forms.Label
    Friend WithEvents cmdAddParentAutoNumber As System.Windows.Forms.Button
    Friend WithEvents txtParentAutoNumber As System.Windows.Forms.TextBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents cmdDefault As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmDatabase))
        Me.tblTable = New System.Data.OleDb.OleDbDataAdapter
        Me.sqlSelect = New System.Data.OleDb.OleDbCommand
        Me.dbConnection = New System.Data.OleDb.OleDbConnection
        Me.cboChildTable = New System.Windows.Forms.ComboBox
        Me.lstChildField = New System.Windows.Forms.ListBox
        Me.cboParentTable = New System.Windows.Forms.ComboBox
        Me.lstParentField = New System.Windows.Forms.ListBox
        Me.txtChildID = New System.Windows.Forms.TextBox
        Me.txtChildName = New System.Windows.Forms.TextBox
        Me.txtChildDescription = New System.Windows.Forms.TextBox
        Me.txtParentChildID = New System.Windows.Forms.TextBox
        Me.txtParentParentID = New System.Windows.Forms.TextBox
        Me.cmdAddChildID = New System.Windows.Forms.Button
        Me.cmdAddChildName = New System.Windows.Forms.Button
        Me.cmdAddChildDescription = New System.Windows.Forms.Button
        Me.cmdAddParentChildID = New System.Windows.Forms.Button
        Me.cmdAddParentParentID = New System.Windows.Forms.Button
        Me.lblChildID = New System.Windows.Forms.Label
        Me.lblChildName = New System.Windows.Forms.Label
        Me.lblChildDescription = New System.Windows.Forms.Label
        Me.lblParentChildID = New System.Windows.Forms.Label
        Me.lblParentParentID = New System.Windows.Forms.Label
        Me.lblChildTable = New System.Windows.Forms.Label
        Me.lblParentTable = New System.Windows.Forms.Label
        Me.cmdOK = New System.Windows.Forms.Button
        Me.cmdCancel = New System.Windows.Forms.Button
        Me.lblCriteriaEquals = New System.Windows.Forms.Label
        Me.txtCriteriaValue = New System.Windows.Forms.TextBox
        Me.cmdAddCriteria = New System.Windows.Forms.Button
        Me.lblCriteria = New System.Windows.Forms.Label
        Me.txtCriteria = New System.Windows.Forms.TextBox
        Me.lblCriteriaValue = New System.Windows.Forms.Label
        Me.cmdClearCriteria = New System.Windows.Forms.Button
        Me.lblParentAutoNumber = New System.Windows.Forms.Label
        Me.cmdAddParentAutoNumber = New System.Windows.Forms.Button
        Me.txtParentAutoNumber = New System.Windows.Forms.TextBox
        Me.Button1 = New System.Windows.Forms.Button
        Me.cmdDefault = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'tblTable
        '
        Me.tblTable.SelectCommand = Me.sqlSelect
        Me.tblTable.TableMappings.AddRange(New System.Data.Common.DataTableMapping() {New System.Data.Common.DataTableMapping("Table", "Child", New System.Data.Common.DataColumnMapping() {New System.Data.Common.DataColumnMapping("DateCreated", "DateCreated"), New System.Data.Common.DataColumnMapping("DateModified", "DateModified"), New System.Data.Common.DataColumnMapping("Description", "Description"), New System.Data.Common.DataColumnMapping("Field1", "Field1"), New System.Data.Common.DataColumnMapping("Field2", "Field2"), New System.Data.Common.DataColumnMapping("Field3", "Field3"), New System.Data.Common.DataColumnMapping("Field4", "Field4"), New System.Data.Common.DataColumnMapping("Field5", "Field5"), New System.Data.Common.DataColumnMapping("Field6", "Field6"), New System.Data.Common.DataColumnMapping("Field7", "Field7"), New System.Data.Common.DataColumnMapping("ID", "ID"), New System.Data.Common.DataColumnMapping("Name", "Name")})})
        '
        'sqlSelect
        '
        Me.sqlSelect.Connection = Me.dbConnection
        '
        'cboChildTable
        '
        Me.cboChildTable.Location = New System.Drawing.Point(25, 25)
        Me.cboChildTable.Name = "cboChildTable"
        Me.cboChildTable.Size = New System.Drawing.Size(121, 21)
        Me.cboChildTable.Sorted = True
        Me.cboChildTable.TabIndex = 0
        '
        'lstChildField
        '
        Me.lstChildField.Location = New System.Drawing.Point(25, 50)
        Me.lstChildField.Name = "lstChildField"
        Me.lstChildField.Size = New System.Drawing.Size(120, 95)
        Me.lstChildField.TabIndex = 1
        '
        'cboParentTable
        '
        Me.cboParentTable.Location = New System.Drawing.Point(25, 175)
        Me.cboParentTable.Name = "cboParentTable"
        Me.cboParentTable.Size = New System.Drawing.Size(121, 21)
        Me.cboParentTable.Sorted = True
        Me.cboParentTable.TabIndex = 2
        '
        'lstParentField
        '
        Me.lstParentField.Location = New System.Drawing.Point(25, 200)
        Me.lstParentField.Name = "lstParentField"
        Me.lstParentField.Size = New System.Drawing.Size(120, 95)
        Me.lstParentField.TabIndex = 3
        '
        'txtChildID
        '
        Me.txtChildID.BackColor = System.Drawing.Color.White
        Me.txtChildID.Location = New System.Drawing.Point(225, 50)
        Me.txtChildID.Name = "txtChildID"
        Me.txtChildID.ReadOnly = True
        Me.txtChildID.Size = New System.Drawing.Size(100, 20)
        Me.txtChildID.TabIndex = 4
        '
        'txtChildName
        '
        Me.txtChildName.BackColor = System.Drawing.Color.White
        Me.txtChildName.Location = New System.Drawing.Point(225, 75)
        Me.txtChildName.Name = "txtChildName"
        Me.txtChildName.ReadOnly = True
        Me.txtChildName.Size = New System.Drawing.Size(100, 20)
        Me.txtChildName.TabIndex = 5
        '
        'txtChildDescription
        '
        Me.txtChildDescription.BackColor = System.Drawing.Color.White
        Me.txtChildDescription.Location = New System.Drawing.Point(225, 100)
        Me.txtChildDescription.Name = "txtChildDescription"
        Me.txtChildDescription.ReadOnly = True
        Me.txtChildDescription.Size = New System.Drawing.Size(100, 20)
        Me.txtChildDescription.TabIndex = 6
        '
        'txtParentChildID
        '
        Me.txtParentChildID.BackColor = System.Drawing.Color.White
        Me.txtParentChildID.Location = New System.Drawing.Point(225, 200)
        Me.txtParentChildID.Name = "txtParentChildID"
        Me.txtParentChildID.ReadOnly = True
        Me.txtParentChildID.Size = New System.Drawing.Size(100, 20)
        Me.txtParentChildID.TabIndex = 7
        '
        'txtParentParentID
        '
        Me.txtParentParentID.BackColor = System.Drawing.Color.White
        Me.txtParentParentID.Location = New System.Drawing.Point(225, 225)
        Me.txtParentParentID.Name = "txtParentParentID"
        Me.txtParentParentID.ReadOnly = True
        Me.txtParentParentID.Size = New System.Drawing.Size(100, 20)
        Me.txtParentParentID.TabIndex = 8
        '
        'cmdAddChildID
        '
        Me.cmdAddChildID.Location = New System.Drawing.Point(175, 50)
        Me.cmdAddChildID.Name = "cmdAddChildID"
        Me.cmdAddChildID.Size = New System.Drawing.Size(25, 23)
        Me.cmdAddChildID.TabIndex = 9
        Me.cmdAddChildID.Text = ">"
        '
        'cmdAddChildName
        '
        Me.cmdAddChildName.Location = New System.Drawing.Point(175, 75)
        Me.cmdAddChildName.Name = "cmdAddChildName"
        Me.cmdAddChildName.Size = New System.Drawing.Size(25, 23)
        Me.cmdAddChildName.TabIndex = 10
        Me.cmdAddChildName.Text = ">"
        '
        'cmdAddChildDescription
        '
        Me.cmdAddChildDescription.Location = New System.Drawing.Point(175, 100)
        Me.cmdAddChildDescription.Name = "cmdAddChildDescription"
        Me.cmdAddChildDescription.Size = New System.Drawing.Size(25, 23)
        Me.cmdAddChildDescription.TabIndex = 11
        Me.cmdAddChildDescription.Text = ">"
        '
        'cmdAddParentChildID
        '
        Me.cmdAddParentChildID.Location = New System.Drawing.Point(175, 200)
        Me.cmdAddParentChildID.Name = "cmdAddParentChildID"
        Me.cmdAddParentChildID.Size = New System.Drawing.Size(25, 23)
        Me.cmdAddParentChildID.TabIndex = 12
        Me.cmdAddParentChildID.Text = ">"
        '
        'cmdAddParentParentID
        '
        Me.cmdAddParentParentID.Location = New System.Drawing.Point(175, 225)
        Me.cmdAddParentParentID.Name = "cmdAddParentParentID"
        Me.cmdAddParentParentID.Size = New System.Drawing.Size(25, 23)
        Me.cmdAddParentParentID.TabIndex = 13
        Me.cmdAddParentParentID.Text = ">"
        '
        'lblChildID
        '
        Me.lblChildID.Location = New System.Drawing.Point(350, 50)
        Me.lblChildID.Name = "lblChildID"
        Me.lblChildID.Size = New System.Drawing.Size(100, 23)
        Me.lblChildID.TabIndex = 14
        Me.lblChildID.Text = "ID"
        Me.lblChildID.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblChildName
        '
        Me.lblChildName.Location = New System.Drawing.Point(350, 75)
        Me.lblChildName.Name = "lblChildName"
        Me.lblChildName.Size = New System.Drawing.Size(100, 23)
        Me.lblChildName.TabIndex = 15
        Me.lblChildName.Text = "Name"
        Me.lblChildName.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblChildDescription
        '
        Me.lblChildDescription.Location = New System.Drawing.Point(350, 100)
        Me.lblChildDescription.Name = "lblChildDescription"
        Me.lblChildDescription.Size = New System.Drawing.Size(100, 23)
        Me.lblChildDescription.TabIndex = 16
        Me.lblChildDescription.Text = "Description"
        Me.lblChildDescription.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblParentChildID
        '
        Me.lblParentChildID.Location = New System.Drawing.Point(350, 200)
        Me.lblParentChildID.Name = "lblParentChildID"
        Me.lblParentChildID.Size = New System.Drawing.Size(100, 23)
        Me.lblParentChildID.TabIndex = 17
        Me.lblParentChildID.Text = "Child ID"
        Me.lblParentChildID.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblParentParentID
        '
        Me.lblParentParentID.Location = New System.Drawing.Point(350, 225)
        Me.lblParentParentID.Name = "lblParentParentID"
        Me.lblParentParentID.Size = New System.Drawing.Size(100, 23)
        Me.lblParentParentID.TabIndex = 18
        Me.lblParentParentID.Text = "Parent ID"
        Me.lblParentParentID.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblChildTable
        '
        Me.lblChildTable.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblChildTable.Location = New System.Drawing.Point(225, 28)
        Me.lblChildTable.Name = "lblChildTable"
        Me.lblChildTable.Size = New System.Drawing.Size(200, 18)
        Me.lblChildTable.TabIndex = 19
        Me.lblChildTable.Text = "Child Table"
        '
        'lblParentTable
        '
        Me.lblParentTable.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblParentTable.Location = New System.Drawing.Point(225, 178)
        Me.lblParentTable.Name = "lblParentTable"
        Me.lblParentTable.Size = New System.Drawing.Size(200, 18)
        Me.lblParentTable.TabIndex = 20
        Me.lblParentTable.Text = "Parent Table"
        '
        'cmdOK
        '
        Me.cmdOK.Location = New System.Drawing.Point(225, 400)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.Size = New System.Drawing.Size(50, 23)
        Me.cmdOK.TabIndex = 21
        Me.cmdOK.Text = "OK"
        '
        'cmdCancel
        '
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Location = New System.Drawing.Point(275, 400)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(50, 23)
        Me.cmdCancel.TabIndex = 22
        Me.cmdCancel.Text = "Cancel"
        '
        'lblCriteriaEquals
        '
        Me.lblCriteriaEquals.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCriteriaEquals.Location = New System.Drawing.Point(175, 275)
        Me.lblCriteriaEquals.Name = "lblCriteriaEquals"
        Me.lblCriteriaEquals.Size = New System.Drawing.Size(25, 23)
        Me.lblCriteriaEquals.TabIndex = 29
        Me.lblCriteriaEquals.Text = "="
        Me.lblCriteriaEquals.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtCriteriaValue
        '
        Me.txtCriteriaValue.BackColor = System.Drawing.Color.White
        Me.txtCriteriaValue.Location = New System.Drawing.Point(225, 275)
        Me.txtCriteriaValue.Name = "txtCriteriaValue"
        Me.txtCriteriaValue.Size = New System.Drawing.Size(100, 20)
        Me.txtCriteriaValue.TabIndex = 30
        '
        'cmdAddCriteria
        '
        Me.cmdAddCriteria.Enabled = False
        Me.cmdAddCriteria.Location = New System.Drawing.Point(175, 300)
        Me.cmdAddCriteria.Name = "cmdAddCriteria"
        Me.cmdAddCriteria.Size = New System.Drawing.Size(25, 23)
        Me.cmdAddCriteria.TabIndex = 31
        Me.cmdAddCriteria.Text = ">"
        '
        'lblCriteria
        '
        Me.lblCriteria.Location = New System.Drawing.Point(25, 300)
        Me.lblCriteria.Name = "lblCriteria"
        Me.lblCriteria.Size = New System.Drawing.Size(100, 23)
        Me.lblCriteria.TabIndex = 32
        Me.lblCriteria.Text = "Criteria"
        Me.lblCriteria.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtCriteria
        '
        Me.txtCriteria.Location = New System.Drawing.Point(25, 325)
        Me.txtCriteria.Multiline = True
        Me.txtCriteria.Name = "txtCriteria"
        Me.txtCriteria.ReadOnly = True
        Me.txtCriteria.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtCriteria.Size = New System.Drawing.Size(300, 70)
        Me.txtCriteria.TabIndex = 33
        '
        'lblCriteriaValue
        '
        Me.lblCriteriaValue.Location = New System.Drawing.Point(350, 275)
        Me.lblCriteriaValue.Name = "lblCriteriaValue"
        Me.lblCriteriaValue.Size = New System.Drawing.Size(100, 23)
        Me.lblCriteriaValue.TabIndex = 34
        Me.lblCriteriaValue.Text = "Criteria Value"
        Me.lblCriteriaValue.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'cmdClearCriteria
        '
        Me.cmdClearCriteria.Location = New System.Drawing.Point(350, 325)
        Me.cmdClearCriteria.Name = "cmdClearCriteria"
        Me.cmdClearCriteria.Size = New System.Drawing.Size(75, 23)
        Me.cmdClearCriteria.TabIndex = 35
        Me.cmdClearCriteria.Text = "Clear"
        '
        'lblParentAutoNumber
        '
        Me.lblParentAutoNumber.Location = New System.Drawing.Point(350, 250)
        Me.lblParentAutoNumber.Name = "lblParentAutoNumber"
        Me.lblParentAutoNumber.Size = New System.Drawing.Size(125, 23)
        Me.lblParentAutoNumber.TabIndex = 38
        Me.lblParentAutoNumber.Text = "AutoNumber (Optional)"
        Me.lblParentAutoNumber.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'cmdAddParentAutoNumber
        '
        Me.cmdAddParentAutoNumber.Location = New System.Drawing.Point(175, 250)
        Me.cmdAddParentAutoNumber.Name = "cmdAddParentAutoNumber"
        Me.cmdAddParentAutoNumber.Size = New System.Drawing.Size(25, 23)
        Me.cmdAddParentAutoNumber.TabIndex = 37
        Me.cmdAddParentAutoNumber.Text = ">"
        '
        'txtParentAutoNumber
        '
        Me.txtParentAutoNumber.BackColor = System.Drawing.Color.White
        Me.txtParentAutoNumber.Location = New System.Drawing.Point(225, 250)
        Me.txtParentAutoNumber.Name = "txtParentAutoNumber"
        Me.txtParentAutoNumber.ReadOnly = True
        Me.txtParentAutoNumber.Size = New System.Drawing.Size(100, 20)
        Me.txtParentAutoNumber.TabIndex = 36
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(150, 250)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(25, 23)
        Me.Button1.TabIndex = 39
        Me.Button1.Text = "<"
        '
        'cmdDefault
        '
        Me.cmdDefault.Location = New System.Drawing.Point(350, 125)
        Me.cmdDefault.Name = "cmdDefault"
        Me.cmdDefault.Size = New System.Drawing.Size(75, 23)
        Me.cmdDefault.TabIndex = 40
        Me.cmdDefault.Text = "Defaults"
        '
        'frmDatabase
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(467, 441)
        Me.Controls.Add(Me.cmdDefault)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.lblParentAutoNumber)
        Me.Controls.Add(Me.cmdAddParentAutoNumber)
        Me.Controls.Add(Me.txtParentAutoNumber)
        Me.Controls.Add(Me.cmdClearCriteria)
        Me.Controls.Add(Me.lblCriteriaValue)
        Me.Controls.Add(Me.txtCriteria)
        Me.Controls.Add(Me.lblCriteria)
        Me.Controls.Add(Me.cmdAddCriteria)
        Me.Controls.Add(Me.txtCriteriaValue)
        Me.Controls.Add(Me.lblCriteriaEquals)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdOK)
        Me.Controls.Add(Me.lblParentTable)
        Me.Controls.Add(Me.lblChildTable)
        Me.Controls.Add(Me.lblParentParentID)
        Me.Controls.Add(Me.lblParentChildID)
        Me.Controls.Add(Me.lblChildDescription)
        Me.Controls.Add(Me.lblChildName)
        Me.Controls.Add(Me.lblChildID)
        Me.Controls.Add(Me.cmdAddParentParentID)
        Me.Controls.Add(Me.cmdAddParentChildID)
        Me.Controls.Add(Me.cmdAddChildDescription)
        Me.Controls.Add(Me.cmdAddChildName)
        Me.Controls.Add(Me.cmdAddChildID)
        Me.Controls.Add(Me.txtParentParentID)
        Me.Controls.Add(Me.txtParentChildID)
        Me.Controls.Add(Me.txtChildDescription)
        Me.Controls.Add(Me.txtChildName)
        Me.Controls.Add(Me.txtChildID)
        Me.Controls.Add(Me.lstParentField)
        Me.Controls.Add(Me.cboParentTable)
        Me.Controls.Add(Me.lstChildField)
        Me.Controls.Add(Me.cboChildTable)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmDatabase"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Custom Table Structure"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

    Dim myCriteriaWhere As String
    Dim myCriteriaParentTableWhere As String
    Dim myCriteriaInsert As String
    Dim myCriteriaValue As String
    Dim myCriteriaSet As String

    Private Sub frmDatabase_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim myDatabase As DataTable
        Dim myTables As DataRow

        Me.dbConnection.ConnectionString = frmDataTree.myConnectString
        'Tables
        myDatabase = funGetTables(Me.dbConnection)
        If Not myDatabase Is Nothing Then
            For Each myTables In myDatabase.Rows
                Me.cboChildTable.Items.Add(myTables.Item(2))
                Me.cboParentTable.Items.Add(myTables.Item(2))
                If myTables.Item(2) = frmDataTree.myChildTable Then
                    Me.cboChildTable.SelectedIndex = Me.cboChildTable.FindString(frmDataTree.myChildTable)
                End If
                If myTables.Item(2) = frmDataTree.myParentTable Then
                    Me.cboParentTable.SelectedIndex = Me.cboChildTable.FindString(frmDataTree.myParentTable)
                End If
            Next
        End If
        'Querys
        Try
            myDatabase = funGetQueries(Me.dbConnection)
            If Not myDatabase Is Nothing Then
                For Each myTables In myDatabase.Rows
                    Me.cboChildTable.Items.Add(myTables.Item(2))
                    Me.cboParentTable.Items.Add(myTables.Item(2))
                    If myTables.Item(2) = frmDataTree.myChildTable Then
                        Me.cboChildTable.SelectedIndex = Me.cboChildTable.FindString(frmDataTree.myChildTable)
                    End If
                    If myTables.Item(2) = frmDataTree.myParentTable Then
                        Me.cboParentTable.SelectedIndex = Me.cboChildTable.FindString(frmDataTree.myParentTable)
                    End If
                Next
            End If
        Catch ex As Exception

        End Try
        Me.myCriteriaWhere = frmDataTree.myCriteriaWhere
        Me.myCriteriaInsert = frmDataTree.myCriteriaInsert
        Me.myCriteriaValue = frmDataTree.myCriteriaValue
        Me.myCriteriaSet = frmDataTree.myCriteriaSet
        Me.myCriteriaParentTableWhere = frmDataTree.myCriteriaParentTableWhere
        Me.txtCriteria.Text = Me.myCriteriaWhere
        Me.txtParentAutoNumber.Text = frmDataTree.myParentTableAutoNumber

    End Sub

    Private Sub cboChildTable_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboChildTable.SelectedIndexChanged
        Dim myTable As DataTable
        Dim myFields As DataRow
        Dim dsTable As New DataSet

        myTable = Me.funGetFields(Me.tblTable, dsTable, Me.cboChildTable.Text)

        Me.lstChildField.Items.Clear()
        Me.txtChildID.Clear()
        Me.txtChildName.Clear()
        Me.txtChildDescription.Clear()
        For Each myFields In myTable.Rows
            Me.lstChildField.Items.Add(myFields.Item(0))
            If myFields.Item(0) = frmDataTree.myChildTableChildID Then
                Me.txtChildID.Text = frmDataTree.myChildTableChildID
            ElseIf myFields.Item(0) = frmDataTree.myChildTableChildName Then
                Me.txtChildName.Text = frmDataTree.myChildTableChildName
            ElseIf myFields.Item(0) = frmDataTree.myChildTableChildDescription Then
                Me.txtChildDescription.Text = frmDataTree.myChildTableChildDescription
            End If
        Next
        Try
            If Me.txtChildID.Text = "" Then
                Me.txtChildID.Text = Me.lstChildField.Items(0)
            End If
            If Me.txtChildName.Text = "" Then
                Me.txtChildName.Text = Me.lstChildField.Items(1)
            End If
            If Me.txtChildDescription.Text = "" Then
                Me.txtChildDescription.Text = Me.lstChildField.Items(2)
            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub cboParentTable_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboParentTable.SelectedIndexChanged
        Dim myTable As DataTable
        Dim myFields As DataRow
        Dim dsTable As New DataSet

        myTable = Me.funGetFields(Me.tblTable, dsTable, Me.cboParentTable.Text)

        Me.lstParentField.Items.Clear()
        Me.txtParentChildID.Clear()
        Me.txtParentParentID.Clear()
        For Each myFields In myTable.Rows
            Me.lstParentField.Items.Add(myFields.Item(0))
            If myFields.Item(0) = frmDataTree.myParentTableChildID Then
                Me.txtParentChildID.Text = frmDataTree.myParentTableChildID
            ElseIf myFields.Item(0) = frmDataTree.myParentTableParentID Then
                Me.txtParentParentID.Text = frmDataTree.myParentTableParentID
            End If
        Next
        Try
            If Me.txtParentChildID.Text = "" Then
                Me.txtParentChildID.Text = Me.lstParentField.Items(0)
            End If
            If Me.txtParentParentID.Text = "" Then
                Me.txtParentParentID.Text = Me.lstParentField.Items(1)
            End If

        Catch ex As Exception

        End Try

    End Sub

    Public Function funGetTables(ByVal myConnection As System.Data.OleDb.OleDbConnection) As DataTable
        Try
            myConnection.Open()
            Dim mySchemaTable As DataTable = _
                myConnection.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, _
                                        New Object() {Nothing, Nothing, Nothing, "TABLE"})
            myConnection.Close()
            Return mySchemaTable
        Catch ex As Exception
            Return Nothing
        End Try

    End Function

    Public Function funGetQueries(ByVal myConnection As System.Data.OleDb.OleDbConnection) As DataTable
        Try
            myConnection.Open()
            Dim mySchemaTable As DataTable = _
                myConnection.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Views, Nothing)
            Return mySchemaTable
        Catch ex As Exception
        Finally
            myConnection.Close()
        End Try

    End Function

    Public Function funGetFields(ByVal myAdapter As System.Data.OleDb.OleDbDataAdapter, ByVal myDataSet As DataSet, ByVal myTableName As String) As DataTable
        ' Create a new DataTable.
        Dim mySchemaField As DataTable = New DataTable("Fields")
        Dim myColumn As DataColumn
        Dim myRow As DataRow
        Dim myTable As DataTable
        Dim myField As DataColumn

        ' Create first column and add to the DataTable.
        myColumn = New DataColumn
        myColumn.DataType = System.Type.GetType("System.String")
        myColumn.ColumnName = "Field"
        mySchemaField.Columns.Add(myColumn)

        'Create Data Adapter Table Mappings
        myAdapter.TableMappings.Clear()
        myAdapter.TableMappings.AddRange( _
            New System.Data.Common.DataTableMapping() { _
            New System.Data.Common.DataTableMapping("Table", myTableName)})
        'Set Data Adapter Column Mappings Automatically
        myAdapter.MissingMappingAction = MissingMappingAction.Passthrough
        myAdapter.SelectCommand.CommandText = _
            " SELECT *" & _
            " FROM [" & myTableName & "]"
        Try
            myAdapter.Fill(myDataSet)

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
        myTable = myDataSet.Tables(myTableName)

        'Add Field Names
        For Each myField In myTable.Columns
            myRow = mySchemaField.NewRow()
            myRow("Field") = myField.ColumnName
            mySchemaField.Rows.Add(myRow)
        Next

        Return mySchemaField

    End Function

    Private Sub cmdAddChildID_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddChildID.Click
        Me.txtChildID.Text = Me.lstChildField.SelectedItem
    End Sub

    Private Sub cmdAddChildName_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddChildName.Click
        Me.txtChildName.Text = Me.lstChildField.SelectedItem
    End Sub

    Private Sub cmdAddChildDescription_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddChildDescription.Click
        Me.txtChildDescription.Text = Me.lstChildField.SelectedItem
    End Sub

    Private Sub cmdAddParentChildID_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddParentChildID.Click
        Me.txtParentChildID.Text = Me.lstParentField.SelectedItem
    End Sub

    Private Sub cmdAddParentParentID_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddParentParentID.Click
        Me.txtParentParentID.Text = Me.lstParentField.SelectedItem
    End Sub

    Private Sub cmdOK_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles cmdOK.MouseDown
        If _
            Me.cboChildTable.Text = "" Or _
            Me.txtChildID.Text = "" Or _
            Me.txtChildName.Text = "" Or _
            Me.txtChildDescription.Text = "" Or _
            Me.cboParentTable.Text = "" Or _
            Me.txtParentChildID.Text = "" Or _
            Me.txtParentParentID.Text = "" _
        Then
            MsgBox("All Fields on the right are required.")
            Me.cmdOK.DialogResult = DialogResult.None
        Else
            'Table Structure
            frmDataTree.myChildTable = Me.cboChildTable.Text
            frmDataTree.myChildTableChildID = Me.txtChildID.Text
            frmDataTree.myChildTableChildName = Me.txtChildName.Text
            frmDataTree.myChildTableChildDescription = Me.txtChildDescription.Text
            frmDataTree.myParentTable = Me.cboParentTable.Text
            frmDataTree.myParentTableChildID = Me.txtParentChildID.Text
            frmDataTree.myParentTableParentID = Me.txtParentParentID.Text
            frmDataTree.myParentTableAutoNumber = Me.txtParentAutoNumber.Text
            'Type Code
            frmDataTree.myCriteriaWhere = Me.myCriteriaWhere
            frmDataTree.myCriteriaParentTableWhere = Me.myCriteriaParentTableWhere
            frmDataTree.myCriteriaInsert = Me.myCriteriaInsert
            frmDataTree.myCriteriaValue = Me.myCriteriaValue
            frmDataTree.myCriteriaSet = Me.myCriteriaSet
            'frmDataTree.myParentTableTypeCode = Me.txtParentTypeCode.Text
            'frmDataTree.myParentTableTypeCodeValue = Me.txtParentTypeCodeValue.Text
            'OK
            Me.cmdOK.DialogResult = DialogResult.OK
        End If

    End Sub

    Private Sub cmdAddCriteria_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddCriteria.Click
        If Me.txtCriteriaValue.Text = "" Then
            MsgBox("Please enter a Criteria Value.")
            Me.txtCriteriaValue.Focus()
            Exit Sub
        End If
        If Me.myCriteriaWhere = "" Then
            Me.myCriteriaWhere = "[" & Me.cboParentTable.Text & "].[" & Me.lstParentField.SelectedItem & "] = " & Me.funConvertDataType(Me.txtCriteriaValue.Text, Me.lstParentField.SelectedItem)
            Me.myCriteriaParentTableWhere = "[ParentTable].[" & Me.lstParentField.SelectedItem & "] = " & Me.funConvertDataType(Me.txtCriteriaValue.Text, Me.lstParentField.SelectedItem)
            Me.myCriteriaInsert = "[" & Me.lstParentField.SelectedItem & "]"
            Me.myCriteriaValue = Me.funConvertDataType(Me.txtCriteriaValue.Text, Me.lstParentField.SelectedItem)
            Me.myCriteriaSet = "[" & Me.cboParentTable.Text & "].[" & Me.lstParentField.SelectedItem & "] = " & Me.funConvertDataType(Me.txtCriteriaValue.Text, Me.lstParentField.SelectedItem)
        Else
            Me.myCriteriaWhere = Me.myCriteriaWhere & " AND " & _
                "[" & Me.cboParentTable.Text & "].[" & Me.lstParentField.SelectedItem & "] = " & Me.funConvertDataType(Me.txtCriteriaValue.Text, Me.lstParentField.SelectedItem)
            Me.myCriteriaParentTableWhere = Me.myCriteriaParentTableWhere & " AND " & _
                "[ParentTable].[" & Me.lstParentField.SelectedItem & "] = " & Me.funConvertDataType(Me.txtCriteriaValue.Text, Me.lstParentField.SelectedItem)
            Me.myCriteriaInsert = Me.myCriteriaInsert & ", " & _
                "[" & Me.lstParentField.SelectedItem & "]"
            Me.myCriteriaValue = Me.myCriteriaValue & ", " & _
                Me.funConvertDataType(Me.txtCriteriaValue.Text, Me.lstParentField.SelectedItem)
            Me.myCriteriaSet = Me.myCriteriaSet & ", " & _
                "[" & Me.cboParentTable.Text & "].[" & Me.lstParentField.SelectedItem & "] = " & Me.funConvertDataType(Me.txtCriteriaValue.Text, Me.lstParentField.SelectedItem)
        End If
        Me.txtCriteria.Text = Me.myCriteriaWhere

    End Sub

    Private Sub lstParentField_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstParentField.SelectedIndexChanged
        If Me.lstParentField.SelectedItem Is Nothing Or Me.lstParentField.SelectedItem = Me.txtParentChildID.Text Or Me.lstParentField.SelectedItem = Me.txtParentParentID.Text Then
            Me.cmdAddCriteria.Enabled = False
        Else
            Me.cmdAddCriteria.Enabled = True
        End If
        Me.txtCriteriaValue.Focus()

    End Sub

    Private Function funConvertDataType(ByVal myValue As String, ByVal myField As String) As String
        Dim myDataSet As New DataSet
        'Dim myTable As DataTable

        'Create Data Adapter Table Mappings
        Me.tblTable.TableMappings.Clear()
        Me.tblTable.TableMappings.AddRange( _
            New System.Data.Common.DataTableMapping() { _
            New System.Data.Common.DataTableMapping("Table", Me.cboParentTable.Text)})
        'Set Data Adapter Column Mappings Automatically
        Me.tblTable.MissingMappingAction = MissingMappingAction.Passthrough
        Me.tblTable.SelectCommand.CommandText = _
            " SELECT " & myField & _
            " FROM [" & Me.cboParentTable.Text & "]"
        Try
            Me.tblTable.Fill(myDataSet)

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
        If myDataSet.Tables(Me.cboParentTable.Text).Columns(0).DataType Is System.Type.GetType("System.String") Then
            'String
            funConvertDataType = "'" & Me.funReplaceSingleQuote(myValue) & "'"
        ElseIf myDataSet.Tables(Me.cboParentTable.Text).Columns(0).DataType Is System.Type.GetType("System.DateTime") Then
            'Date
            funConvertDataType = "#" & myValue & "#"
        Else
            'Number
            funConvertDataType = myValue
        End If

    End Function

    Private Function funReplaceSingleQuote(ByVal myString As String) As String
        Dim intCharacter As Integer = 1
        Dim intLength As Integer = 1

        If Not myString Is Nothing Then
            If InStr(myString, "'") = 0 Then
                funReplaceSingleQuote = myString
                Exit Function
            End If
            intLength = myString.Length
            While intCharacter <= intLength
                If Mid(myString, intCharacter, 1) = "'" Then
                    'Add a Single Quote to every Single Quote ''
                    myString = Mid(myString, 1, intCharacter) & "'" & Mid(myString, intCharacter + 1)
                    intCharacter = intCharacter + 1
                    intLength = intLength + 1
                End If
                intCharacter = intCharacter + 1
            End While
            funReplaceSingleQuote = myString
        End If

    End Function

    Private Sub txtCriteriaValue_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCriteriaValue.KeyPress
        If e.KeyChar = Microsoft.VisualBasic.ChrW(13) Then
            Me.cmdAddCriteria_Click(Nothing, Nothing)
            e.Handled = True
        End If
    End Sub

    Private Sub cmdClearCriteria_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClearCriteria.Click
        Me.txtCriteria.Text = ""
        Me.myCriteriaWhere = ""
        Me.myCriteriaInsert = ""
        Me.myCriteriaValue = ""
        Me.myCriteriaSet = ""
        Me.myCriteriaParentTableWhere = ""
    End Sub

    Private Sub cmdAddParentAutoNumber_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddParentAutoNumber.Click
        Me.txtParentAutoNumber.Text = Me.lstParentField.SelectedItem

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.txtParentAutoNumber.Text = ""

    End Sub

End Class
