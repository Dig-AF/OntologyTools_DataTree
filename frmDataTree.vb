'**************************************************
' Data Tree
'
' Developed By:     John Grindle
' Developed For:    Silver Bullet Solutions, Inc.
' Version:          Data Tree Version 1.61
' Date:             July 16, 2012
'
'**************************************************

Public Class frmDataTree
    Inherits System.Windows.Forms.Form

#Region "Windows Form Designer"

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
    'Data
    Friend WithEvents tblChild As System.Data.OleDb.OleDbDataAdapter
    Friend WithEvents dsChild As DataTree.dsChild
    Friend WithEvents dsTree As DataTree.dsTree
    Friend WithEvents dsNode As DataTree.dsNode
    'File Browser
    Friend WithEvents filDatabase As System.Windows.Forms.OpenFileDialog
    'Data Tree
    Friend WithEvents lblSelected As System.Windows.Forms.Label
    Friend WithEvents treDataTree As System.Windows.Forms.TreeView
    Friend WithEvents grdChildren As System.Windows.Forms.DataGrid
    Friend WithEvents txtDescription As System.Windows.Forms.TextBox
    'Menu Bar
    Friend WithEvents mnuDataTree As System.Windows.Forms.MainMenu
    Friend WithEvents mnuFile As System.Windows.Forms.MenuItem
    Friend WithEvents mnuOpenDatabase As System.Windows.Forms.MenuItem
    Friend WithEvents mnuCreateTree As System.Windows.Forms.MenuItem
    Friend WithEvents mnuExit As System.Windows.Forms.MenuItem
    Friend WithEvents mnuEdit As System.Windows.Forms.MenuItem
    Friend WithEvents mnuUndo As System.Windows.Forms.MenuItem
    Friend WithEvents mnuRedo As System.Windows.Forms.MenuItem
    Friend WithEvents mnuCut As System.Windows.Forms.MenuItem
    Friend WithEvents mnuCopy As System.Windows.Forms.MenuItem
    Friend WithEvents mnuPaste As System.Windows.Forms.MenuItem
    Friend WithEvents mnuNode As System.Windows.Forms.MenuItem
    Friend WithEvents mnuNewNode As System.Windows.Forms.MenuItem
    Friend WithEvents mnuAddNode As System.Windows.Forms.MenuItem
    Friend WithEvents mnuEditNode As System.Windows.Forms.MenuItem
    Friend WithEvents mnuDeleteNode As System.Windows.Forms.MenuItem
    Friend WithEvents mnuHelp As System.Windows.Forms.MenuItem
    Friend WithEvents mnuContents As System.Windows.Forms.MenuItem
    Friend WithEvents mnuAbout As System.Windows.Forms.MenuItem
    Friend WithEvents mnuAccess As System.Windows.Forms.MenuItem
    Friend WithEvents mnuPasteToRoot As System.Windows.Forms.MenuItem
    Friend WithEvents mnuSQLServer As System.Windows.Forms.MenuItem
    Friend WithEvents mnuOracle As System.Windows.Forms.MenuItem
    Friend WithEvents mnuODBC As System.Windows.Forms.MenuItem
    Friend WithEvents mnuDatabaseCustom As System.Windows.Forms.MenuItem
    Friend WithEvents sqlSelectCommand As System.Data.OleDb.OleDbCommand
    Friend WithEvents sqlInsertCommand As System.Data.OleDb.OleDbCommand
    Friend WithEvents sqlUpdateCommand As System.Data.OleDb.OleDbCommand
    Friend WithEvents sqlDeleteCommand As System.Data.OleDb.OleDbCommand
    Friend WithEvents imgNode As System.Windows.Forms.ImageList
    Friend WithEvents mnuTreeView As System.Windows.Forms.ContextMenu
    Friend WithEvents mnuDataGridCustomize As System.Windows.Forms.MenuItem
    Friend WithEvents mnuDataGridPaste As System.Windows.Forms.MenuItem
    Friend WithEvents mnuDataGridAddNode As System.Windows.Forms.MenuItem
    Friend WithEvents mnuDataGridProperties As System.Windows.Forms.MenuItem
    Friend WithEvents mnuTreeViewExpand As System.Windows.Forms.MenuItem
    Friend WithEvents mnuTreeViewCut As System.Windows.Forms.MenuItem
    Friend WithEvents mnuTreeViewCopy As System.Windows.Forms.MenuItem
    Friend WithEvents mnuTreeViewPaste As System.Windows.Forms.MenuItem
    Friend WithEvents mnuTreeViewPasteToRoot As System.Windows.Forms.MenuItem
    Friend WithEvents mnuTreeViewNewNode As System.Windows.Forms.MenuItem
    Friend WithEvents mnuTreeViewEditNode As System.Windows.Forms.MenuItem
    Friend WithEvents mnuTreeViewAddNode As System.Windows.Forms.MenuItem
    Friend WithEvents mnuTreeViewDeleteNode As System.Windows.Forms.MenuItem
    Friend WithEvents mnuTreeViewProperties As System.Windows.Forms.MenuItem
    Friend WithEvents mnuFind As System.Windows.Forms.MenuItem
    Friend WithEvents mnuTreeViewFind As System.Windows.Forms.MenuItem
    Friend WithEvents mnuRefresh As System.Windows.Forms.MenuItem
    Friend WithEvents mnuTreeViewSeparator1 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuTreeViewSeparator2 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuTreeViewSeparator3 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuDataGridSeparator1 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuDataGridSeparator2 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuFileSeparator1 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuFileSeparator2 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuFileSeparator3 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuEditSeparator1 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuEditSeparator2 As System.Windows.Forms.MenuItem
    Friend WithEvents dbConnection As System.Data.OleDb.OleDbConnection
    Public WithEvents mnuDataGrid As System.Windows.Forms.ContextMenu
    Friend WithEvents mnuDuplicate As System.Windows.Forms.MenuItem
    Friend WithEvents mnuTreeViewDuplicate As System.Windows.Forms.MenuItem
    Friend WithEvents mnuOpen As System.Windows.Forms.MenuItem
    Friend WithEvents mnuSaveAs As System.Windows.Forms.MenuItem
    Friend WithEvents mnuSave As System.Windows.Forms.MenuItem
    Friend WithEvents mnuFileSeparator4 As System.Windows.Forms.MenuItem
    Friend WithEvents filSaveConfigurationFile As System.Windows.Forms.SaveFileDialog
    Friend WithEvents filOpenConfigurationFile As System.Windows.Forms.OpenFileDialog
    Friend WithEvents mnuFileSeparator5 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuClose As System.Windows.Forms.MenuItem
    Friend WithEvents mnuNew As System.Windows.Forms.MenuItem
    Friend WithEvents mnuDataGridColumnWidth As System.Windows.Forms.MenuItem
    Friend WithEvents mnuDataGridRowHeight As System.Windows.Forms.MenuItem
    Friend WithEvents mnuMySQL As System.Windows.Forms.MenuItem
    Friend WithEvents mnuExportOrganizationChart As System.Windows.Forms.MenuItem
    Friend WithEvents mnuExportIndentedHierarchy As System.Windows.Forms.MenuItem
    Friend WithEvents filExport As System.Windows.Forms.SaveFileDialog
    Friend WithEvents mnuExportOrganizationChartText As System.Windows.Forms.MenuItem
    Friend WithEvents mnuExportIndentedHierarchyText As System.Windows.Forms.MenuItem
    Friend WithEvents filExportText As System.Windows.Forms.SaveFileDialog
    Friend WithEvents mnuExport As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuTableCustom As System.Windows.Forms.MenuItem
    Friend WithEvents mnuRemoveNode As System.Windows.Forms.MenuItem
    Friend WithEvents mnuTreeViewRemoveNode As System.Windows.Forms.MenuItem
    Friend WithEvents mnuWordWrap As System.Windows.Forms.MenuItem

    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmDataTree))
        Me.dbConnection = New System.Data.OleDb.OleDbConnection()
        Me.tblChild = New System.Data.OleDb.OleDbDataAdapter()
        Me.sqlDeleteCommand = New System.Data.OleDb.OleDbCommand()
        Me.sqlInsertCommand = New System.Data.OleDb.OleDbCommand()
        Me.sqlSelectCommand = New System.Data.OleDb.OleDbCommand()
        Me.sqlUpdateCommand = New System.Data.OleDb.OleDbCommand()
        Me.dsChild = New DataTree.dsChild()
        Me.dsTree = New DataTree.dsTree()
        Me.dsNode = New DataTree.dsNode()
        Me.filDatabase = New System.Windows.Forms.OpenFileDialog()
        Me.lblSelected = New System.Windows.Forms.Label()
        Me.treDataTree = New System.Windows.Forms.TreeView()
        Me.mnuTreeView = New System.Windows.Forms.ContextMenu()
        Me.mnuTreeViewExpand = New System.Windows.Forms.MenuItem()
        Me.mnuTreeViewFind = New System.Windows.Forms.MenuItem()
        Me.mnuTreeViewSeparator1 = New System.Windows.Forms.MenuItem()
        Me.mnuTreeViewCut = New System.Windows.Forms.MenuItem()
        Me.mnuTreeViewCopy = New System.Windows.Forms.MenuItem()
        Me.mnuTreeViewPaste = New System.Windows.Forms.MenuItem()
        Me.mnuTreeViewDuplicate = New System.Windows.Forms.MenuItem()
        Me.mnuTreeViewPasteToRoot = New System.Windows.Forms.MenuItem()
        Me.mnuTreeViewSeparator2 = New System.Windows.Forms.MenuItem()
        Me.mnuTreeViewNewNode = New System.Windows.Forms.MenuItem()
        Me.mnuTreeViewEditNode = New System.Windows.Forms.MenuItem()
        Me.mnuTreeViewAddNode = New System.Windows.Forms.MenuItem()
        Me.mnuTreeViewRemoveNode = New System.Windows.Forms.MenuItem()
        Me.mnuTreeViewDeleteNode = New System.Windows.Forms.MenuItem()
        Me.mnuTreeViewSeparator3 = New System.Windows.Forms.MenuItem()
        Me.mnuTreeViewProperties = New System.Windows.Forms.MenuItem()
        Me.grdChildren = New System.Windows.Forms.DataGrid()
        Me.mnuDataGrid = New System.Windows.Forms.ContextMenu()
        Me.mnuDataGridRowHeight = New System.Windows.Forms.MenuItem()
        Me.mnuDataGridColumnWidth = New System.Windows.Forms.MenuItem()
        Me.mnuDataGridCustomize = New System.Windows.Forms.MenuItem()
        Me.mnuDataGridSeparator1 = New System.Windows.Forms.MenuItem()
        Me.mnuDataGridPaste = New System.Windows.Forms.MenuItem()
        Me.mnuDataGridAddNode = New System.Windows.Forms.MenuItem()
        Me.mnuDataGridSeparator2 = New System.Windows.Forms.MenuItem()
        Me.mnuDataGridProperties = New System.Windows.Forms.MenuItem()
        Me.txtDescription = New System.Windows.Forms.TextBox()
        Me.mnuDataTree = New System.Windows.Forms.MainMenu(Me.components)
        Me.mnuFile = New System.Windows.Forms.MenuItem()
        Me.mnuNew = New System.Windows.Forms.MenuItem()
        Me.mnuOpen = New System.Windows.Forms.MenuItem()
        Me.mnuClose = New System.Windows.Forms.MenuItem()
        Me.mnuFileSeparator1 = New System.Windows.Forms.MenuItem()
        Me.mnuOpenDatabase = New System.Windows.Forms.MenuItem()
        Me.mnuAccess = New System.Windows.Forms.MenuItem()
        Me.mnuSQLServer = New System.Windows.Forms.MenuItem()
        Me.mnuMySQL = New System.Windows.Forms.MenuItem()
        Me.mnuOracle = New System.Windows.Forms.MenuItem()
        Me.mnuODBC = New System.Windows.Forms.MenuItem()
        Me.mnuDatabaseCustom = New System.Windows.Forms.MenuItem()
        Me.mnuFileSeparator2 = New System.Windows.Forms.MenuItem()
        Me.mnuSave = New System.Windows.Forms.MenuItem()
        Me.mnuSaveAs = New System.Windows.Forms.MenuItem()
        Me.mnuExport = New System.Windows.Forms.MenuItem()
        Me.mnuExportOrganizationChart = New System.Windows.Forms.MenuItem()
        Me.mnuExportIndentedHierarchy = New System.Windows.Forms.MenuItem()
        Me.MenuItem1 = New System.Windows.Forms.MenuItem()
        Me.mnuExportOrganizationChartText = New System.Windows.Forms.MenuItem()
        Me.mnuExportIndentedHierarchyText = New System.Windows.Forms.MenuItem()
        Me.mnuFileSeparator3 = New System.Windows.Forms.MenuItem()
        Me.mnuTableCustom = New System.Windows.Forms.MenuItem()
        Me.mnuFileSeparator4 = New System.Windows.Forms.MenuItem()
        Me.mnuRefresh = New System.Windows.Forms.MenuItem()
        Me.mnuCreateTree = New System.Windows.Forms.MenuItem()
        Me.mnuFileSeparator5 = New System.Windows.Forms.MenuItem()
        Me.mnuExit = New System.Windows.Forms.MenuItem()
        Me.mnuEdit = New System.Windows.Forms.MenuItem()
        Me.mnuUndo = New System.Windows.Forms.MenuItem()
        Me.mnuRedo = New System.Windows.Forms.MenuItem()
        Me.mnuEditSeparator1 = New System.Windows.Forms.MenuItem()
        Me.mnuCut = New System.Windows.Forms.MenuItem()
        Me.mnuCopy = New System.Windows.Forms.MenuItem()
        Me.mnuPaste = New System.Windows.Forms.MenuItem()
        Me.mnuDuplicate = New System.Windows.Forms.MenuItem()
        Me.mnuPasteToRoot = New System.Windows.Forms.MenuItem()
        Me.mnuEditSeparator2 = New System.Windows.Forms.MenuItem()
        Me.mnuFind = New System.Windows.Forms.MenuItem()
        Me.mnuWordWrap = New System.Windows.Forms.MenuItem()
        Me.mnuNode = New System.Windows.Forms.MenuItem()
        Me.mnuNewNode = New System.Windows.Forms.MenuItem()
        Me.mnuEditNode = New System.Windows.Forms.MenuItem()
        Me.mnuAddNode = New System.Windows.Forms.MenuItem()
        Me.mnuRemoveNode = New System.Windows.Forms.MenuItem()
        Me.mnuDeleteNode = New System.Windows.Forms.MenuItem()
        Me.mnuHelp = New System.Windows.Forms.MenuItem()
        Me.mnuContents = New System.Windows.Forms.MenuItem()
        Me.mnuAbout = New System.Windows.Forms.MenuItem()
        Me.imgNode = New System.Windows.Forms.ImageList(Me.components)
        Me.filSaveConfigurationFile = New System.Windows.Forms.SaveFileDialog()
        Me.filOpenConfigurationFile = New System.Windows.Forms.OpenFileDialog()
        Me.filExport = New System.Windows.Forms.SaveFileDialog()
        Me.filExportText = New System.Windows.Forms.SaveFileDialog()
        CType(Me.dsChild, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dsTree, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dsNode, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grdChildren, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'tblChild
        '
        Me.tblChild.DeleteCommand = Me.sqlDeleteCommand
        Me.tblChild.InsertCommand = Me.sqlInsertCommand
        Me.tblChild.SelectCommand = Me.sqlSelectCommand
        Me.tblChild.UpdateCommand = Me.sqlUpdateCommand
        '
        'sqlDeleteCommand
        '
        Me.sqlDeleteCommand.Connection = Me.dbConnection
        '
        'sqlInsertCommand
        '
        Me.sqlInsertCommand.Connection = Me.dbConnection
        '
        'sqlSelectCommand
        '
        Me.sqlSelectCommand.Connection = Me.dbConnection
        '
        'sqlUpdateCommand
        '
        Me.sqlUpdateCommand.Connection = Me.dbConnection
        '
        'dsChild
        '
        Me.dsChild.DataSetName = "dsChild"
        Me.dsChild.Locale = New System.Globalization.CultureInfo("en-US")
        Me.dsChild.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'dsTree
        '
        Me.dsTree.DataSetName = "dsTree"
        Me.dsTree.Locale = New System.Globalization.CultureInfo("en-US")
        Me.dsTree.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'dsNode
        '
        Me.dsNode.DataSetName = "dsNode"
        Me.dsNode.Locale = New System.Globalization.CultureInfo("en-US")
        Me.dsNode.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'filDatabase
        '
        Me.filDatabase.DefaultExt = "mdb"
        Me.filDatabase.Filter = "Micorsoft Office Access|*.mdb"
        Me.filDatabase.Title = "Open Database Connection"
        '
        'lblSelected
        '
        Me.lblSelected.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblSelected.Location = New System.Drawing.Point(0, 0)
        Me.lblSelected.Name = "lblSelected"
        Me.lblSelected.Size = New System.Drawing.Size(600, 25)
        Me.lblSelected.TabIndex = 2
        Me.lblSelected.Text = "Data Tree"
        Me.lblSelected.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'treDataTree
        '
        Me.treDataTree.AllowDrop = True
        Me.treDataTree.ContextMenu = Me.mnuTreeView
        Me.treDataTree.Cursor = System.Windows.Forms.Cursors.Arrow
        Me.treDataTree.Indent = 20
        Me.treDataTree.ItemHeight = 14
        Me.treDataTree.LabelEdit = True
        Me.treDataTree.Location = New System.Drawing.Point(0, 25)
        Me.treDataTree.Name = "treDataTree"
        Me.treDataTree.Size = New System.Drawing.Size(172, 347)
        Me.treDataTree.TabIndex = 1
        '
        'mnuTreeView
        '
        Me.mnuTreeView.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuTreeViewExpand, Me.mnuTreeViewFind, Me.mnuTreeViewSeparator1, Me.mnuTreeViewCut, Me.mnuTreeViewCopy, Me.mnuTreeViewPaste, Me.mnuTreeViewDuplicate, Me.mnuTreeViewPasteToRoot, Me.mnuTreeViewSeparator2, Me.mnuTreeViewNewNode, Me.mnuTreeViewEditNode, Me.mnuTreeViewAddNode, Me.mnuTreeViewRemoveNode, Me.mnuTreeViewDeleteNode, Me.mnuTreeViewSeparator3, Me.mnuTreeViewProperties})
        '
        'mnuTreeViewExpand
        '
        Me.mnuTreeViewExpand.Index = 0
        Me.mnuTreeViewExpand.Text = "Expand / Collapse"
        Me.mnuTreeViewExpand.Visible = False
        '
        'mnuTreeViewFind
        '
        Me.mnuTreeViewFind.Enabled = False
        Me.mnuTreeViewFind.Index = 1
        Me.mnuTreeViewFind.Text = "Find"
        '
        'mnuTreeViewSeparator1
        '
        Me.mnuTreeViewSeparator1.Enabled = False
        Me.mnuTreeViewSeparator1.Index = 2
        Me.mnuTreeViewSeparator1.Text = "-"
        '
        'mnuTreeViewCut
        '
        Me.mnuTreeViewCut.Enabled = False
        Me.mnuTreeViewCut.Index = 3
        Me.mnuTreeViewCut.Text = "Cut"
        '
        'mnuTreeViewCopy
        '
        Me.mnuTreeViewCopy.Enabled = False
        Me.mnuTreeViewCopy.Index = 4
        Me.mnuTreeViewCopy.Text = "Copy"
        '
        'mnuTreeViewPaste
        '
        Me.mnuTreeViewPaste.Enabled = False
        Me.mnuTreeViewPaste.Index = 5
        Me.mnuTreeViewPaste.Text = "Paste"
        '
        'mnuTreeViewDuplicate
        '
        Me.mnuTreeViewDuplicate.Enabled = False
        Me.mnuTreeViewDuplicate.Index = 6
        Me.mnuTreeViewDuplicate.Text = "Paste As Duplicate"
        '
        'mnuTreeViewPasteToRoot
        '
        Me.mnuTreeViewPasteToRoot.Enabled = False
        Me.mnuTreeViewPasteToRoot.Index = 7
        Me.mnuTreeViewPasteToRoot.Text = "Paste to Root"
        '
        'mnuTreeViewSeparator2
        '
        Me.mnuTreeViewSeparator2.Enabled = False
        Me.mnuTreeViewSeparator2.Index = 8
        Me.mnuTreeViewSeparator2.Text = "-"
        '
        'mnuTreeViewNewNode
        '
        Me.mnuTreeViewNewNode.Enabled = False
        Me.mnuTreeViewNewNode.Index = 9
        Me.mnuTreeViewNewNode.Text = "New Node"
        '
        'mnuTreeViewEditNode
        '
        Me.mnuTreeViewEditNode.Enabled = False
        Me.mnuTreeViewEditNode.Index = 10
        Me.mnuTreeViewEditNode.Text = "Edit Node"
        '
        'mnuTreeViewAddNode
        '
        Me.mnuTreeViewAddNode.Enabled = False
        Me.mnuTreeViewAddNode.Index = 11
        Me.mnuTreeViewAddNode.Text = "Add Node"
        '
        'mnuTreeViewRemoveNode
        '
        Me.mnuTreeViewRemoveNode.Enabled = False
        Me.mnuTreeViewRemoveNode.Index = 12
        Me.mnuTreeViewRemoveNode.Text = "Remove Node"
        '
        'mnuTreeViewDeleteNode
        '
        Me.mnuTreeViewDeleteNode.Enabled = False
        Me.mnuTreeViewDeleteNode.Index = 13
        Me.mnuTreeViewDeleteNode.Text = "Delete Node"
        '
        'mnuTreeViewSeparator3
        '
        Me.mnuTreeViewSeparator3.Index = 14
        Me.mnuTreeViewSeparator3.Text = "-"
        Me.mnuTreeViewSeparator3.Visible = False
        '
        'mnuTreeViewProperties
        '
        Me.mnuTreeViewProperties.Index = 15
        Me.mnuTreeViewProperties.Text = "Properties"
        Me.mnuTreeViewProperties.Visible = False
        '
        'grdChildren
        '
        Me.grdChildren.AllowNavigation = False
        Me.grdChildren.BackgroundColor = System.Drawing.SystemColors.Window
        Me.grdChildren.CaptionVisible = False
        Me.grdChildren.ContextMenu = Me.mnuDataGrid
        Me.grdChildren.DataMember = ""
        Me.grdChildren.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grdChildren.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.grdChildren.Location = New System.Drawing.Point(175, 25)
        Me.grdChildren.Name = "grdChildren"
        Me.grdChildren.RowHeaderWidth = 30
        Me.grdChildren.Size = New System.Drawing.Size(425, 347)
        Me.grdChildren.TabIndex = 5
        '
        'mnuDataGrid
        '
        Me.mnuDataGrid.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuDataGridRowHeight, Me.mnuDataGridColumnWidth, Me.mnuDataGridCustomize, Me.mnuDataGridSeparator1, Me.mnuDataGridPaste, Me.mnuDataGridAddNode, Me.mnuDataGridSeparator2, Me.mnuDataGridProperties})
        '
        'mnuDataGridRowHeight
        '
        Me.mnuDataGridRowHeight.Index = 0
        Me.mnuDataGridRowHeight.Text = "Row Height"
        '
        'mnuDataGridColumnWidth
        '
        Me.mnuDataGridColumnWidth.Index = 1
        Me.mnuDataGridColumnWidth.Text = "Column Width"
        '
        'mnuDataGridCustomize
        '
        Me.mnuDataGridCustomize.Index = 2
        Me.mnuDataGridCustomize.Text = "Customize"
        '
        'mnuDataGridSeparator1
        '
        Me.mnuDataGridSeparator1.Index = 3
        Me.mnuDataGridSeparator1.Text = "-"
        '
        'mnuDataGridPaste
        '
        Me.mnuDataGridPaste.Index = 4
        Me.mnuDataGridPaste.Text = "Paste"
        Me.mnuDataGridPaste.Visible = False
        '
        'mnuDataGridAddNode
        '
        Me.mnuDataGridAddNode.Index = 5
        Me.mnuDataGridAddNode.Text = "Add Node"
        Me.mnuDataGridAddNode.Visible = False
        '
        'mnuDataGridSeparator2
        '
        Me.mnuDataGridSeparator2.Index = 6
        Me.mnuDataGridSeparator2.Text = "-"
        Me.mnuDataGridSeparator2.Visible = False
        '
        'mnuDataGridProperties
        '
        Me.mnuDataGridProperties.Index = 7
        Me.mnuDataGridProperties.Text = "Properties"
        Me.mnuDataGridProperties.Visible = False
        '
        'txtDescription
        '
        Me.txtDescription.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.txtDescription.Location = New System.Drawing.Point(0, 375)
        Me.txtDescription.Multiline = True
        Me.txtDescription.Name = "txtDescription"
        Me.txtDescription.ReadOnly = True
        Me.txtDescription.Size = New System.Drawing.Size(600, 25)
        Me.txtDescription.TabIndex = 4
        '
        'mnuDataTree
        '
        Me.mnuDataTree.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuFile, Me.mnuEdit, Me.mnuNode, Me.mnuHelp})
        '
        'mnuFile
        '
        Me.mnuFile.Index = 0
        Me.mnuFile.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuNew, Me.mnuOpen, Me.mnuClose, Me.mnuFileSeparator1, Me.mnuOpenDatabase, Me.mnuFileSeparator2, Me.mnuSave, Me.mnuSaveAs, Me.mnuExport, Me.mnuFileSeparator3, Me.mnuTableCustom, Me.mnuFileSeparator4, Me.mnuRefresh, Me.mnuCreateTree, Me.mnuFileSeparator5, Me.mnuExit})
        Me.mnuFile.Text = "&File"
        '
        'mnuNew
        '
        Me.mnuNew.Index = 0
        Me.mnuNew.Text = "&New..."
        '
        'mnuOpen
        '
        Me.mnuOpen.Index = 1
        Me.mnuOpen.Shortcut = System.Windows.Forms.Shortcut.CtrlO
        Me.mnuOpen.Text = "&Open..."
        '
        'mnuClose
        '
        Me.mnuClose.Enabled = False
        Me.mnuClose.Index = 2
        Me.mnuClose.Text = "&Close"
        '
        'mnuFileSeparator1
        '
        Me.mnuFileSeparator1.Index = 3
        Me.mnuFileSeparator1.Text = "-"
        Me.mnuFileSeparator1.Visible = False
        '
        'mnuOpenDatabase
        '
        Me.mnuOpenDatabase.Index = 4
        Me.mnuOpenDatabase.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuAccess, Me.mnuSQLServer, Me.mnuMySQL, Me.mnuOracle, Me.mnuODBC, Me.mnuDatabaseCustom})
        Me.mnuOpenDatabase.Text = "Open &Database Connection"
        '
        'mnuAccess
        '
        Me.mnuAccess.Index = 0
        Me.mnuAccess.Text = "Microsoft Access"
        '
        'mnuSQLServer
        '
        Me.mnuSQLServer.Index = 1
        Me.mnuSQLServer.Text = "SQL Server"
        '
        'mnuMySQL
        '
        Me.mnuMySQL.Enabled = False
        Me.mnuMySQL.Index = 2
        Me.mnuMySQL.Text = "MySQL"
        '
        'mnuOracle
        '
        Me.mnuOracle.Enabled = False
        Me.mnuOracle.Index = 3
        Me.mnuOracle.Text = "Oracle"
        '
        'mnuODBC
        '
        Me.mnuODBC.Enabled = False
        Me.mnuODBC.Index = 4
        Me.mnuODBC.Text = "ODBC"
        '
        'mnuDatabaseCustom
        '
        Me.mnuDatabaseCustom.Enabled = False
        Me.mnuDatabaseCustom.Index = 5
        Me.mnuDatabaseCustom.Text = "Custom ..."
        '
        'mnuFileSeparator2
        '
        Me.mnuFileSeparator2.Index = 5
        Me.mnuFileSeparator2.Text = "-"
        '
        'mnuSave
        '
        Me.mnuSave.Enabled = False
        Me.mnuSave.Index = 6
        Me.mnuSave.Shortcut = System.Windows.Forms.Shortcut.CtrlS
        Me.mnuSave.Text = "&Save"
        '
        'mnuSaveAs
        '
        Me.mnuSaveAs.Enabled = False
        Me.mnuSaveAs.Index = 7
        Me.mnuSaveAs.Text = "Save &As..."
        '
        'mnuExport
        '
        Me.mnuExport.Enabled = False
        Me.mnuExport.Index = 8
        Me.mnuExport.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuExportOrganizationChart, Me.mnuExportIndentedHierarchy, Me.MenuItem1, Me.mnuExportOrganizationChartText, Me.mnuExportIndentedHierarchyText})
        Me.mnuExport.Text = "Export"
        '
        'mnuExportOrganizationChart
        '
        Me.mnuExportOrganizationChart.Enabled = False
        Me.mnuExportOrganizationChart.Index = 0
        Me.mnuExportOrganizationChart.Text = "Organization Chart (Visio .xls)"
        Me.mnuExportOrganizationChart.Visible = False
        '
        'mnuExportIndentedHierarchy
        '
        Me.mnuExportIndentedHierarchy.Enabled = False
        Me.mnuExportIndentedHierarchy.Index = 1
        Me.mnuExportIndentedHierarchy.Text = "Indented Hierarchy (Excel .xls)"
        Me.mnuExportIndentedHierarchy.Visible = False
        '
        'MenuItem1
        '
        Me.MenuItem1.Enabled = False
        Me.MenuItem1.Index = 2
        Me.MenuItem1.Text = "-"
        Me.MenuItem1.Visible = False
        '
        'mnuExportOrganizationChartText
        '
        Me.mnuExportOrganizationChartText.Enabled = False
        Me.mnuExportOrganizationChartText.Index = 3
        Me.mnuExportOrganizationChartText.Text = "Organization Chart (Visio .csv)"
        '
        'mnuExportIndentedHierarchyText
        '
        Me.mnuExportIndentedHierarchyText.Index = 4
        Me.mnuExportIndentedHierarchyText.Text = "Indented Hierarchy (Excel .csv)"
        '
        'mnuFileSeparator3
        '
        Me.mnuFileSeparator3.Index = 9
        Me.mnuFileSeparator3.Text = "-"
        '
        'mnuTableCustom
        '
        Me.mnuTableCustom.Enabled = False
        Me.mnuTableCustom.Index = 10
        Me.mnuTableCustom.Shortcut = System.Windows.Forms.Shortcut.CtrlT
        Me.mnuTableCustom.Text = "&Table Structure ..."
        '
        'mnuFileSeparator4
        '
        Me.mnuFileSeparator4.Index = 11
        Me.mnuFileSeparator4.Text = "-"
        '
        'mnuRefresh
        '
        Me.mnuRefresh.Enabled = False
        Me.mnuRefresh.Index = 12
        Me.mnuRefresh.Shortcut = System.Windows.Forms.Shortcut.F5
        Me.mnuRefresh.Text = "&Refresh"
        '
        'mnuCreateTree
        '
        Me.mnuCreateTree.Enabled = False
        Me.mnuCreateTree.Index = 13
        Me.mnuCreateTree.Shortcut = System.Windows.Forms.Shortcut.F6
        Me.mnuCreateTree.Text = "&Create Tree"
        '
        'mnuFileSeparator5
        '
        Me.mnuFileSeparator5.Index = 14
        Me.mnuFileSeparator5.Text = "-"
        '
        'mnuExit
        '
        Me.mnuExit.Index = 15
        Me.mnuExit.Text = "E&xit"
        '
        'mnuEdit
        '
        Me.mnuEdit.Index = 1
        Me.mnuEdit.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuUndo, Me.mnuRedo, Me.mnuEditSeparator1, Me.mnuCut, Me.mnuCopy, Me.mnuPaste, Me.mnuDuplicate, Me.mnuPasteToRoot, Me.mnuEditSeparator2, Me.mnuFind, Me.mnuWordWrap})
        Me.mnuEdit.Text = "&Edit"
        '
        'mnuUndo
        '
        Me.mnuUndo.Enabled = False
        Me.mnuUndo.Index = 0
        Me.mnuUndo.Shortcut = System.Windows.Forms.Shortcut.CtrlZ
        Me.mnuUndo.Text = "&Undo"
        Me.mnuUndo.Visible = False
        '
        'mnuRedo
        '
        Me.mnuRedo.Enabled = False
        Me.mnuRedo.Index = 1
        Me.mnuRedo.Shortcut = System.Windows.Forms.Shortcut.CtrlY
        Me.mnuRedo.Text = "&Redo"
        Me.mnuRedo.Visible = False
        '
        'mnuEditSeparator1
        '
        Me.mnuEditSeparator1.Enabled = False
        Me.mnuEditSeparator1.Index = 2
        Me.mnuEditSeparator1.Text = "-"
        Me.mnuEditSeparator1.Visible = False
        '
        'mnuCut
        '
        Me.mnuCut.Enabled = False
        Me.mnuCut.Index = 3
        Me.mnuCut.Shortcut = System.Windows.Forms.Shortcut.CtrlX
        Me.mnuCut.Text = "Cu&t"
        '
        'mnuCopy
        '
        Me.mnuCopy.Enabled = False
        Me.mnuCopy.Index = 4
        Me.mnuCopy.Shortcut = System.Windows.Forms.Shortcut.CtrlC
        Me.mnuCopy.Text = "&Copy"
        '
        'mnuPaste
        '
        Me.mnuPaste.Enabled = False
        Me.mnuPaste.Index = 5
        Me.mnuPaste.Shortcut = System.Windows.Forms.Shortcut.CtrlV
        Me.mnuPaste.Text = "&Paste"
        '
        'mnuDuplicate
        '
        Me.mnuDuplicate.Enabled = False
        Me.mnuDuplicate.Index = 6
        Me.mnuDuplicate.Shortcut = System.Windows.Forms.Shortcut.CtrlD
        Me.mnuDuplicate.Text = "Paste As &Duplicate"
        '
        'mnuPasteToRoot
        '
        Me.mnuPasteToRoot.Enabled = False
        Me.mnuPasteToRoot.Index = 7
        Me.mnuPasteToRoot.Shortcut = System.Windows.Forms.Shortcut.CtrlR
        Me.mnuPasteToRoot.Text = "Paste To &Root"
        '
        'mnuEditSeparator2
        '
        Me.mnuEditSeparator2.Index = 8
        Me.mnuEditSeparator2.Text = "-"
        '
        'mnuFind
        '
        Me.mnuFind.Enabled = False
        Me.mnuFind.Index = 9
        Me.mnuFind.Shortcut = System.Windows.Forms.Shortcut.CtrlF
        Me.mnuFind.Text = "&Find"
        '
        'mnuWordWrap
        '
        Me.mnuWordWrap.Checked = True
        Me.mnuWordWrap.Index = 10
        Me.mnuWordWrap.Text = "Word Wrap"
        '
        'mnuNode
        '
        Me.mnuNode.Index = 2
        Me.mnuNode.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuNewNode, Me.mnuEditNode, Me.mnuAddNode, Me.mnuRemoveNode, Me.mnuDeleteNode})
        Me.mnuNode.Text = "&Node"
        '
        'mnuNewNode
        '
        Me.mnuNewNode.Enabled = False
        Me.mnuNewNode.Index = 0
        Me.mnuNewNode.Shortcut = System.Windows.Forms.Shortcut.CtrlN
        Me.mnuNewNode.Text = "&New Node"
        '
        'mnuEditNode
        '
        Me.mnuEditNode.Enabled = False
        Me.mnuEditNode.Index = 1
        Me.mnuEditNode.Shortcut = System.Windows.Forms.Shortcut.CtrlE
        Me.mnuEditNode.Text = "&Edit Node"
        '
        'mnuAddNode
        '
        Me.mnuAddNode.Enabled = False
        Me.mnuAddNode.Index = 2
        Me.mnuAddNode.Shortcut = System.Windows.Forms.Shortcut.Ins
        Me.mnuAddNode.Text = "&Add Node"
        '
        'mnuRemoveNode
        '
        Me.mnuRemoveNode.Enabled = False
        Me.mnuRemoveNode.Index = 3
        Me.mnuRemoveNode.Shortcut = System.Windows.Forms.Shortcut.CtrlDel
        Me.mnuRemoveNode.Text = "&Remove Node"
        '
        'mnuDeleteNode
        '
        Me.mnuDeleteNode.Enabled = False
        Me.mnuDeleteNode.Index = 4
        Me.mnuDeleteNode.Text = "&Delete Node"
        '
        'mnuHelp
        '
        Me.mnuHelp.Index = 3
        Me.mnuHelp.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuContents, Me.mnuAbout})
        Me.mnuHelp.Text = "&Help"
        '
        'mnuContents
        '
        Me.mnuContents.Index = 0
        Me.mnuContents.Shortcut = System.Windows.Forms.Shortcut.F1
        Me.mnuContents.Text = "&Contents"
        '
        'mnuAbout
        '
        Me.mnuAbout.Index = 1
        Me.mnuAbout.Text = "&About Data Tree"
        '
        'imgNode
        '
        Me.imgNode.ImageStream = CType(resources.GetObject("imgNode.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.imgNode.TransparentColor = System.Drawing.Color.Transparent
        Me.imgNode.Images.SetKeyName(0, "")
        Me.imgNode.Images.SetKeyName(1, "")
        Me.imgNode.Images.SetKeyName(2, "")
        '
        'filSaveConfigurationFile
        '
        Me.filSaveConfigurationFile.DefaultExt = "dtc"
        Me.filSaveConfigurationFile.Filter = "Data Tree Configuration (*.dtc)|*.dtc"
        Me.filSaveConfigurationFile.Title = "Save Configuration File"
        '
        'filOpenConfigurationFile
        '
        Me.filOpenConfigurationFile.DefaultExt = "dtc"
        Me.filOpenConfigurationFile.Filter = "Data Tree Configuration (*.dtc)|*.dtc"
        Me.filOpenConfigurationFile.Title = "Open Configuration File"
        '
        'filExport
        '
        Me.filExport.DefaultExt = "xls"
        Me.filExport.Filter = "Microsoft Office Excel Workbook (*.xls) |*.xls"
        Me.filExport.Title = "Export ..."
        '
        'filExportText
        '
        Me.filExportText.DefaultExt = "csv"
        Me.filExportText.Filter = "CSV (Comma delimited) (*.csv) |*.csv"
        Me.filExportText.Title = "Export ..."
        '
        'frmDataTree
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(592, 400)
        Me.Controls.Add(Me.treDataTree)
        Me.Controls.Add(Me.lblSelected)
        Me.Controls.Add(Me.txtDescription)
        Me.Controls.Add(Me.grdChildren)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Menu = Me.mnuDataTree
        Me.MinimumSize = New System.Drawing.Size(125, 125)
        Me.Name = "frmDataTree"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Data Tree"
        CType(Me.dsChild, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dsTree, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dsNode, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grdChildren, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

#Region "Global Variables"
    'Server
    Public Shared strUserName As String
    Public Shared strPassword As String
    Public Shared strDatabase As String
    Public Shared strServer As String
    'Default Table Structure
    'Dim strDatabase As String
    Dim myConfigurationFile As String
    Dim myExportFile As String
    Public Shared myConnectString As String
    Public Shared myChildTable As String = "Child"
    Public Shared myParentTable As String = "Parent"
    Public Shared myChildTableChildID As String = "ID"
    Public Shared myChildTableChildName As String = "Name"
    Public Shared myChildTableChildDescription As String = "Description"
    Public Shared myParentTableChildID As String = "ID"
    Public Shared myParentTableParentID As String = "ParentID"
    Public Shared myCriteriaWhere As String
    Public Shared myCriteriaParentTableWhere As String
    Public Shared myCriteriaInsert As String
    Public Shared myCriteriaValue As String
    Public Shared myCriteriaSet As String
    Public Shared myParentTableAutoNumber As String = ""
    Dim intChildIDColumn As Integer
    'SQL Statements
    Dim sqlSelect As String
    Dim sqlFrom As String
    Dim sqlWhere As String
    Dim sqlOrder As String
    Dim sqlUpdate As String
    Dim sqlSet As String
    Dim sqlInsert As String
    Dim sqlValues As String
    Dim sqlDelete As String
    'Nodes
    Dim mySelectedNode As TreeNode
    Dim myFoundNode As TreeNode
    Dim myCutNode As TreeNode
    Dim myCopyNode As TreeNode
    Dim bolIsCollapsing As Boolean
    'Drag and Drop
    Dim myDragNode As New TreeNode
    Dim mySourceNode As New TreeNode
    Dim myDestinationNode As New TreeNode
    Dim myHoverNode As New TreeNode
    Dim myHoverTime As DateTime
    Dim myStartX As Integer
    Dim myStartY As Integer
    'Dim myCursor As System.Windows.Forms.Cursor
    'DataGrid
    Public Shared myColumnList As New CheckedListBox
    Dim myDataTable As DataTable
    Dim myDataView As DataView
    Dim intRowHeight As Integer
    Dim intColumnNumber As Integer
    Dim bolMouseOverColumnHeader As Boolean
    Dim bolMouseOverRowHeader As Boolean
    'Configuration File
    Dim bolSaveConfigurationFile As Boolean
    'Export File
    Dim intRow As Integer
    'Constants
    Const cintFormWidthOffset As Integer = 10
    Const cintFormHeightOffset As Integer = 54
    Const cintMenuHeightOffset As Integer = 25
    Const cintMinimumWidth As Integer = 50
    Const cintMinimumHeight As Integer = 53

#End Region

#Region "Menu Bar"

    'File
    Private Sub mnuNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuNew.Click
        Dim myNewDataTree As New frmDataTree

        myNewDataTree.Show()
        myNewDataTree.Top = Me.Top + 25
        myNewDataTree.Left = Me.Left + 25

    End Sub

    Private Sub mnuAccess_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAccess.Click
        'Open Database Connection
        Try
            If Me.funGetAccessDatabase() Then
                Me.funEnableFileMenuItems(False)
                Me.funCreateComponents()
                Me.bolSaveConfigurationFile = True
            Else
                'Reset
                strDatabase = ""
                Me.funResetTree()
                Me.funResetGlobals()
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub mnuSQLServer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSQLServer.Click
        'Open Database Connection
        Try
            Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
            If Me.funGetSQLDatabase() Then
                Me.funEnableFileMenuItems(False)
                Me.funCreateComponents()
                Me.bolSaveConfigurationFile = True
            Else
                'Reset
                strDatabase = ""
                Me.funResetTree()
                Me.funResetGlobals()
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub mnuMySQL_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuMySQL.Click
        'Open Database Connection
        If Me.funGetMySQLDatabase() Then
            Me.funEnableFileMenuItems(False)
            Me.funCreateComponents()
            Me.bolSaveConfigurationFile = True
        End If

    End Sub

    Private Sub mnuExportOrganizationChart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuExportOrganizationChart.Click
        Dim myResponse As DialogResult

        'Browse for Access Database
        Me.filExport.FileName = ""
        Me.filExport.InitialDirectory = Application.StartupPath
        myResponse = Me.filExport.ShowDialog()
        myExportFile = Me.filExport.FileName

        '***** REMOVED 08/19/08 *****
        'If myResponse = DialogResult.OK Then
        '    Try
        '        'Create Excel Workbook
        '        Dim xlApp As Excel.Application
        '        Dim xlBook As Excel.Workbook
        '        Dim xlSheet As Excel.Worksheet
        '        Dim xlSheets As Excel.Sheets
        '        xlApp = CType(CreateObject("Excel.Application"), Excel.Application)
        '        Try
        '            xlBook = CType(xlApp.Workbooks.Open(myExportFile), Excel.Workbook)
        '        Catch ex As Exception
        '            xlBook = CType(xlApp.Workbooks.Add, Excel.Workbook)
        '            xlBook.SaveAs(myExportFile)
        '        End Try
        '        xlSheet = CType(xlBook.Worksheets(1), Excel.Worksheet)
        '        xlSheet.Cells.Clear()
        '        'Minimize Excel Workbook
        '        xlApp.WindowState = Excel.XlWindowState.xlMinimized
        '        'Create Excel Sheet
        '        Me.funExportOrganizationChart(xlSheet)
        '        'Save Excel Workbook
        '        xlBook.Save()
        '        If MsgBox("The Excel file has been created.  Would you like to import into Visio now?  (Click No to view the Excel file.)", MsgBoxStyle.YesNo, "Export to Excel") = MsgBoxResult.No Then
        '            'Show Excel Workbook
        '            xlApp.Application.Visible = True
        '            xlBook.Application.Visible = True
        '            xlSheet.Application.Visible = True
        '            xlApp.WindowState = Excel.XlWindowState.xlMaximized
        '        Else
        '            Try
        '                'Quit Excel
        '                xlApp.Application.Quit()
        '                'Create Visio Document
        '                Dim viApp As Microsoft.Office.Interop.Visio.Application
        '                Dim viDoc As Microsoft.Office.Interop.Visio.Document
        '                viApp = CType(CreateObject("Visio.Application"), Microsoft.Office.Interop.Visio.Application)
        '                viDoc = CType(viApp.Documents.AddEx("orgch_u.vst", 0, 0), Microsoft.Office.Interop.Visio.Document)
        '                'viDoc = CType(viApp.Documents.Open(myExportFile), Microsoft.Office.Interop.Visio.Document)
        '                'Minimize Visio Document
        '                viApp.Window.WindowState = Microsoft.Office.Interop.Visio.VisWindowStates.visWSMinimized
        '                MsgBox("The data has been exported to:" & vbCr & myExportFile.ToCharArray & "." & vbCr & vbCr & _
        '                    "Click Organization Chart > Import Organization Data ...", _
        '                    MsgBoxStyle.OKOnly, "Import into Visio")
        '                'Show Visio Document
        '                viApp.Application.Visible = True
        '                viDoc.Application.Visible = True
        '                viApp.Window.WindowState = Microsoft.Office.Interop.Visio.VisWindowStates.visWSMaximized
        '                'Save Visio Document
        '                'viDoc.SaveAs(myExportFile)
        '                'viApp.Application.Quit()
        '                viApp = Nothing
        '            Catch ex As Exception
        '                MsgBox("Unable to Import into Visio.  Please open Visio and import it manually.")
        '            End Try
        '        End If
        '        xlApp = Nothing
        '    Catch ex As Exception
        '        MsgBox("Error exporting Excel file." & vbCr & ex.Message)
        '    End Try
        'End If

    End Sub

    Private Sub mnuExportIndentedHierarchy_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuExportIndentedHierarchy.Click
        Dim myResponse As DialogResult

        'Browse for Access Database
        Me.filExport.FileName = ""
        Me.filExport.InitialDirectory = Application.StartupPath
        myResponse = Me.filExport.ShowDialog()
        myExportFile = Me.filExport.FileName

        '***** REMOVED 08/19/08 *****
        'If myResponse = DialogResult.OK Then
        '    Try
        '        'Create Excel Workbook
        '        Dim xlApp As Excel.Application
        '        Dim xlBook As Excel.Workbook
        '        Dim xlSheet As Excel.Worksheet
        '        xlApp = CType(CreateObject("Excel.Application"), Excel.Application)
        '        Try
        '            xlBook = CType(xlApp.Workbooks.Open(myExportFile), Excel.Workbook)
        '        Catch ex As Exception
        '            xlBook = CType(xlApp.Workbooks.Add, Excel.Workbook)
        '            xlBook.SaveAs(myExportFile)
        '        End Try
        '        xlSheet = CType(xlBook.Worksheets(1), Excel.Worksheet)
        '        xlSheet.Cells.Clear()
        '        'Minimize Excel Workbook
        '        xlApp.WindowState = Excel.XlWindowState.xlMinimized
        '        'Create Excel Sheet
        '        Me.funExportIndentedHierarchy(xlSheet)
        '        'Save Excel Workbook
        '        xlBook.Save()
        '        If MsgBox("The Excel file has been created.  Would you like to view it now?", MsgBoxStyle.YesNo, "Export to Excel") = MsgBoxResult.Yes Then
        '            'Show Excel Workbook
        '            xlApp.Application.Visible = True
        '            xlBook.Application.Visible = True
        '            xlSheet.Application.Visible = True
        '            xlApp.WindowState = Excel.XlWindowState.xlMaximized
        '        Else
        '            'Quit Excel
        '            xlApp.Application.Quit()
        '        End If
        '        xlApp = Nothing
        '    Catch ex As Exception
        '        MsgBox("Error exporting Excel file." & vbCr & ex.Message)
        '    End Try
        'End If

    End Sub

    Private Sub mnuExportOrganizationChartText_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuExportOrganizationChartText.Click
        Dim myResponse As DialogResult

        'Browse for Access Database
        Me.filExportText.FileName = ""
        Me.filExportText.InitialDirectory = Application.StartupPath
        myResponse = Me.filExportText.ShowDialog()
        myExportFile = Me.filExportText.FileName

        '***** REMOVED 08/19/08 *****
        'If myResponse = DialogResult.OK Then
        '    Try
        '        'Create Export File
        '        Dim sw As System.IO.StreamWriter
        '        Try
        '            sw = System.IO.File.CreateText(myExportFile)
        '        Catch ex As Exception
        '            'Default Configuration File
        '            MsgBox(ex.Message)
        '            Exit Sub
        '        End Try
        '        'Create Export File
        '        Me.funExportOrganizationChartText(sw)
        '        If MsgBox( _
        '            "The CSV file has been created: " & myExportFile & vbCr & vbCr & _
        '            "Would you like to import into Visio now?", MsgBoxStyle.YesNo, "Export to Excel") = MsgBoxResult.No Then
        '            'Show Export File
        '            sw.Flush()
        '            sw.Close()
        '        Else
        '            'Quit Export File
        '            sw.Flush()
        '            sw.Close()
        '            'Create Visio Document
        '            Try
        '                Dim viApp As Microsoft.Office.Interop.Visio.Application
        '                Dim viDoc As Microsoft.Office.Interop.Visio.Document
        '                viApp = CType(CreateObject("Visio.Application"), Microsoft.Office.Interop.Visio.Application)
        '                viDoc = CType(viApp.Documents.AddEx("orgch_u.vst", 0, 0), Microsoft.Office.Interop.Visio.Document)
        '                'viDoc = CType(viApp.Documents.Open(myExportFile), Microsoft.Office.Interop.Visio.Document)
        '                'Minimize Visio Document
        '                viApp.Window.WindowState = Microsoft.Office.Interop.Visio.VisWindowStates.visWSMinimized
        '                MsgBox("The data has been exported to:" & vbCr & myExportFile.ToCharArray & "." & vbCr & vbCr & _
        '                    "Click Organization Chart > Import Organization Data ...", _
        '                    MsgBoxStyle.OKOnly, "Import into Visio")
        '                'Show Visio Document
        '                viApp.Application.Visible = True
        '                viDoc.Application.Visible = True
        '                viApp.Window.WindowState = Microsoft.Office.Interop.Visio.VisWindowStates.visWSMaximized
        '                'Save Visio Document
        '                'viDoc.SaveAs(myExportFile)
        '                'viApp.Application.Quit()
        '                viApp = Nothing
        '            Catch ex As Exception
        '                MsgBox("Unable to Import into Visio.  Please open Visio and import it manually.")
        '            End Try
        '        End If
        '    Catch ex As Exception
        '        MsgBox("Error exporting Text File." & vbCr & ex.Message)
        '    End Try
        'End If

    End Sub

    Private Sub mnuExportIndentedHierarchyText_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuExportIndentedHierarchyText.Click
        Dim myResponse As DialogResult

        'Browse for Access Database
        Me.filExportText.FileName = ""
        Me.filExportText.InitialDirectory = Application.StartupPath
        myResponse = Me.filExportText.ShowDialog()
        myExportFile = Me.filExportText.FileName

        If myResponse = DialogResult.OK Then
            Try
                'Create Export File
                Dim sw As System.IO.StreamWriter
                Try
                    sw = System.IO.File.CreateText(myExportFile)
                Catch ex As Exception
                    'Default Configuration File
                    MsgBox(ex.Message)
                    Exit Sub
                End Try
                'Create Export File
                Me.funExportIndentedHierarchyText(sw)
                MsgBox("The CSV file has been created: " & myExportFile)
                sw.Flush()
                sw.Close()
            Catch ex As Exception
                MsgBox("Error exporting Text File." & vbCr & ex.Message)
            End Try
        End If

    End Sub

    Private Sub mnuTableDefault_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.funSaveChanges()
        'Table Structure
        myChildTable = "Child"
        myParentTable = "Parent"
        myChildTableChildID = "ID"
        myChildTableChildName = "Name"
        myChildTableChildDescription = "Description"
        myParentTableChildID = "ID"
        myParentTableParentID = "ParentID"

        'Me.funCheckTableMenu(Me.mnuTableDefault)

        Me.lblSelected.Text = "Table Structure Set to Child-Parent"

        Me.funCreateComponents()

    End Sub

    Private Sub mnuChildParent_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.funSaveChanges()
        'Table Structure
        myChildTable = "Child"
        myParentTable = "Parent"
        myChildTableChildID = "ID"
        myChildTableChildName = "Name"
        myChildTableChildDescription = "Description"
        myParentTableChildID = "ID"
        myParentTableParentID = "ParentID"

        'Me.funCheckTableMenu(Me.mnuChildParent)

        Me.lblSelected.Text = "Table Structure Set to Child-Parent"

        Me.funCreateComponents()

    End Sub

    Private Sub mnuTableTableAssociation_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.funSaveChanges()
        'Table Structure
        myChildTable = "Table"
        myParentTable = "Table Association"
        myChildTableChildID = "Table ID"
        myChildTableChildName = "Table Name"
        myChildTableChildDescription = "Table Description"
        myParentTableChildID = "Table Child ID"
        myParentTableParentID = "Table Parent ID"

        'Me.funCheckTableMenu(Me.mnuTableTableAssociation)

        Me.lblSelected.Text = "Table Structure Set to Table-Table Association"

        Me.funCreateComponents()

    End Sub

    Private Sub mnuHierarchy_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.funSaveChanges()
        'Table Structure
        myChildTable = "Hierarchy"
        myParentTable = "Hierarchy"
        myChildTableChildID = "ChildID"
        myChildTableChildName = "Name"
        myChildTableChildDescription = "Description"
        myParentTableChildID = "ChildID"
        myParentTableParentID = "ParentID"

        'Me.funCheckTableMenu(Me.mnuHierarchy)

        Me.lblSelected.Text = "Table Structure Set to Hierarchy"

        Me.funCreateComponents()

    End Sub

    Private Sub mnuTableCustom_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuTableCustom.Click
        Dim myTables As New frmDatabase

        Me.funSaveChanges()
        If myTables.ShowDialog(Me) = DialogResult.OK Then
            'Me.funCheckTableMenu(Me.mnuTableCustom)
            Me.lblSelected.Text = "Table Structure Set to Custom"
            Me.funCreateComponents()
            Me.bolSaveConfigurationFile = True
        End If
        myTables.Dispose()

    End Sub

    Private Sub mnuCreateTree_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCreateTree.Click
        'Create Form Components
        Me.funCreateComponents()

    End Sub

    Private Sub mnuRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuRefresh.Click
        Dim myNode As TreeNode

        Me.funSaveChanges()
        'Refresh Tree Colors
        For Each myNode In Me.treDataTree.Nodes
            Me.funResetSearchTree(myNode)
        Next myNode

    End Sub

    Private Sub mnuExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuExit.Click
        If Me.funSaveConfigurationChanges = DialogResult.OK Then
            'Me.funSaveDefaultConfigurationFile()
            Me.Close()
        End If

    End Sub

    'Edit
    Private Sub mnuCut_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCut.Click
        Me.funCutNode()

    End Sub

    Private Sub mnuCopy_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCopy.Click
        Me.funCopyNode()

    End Sub

    Private Sub mnuPaste_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuPaste.Click
        Me.funPasteNode()

    End Sub

    Private Sub mnuPasteToRoot_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuPasteToRoot.Click
        Me.funPasteNode(True)

    End Sub

    Private Sub mnuDuplicate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDuplicate.Click
        Me.funPasteNode(False, True)

    End Sub

    Private Sub mnuFind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFind.Click
        Me.funSearch()

    End Sub

    'Node
    Private Sub mnuNewNode_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuNewNode.Click
        'New (Ctrl-N)
        Me.funAddRootNode(Me.treDataTree, Me.dsTree.Tables(myChildTable), myChildTableChildID, myChildTableChildName)

    End Sub

    Private Sub mnuEditNode_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuEditNode.Click
        'Edit (Ctrl-E)
        If Not Me.treDataTree.SelectedNode Is Nothing Then
            Me.treDataTree.LabelEdit = True
            Me.treDataTree.SelectedNode.BeginEdit()
        End If

    End Sub

    Private Sub mnuAddNode_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAddNode.Click
        'Add (Ins)
        Me.funSaveChanges()
        If Not mySelectedNode Is Nothing Then
            Me.funAddNode(mySelectedNode)
        End If

    End Sub

    Private Sub mnuRemoveNode_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuRemoveNode.Click
        Dim myNode As TreeNode
        Dim myDeletedNode As TreeNode
        Dim myParentNodeName As String
        'Delete (Del)
        If Not mySelectedNode Is Nothing Then
            If mySelectedNode.Parent Is Nothing Then
                myParentNodeName = "."
            Else
                myParentNodeName = " from '" & mySelectedNode.Parent.Text & "."
            End If
            If MsgBox("This will REMOVE '" & mySelectedNode.Text & "' and its entire Sub Tree" & myParentNodeName & vbCr & vbCr & _
                    "This node will not be deleted.  It may become a Root Node when the tree is recreated." & vbCr & vbCr & _
                    "Do you want to continue?", MsgBoxStyle.YesNo + MsgBoxStyle.Question, "Remove Node") = MsgBoxResult.Yes Then
                myDeletedNode = mySelectedNode
                Me.funRemoveNode(mySelectedNode)
                myNode = Me.treDataTree.Nodes(0)
                'While Not myNode Is Nothing
                '    If myNode.Tag = myDeletedNode.Tag Then
                '        myNode.Remove()
                '        myNode = Me.treDataTree.Nodes(0)
                '    End If
                '    myNode = myNode.NextVisibleNode
                'End While
                mySelectedNode = Me.treDataTree.SelectedNode
            End If
        End If

    End Sub

    Private Sub mnuDeleteNode_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDeleteNode.Click
        Dim myNode As TreeNode
        Dim myDeletedNode As TreeNode
        'Delete (Del)
        If Not mySelectedNode Is Nothing Then
            If MsgBox("This will completely DELETE '" & mySelectedNode.Text & "', all duplicate nodes, and all subtree nodes from the database." & vbCr & vbCr & _
                    "This is a major change.  ALL references to this node will be deleted." & vbCr & _
                    "THERE MAY BE SOME REFERENCES NOT SEEN HERE." & vbCr & vbCr & _
                    "Do you want to continue?", MsgBoxStyle.YesNo + MsgBoxStyle.Critical, "Delete Node") = MsgBoxResult.Yes Then
                myDeletedNode = mySelectedNode
                Me.funDeleteNode(mySelectedNode)
                myNode = Me.treDataTree.Nodes(0)
                While Not myNode Is Nothing
                    If myNode.Tag = myDeletedNode.Tag Then
                        myNode.Remove()
                        myNode = Me.treDataTree.Nodes(0)
                    End If
                    myNode = myNode.NextVisibleNode
                End While
                mySelectedNode = Me.treDataTree.SelectedNode
            End If
        End If

    End Sub

    'Help
    Private Sub mnuContents_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuContents.Click
        Me.funHelpContents()

    End Sub

    Private Sub mnuAbout_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAbout.Click
        MsgBox( _
            "Data Tree Version 1.61" & vbCr & vbCr & _
            "Silver Bullet Solutions, Inc." & vbCr & _
            "July 16, 2012", , _
            "Data Tree Version 1.61")

    End Sub

    'Functions
    Private Sub funCheckTableMenu(ByVal MyMenuItem As MenuItem)
        'Table
        'Me.mnuTableDefault.Checked = False
        'Me.mnuTableTypical.Checked = False
        'Me.mnuTableCustom.Checked = False
        'Table Typical
        'Me.mnuChildParent.Checked = False
        'Me.mnuTableTableAssociation.Checked = False
        'Me.mnuHierarchy.Checked = False
        'Checked Table
        MyMenuItem.Checked = True

    End Sub

    Private Sub funEnableFileMenuItems(ByVal bolEnabled As Boolean)
        Me.mnuOpen.Enabled = bolEnabled
        Me.mnuSave.Enabled = bolEnabled
        Me.mnuSaveAs.Enabled = bolEnabled
        'Me.mnuSetTable.Enabled = bolEnabled
        Me.mnuRefresh.Enabled = bolEnabled
        Me.mnuCreateTree.Enabled = bolEnabled
        Me.mnuExport.Enabled = bolEnabled
        Me.mnuTableCustom.Enabled = bolEnabled

    End Sub

    Private Sub funEnableEditMenuItems(ByVal bolEnabled As Boolean)
        Me.mnuCut.Enabled = bolEnabled
        Me.mnuCopy.Enabled = bolEnabled
        Me.mnuPaste.Enabled = bolEnabled
        Me.mnuPasteToRoot.Enabled = bolEnabled
        Me.mnuDuplicate.Enabled = bolEnabled
        Me.mnuEditSeparator2.Enabled = bolEnabled
        Me.mnuFind.Enabled = bolEnabled

        Me.mnuTreeViewFind.Enabled = bolEnabled
        Me.mnuTreeViewSeparator1.Enabled = bolEnabled
        Me.mnuTreeViewCut.Enabled = bolEnabled
        Me.mnuTreeViewCopy.Enabled = bolEnabled
        Me.mnuTreeViewPaste.Enabled = bolEnabled
        Me.mnuTreeViewPasteToRoot.Enabled = bolEnabled
        Me.mnuTreeViewDuplicate.Enabled = bolEnabled

        If myCopyNode Is Nothing And myCutNode Is Nothing Then
            Me.mnuPaste.Enabled = False
            Me.mnuPasteToRoot.Enabled = False
            Me.mnuDuplicate.Enabled = False
            Me.mnuTreeViewPaste.Enabled = False
            Me.mnuTreeViewPasteToRoot.Enabled = False
            Me.mnuTreeViewDuplicate.Enabled = False
        End If
        If mySelectedNode Is Nothing Then
            Me.mnuCut.Enabled = False
            Me.mnuCopy.Enabled = False
            Me.mnuPaste.Enabled = False
            Me.mnuDuplicate.Enabled = False
            Me.mnuTreeViewCut.Enabled = False
            Me.mnuTreeViewCopy.Enabled = False
            Me.mnuTreeViewPaste.Enabled = False
            Me.mnuTreeViewPasteToRoot.Enabled = False
            Me.mnuTreeViewDuplicate.Enabled = False
        End If

    End Sub

    Private Sub funEnableNodeMenuItems(ByVal bolEnabled As Boolean)
        Me.mnuNewNode.Enabled = bolEnabled
        Me.mnuAddNode.Enabled = bolEnabled
        Me.mnuEditNode.Enabled = bolEnabled
        Me.mnuDeleteNode.Enabled = bolEnabled
        Me.mnuRemoveNode.Enabled = bolEnabled

        Me.mnuTreeViewSeparator2.Enabled = bolEnabled
        Me.mnuTreeViewAddNode.Enabled = bolEnabled
        Me.mnuTreeViewNewNode.Enabled = bolEnabled
        Me.mnuTreeViewEditNode.Enabled = bolEnabled
        Me.mnuTreeViewDeleteNode.Enabled = bolEnabled
        Me.mnuTreeViewRemoveNode.Enabled = bolEnabled

        If mySelectedNode Is Nothing Then
            Me.mnuAddNode.Enabled = False
            Me.mnuEditNode.Enabled = False
            Me.mnuDeleteNode.Enabled = False
            Me.mnuRemoveNode.Enabled = False
            Me.mnuTreeViewAddNode.Enabled = False
            Me.mnuTreeViewEditNode.Enabled = False
            Me.mnuTreeViewDeleteNode.Enabled = False
            Me.mnuTreeViewRemoveNode.Enabled = False
        End If

    End Sub

#End Region

#Region "Configuration File"

    Private Sub mnuOpen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuOpen.Click
        Me.funOpenConfigurationFile()

    End Sub

    Private Sub mnuSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSave.Click
        Me.funSaveConfigurationFile()

    End Sub

    Private Sub mnuSaveAs_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSaveAs.Click
        Dim myResponse As DialogResult
        'Dim strConfigurationFile As String

        'Browse for Access Database
        Me.filSaveConfigurationFile.FileName = ""
        Me.filSaveConfigurationFile.InitialDirectory = Application.StartupPath
        myResponse = Me.filSaveConfigurationFile.ShowDialog()
        myConfigurationFile = Me.filSaveConfigurationFile.FileName

        If myResponse = DialogResult.OK Then
            Me.mnuSave_Click(sender, e)
        End If


    End Sub

    Private Sub funSaveDefaultConfigurationFile()
        'Browse for Access Database
        myConfigurationFile = Application.StartupPath & "\Default.dtc"
        Me.funSaveConfigurationFile()

    End Sub

    Private Sub mnuClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuClose.Click
        If Me.funSaveConfigurationChanges = DialogResult.OK Then
            myConfigurationFile = ""
            'Reset
            Me.funResetTree()
            Me.funResetGlobals()
            ''Reset Variables
            'myChildTable = "Child"
            'myParentTable = "Parent"
            'myChildTableChildID = "ID"
            'myChildTableChildName = "Name"
            'myChildTableChildDescription = "Description"
            'myParentTableChildID = "ID"
            'myCriteriaWhere = ""
            'myCriteriaParentTableWhere = ""
            'myCriteriaInsert = ""
            'myCriteriaValue = ""
            'myCriteriaSet = ""
            'myParentTableParentID = "ParentID"
            'myParentTableAutoNumber = ""

            Me.mnuClose.Enabled = False
            Me.mnuExport.Enabled = False
            Me.bolSaveConfigurationFile = False
        End If

    End Sub

    Private Function funSaveConfigurationChanges() As DialogResult
        Dim myResponse As DialogResult

        If Me.bolSaveConfigurationFile And Me.myConfigurationFile <> "" Then
            myResponse = MsgBox("Do you want to save changes to " & myConfigurationFile & "?", MsgBoxStyle.YesNoCancel)
            If myResponse = DialogResult.Yes Then
                Me.mnuSave_Click(Nothing, Nothing)
            ElseIf myResponse = DialogResult.Cancel Then
                funSaveConfigurationChanges = DialogResult.Cancel
                Exit Function
            End If
        End If
        Me.bolSaveConfigurationFile = False
        funSaveConfigurationChanges = DialogResult.OK

    End Function

#End Region

#Region "Shortcut Menus"

    Private Sub mnuTreeViewFind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuTreeViewFind.Click
        Me.funSearch(mySelectedNode)

    End Sub

    Private Sub mnuDataGridCustomize_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDataGridCustomize.Click
        Dim myDataGridForm As New frmDataGrid
        'Dim myMenuItem As MenuItem
        'Dim myColumn As DataGridColumnStyle

        Me.funSaveChanges()
        If myDataGridForm.ShowDialog(Me) = DialogResult.OK Then
            Me.funSetColumnOrder(Me.grdChildren, Me.dsChild.Tables(myChildTable), Me.myColumnList, Me.mnuDataGrid)
            Me.funShowColumns()
            Me.bolSaveConfigurationFile = True
            'For Each myMenuItem In Me.mnuDataGrid.MenuItems
            '    If myMenuItem.Index > 1 Then
            '        myColumn = Me.grdChildren.TableStyles(myChildTable).GridColumnStyles(myMenuItem.Text)
            '        If Me.myColumnList.CheckedItems.Contains(myMenuItem.Text) Then
            '            myMenuItem.Checked = True
            '            If myColumn.Width = 0 Then
            '                myColumn.Width = 75
            '            End If
            '        Else
            '            myMenuItem.Checked = False
            '            myColumn.Width = 0
            '        End If
            '    End If
            'Next
        End If
        myDataGridForm.Dispose()

    End Sub

    Private Sub mnuTreeViewCut_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuTreeViewCut.Click
        Me.funCutNode()

    End Sub

    Private Sub mnuTreeViewCopy_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuTreeViewCopy.Click
        Me.funCopyNode()

    End Sub

    Private Sub mnuTreeViewPaste_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuTreeViewPaste.Click
        Me.funPasteNode()

    End Sub

    Private Sub mnuTreeViewPasteToRoot_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuTreeViewPasteToRoot.Click
        Me.funPasteNode(True)

    End Sub

    Private Sub mnuTreeViewDuplicate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuTreeViewDuplicate.Click
        Me.funPasteNode(False, True)

    End Sub

    Private Sub mnuTreeViewNewNode_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuTreeViewNewNode.Click
        Me.mnuNewNode_Click(sender, e)

    End Sub

    Private Sub mnuTreeViewEditNode_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuTreeViewEditNode.Click
        Me.mnuEditNode_Click(sender, e)

    End Sub

    Private Sub mnuTreeViewAddNode_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuTreeViewAddNode.Click
        Me.mnuAddNode_Click(sender, e)

    End Sub

    Private Sub mnuTreeViewRemoveNode_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuTreeViewRemoveNode.Click
        Me.mnuRemoveNode_Click(sender, e)

    End Sub

    Private Sub mnuTreeViewDeleteNode_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuTreeViewDeleteNode.Click
        Me.mnuDeleteNode_Click(sender, e)

    End Sub

    Private Sub grdChildren_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles grdChildren.MouseDown
        'DataGrid
        If e.Button = MouseButtons.Left And e.Y <= 20 Then
            Me.treDataTree.Focus()
        ElseIf e.Button = MouseButtons.Right And e.Y <= 20 Then
            'Column Header
            Me.grdChildren.ContextMenu = Me.mnuDataGrid
            Me.mnuDataGridCustomize.Visible = True
            Me.mnuDataGridRowHeight.Visible = True
            Me.mnuDataGridColumnWidth.Visible = True
            Me.mnuDataGridPaste.Visible = False
            Me.mnuDataGridAddNode.Visible = False
            Me.mnuDataGridProperties.Visible = False
            'Me.mnuDataGridSeparator1.Visible = False
            Me.mnuDataGridSeparator2.Visible = False
        ElseIf e.Button = MouseButtons.Right And e.Y > 20 Then
            'DataGrid Children
            Me.grdChildren.ContextMenu = Nothing
            Me.mnuDataGridCustomize.Visible = False
            Me.mnuDataGridRowHeight.Visible = False
            Me.mnuDataGridColumnWidth.Visible = False
            Me.mnuDataGridPaste.Visible = True
            Me.mnuDataGridAddNode.Visible = True
            Me.mnuDataGridProperties.Visible = True
            Me.mnuDataGridSeparator1.Visible = False
            Me.mnuDataGridSeparator2.Visible = True
        End If

        Dim myHitTest As System.Windows.Forms.DataGrid.HitTestInfo
        myHitTest = Me.grdChildren.HitTest(e.X, e.Y)
        intColumnNumber = myHitTest.Column

    End Sub

    Private Sub mnuDataGridRowHeight_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDataGridRowHeight.Click
        Dim myTableStyle As New DataGridTableStyle
        Dim intRowHeight As Integer
        Try
            myTableStyle = Me.grdChildren.TableStyles(0)
            intRowHeight = InputBox("Row Height", "Row Height", myTableStyle.PreferredRowHeight)
            myTableStyle.PreferredRowHeight = intRowHeight
            Me.grdChildren.TableStyles.Clear()
            Me.grdChildren.TableStyles.Add(myTableStyle)
            Me.bolSaveConfigurationFile = True
        Catch ex As Exception

        End Try

    End Sub

    Private Sub mnuDataGridColumnWidth_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDataGridColumnWidth.Click
        Dim myTableStyle As New DataGridTableStyle
        Dim intColumnWidth As Integer
        Try
            myTableStyle = Me.grdChildren.TableStyles(0)
            intColumnWidth = InputBox("Column Width", "Column Width", myTableStyle.GridColumnStyles(intColumnNumber).Width)
            myTableStyle.GridColumnStyles(intColumnNumber).Width = intColumnWidth
            Me.grdChildren.TableStyles.Clear()
            Me.grdChildren.TableStyles.Add(myTableStyle)
            Me.bolSaveConfigurationFile = True
        Catch ex As Exception

        End Try

    End Sub

#End Region

#Region "Form Events"

    'Select Node
    Private Sub treDataTree_AfterSelect(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles treDataTree.AfterSelect
        Try
            Me.funSaveChanges()
            'Remove Node Place Holder
            If Me.treDataTree.SelectedNode.Nodes(0).Text = "NodePlaceHolder" Then
                Me.treDataTree.SelectedNode.Nodes(0).Remove()
            End If
        Catch ex As Exception
            'MsgBox("Remove Node Place Holder: " & ex.Message)
        End Try
        Try
            mySelectedNode = Me.treDataTree.SelectedNode
            'Node Path
            Me.lblSelected.Text = _
                mySelectedNode.FullPath
            'Display Children
            Try
                Me.funDisplayDescription(mySelectedNode.Tag)
                Me.funDisplayChildren(mySelectedNode.Tag)
            Catch ex As Exception
                MsgBox("Display Children: " & ex.Message)
            End Try
            'Expand Node (Select Node)
            If mySelectedNode.GetNodeCount(True) = 0 Then
                Me.funCreateSubTree(mySelectedNode, Me.dsChild.Tables(myChildTable), myChildTableChildID, myChildTableChildName)
            End If
            'Collapse Node (Click Minus)
            If Not bolIsCollapsing Then
                mySelectedNode.Expand()
            End If
            Me.grdChildren.CurrentRowIndex = Nothing
            'Menu Bar
            Me.funEnableEditMenuItems(True)
            Me.funEnableNodeMenuItems(True)
        Catch ex As Exception
            'MsgBox("Data Tree Select: " & ex.Message)
        End Try

    End Sub

    'Expand
    Private Sub treDataTree_BeforeExpand(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeViewCancelEventArgs) Handles treDataTree.BeforeExpand
        'Select Node to Remove Place Holder
        Me.treDataTree.SelectedNode = e.Node

    End Sub

    'Collapse
    Private Sub treDataTree_BeforeCollapse(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeViewCancelEventArgs) Handles treDataTree.BeforeCollapse
        'Click Minus
        bolIsCollapsing = True

    End Sub

    Private Sub treDataTree_AfterCollapse(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles treDataTree.AfterCollapse
        'Node Selected
        bolIsCollapsing = False

    End Sub

    'Update Node
    Private Sub treDataTree_BeforeLabelEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.NodeLabelEditEventArgs) Handles treDataTree.BeforeLabelEdit
        Me.funEnableEditMenuItems(False)
        Me.funEnableNodeMenuItems(False)

    End Sub

    Private Sub treDataTree_AfterLabelEdit(ByVal sender As System.Object, ByVal e As System.Windows.Forms.NodeLabelEditEventArgs) Handles treDataTree.AfterLabelEdit
        Dim myNode As TreeNode
        Dim myCurrentNode As TreeNode = e.Node

        If e.Label = Nothing Or e.Label = "" Then
            e.CancelEdit = True
        Else
            Me.funUpdateTree(e.Node, e.Label)
            'Node Path
            If mySelectedNode.Parent Is Nothing Then
                Me.lblSelected.Text = e.Label
            Else
                Me.lblSelected.Text = _
                    mySelectedNode.Parent.FullPath & "\" & e.Label
            End If
            'Reset Tree 
            For Each myNode In Me.treDataTree.Nodes
                Me.funResetSearchTree(myNode)
            Next myNode
            'Change Node ID
            If myChildTableChildName = myChildTableChildID Then
                'Find Node 
                For Each myNode In Me.treDataTree.Nodes
                    Me.funSearchTree(mySelectedNode.Tag, myNode, "ReplaceNodeID", e.Label)
                Next myNode
            End If
            'Change Node Name
            For Each myNode In Me.treDataTree.Nodes
                Me.funSearchTree(e.Node.Tag, myNode, "ReplaceNode", e.Label)
            Next myNode
            'Change Node Description
            If myChildTableChildName = myChildTableChildDescription Then
                Me.txtDescription.Text = e.Label
            End If
        End If
        Me.funEnableEditMenuItems(True)
        Me.funEnableNodeMenuItems(True)
        Me.treDataTree.SelectedNode = myCurrentNode

    End Sub

    'Select Current Node
    Private Sub treDataTree_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles treDataTree.Enter
        Try
            mySelectedNode = Me.treDataTree.SelectedNode
            Me.funDisplayDescription(mySelectedNode.Tag)
        Catch ex As Exception
            'MsgBox("Data Tree Enter: " & ex.Message)
        End Try

    End Sub

    'Display Child Description
    Private Sub grdChildren_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdChildren.CurrentCellChanged
        Dim intRow As Integer
        Dim myNode As TreeNode

        Try
            intRow = Me.grdChildren.CurrentRowIndex
            For Each myNode In Me.treDataTree.SelectedNode.Nodes
                If myNode.Tag = Me.grdChildren.Item(Me.grdChildren.CurrentRowIndex, intChildIDColumn) Then
                    mySelectedNode = myNode
                    Exit For
                End If
            Next
            Me.funDisplayDescription(mySelectedNode.Tag)
        Catch ex As Exception
            'MsgBox("DataGrid Children Click: " & ex.Message)
        End Try

    End Sub

    'Select Child
    Private Sub grdChildren_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdChildren.DoubleClick
        Dim intRow As Integer

        Try
            bolMouseOverRowHeader = True
            If bolMouseOverColumnHeader Or bolMouseOverRowHeader Then
                Exit Sub
            End If
            intRow = Me.grdChildren.CurrentRowIndex
            Me.treDataTree.SelectedNode.Expand()
            Me.treDataTree.SelectedNode = Me.treDataTree.SelectedNode.Nodes(intRow)
            Me.treDataTree.Focus()
        Catch ex As Exception
            'MsgBox("DataGrid Children Double Click: " & ex.Message)
            'Reset DataTree and DataGrid
            Me.grdChildren.CurrentRowIndex = Nothing
            Me.treDataTree.SelectedNode = mySelectedNode
            Me.treDataTree.Focus()
        End Try

    End Sub

    'Menu Items
    Private Sub grdChildren_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdChildren.Enter
        Dim intRow As Integer
        Dim myNode As TreeNode

        'Disable
        Me.funEnableEditMenuItems(False)
        Me.funEnableNodeMenuItems(False)
        Try
            intRow = Me.grdChildren.CurrentRowIndex
            For Each myNode In Me.treDataTree.SelectedNode.Nodes
                If myNode.Tag = Me.grdChildren.Item(Me.grdChildren.CurrentRowIndex, intChildIDColumn) Then
                    mySelectedNode = myNode
                    Exit For
                End If
            Next
            Me.funDisplayDescription(mySelectedNode.Tag)
        Catch ex As Exception
            'MsgBox("DataGrid Children Enter: " & ex.Message)
        End Try

    End Sub

    Private Sub grdChildren_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdChildren.Leave
        'Enable
        Me.funEnableEditMenuItems(True)
        Me.funEnableNodeMenuItems(True)

    End Sub

    Private Sub frmDataTree_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        If Me.funSaveConfigurationChanges = DialogResult.OK Then
            'Save Changes to DataGrid
            Me.funSaveChanges()
            'Me.funSaveDefaultConfigurationFile()
        Else
            e.Cancel = True
        End If

    End Sub

    'Edit Description
    Private Sub txtDescription_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDescription.DoubleClick
        If Not mySelectedNode Is Nothing Then
            Me.txtDescription.ReadOnly = False
        End If

    End Sub

    'Update Description
    Private Sub txtDescription_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtDescription.KeyPress
        Dim myNode As TreeNode

        'Hit Return to Save Description
        If e.KeyChar = Microsoft.VisualBasic.ChrW(13) And Not mySelectedNode Is Nothing Then
            Me.funUpdateTree(mySelectedNode, , Me.txtDescription.Text)
            Me.txtDescription.ReadOnly = True
            'Change Node ID
            If myChildTableChildDescription = myChildTableChildID Then
                'Reset Tree 
                For Each myNode In Me.treDataTree.Nodes
                    Me.funResetSearchTree(myNode)
                Next myNode
                'Find Node 
                For Each myNode In Me.treDataTree.Nodes
                    Me.funSearchTree(mySelectedNode.Tag, myNode, "ReplaceNodeID", Me.txtDescription.Text)
                Next myNode
            End If
            'Change Node Name
            If myChildTableChildDescription = myChildTableChildName Then
                'Reset Tree 
                For Each myNode In Me.treDataTree.Nodes
                    Me.funResetSearchTree(myNode)
                Next myNode
                'Find Node 
                For Each myNode In Me.treDataTree.Nodes
                    Me.funSearchTree(mySelectedNode.Tag, myNode, "ReplaceNode", Me.txtDescription.Text)
                Next myNode
            End If
        End If

    End Sub

#End Region

#Region "Tree"

    'Create Tree
    Private Sub funCreateTree(ByVal myTree As TreeView, ByVal myTable As DataTable, ByVal myID As String, ByVal myName As String)
        Dim myNode As TreeNode
        Dim myRow As DataRow
        'Create Tree
        myTree.Nodes.Clear()
        For Each myRow In myTable.Rows
            myNode = New TreeNode
            If Not IsDBNull(myRow.Item(myID)) Then
                myNode.Tag = myRow.Item(myID)
                Try
                    myNode.Text = myRow.Item(myName)
                    If myNode.Text = "" Then
                        myNode.Text = "(null)"
                    End If
                Catch ex As Exception
                    myNode.Text = "(null)"
                End Try
                myTree.Nodes.Add(myNode)
                myNode.Nodes.Add("NodePlaceHolder")
            End If
        Next

    End Sub

    'Create Sub Tree
    Private Sub funCreateSubTree(ByVal mySelectedNode As TreeNode, ByVal myTable As DataTable, ByVal myID As String, ByVal myName As String)
        Dim myNode As TreeNode
        Dim myRow As DataRow
        'Add Sub Tree
        For Each myRow In myTable.Rows
            myNode = New TreeNode
            If Not IsDBNull(myRow.Item(myID)) Then
                myNode.Tag = myRow.Item(myID)
                Try
                    myNode.Text = myRow.Item(myName)
                    If myNode.Text = "" Then
                        myNode.Text = "(null)"
                    End If
                Catch ex As Exception
                    myNode.Text = "(null)"
                End Try
                mySelectedNode.Nodes.Add(myNode)
                myNode.Nodes.Add("NodePlaceHolder")
            End If
        Next

    End Sub

    'Destroy Tree
    Private Sub funDestroyTree(ByVal myTree As TreeView)
        'Destroy Tree
        myTree.Nodes.Clear()

    End Sub

    'Destroy Sub Tree
    Private Sub funDestroySubTree(ByVal myNode As TreeNode)
        'Destroy Sub Tree
        myNode.Nodes.Clear()

    End Sub

    'Update Node
    Private Sub funUpdateTree(ByVal myNode As TreeNode, Optional ByVal myLabel As String = Nothing, Optional ByVal myDescription As String = Nothing)
        'Modify Select Statement
        sqlFrom = _
            " FROM [" & myChildTable & "] "
        sqlWhere = _
            " WHERE [" & myChildTable & "].[" & myChildTableChildID & "] = " & Me.funConvertDataType(myNode.Tag, myChildTableChildID)
        'Get Selected Node
        Me.tblChild.SelectCommand.CommandText = _
            sqlSelect & _
            sqlFrom & _
            sqlWhere
        Me.dsTree.Clear()
        Me.tblChild.Fill(Me.dsTree)
        'Update Name
        Me.dsTree.Tables(myChildTable).Rows(0).Item(myChildTableChildName) = myLabel
        'Find/Replace Single Quotes
        myLabel = Me.funReplaceSingleQuote(myLabel)
        myDescription = Me.funReplaceSingleQuote(myDescription)
        'Create Update Statement
        sqlUpdate = _
            " UPDATE [" & myChildTable & "]"
        If myLabel <> Nothing And myDescription <> Nothing Then
            sqlSet = _
                " SET [" & myChildTableChildName & "] = " & Me.funConvertDataType(myLabel, myChildTableChildName) & ", " & _
                     "[" & myChildTableChildDescription & "] = " & Me.funConvertDataType(myDescription, myChildTableChildDescription) & ""
        ElseIf myLabel <> Nothing Then
            sqlSet = _
                " SET [" & myChildTableChildName & "] = " & Me.funConvertDataType(myLabel, myChildTableChildName)
        ElseIf myDescription <> Nothing Then
            sqlSet = _
                " SET [" & myChildTableChildDescription & "] = " & Me.funConvertDataType(myDescription, myChildTableChildDescription)
        Else
            Exit Sub
        End If
        'Update Selected Node
        Me.tblChild.UpdateCommand.CommandText = _
            sqlUpdate & _
            sqlSet & _
            sqlWhere
        Try
            Me.tblChild.Update(Me.dsTree)
        Catch ex As Exception
            MsgBox("Update Tree: " & ex.Message)
        End Try
        Me.funDisplayChildren(mySelectedNode.Tag)

    End Sub

    'Add Root Node
    Private Sub funAddRootNode(ByVal myTree As TreeView, ByVal myTable As DataTable, ByVal myID As String, ByVal myName As String)
        Dim myNode As TreeNode
        Dim myRow As DataRow
        'Create Node
        myNode = New TreeNode
        myNode.Tag = funGetNextID(myChildTableChildID)
        myNode.Text = "New Item"
        'Modify Select Statement
        sqlFrom = _
            " FROM [" & myChildTable & "] "
        'Get Selected Node
        Me.tblChild.SelectCommand.CommandText = _
            sqlSelect & _
            sqlFrom
        Me.dsTree.Clear()
        Me.tblChild.Fill(Me.dsTree)
        'Add Row
        myRow = Me.dsTree.Tables(myChildTable).NewRow
        myRow(myChildTableChildID) = myNode.Tag
        myRow(myChildTableChildName) = myNode.Text
        Me.dsTree.Tables(myChildTable).Rows.Add(myRow)
        'Insert Into Database
        sqlInsert = _
            " INSERT INTO [" & myChildTable & "]" & _
                "([" & myChildTableChildID & "], [" & myChildTableChildName & "])"
        sqlValues = _
            " VALUES (" & Me.funConvertDataType(myNode.Tag, myChildTableChildID) & ", " & Me.funConvertDataType(myNode.Text, myChildTableChildName) & ")"
        Me.tblChild.InsertCommand.CommandText = _
            sqlInsert & _
            sqlValues
        Me.dbConnection.Open()
        Try
            Me.tblChild.InsertCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Add Root Node: " & ex.Message)
        End Try
        Me.dbConnection.Close()
        'Add Node
        myTree.Nodes.Add(myNode)
        Me.treDataTree.SelectedNode = myNode
        Me.funDisplayChildren(mySelectedNode.Tag)
        Me.treDataTree.SelectedNode.BeginEdit()

    End Sub

    'Add Child Node
    Private Sub funAddNode(ByVal myParentNode As TreeNode, Optional ByVal myDataRow As DataRow = Nothing)
        Dim myNode As TreeNode
        Dim myID As Integer
        Dim myName As String
        Dim intCounter As Integer

        'Create Node
        myNode = New TreeNode
        myID = funGetNextID(myChildTableChildID)
        myNode.Tag = myID
        If myDataRow Is Nothing Then
            myName = "New Item"
        Else
            myName = myDataRow(myChildTableChildName)
            myName = Me.funReplaceSingleQuote(myName)
        End If
        myNode.Text = myName



        ''Modify Select Statement
        'sqlFrom = _
        '    " FROM [" & myChildTable & "] "
        ''Get Selected Node
        'Me.tblChild.SelectCommand.CommandText = _
        '    sqlSelect & _
        '    sqlFrom
        'Me.dsTree.Clear()
        'Me.tblChild.Fill(Me.dsTree)
        ''Add Child
        'myRow = Me.dsTree.Tables(myChildTable).NewRow
        'myRow(myChildTableChildID) = myNode.Tag
        'myRow(myChildTableChildName) = myNode.Text
        'If myParentTable = myChildTable Then
        '    'myRow(myParentTableParentID) = mySelectedNode.Tag
        'End If
        'Me.dsTree.Tables(myChildTable).Rows.Add(myRow)
        'Insert Child
        If myParentTable = myChildTable Then
            sqlInsert = _
                " INSERT INTO [" & myChildTable & "]" & _
                    "([" & myChildTableChildID & "], [" & myChildTableChildName & "], [" & myParentTableParentID & "])"
            sqlValues = _
                " VALUES (" & Me.funConvertDataType(myNode.Tag, myChildTableChildID) & ", " & Me.funConvertDataType(myName, myChildTableChildName) & ", " & Me.funConvertDataType(myParentNode.Tag, myChildTableChildID) & ")"
        Else
            sqlInsert = _
                " INSERT INTO [" & myChildTable & "]" & _
                    "([" & myChildTableChildID & "], [" & myChildTableChildName & "])"
            sqlValues = _
                " VALUES (" & Me.funConvertDataType(myNode.Tag, myChildTableChildID) & ", " & Me.funConvertDataType(myName, myChildTableChildName) & ")"
        End If
        Me.tblChild.InsertCommand.CommandText = _
            sqlInsert & _
            sqlValues
        Me.dbConnection.Open()
        Try
            Me.tblChild.InsertCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Add Node (Child): " & ex.Message)
        End Try
        Me.dbConnection.Close()
        'Insert Parent
        If myParentTable <> myChildTable And Not myParentNode Is Nothing Then
            'sqlInsert = _
            '    " INSERT INTO [" & myParentTable & "]" & _
            '        "([" & myParentTableChildID & "], [" & myParentTableParentID & "])"
            'sqlValues = _
            '    " VALUES (" & Me.funConvertDataType(myNode.Tag, myChildTableChildID) & ", " & Me.funConvertDataType(myParentNode.Tag, myChildTableChildID) & ")"
            'Me.tblChild.InsertCommand.CommandText = _
            '    sqlInsert & _
            '    sqlValues
            Me.funAddParentNode(myNode, myParentNode)
            Try
                Me.dbConnection.Open()
                Me.tblChild.InsertCommand.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox("Add Node (Parent): " & ex.Message)
            Finally
                Me.dbConnection.Close()
            End Try
        End If
        'Update Child Node
        If Not myDataRow Is Nothing Then
            sqlUpdate = _
                " UPDATE [" & myChildTable & "]"
            sqlWhere = _
                " WHERE [" & myChildTableChildID & "] = " & Me.funConvertDataType(myNode.Tag, myChildTableChildID)
            sqlSet = ""
            For intCounter = 0 To myDataRow.Table.Columns.Count - 1
                If myDataRow.Table.Columns(intCounter).ColumnName <> myChildTableChildID And _
                    Not IsDBNull(myDataRow(intCounter)) And _
                    myDataRow.Table.Columns(intCounter).ColumnName <> "DateCreated" And _
                    myDataRow.Table.Columns(intCounter).ColumnName <> "CreationDate" And _
                    myDataRow.Table.Columns(intCounter).ColumnName <> "DateModified" And _
                    myDataRow.Table.Columns(intCounter).ColumnName <> "ModificationDate" _
                Then
                    sqlSet = sqlSet & IIf(sqlSet = "", " SET ", ", ") & _
                        "[" & myChildTable & "].[" & myDataRow.Table.Columns(intCounter).ColumnName & "] = " & Me.funConvertDataType(myDataRow(intCounter), myDataRow.Table.Columns(intCounter).ColumnName)
                End If
            Next
            Me.tblChild.UpdateCommand.CommandText = _
                sqlUpdate & _
                sqlSet & _
                sqlWhere
            Try
                Me.dbConnection.Open()
                Me.tblChild.UpdateCommand.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox("Add Node Update Query: " & ex.Message)
            Finally
                Me.dbConnection.Close()
            End Try
        End If
        'Add Node
        If myParentNode Is Nothing Then
            Me.treDataTree.Nodes.Add(myNode)
        Else
            myParentNode.Nodes.Add(myNode)
        End If
        Me.funDisplayChildren(mySelectedNode.Tag)
        Me.treDataTree.SelectedNode = myNode
        mySelectedNode = myNode
        mySelectedNode.BeginEdit()

    End Sub

    'Remove Node
    Private Sub funRemoveNode(ByVal myNode As TreeNode)
        'Dim intNodeCount As Integer
        ''Remove Children
        'While myNode.GetNodeCount(False) > 0
        '    funRemoveNode(myNode.Nodes(0))
        'End While
        'Remove Relationship
        If Not myNode.Parent Is Nothing Then
            sqlDelete = _
                " DELETE"
            sqlFrom = _
                " FROM [" & myParentTable & "]"
            sqlWhere = _
                " WHERE [" & myParentTableChildID & "] = " & Me.funConvertDataType(myNode.Tag, myChildTableChildID) & _
                    " AND [" & myParentTableParentID & "] = " & Me.funConvertDataType(myNode.Parent.Tag, myChildTableChildID)
            If Me.myCriteriaWhere <> "" Then
                sqlWhere = _
                    " WHERE [" & myParentTable & "].[" & myParentTableChildID & "] = " & Me.funConvertDataType(myNode.Tag, myChildTableChildID) & _
                        " AND [" & myParentTable & "].[" & myParentTableParentID & "] = " & Me.funConvertDataType(myNode.Parent.Tag, myChildTableChildID) & _
                        " AND " & Me.myCriteriaWhere
            End If
            Me.tblChild.DeleteCommand.CommandText = _
                sqlDelete & _
                sqlFrom & _
                sqlWhere
            Me.dbConnection.Open()
            Try
                Me.tblChild.DeleteCommand.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox("Delete Node: " & ex.Message)
            End Try
            Me.dbConnection.Close()

        End If
        'Remove Tree Node
        Me.treDataTree.SelectedNode = myNode
        Me.treDataTree.SelectedNode = myNode.Parent
        Me.treDataTree.Focus()
        Try
            myNode.Remove()
        Catch ex As Exception
            'Error Removing NodePlaceHolder
            'MsgBox("Error Removing: " & myNode.Text)
        End Try

    End Sub

    'Delete Node
    Private Sub funDeleteNode(ByVal myNode As TreeNode)
        'Dim intNodeCount As Integer
        'Delete Children
        While myNode.GetNodeCount(False) > 0
            funDeleteNode(myNode.Nodes(0))
        End While
        'Delete Relationship First
        'If Not myNode.Parent Is Nothing Then
        sqlDelete = _
            " DELETE"
        sqlFrom = _
            " FROM [" & myParentTable & "]"
        sqlWhere = _
            " WHERE [" & myParentTableChildID & "] = " & Me.funConvertDataType(myNode.Tag, myChildTableChildID) & _
                " OR [" & myParentTableParentID & "] = " & Me.funConvertDataType(myNode.Tag, myChildTableChildID)
        'If Me.myCriteriaWhere <> "" Then
        '    sqlWhere = _
        '        " WHERE [" & myParentTable & "].[" & myParentTableChildID & "] = " & Me.funConvertDataType(myNode.Tag, myChildTableChildID) & _
        '            " AND [" & myParentTable & "].[" & myParentTableParentID & "] = " & Me.funConvertDataType(myNode.Parent.Tag, myChildTableChildID) & _
        '            " AND " & Me.myCriteriaWhere
        'End If
        Me.tblChild.DeleteCommand.CommandText = _
            sqlDelete & _
            sqlFrom & _
            sqlWhere
        Me.dbConnection.Open()
        Try
            Me.tblChild.DeleteCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Delete Node: " & ex.Message)
        End Try
        Me.dbConnection.Close()

        'End If

        'Delete Node from Database
        sqlDelete = _
            " DELETE"
        sqlFrom = _
            " FROM [" & myChildTable & "]"
        sqlWhere = _
            " WHERE [" & myChildTableChildID & "] = " & Me.funConvertDataType(myNode.Tag, myChildTableChildID)
        Me.tblChild.DeleteCommand.CommandText = _
            sqlDelete & _
            sqlFrom & _
            sqlWhere
        Me.dbConnection.Open()
        Try
            Me.tblChild.DeleteCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Delete Node: " & ex.Message)
        End Try
        Me.dbConnection.Close()
        'Remove Tree Node
        Me.treDataTree.SelectedNode = myNode
        Me.treDataTree.SelectedNode = myNode.Parent
        Me.treDataTree.Focus()
        Try
            myNode.Remove()
        Catch ex As Exception
            'Error Removing NodePlaceHolder
            'MsgBox("Error Removing: " & myNode.Text)
        End Try

    End Sub

    'Add Parent Node (Root to Child)
    Private Sub funAddParentNode(ByVal myNode As TreeNode, ByVal myNewParentNode As TreeNode)
        'Insert into ParentTable
        If Me.myParentTableAutoNumber = "" And Me.myCriteriaWhere = "" Then
            sqlInsert = _
                " INSERT INTO [" & myParentTable & "]" & _
                    "([" & myParentTableChildID & "], [" & myParentTableParentID & "])"
            sqlValues = _
                " VALUES (" & Me.funConvertDataType(myNode.Tag, myChildTableChildID) & ", " & Me.funConvertDataType(myNewParentNode.Tag, myChildTableChildID) & ")"
        ElseIf Me.myParentTableAutoNumber <> "" And Me.myCriteriaWhere = "" Then
            sqlInsert = _
                " INSERT INTO [" & myParentTable & "]" & _
                    "([" & myParentTableAutoNumber & "], [" & myParentTableChildID & "], [" & myParentTableParentID & "])"
            sqlValues = _
                " VALUES (" & Me.funGetNextParentAutoNumber(Me.myParentTableAutoNumber) & ", " & Me.funConvertDataType(myNode.Tag, myChildTableChildID) & ", " & Me.funConvertDataType(myNewParentNode.Tag, myChildTableChildID) & ")"
        ElseIf Me.myParentTableAutoNumber = "" And Me.myCriteriaWhere <> "" Then
            sqlInsert = _
                " INSERT INTO [" & myParentTable & "]" & _
                    "([" & myParentTableChildID & "], [" & myParentTableParentID & "], " & Me.myCriteriaInsert & ")"
            sqlValues = _
                " VALUES (" & Me.funConvertDataType(myNode.Tag, myChildTableChildID) & ", " & Me.funConvertDataType(myNewParentNode.Tag, myChildTableChildID) & ", " & Me.myCriteriaValue & ")"
        ElseIf Me.myParentTableAutoNumber <> "" And Me.myCriteriaWhere <> "" Then
            sqlInsert = _
                " INSERT INTO [" & myParentTable & "]" & _
                    "([" & myParentTableAutoNumber & "], [" & myParentTableChildID & "], [" & myParentTableParentID & "], " & Me.myCriteriaInsert & ")"
            sqlValues = _
                " VALUES (" & Me.funGetNextParentAutoNumber(Me.myParentTableAutoNumber) & ", " & Me.funConvertDataType(myNode.Tag, myChildTableChildID) & ", " & Me.funConvertDataType(myNewParentNode.Tag, myChildTableChildID) & ", " & Me.myCriteriaValue & ")"
        End If

        Me.tblChild.InsertCommand.CommandText = _
            sqlInsert & _
            sqlValues

    End Sub

    'Delete Parent Node (Child to Root)
    Private Sub funDeleteParentNode(ByVal myNode As TreeNode, ByVal myNewParentNode As TreeNode)
        'Delete from ParentTable
        sqlDelete = _
            " DELETE"
        sqlFrom = _
            " FROM [" & myParentTable & "]"
        sqlWhere = _
            " WHERE [" & myParentTableChildID & "] = " & Me.funConvertDataType(myNode.Tag, myChildTableChildID) & _
                " AND [" & myParentTableParentID & "] = " & Me.funConvertDataType(myNode.Parent.Tag, myChildTableChildID)
        If Me.myCriteriaWhere <> "" Then
            sqlWhere = _
                " WHERE [" & myParentTable & "].[" & myParentTableChildID & "] = " & Me.funConvertDataType(myNode.Tag, myChildTableChildID) & _
                    " AND [" & myParentTable & "].[" & myParentTableParentID & "] = " & Me.funConvertDataType(myNode.Parent.Tag, myChildTableChildID) & _
                    " AND " & Me.myCriteriaWhere
        End If
        Me.tblChild.DeleteCommand.CommandText = _
            sqlDelete & _
            sqlFrom & _
            sqlWhere

    End Sub

    'Update Parent Node (Move Child to New Parent)
    Private Sub funUpdateParentNode(ByVal myNode As TreeNode, ByVal myNewParentNode As TreeNode)
        Dim myParentNodeID As String

        'Update ParentTable
        If myNewParentNode Is Nothing Then
            myParentNodeID = "Null"
        Else
            myParentNodeID = myNewParentNode.Tag
        End If
        sqlUpdate = _
            " UPDATE [" & myParentTable & "]"
        sqlSet = _
            " SET [" & myParentTable & "].[" & myParentTableParentID & "] = " & Me.funConvertDataType(myParentNodeID, myChildTableChildID)
        If myChildTable = myParentTable Then
            sqlWhere = _
                " WHERE [" & myParentTable & "].[" & myChildTableChildID & "] = " & Me.funConvertDataType(myNode.Tag, myChildTableChildID)
            'Type Code ???
        Else
            sqlWhere = _
                " WHERE [" & myParentTable & "].[" & myParentTableChildID & "] = " & Me.funConvertDataType(myNode.Tag, myChildTableChildID) & _
                    " AND [" & myParentTable & "].[" & myParentTableParentID & "] = " & Me.funConvertDataType(myNode.Parent.Tag, myChildTableChildID)
            If Me.myCriteriaWhere <> "" Then
                sqlSet = sqlSet & ", " & Me.myCriteriaSet
                sqlWhere = _
                    " WHERE [" & myParentTable & "].[" & myParentTableChildID & "] = " & Me.funConvertDataType(myNode.Tag, myChildTableChildID) & _
                        " AND [" & myParentTable & "].[" & myParentTableParentID & "] = " & Me.funConvertDataType(myNode.Parent.Tag, myChildTableChildID) & _
                        " AND " & Me.myCriteriaWhere
            End If
        End If
        'Update Selected Node
        Me.tblChild.UpdateCommand.CommandText = _
            sqlUpdate & _
            sqlSet & _
            sqlWhere

    End Sub

    'Change Parent Node
    Private Sub funChangeParentNode(ByVal myNode As TreeNode, ByVal myNewParentNode As TreeNode)
        Dim myParentCheck As TreeNode = myNewParentNode
        Try
            While Not myParentCheck Is Nothing
                If myParentCheck.Tag = myNode.Tag Then
                    MsgBox("Action can not be completed because it will create a circular reference to this node.")
                    Exit Sub
                End If
                myParentCheck = myParentCheck.Parent
            End While
        Catch ex As Exception
            MsgBox("Change Parent Node: " & ex.Message)
        End Try
        'Update Parent
        Try
            Me.treDataTree.SelectedNode = myNewParentNode
            Me.dbConnection.Open()
            If myChildTable = myParentTable Then
                Me.funUpdateParentNode(myNode, myNewParentNode)
                Me.tblChild.UpdateCommand.ExecuteNonQuery()
            ElseIf myNode.Parent Is Nothing And Not myNewParentNode Is Nothing Then
                Me.funAddParentNode(myNode, myNewParentNode)
                Me.tblChild.InsertCommand.ExecuteNonQuery()
            ElseIf myNewParentNode Is Nothing And Not myNode.Parent Is Nothing Then
                Me.funDeleteParentNode(myNode, myNewParentNode)
                Me.tblChild.DeleteCommand.ExecuteNonQuery()
            ElseIf Not myNode.Parent Is myNewParentNode Then
                Me.funUpdateParentNode(myNode, myNewParentNode)
                Me.tblChild.UpdateCommand.ExecuteNonQuery()
            End If
            'Move Node
            If Not myNode Is Nothing Then
                myNode.Remove()
            End If
            If myNewParentNode Is Nothing Then
                Me.treDataTree.Nodes.Add(myNode)
                Me.treDataTree.SelectedNode = myNode
            Else
                myNewParentNode.Nodes.Add(myNode)
                Me.treDataTree.SelectedNode = myNewParentNode
            End If
            Me.funDisplayChildren(mySelectedNode.Tag)
            Me.treDataTree.SelectedNode = myNode
        Catch ex As Exception
            MsgBox("Change Parent Node: " & ex.Message)
        Finally
            Me.dbConnection.Close()
        End Try

    End Sub

    'Cut Node
    Private Sub funCutNode()
        If Not mySelectedNode Is Nothing Then
            myCutNode = mySelectedNode
            myCopyNode = Nothing
            myCutNode.ForeColor = Color.Gray
            Me.treDataTree.SelectedNode = myCutNode.Parent
            Me.funEnableEditMenuItems(True)
        End If

    End Sub

    'Copy Node
    Private Sub funCopyNode()
        If Not mySelectedNode Is Nothing Then
            myCopyNode = mySelectedNode
            If Not myCutNode Is Nothing Then
                myCutNode.ForeColor = Color.Black
                myCutNode = Nothing
            End If
            Me.funEnableEditMenuItems(True)
        End If

    End Sub

    'Paste Node
    Private Sub funPasteNode(Optional ByVal bolPasteAsRoot As Boolean = False, Optional ByVal bolDuplicate As Boolean = False)
        If Not mySelectedNode Is Nothing Or bolPasteAsRoot Then
            If Not myCutNode Is Nothing Then
                'Cut Paste
                If bolPasteAsRoot Then
                    Me.funChangeParentNode(myCutNode, Nothing)
                Else
                    If Not mySelectedNode Is Nothing Then
                        Me.funChangeParentNode(myCutNode, mySelectedNode)
                    End If
                End If
                myCutNode.ForeColor = Color.Black
                myCutNode = Nothing
            ElseIf Not myCopyNode Is Nothing Then
                'Copy Paste
                If bolDuplicate Then
                    If MsgBox("This will DUPLICATE Node '" & myCopyNode.Text & "'." & vbCr & vbCr & _
                        "Duplicate creates a copy of the original node.  A change to one node will affect all other copies." & vbCr & _
                        "Would you like to continue?", MsgBoxStyle.YesNo + MsgBoxStyle.Question, "Duplicate Node") = MsgBoxResult.Yes _
                    Then
                        Me.funDuplicateNode(myCopyNode, mySelectedNode, bolDuplicate)
                    End If
                Else
                    If MsgBox("This will COPY Node '" & myCopyNode.Text & "' and the entire Sub Tree to '" & IIf(bolPasteAsRoot, "Root", mySelectedNode.Text) & "'." & vbCr & vbCr & _
                        "Copy creates new nodes that will NOT be associated with the original nodes." & vbCr & _
                        "Would you like to continue?", MsgBoxStyle.YesNo + MsgBoxStyle.Question, "Duplicate Node") = MsgBoxResult.Yes _
                    Then
                        If bolPasteAsRoot Then
                            Me.funDuplicateNode(myCopyNode, Nothing, bolDuplicate)
                        Else
                            Me.funDuplicateNode(myCopyNode, mySelectedNode, bolDuplicate)
                        End If
                    End If
                End If
            End If
        End If

    End Sub

    Private Sub funDuplicateNode(ByVal mySourceNode As TreeNode, ByVal myParentNode As TreeNode, Optional ByVal bolDuplicate As Boolean = False)
        Dim myChildNode As TreeNode
        Dim myParentCheck As TreeNode = myParentNode
        Try
            'Check for Circular Reference to Child Node
            While Not myParentCheck Is Nothing
                If myParentCheck.Tag = mySourceNode.Tag Then
                    MsgBox("Action can not be completed because it will create a circular reference to this node.")
                    Exit Sub
                End If
                myParentCheck = myParentCheck.Parent
            End While
        Catch ex As Exception
            MsgBox("Duplicate Node: " & ex.Message)
        End Try
        If Not mySourceNode Is myParentNode Then
            If bolDuplicate Then
                Try
                    Me.dbConnection.Open()
                    Me.funAddParentNode(mySourceNode, myParentNode)
                    Me.tblChild.InsertCommand.ExecuteNonQuery()
                Catch ex As Exception
                Finally
                    Me.dbConnection.Close()
                End Try
                myParentNode.Nodes.Add(mySourceNode.Clone)
            Else
                'Modify Select Statement
                sqlFrom = _
                    " FROM [" & myChildTable & "] "
                sqlWhere = _
                    " WHERE [" & myChildTable & "].[" & myChildTableChildID & "] = " & Me.funConvertDataType(mySourceNode.Tag, myChildTableChildID)
                'Get Selected Node
                Me.tblChild.SelectCommand.CommandText = _
                    sqlSelect & _
                    sqlFrom & _
                    sqlWhere
                Me.dsNode.Clear()
                Me.tblChild.Fill(Me.dsNode)
                Me.funAddNode(myParentNode, Me.dsNode.Tables(myChildTable).Rows(0))
                myParentNode = mySelectedNode
                myParentNode.EndEdit(True)
                Me.treDataTree.SelectedNode = mySourceNode
                For Each myChildNode In mySourceNode.Nodes
                    Me.funDuplicateNode(myChildNode, myParentNode)
                Next
            End If
        End If

    End Sub

    Private Function funGetNextID(ByVal myField As String) As String
        Dim myColumn As DataColumn

        'Modify Select Statement
        sqlFrom = _
            " FROM [" & myChildTable & "] "
        sqlOrder = _
            " ORDER BY [" & myChildTable & "].[" & myChildTableChildID & "] DESC"
        'Get Selected Node
        Me.tblChild.SelectCommand.CommandText = _
            sqlSelect & _
            sqlFrom & _
            sqlOrder
        Me.dsTree.Clear()
        Me.tblChild.Fill(Me.dsTree)
        'Get Data Type
        myColumn = Me.dsTree.Tables(myChildTable).Columns(myField)
        Try
            If myColumn.DataType Is System.Type.GetType("System.String") Then
                'Prompt user for Text ID
                funGetNextID = "Null"
                funGetNextID = InputBox("What is the ID (Primary Key) for this Node?" & vbCr & vbCr & "The Last ID is listed below.", "Next ID", Me.dsTree.Tables(myChildTable).Rows(0).Item(myChildTableChildID))
            Else
                'Automatically Increment ID
                funGetNextID = 0
                funGetNextID = Me.dsTree.Tables(myChildTable).Rows(0).Item(myChildTableChildID) + 1
            End If
        Catch ex As Exception
        End Try

    End Function

    Private Function funGetNextParentAutoNumber(ByVal myField As String) As String
        Dim sqlSelect As String
        Dim sqlFrom As String
        Dim sqlOrder As String
        Dim tblAutoNumber As New DataTable
        Dim colAutoNumber As DataColumn
        Dim rowAutoNumber As DataRow

        'Modify Select Statement
        sqlSelect = _
            " SELECT *"
        sqlFrom = _
            " FROM [" & myParentTable & "] "
        sqlOrder = _
            " ORDER BY [" & myParentTable & "].[" & myField & "] DESC"
        'Get Selected Node
        Me.tblChild.SelectCommand.CommandText = _
            sqlSelect & _
            sqlFrom & _
            sqlOrder
        tblAutoNumber.Clear()
        Me.tblChild.Fill(tblAutoNumber)

        'Get Data Type
        colAutoNumber = tblAutoNumber.Columns(myField)
        Try
            If colAutoNumber.DataType Is System.Type.GetType("System.String") Then
                'Prompt user for Text ID
                funGetNextParentAutoNumber = "Null"
                funGetNextParentAutoNumber = InputBox("What is the ID (Primary Key) for this Node?" & vbCr & vbCr & "The Last ID is listed below.", "Next ID", Me.dsTree.Tables(myChildTable).Rows(0).Item(myChildTableChildID))
            Else
                'Automatically Increment ID
                funGetNextParentAutoNumber = 0
                rowAutoNumber = tblAutoNumber.Rows(0)
                funGetNextParentAutoNumber = rowAutoNumber(myField) + 1
            End If
        Catch ex As Exception
        End Try

    End Function

    'Get Children
    Private Sub funDisplayChildren(ByVal MyID As String)
        If MyID Is Nothing Then
            Me.dsChild.Clear()
            Exit Sub
        End If
        'Modify Select Statement
        sqlFrom = _
            " FROM ([" & myChildTable & "] " & _
                " LEFT OUTER JOIN [" & myParentTable & "] AS ParentTable" & _
                " ON [" & myChildTable & "].[" & myChildTableChildID & "] = [ParentTable].[" & myParentTableChildID & "])"
        sqlWhere = _
            " WHERE [ParentTable].[" & myParentTableParentID & "] = " & Me.funConvertDataType(MyID, myChildTableChildID)
        If Me.myCriteriaWhere <> "" Then
            sqlWhere = _
                " WHERE [ParentTable].[" & myParentTableParentID & "] = " & Me.funConvertDataType(MyID, myChildTableChildID) & _
                    " AND " & Me.myCriteriaParentTableWhere
        End If
        sqlOrder = _
            " ORDER BY [" & myChildTable & "].[" & myChildTableChildName & "]"
        'Display Children of Selected Node
        Me.tblChild.SelectCommand.CommandText = _
            sqlSelect & _
            sqlFrom & _
            sqlWhere & _
            sqlOrder
        Me.dsChild.Clear()
        Me.tblChild.Fill(Me.dsChild)
        'Set DataGridChildren to a formatted DataView of DataSetChild
        'Me.grdChildren.DataSource = myDataView

    End Sub

    'Get Description
    Private Sub funDisplayDescription(ByVal MyID As String)
        If MyID Is Nothing Then
            Me.txtDescription = Nothing
            Exit Sub
        End If
        'Modify Select Statement
        sqlFrom = _
            " FROM [" & myChildTable & "] "
        sqlWhere = _
            " WHERE [" & myChildTable & "].[" & myChildTableChildID & "] = " & Me.funConvertDataType(MyID, myChildTableChildID)
        'Get Selected Node
        Me.tblChild.SelectCommand.CommandText = _
            sqlSelect & _
            sqlFrom & _
            sqlWhere
        Me.dsNode.Clear()
        Me.tblChild.Fill(Me.dsNode)
        With Me.dsNode.Tables(myChildTable)
            If .Rows.Count > 0 Then
                'Display Description
                Me.txtDescription.Text = .Rows(0).Item(myChildTableChildDescription).ToString
            End If
        End With
        Me.txtDescription.ReadOnly = True

    End Sub

    Private Function funGetDescription(ByVal MyID As String) As String
        If MyID Is Nothing Then
            funGetDescription = ""
            Exit Function
        End If
        'Modify Select Statement
        sqlFrom = _
            " FROM [" & myChildTable & "] "
        sqlWhere = _
            " WHERE [" & myChildTable & "].[" & myChildTableChildID & "] = " & Me.funConvertDataType(MyID, myChildTableChildID)
        'Get Selected Node
        Me.tblChild.SelectCommand.CommandText = _
            sqlSelect & _
            sqlFrom & _
            sqlWhere
        Me.dsNode.Clear()
        Me.tblChild.Fill(Me.dsNode)
        With Me.dsNode.Tables(myChildTable)
            If .Rows.Count > 0 Then
                'Display Description
                funGetDescription = .Rows(0).Item(myChildTableChildDescription).ToString
            End If
        End With

    End Function

    Private Sub funResetSearchTree(ByVal myRootNode As TreeNode)
        Dim myNode As TreeNode

        If Not myRootNode Is Nothing Then
            myRootNode.ForeColor = Color.Black
            For Each myNode In myRootNode.Nodes
                funResetSearchTree(myNode)
            Next
        End If

    End Sub

    Private Sub funSearchTree(ByVal myFind As String, ByVal myStartNode As TreeNode, Optional ByVal mySearchType As String = "Find", Optional ByVal myReplaceWith As String = Nothing)
        Dim myNode As TreeNode

        If Not myStartNode Is Nothing Then
            If mySearchType = "Find" Then
                'Find
                Me.treDataTree.SelectedNode = myStartNode
                If UCase(myStartNode.Text) Like "*" & UCase(myFind) & "*" Then
                    myStartNode.ForeColor = Color.Blue
                    myFoundNode = myStartNode
                End If
            ElseIf mySearchType = "Replace" Then
                'Replace
                Me.treDataTree.SelectedNode = myStartNode
                If UCase(myStartNode.Text) Like "*" & UCase(myFind) & "*" Then
                    myStartNode.Text = myReplaceWith
                    myStartNode.ForeColor = Color.Blue
                    myFoundNode = myStartNode
                End If
            ElseIf mySearchType = "ReplaceNode" Then
                'Replace Node (After Edit)
                If myStartNode.Tag = myFind Then
                    myStartNode.Text = myReplaceWith
                    If myStartNode.Text = "" Then
                        myStartNode.Text = "(null)"
                    End If
                    myStartNode.ForeColor = Color.Green
                    myFoundNode = myStartNode
                End If
            ElseIf mySearchType = "ReplaceNodeID" Then
                'Replace Node (After Edit)
                If myStartNode.Tag = myFind Then
                    myStartNode.Tag = myReplaceWith
                    If myStartNode.Tag = "" Then
                        myStartNode.Tag = "(null)"
                    End If
                    myStartNode.ForeColor = Color.Green
                    'myStartNode.NodeFont = New Font(Me.treDataTree.Font, FontStyle.Bold)
                    myFoundNode = myStartNode
                End If
            End If
            'Children
            For Each myNode In myStartNode.Nodes
                funSearchTree(myFind, myNode, mySearchType, myReplaceWith)
            Next
        End If

    End Sub

#End Region

#Region "Resize"
    'Resize
    Private Sub frmDataTree_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Resize
        'Form Minimize
        If Me.Height < 125 Then
            Exit Sub
        End If
        'Maximum Text Description Height
        If Me.txtDescription.Height > Me.Height - Me.cintMinimumHeight Then
            Me.txtDescription.Top = cintMinimumHeight
            Me.txtDescription.Height = Me.Height - Me.txtDescription.Top - cintFormHeightOffset
        End If
        'Resize Height
        Me.grdChildren.Height = Me.txtDescription.Top - cintMenuHeightOffset - 3
        Me.treDataTree.Height = Me.txtDescription.Top - cintMenuHeightOffset - 3
        'Maximum Data Tree Width
        If Me.treDataTree.Width > Me.Width - cintMinimumWidth Then
            Me.treDataTree.Width = Me.Width - cintMinimumWidth
            Me.grdChildren.Left = Me.Width - cintMinimumWidth + 3
        Else
            Me.grdChildren.Width = Me.Width - Me.treDataTree.Width - cintFormWidthOffset
        End If
        'Resize Width
        Me.lblSelected.Width = Me.Width
        Me.txtDescription.Width = Me.Width

    End Sub

    Private Sub frmDataTree_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseMove
        If e.X >= Me.treDataTree.Width And e.X <= Me.grdChildren.Left And e.Button = MouseButtons.None Then
            'Change Mouse Pointer
            Me.Cursor = System.Windows.Forms.Cursors.SizeWE
        ElseIf e.Y >= Me.treDataTree.Height + Me.lblSelected.Height And e.Y <= Me.txtDescription.Top And e.Button = MouseButtons.None Then
            'Change Mouse Pointer
            Me.Cursor = System.Windows.Forms.Cursors.SizeNS
        ElseIf Me.Cursor Is System.Windows.Forms.Cursors.SizeNS And e.Button = MouseButtons.Left Then
            'Move Horizontal Seperator
            If e.Y < cintMinimumHeight Then
                Me.txtDescription.Top = cintMinimumHeight
                Me.txtDescription.Height = Me.Height - Me.txtDescription.Top - cintFormHeightOffset
                Me.treDataTree.Height = Me.txtDescription.Top - cintMenuHeightOffset - 3
                Me.grdChildren.Height = Me.txtDescription.Top - cintMenuHeightOffset - 3
            ElseIf e.Y > Me.Height - cintMinimumHeight - cintMenuHeightOffset Then
                Me.txtDescription.Top = Me.Height - cintMinimumHeight - cintMenuHeightOffset
                Me.txtDescription.Height = Me.Height - Me.txtDescription.Top - cintFormHeightOffset
                Me.treDataTree.Height = Me.txtDescription.Top - cintMenuHeightOffset - 3
                Me.grdChildren.Height = Me.txtDescription.Top - cintMenuHeightOffset - 3
            Else
                Me.txtDescription.Top = e.Y
                Me.txtDescription.Height = Me.Height - Me.txtDescription.Top - cintFormHeightOffset
                Me.treDataTree.Height = e.Y - cintMenuHeightOffset - 3
                Me.grdChildren.Height = e.Y - cintMenuHeightOffset - 3
            End If
        ElseIf Me.Cursor Is System.Windows.Forms.Cursors.SizeWE And e.Button = MouseButtons.Left Then
            'Move Vertical Seperator
            If e.X < cintMinimumWidth Then
                Me.treDataTree.Width = cintMinimumWidth
                Me.grdChildren.Left = cintMinimumWidth + 3
            ElseIf e.X > Me.Width - cintMinimumWidth Then
                Me.treDataTree.Width = Me.Width - cintMinimumWidth
                Me.grdChildren.Left = Me.Width - cintMinimumWidth + 3
            Else
                Me.treDataTree.Width = e.X
                Me.grdChildren.Left = e.X + 3
            End If
            Me.grdChildren.Width = Me.Width - Me.treDataTree.Width - cintFormWidthOffset
        End If

    End Sub

    Private Sub grdChildren_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles grdChildren.MouseMove
        'Restor Default Mouse After Resize
        'Me.Cursor = myCursor
        Me.Cursor = System.Windows.Forms.Cursors.Default
        'If e.X < 30 Then
        '    bolMouseOverRowHeader = True
        'Else
        '    bolMouseOverRowHeader = False
        'End If
        If e.Y < 20 Then
            bolMouseOverColumnHeader = True
        Else
            bolMouseOverColumnHeader = False
        End If

    End Sub

#End Region

#Region "Drag and Drop"

    'Select Source Node
    Private Sub treDataTree_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles treDataTree.MouseDown
        mySourceNode = Me.treDataTree.GetNodeAt(e.X, e.Y)
        myStartX = e.X
        myStartY = e.Y
        If mySourceNode Is Nothing Then
            Me.mnuTreeViewFind.Enabled = False
            Me.mnuTreeViewSeparator1.Enabled = False
            Me.mnuTreeViewCut.Enabled = False
            Me.mnuTreeViewCopy.Enabled = False
            Me.mnuTreeViewPaste.Enabled = False
            'Me.mnuTreeViewPasteToRoot.Enabled = False
            Me.mnuTreeViewDuplicate.Enabled = False
            Me.mnuTreeViewSeparator2.Enabled = False
            Me.mnuTreeViewAddNode.Enabled = False
            'Me.mnuTreeViewNewNode.Enabled = False
            Me.mnuTreeViewEditNode.Enabled = False
            Me.mnuTreeViewDeleteNode.Enabled = False
        Else
            Me.funEnableNodeMenuItems(True)
            Me.funEnableEditMenuItems(True)
        End If
        If e.Button = MouseButtons.Right Then
            Me.treDataTree.SelectedNode = mySourceNode
        End If

    End Sub

    'Begin Drag and Drop
    Private Sub treDataTree_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles treDataTree.MouseMove
        'Restore Default Cursor After Resize
        'Me.Cursor = myCursor
        Me.Cursor = System.Windows.Forms.Cursors.Default
        'Start Drag and Drop
        Try
            If e.Button = MouseButtons.Left Then
                Me.treDataTree.SelectedNode = mySourceNode
                If myStartX > e.X + 5 Or myStartY > e.Y + 5 Or _
                    myStartX < e.X - 5 Or myStartY < e.Y - 5 _
                Then
                    Me.treDataTree.DoDragDrop(mySourceNode, DragDropEffects.Copy)
                End If
            End If
        Catch ex As Exception
            'MsgBox("Data Tree Mouse Move: " & ex.Message)
        End Try

    End Sub

    'Change Mouse Icon
    Private Sub treDataTree_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles treDataTree.DragEnter
        'Set Drag and Drop Effect (Copy, Move)
        e.Effect = DragDropEffects.Copy

    End Sub

    Private Sub treDataTree_DragLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles treDataTree.DragLeave
        Me.mnuRefresh_Click(sender, e)

    End Sub

    'Drop At Destination Node
    Private Sub treDataTree_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles treDataTree.MouseUp
        myDestinationNode = Me.treDataTree.GetNodeAt(e.X, e.Y)

    End Sub

    'Execute Drag and Drop
    Private Sub treDataTree_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles treDataTree.DragDrop
        'Convert Screen Location to TreeView Location
        Dim myPoint As New Point(e.X, e.Y)
        myPoint = Me.treDataTree.PointToClient(myPoint)
        myDestinationNode = Me.treDataTree.GetNodeAt(myPoint)
        'Change Parent
        If Not mySourceNode Is myDestinationNode Then
            funChangeParentNode(mySourceNode, myDestinationNode)
        End If

    End Sub

    'Highlight Nodes
    Private Sub treDataTree_DragOver(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles treDataTree.DragOver
        'Convert Screen Location to TreeView Location
        Dim myPoint As New Point(e.X, e.Y)
        myPoint = Me.treDataTree.PointToClient(myPoint)
        If Not myHoverNode Is Me.treDataTree.GetNodeAt(myPoint) Then
            myHoverNode = Me.treDataTree.GetNodeAt(myPoint)
            myHoverTime = Now
            'Refresh Nodes
            Dim myNode As TreeNode
            myNode = Me.treDataTree.Nodes(0)
            While Not myNode Is Nothing
                myNode.ForeColor = Color.Black
                'myNode.BackColor = Color.White
                myNode = myNode.NextVisibleNode()
            End While
            'Highlight Node
            myHoverNode.ForeColor = Color.Red
            'myHoverNode.BackColor = Color.FromKnownColor(KnownColor.Highlight)
        ElseIf myHoverTime.AddSeconds(1) < Now Then
            'Select Node
            Me.treDataTree.SelectedNode = myHoverNode
            Me.treDataTree.SelectedNode = mySourceNode
            myHoverTime = Now.AddMinutes(1)
        End If

    End Sub

#End Region

#Region "Functions"

    'Connect to Database
    Private Function funGetAccessDatabase(Optional ByVal bolConfigurationFile As Boolean = False) As Boolean
        'Dim strConfigurationFile As String

        'Browse for Access Database
        If Not bolConfigurationFile Then
            Me.filDatabase.FileName = ""
            Me.filDatabase.InitialDirectory = Application.StartupPath
            Me.filDatabase.ShowDialog()
            strDatabase = Me.filDatabase.FileName
            If strDatabase = "" Then
                funGetAccessDatabase = False
                Exit Function
            End If
            'Create Database Connection
            Me.dbConnection.ConnectionString = _
                " Provider=Microsoft.Jet.OLEDB.4.0;" & _
                " Data Source=" & strDatabase & ";"
        Else
            strDatabase = Mid(Me.dbConnection.ConnectionString, InStr(Me.dbConnection.ConnectionString, "Data Source=") + 12)
        End If
        myConnectString = Me.dbConnection.ConnectionString
        'Test Database Connection
        Try
            Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
            Me.lblSelected.Text = "Searching for Database ..."
            Me.txtDescription.Text = ""
            Me.dbConnection.Open()
            Me.dbConnection.Close()
            Me.Cursor = System.Windows.Forms.Cursors.Default
            Me.lblSelected.Text = "Database Connected"
            Me.txtDescription.Text = "Next Step: Set Table Structure"
            funGetAccessDatabase = True
        Catch ex As Exception
            Me.lblSelected.Text = "Database Not Found"
            Me.txtDescription.Text = "Next Step: Open Database Connection"
            funGetAccessDatabase = False
        End Try
        While InStr(strDatabase, "\") > 0
            strDatabase = Mid(strDatabase, InStr(strDatabase, "\") + 1)
        End While
        If InStr(strDatabase, ";") > 0 Then
            strDatabase = Mid(strDatabase, 1, InStr(strDatabase, ";") - 1)
        End If
        'Me.funSetFormName(myConfigurationFile)

    End Function

    'Connect to Database
    Private Function funGetSQLDatabase(Optional ByVal bolConfigurationFile As Boolean = False) As Boolean
        'Dim strConfigurationFile As String
        Dim myDataServerForm As New frmDataServer

        'Test Connection
        'Test Database Connection
        If bolConfigurationFile Then
            Try
                Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
                Me.lblSelected.Text = "Searching for Database ..."
                Me.txtDescription.Text = ""
                Me.dbConnection.Open()
                Me.dbConnection.Close()
                Me.Cursor = System.Windows.Forms.Cursors.Default
                Me.lblSelected.Text = "Database Connected"
                Me.txtDescription.Text = "Next Step: Set Table Structure"
                funGetSQLDatabase = True
            Catch
                Me.lblSelected.Text = "Database Not Found"
                Me.txtDescription.Text = "Next Step: Open Database Connection"
                bolConfigurationFile = False
            End Try
        End If
        If Not bolConfigurationFile Then
            'Set SQL Server Login
            If myDataServerForm.ShowDialog(Me) = DialogResult.OK Then
                Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
                Me.bolSaveConfigurationFile = True
                'Create Database Connection
                Me.dbConnection.ConnectionString = _
                    "Provider=SQLOLEDB.1;" & _
                    "Persist Security Info=False;" & _
                    "User ID=" & strUserName & ";" & _
                    "Initial Catalog=" & strDatabase & ";" & _
                    "Data Source=" & strServer & ";" & _
                    "Use Procedure for Prepare=1;" & _
                    "Auto Translate=True;" & _
                    "Packet Size=4096;" & _
                    "Workstation ID=;" & _
                    "Use Encryption for Data=False;" & _
                    "Tag with column collation when possible=False;" & _
                    "Password=" & strPassword & ""
                myConnectString = Me.dbConnection.ConnectionString
                'Test Database Connection
                Try
                    Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
                    Me.lblSelected.Text = "Searching for Database ..."
                    Me.txtDescription.Text = ""
                    Me.dbConnection.Open()
                    Me.dbConnection.Close()
                    Me.Cursor = System.Windows.Forms.Cursors.Default
                    Me.lblSelected.Text = "Database Connected"
                    Me.txtDescription.Text = "Next Step: Set Table Structure"
                    funGetSQLDatabase = True
                Catch ex As Exception
                    Me.lblSelected.Text = "Database Not Found: " & vbCr & ex.Message
                    Me.txtDescription.Text = "Next Step: Open Database Connection"
                    funGetSQLDatabase = False
                End Try
                While InStr(strDatabase, "\") > 0
                    strDatabase = Mid(strDatabase, InStr(strDatabase, "\") + 1)
                End While
            End If
        End If
        'Me.funSetFormName(myConfigurationFile)
        myDataServerForm.Dispose()

    End Function

    'Connect to Database
    Private Function funGetMySQLDatabase(Optional ByVal bolConfigurationFile As Boolean = False) As Boolean
        'Dim strConfigurationFile As String
        Dim myDataServerForm As New frmDataServer

        'Test Connection
        'Test Database Connection
        If bolConfigurationFile Then
            Try
                Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
                Me.lblSelected.Text = "Searching for Database ..."
                Me.txtDescription.Text = ""
                Me.dbConnection.Open()
                Me.dbConnection.Close()
                Me.Cursor = System.Windows.Forms.Cursors.Default
                Me.lblSelected.Text = "Database Connected"
                Me.txtDescription.Text = "Next Step: Set Table Structure"
                funGetMySQLDatabase = True
            Catch
                Me.lblSelected.Text = "Database Not Found"
                Me.txtDescription.Text = "Next Step: Open Database Connection"
                bolConfigurationFile = False
            End Try
        End If
        If Not bolConfigurationFile Then
            'Set SQL Server Login
            If myDataServerForm.ShowDialog(Me) = DialogResult.OK Then
                Me.bolSaveConfigurationFile = True
                'Create Database Connection
                Me.dbConnection.ConnectionString = _
                    "Provider=SQLOLEDB.1;" & _
                    "Server=" & strServer & ";" & _
                    "Database=" & strDatabase & ";" & _
                    "User ID=" & strUserName & ";" & _
                    "Password=" & strPassword & ""
                myConnectString = Me.dbConnection.ConnectionString
                'Test Database Connection
                Try
                    Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
                    Me.lblSelected.Text = "Searching for Database ..."
                    Me.txtDescription.Text = ""
                    Me.dbConnection.Open()
                    Me.dbConnection.Close()
                    Me.Cursor = System.Windows.Forms.Cursors.Default
                    Me.lblSelected.Text = "Database Connected"
                    Me.txtDescription.Text = "Next Step: Set Table Structure"
                    funGetMySQLDatabase = True
                Catch ex As Exception
                    Me.lblSelected.Text = "Database Not Found: " & vbCr & ex.Message
                    Me.txtDescription.Text = "Next Step: Open Database Connection"
                    funGetMySQLDatabase = False
                End Try
                While InStr(strDatabase, "\") > 0
                    strDatabase = Mid(strDatabase, InStr(strDatabase, "\") + 1)
                End While
            End If
        End If
        'Me.funSetFormName(myConfigurationFile)
        myDataServerForm.Dispose()

    End Function

    Private Sub funOpenConfigurationFile(Optional ByVal bolOpenDefault As Boolean = False)
        Dim myResponse As DialogResult
        Dim strConfigurationFile As String

        If Me.funSaveConfigurationChanges = DialogResult.OK Then
            'Browse for Access Database
            If bolOpenDefault Then
                strConfigurationFile = Application.StartupPath & "\Default.dtc"
            Else
                Me.filOpenConfigurationFile.FileName = ""
                Me.filOpenConfigurationFile.InitialDirectory = Application.StartupPath
                myResponse = Me.filOpenConfigurationFile.ShowDialog()
                strConfigurationFile = Me.filOpenConfigurationFile.FileName
                If strConfigurationFile = "" Then
                    Exit Sub
                End If
            End If

            ' Open the file to read from.
            Dim sr As System.IO.StreamReader
            Try
                sr = System.IO.File.OpenText(strConfigurationFile)
            Catch ex As Exception
                'No Default Configuration File
                Exit Sub
            End Try

            If bolOpenDefault Then
                strConfigurationFile = ""
                bolOpenDefault = False
            End If

            Try
                'Database
                sr.ReadLine()
                'Set ConnectionString only if not set
                'If Me.dbConnection.ConnectionString = "" Then
                '    Me.dbConnection.ConnectionString = sr.ReadLine
                'Else
                '    sr.ReadLine()
                'End If
                'Always set ConnectionString
                Me.dbConnection.ConnectionString = sr.ReadLine
                'Parse ConnectionString

                Dim strDatabaseConnection As String
                Dim intStart As Integer
                Dim intEnd As Integer
                strDatabaseConnection = Me.dbConnection.ConnectionString
                If strDatabaseConnection Like "Provider=SQLOLEDB*" Then
                    'UserName
                    intStart = InStr(strDatabaseConnection, "User ID=") + 8
                    intEnd = InStr(intStart, strDatabaseConnection, ";") - 0
                    Me.strUserName = Mid(strDatabaseConnection, intStart, intEnd - intStart)
                    'Password
                    intStart = InStr(strDatabaseConnection, "Password=") + 9
                    intEnd = InStr(intStart, strDatabaseConnection, ";") - 0
                    Me.strPassword = Mid(strDatabaseConnection, intStart, intEnd - intStart)
                    'Database
                    intStart = InStr(strDatabaseConnection, "Initial Catalog=") + 16
                    intEnd = InStr(intStart, strDatabaseConnection, ";") - 0
                    Me.strDatabase = Mid(strDatabaseConnection, intStart, intEnd - intStart)
                    'Server
                    intStart = InStr(strDatabaseConnection, "Data Source=") + 12
                    intEnd = InStr(intStart, strDatabaseConnection, ";") - 0
                    Me.strServer = Mid(strDatabaseConnection, intStart, intEnd - intStart)
                Else
                    Me.strUserName = ""
                    Me.strPassword = ""
                    Me.strDatabase = ""
                    Me.strServer = ""
                End If
                Me.lblSelected.Text = "Configuration File Loaded"
                Me.txtDescription.Text = "Next Step: Open Database Connection"
                myConnectString = Me.dbConnection.ConnectionString
            Catch ex As Exception
                'Database Not Found
                MsgBox("Database Not Found")
                Me.dbConnection.ConnectionString = ""
            End Try

            Try
                'Child Table
                sr.ReadLine()
                myChildTable = sr.ReadLine
                myChildTableChildID = sr.ReadLine
                myChildTableChildName = sr.ReadLine
                myChildTableChildDescription = sr.ReadLine
                'Parent Table
                sr.ReadLine()
                myParentTable = sr.ReadLine
                myParentTableChildID = sr.ReadLine
                myParentTableParentID = sr.ReadLine
                myParentTableAutoNumber = sr.ReadLine
                'Criteria
                If myParentTableAutoNumber = "---Criteria---" Then
                    myParentTableAutoNumber = ""
                Else
                    sr.ReadLine()
                End If
                myCriteriaWhere = sr.ReadLine
                myCriteriaInsert = sr.ReadLine
                myCriteriaValue = sr.ReadLine
                myCriteriaSet = sr.ReadLine
                myCriteriaParentTableWhere = sr.ReadLine
                myConfigurationFile = strConfigurationFile
            Catch ex As Exception
                'Table Structure Error
                MsgBox("Table Structure Not Found")
                Exit Sub
            End Try

            Try
                'Open Database Connection
                If myConnectString Like "Provider=Microsoft*" Then
                    If Me.funGetAccessDatabase(True) Then
                        Me.funEnableFileMenuItems(False)
                        Me.funCreateComponents()
                    End If
                ElseIf myConnectString Like "Provider*" Then
                    If Me.funGetSQLDatabase(True) Then
                        Me.funEnableFileMenuItems(False)
                        Me.funCreateComponents()
                    End If
                End If
            Catch ex As Exception
                'Database Connection Error
                'MsgBox("Unable to connect to Database")
            End Try

            'Rows
            sr.ReadLine()
            Try
                intRowHeight = sr.ReadLine
            Catch ex As Exception
                intRowHeight = 30
            End Try

            Try
                'Columns
                sr.ReadLine()
                Me.myColumnList.Items.Clear()
                Do While sr.Peek() >= 0
                    Dim myItem As String
                    Dim IsChecked As CheckState
                    myItem = sr.ReadLine
                    Me.grdChildren.TableStyles(0).GridColumnStyles(myItem).Width = sr.ReadLine
                    If Me.grdChildren.TableStyles(0).GridColumnStyles(myItem).Width > 0 Then
                        IsChecked = CheckState.Checked
                    Else
                        IsChecked = CheckState.Unchecked
                    End If
                    Me.myColumnList.Items.Add(myItem, IsChecked)
                Loop
                Me.funSetColumnOrder(Me.grdChildren, Me.dsChild.Tables(myChildTable), Me.myColumnList, Me.mnuDataGrid)
                Me.funShowColumns()
            Catch ex As Exception
                'Data Grid Layout Error
                'MsgBox("Data Grid Layout Not Found")
            End Try

            Me.mnuClose.Enabled = True
            Me.mnuExport.Enabled = True
            Me.bolSaveConfigurationFile = False
            Try
                sr.Close()
            Catch ex As Exception

            End Try
        End If

    End Sub

    Private Sub funSaveConfigurationFile()
        Dim strConfigurationFile As String

        If Not myConfigurationFile = "" Then
            strConfigurationFile = myConfigurationFile
            ' Create a file to write to.
            Dim sw As System.IO.StreamWriter
            Try
                sw = System.IO.File.CreateText(strConfigurationFile)
            Catch ex As Exception
                'Default Configuration File
                MsgBox(ex.Message)
                Exit Sub
            End Try
            'Database
            sw.WriteLine("---Database---")
            sw.Write(Trim(Me.dbConnection.ConnectionString))
            sw.WriteLine("Password=" & Me.strPassword & ";")
            'Table Structure
            sw.WriteLine("---Child Table---")
            sw.WriteLine(myChildTable)
            sw.WriteLine(myChildTableChildID)
            sw.WriteLine(myChildTableChildName)
            sw.WriteLine(myChildTableChildDescription)
            sw.WriteLine("---Parent Table---")
            sw.WriteLine(myParentTable)
            sw.WriteLine(myParentTableChildID)
            sw.WriteLine(myParentTableParentID)
            sw.WriteLine(myParentTableAutoNumber)
            'Criteria
            sw.WriteLine("---Criteria---")
            sw.WriteLine(myCriteriaWhere)
            sw.WriteLine(myCriteriaInsert)
            sw.WriteLine(myCriteriaValue)
            sw.WriteLine(myCriteriaSet)
            sw.WriteLine(myCriteriaParentTableWhere)
            'Rows
            sw.WriteLine("---Rows---")
            Try
                sw.WriteLine(Me.grdChildren.TableStyles(0).PreferredRowHeight)
            Catch ex As Exception
                sw.WriteLine(75)
            End Try
            'Columns
            sw.WriteLine("---Columns---")
            Dim intCounter As Integer
            'Create Column List Box
            For intCounter = 0 To Me.myColumnList.Items.Count - 1
                sw.WriteLine(Me.myColumnList.Items(intCounter))
                'sw.WriteLine(Me.myColumnList.CheckedItems.Contains(Me.myColumnList.Items(intCounter)))
                sw.WriteLine(Me.grdChildren.TableStyles(0).GridColumnStyles(intCounter).Width)
            Next
            sw.Flush()
            sw.Close()
            Me.mnuClose.Enabled = True
            Me.mnuExport.Enabled = True
        Else
            Me.mnuSaveAs_Click(Nothing, Nothing)
        End If
        Me.funSetFormName(myConfigurationFile)
        Me.bolSaveConfigurationFile = False

    End Sub

    '***** REMOVED 08/19/08 *****
    'Private Sub funExportOrganizationChart(ByVal xlSheet As Excel.Worksheet)
    '    Dim myTable As DataTable
    '    Dim myRow As DataRow
    '    Dim intRow As Integer
    '    Dim intColumn As Integer

    '    'Write Header
    '    xlSheet.Cells(1, 1) = "Unique_ID"
    '    xlSheet.Cells(1, 2) = "Name"
    '    xlSheet.Cells(1, 3) = "Title"
    '    xlSheet.Cells(1, 4) = "Reports_To"
    '    xlSheet.Cells(1, 5) = "Master_Shape"
    '    'Get Root(s)
    '    sqlSelect = _
    '            " SELECT *"
    '    sqlFrom = _
    '        " FROM ([" & myChildTable & "]" & _
    '            " LEFT OUTER JOIN [" & myParentTable & "] AS ParentTable" & _
    '            " ON [" & myChildTable & "].[" & myChildTableChildID & "] = [ParentTable].[" & myParentTableChildID & "])"
    '    sqlWhere = _
    '        " WHERE [ParentTable].[" & myParentTableParentID & "] IS NULL"
    '    If Me.myCriteriaWhere <> "" Then
    '        sqlWhere = _
    '            " WHERE [" & myChildTable & "].[" & myChildTableChildID & "] Not In (" & _
    '                " SELECT [" & myParentTable & "].[" & myParentTableChildID & "]" & _
    '                " FROM [" & myParentTable & "]" & _
    '                " WHERE " & Me.myCriteriaWhere & ")"
    '    End If
    '    sqlOrder = _
    '        " ORDER BY [" & myChildTableChildName & "]"
    '    sqlOrder = ""
    '    Me.tblChild.SelectCommand.CommandText = _
    '        sqlSelect & _
    '        sqlFrom & _
    '        sqlWhere & _
    '        sqlOrder
    '    Me.dsNode.Clear()
    '    Me.tblChild.Fill(Me.dsNode)
    '    myTable = Me.dsNode.Tables(myChildTable)
    '    'Write Root(s)
    '    intRow = 1
    '    For Each myRow In myTable.Rows
    '        If Not IsDBNull(myRow.Item(myChildTableChildID)) Then
    '            intRow = intRow + 1
    '            xlSheet.Cells(intRow, 1) = "ID" & myRow.Item(myChildTableChildID)
    '            xlSheet.Cells(intRow, 2) = myRow.Item(myChildTableChildName)
    '            xlSheet.Cells(intRow, 3) = ""
    '            xlSheet.Cells(intRow, 4) = ""
    '            xlSheet.Cells(intRow, 5) = 1
    '        End If
    '    Next
    '    'Get Tree
    '    sqlSelect = _
    '            " SELECT *"
    '    sqlFrom = _
    '        " FROM ([" & myChildTable & "]" & _
    '            " LEFT OUTER JOIN [" & myParentTable & "]" & _
    '            " ON [" & myChildTable & "].[" & myChildTableChildID & "] = [" & myParentTable & "].[" & myParentTableChildID & "])"
    '    If Me.myCriteriaWhere <> "" Then
    '        sqlWhere = _
    '            " WHERE " & Me.myCriteriaWhere & ""
    '    End If
    '    sqlOrder = _
    '        " ORDER BY [" & myChildTableChildName & "]"
    '    sqlOrder = ""
    '    Me.tblChild.SelectCommand.CommandText = _
    '        sqlSelect & _
    '        sqlFrom & _
    '        sqlWhere & _
    '        sqlOrder
    '    Me.dsNode.Clear()
    '    Me.tblChild.Fill(Me.dsNode)
    '    myTable = Me.dsNode.Tables(myChildTable)
    '    'Write Tree
    '    For Each myRow In myTable.Rows
    '        If Not IsDBNull(myRow.Item(myChildTableChildID)) Then
    '            intRow = intRow + 1
    '            xlSheet.Cells(intRow, 1) = "ID" & myRow.Item(myChildTableChildID)
    '            xlSheet.Cells(intRow, 2) = myRow.Item(myChildTableChildName)
    '            xlSheet.Cells(intRow, 3) = ""
    '            xlSheet.Cells(intRow, 4) = "ID" & myRow.Item(myParentTableParentID)
    '            xlSheet.Cells(intRow, 5) = 2
    '        End If
    '    Next

    'End Sub

    Private Sub funExportOrganizationChartText(ByVal sw As System.IO.StreamWriter)
        Dim myTable As DataTable
        Dim myRow As DataRow
        Dim intRow As Integer
        'Dim intColumn As Integer

        'Write Header
        sw.WriteLine("Unique_ID,Name,Title,Reports_To,Master_Shape")
        'Get Root(s)
        sqlSelect = _
                " SELECT *"
        sqlFrom = _
            " FROM ([" & myChildTable & "]" & _
                " LEFT OUTER JOIN [" & myParentTable & "] AS ParentTable" & _
                " ON [" & myChildTable & "].[" & myChildTableChildID & "] = [ParentTable].[" & myParentTableChildID & "])"
        sqlWhere = _
            " WHERE [ParentTable].[" & myParentTableParentID & "] IS NULL"
        If Me.myCriteriaWhere <> "" Then
            sqlWhere = _
                " WHERE [" & myChildTable & "].[" & myChildTableChildID & "] Not In (" & _
                    " SELECT [" & myParentTable & "].[" & myParentTableChildID & "]" & _
                    " FROM [" & myParentTable & "]" & _
                    " WHERE " & Me.myCriteriaWhere & ")"
        End If
        sqlOrder = _
            " ORDER BY [" & myChildTableChildName & "]"
        sqlOrder = ""
        Me.tblChild.SelectCommand.CommandText = _
            sqlSelect & _
            sqlFrom & _
            sqlWhere & _
            sqlOrder
        Me.dsNode.Clear()
        Me.tblChild.Fill(Me.dsNode)
        myTable = Me.dsNode.Tables(myChildTable)
        'Write Root(s)
        intRow = 1
        For Each myRow In myTable.Rows
            If Not IsDBNull(myRow.Item(myChildTableChildID)) Then
                intRow = intRow + 1
                sw.WriteLine( _
                    "ID" & myRow.Item(myChildTableChildID) & "," & _
                    myRow.Item(myChildTableChildName) & "," & _
                    "," & _
                    "," & _
                    "1")
            End If
        Next
        'Get Tree
        sqlSelect = _
                " SELECT *"
        sqlFrom = _
            " FROM ([" & myChildTable & "]" & _
                " LEFT OUTER JOIN [" & myParentTable & "]" & _
                " ON [" & myChildTable & "].[" & myChildTableChildID & "] = [" & myParentTable & "].[" & myParentTableChildID & "])"
        If Me.myCriteriaWhere <> "" Then
            sqlWhere = _
                " WHERE " & Me.myCriteriaWhere & ""
        End If
        sqlOrder = _
            " ORDER BY [" & myChildTableChildName & "]"
        sqlOrder = ""
        Me.tblChild.SelectCommand.CommandText = _
            sqlSelect & _
            sqlFrom & _
            sqlWhere & _
            sqlOrder
        Me.dsNode.Clear()
        Me.tblChild.Fill(Me.dsNode)
        myTable = Me.dsNode.Tables(myChildTable)
        'Write Tree
        For Each myRow In myTable.Rows
            If Not IsDBNull(myRow.Item(myChildTableChildID)) Then
                intRow = intRow + 1
                sw.WriteLine( _
                    "ID" & myRow.Item(myChildTableChildID) & "," & _
                    myRow.Item(myChildTableChildName) & "," & _
                    "," & _
                    "ID" & myRow.Item(myParentTableParentID) & "," & _
                    "2")
            End If
        Next
        sw.Flush()

    End Sub

    '***** REMOVED 08/19/08 *****
    'Private Sub funExportIndentedHierarchy(ByVal xlSheet As Excel.Worksheet, Optional ByVal myStartNode As TreeNode = Nothing)
    '    Dim myNode As TreeNode

    '    If Me.treDataTree.GetNodeCount(True) = 0 Then
    '        MsgBox("Please Create a Tree first.")
    '        Exit Sub
    '    End If
    '    intRow = 1
    '    'Header Row
    '    xlSheet.Cells(intRow, 1) = "ID"
    '    xlSheet.Cells(intRow, 2) = "Tier1"
    '    xlSheet.Cells(intRow, 3) = "Tier2"
    '    xlSheet.Cells(intRow, 4) = "Tier3"
    '    xlSheet.Cells(intRow, 5) = "Tier4"
    '    xlSheet.Cells(intRow, 6) = "Tier5"
    '    xlSheet.Cells(intRow, 7) = "Tier6"
    '    xlSheet.Cells(intRow, 8) = "Tier7"
    '    xlSheet.Cells(intRow, 9) = "Tier8"
    '    xlSheet.Cells(intRow, 10) = "Tier9"
    '    xlSheet.Cells(intRow, 11) = "Description"
    '    intRow = intRow + 1
    '    'Node
    '    If myStartNode Is Nothing Then
    '        'Entire Tree
    '        For Each myNode In Me.treDataTree.Nodes
    '            Me.funExportIndentedHierarchyTree(xlSheet, myNode, 0)
    '        Next myNode
    '        Me.treDataTree.SelectedNode = Me.treDataTree.Nodes(0)
    '    Else
    '        'SubTree Only
    '        Me.funExportIndentedHierarchyTree(xlSheet, myStartNode, 0)
    '        Me.treDataTree.SelectedNode = myStartNode
    '    End If

    'End Sub

    '***** REMOVED 08/19/08 *****
    'Private Sub funExportIndentedHierarchyTree(ByVal xlSheet As Excel.Worksheet, ByVal myStartNode As TreeNode, ByVal intLevel As Integer)
    '    Dim myNode As TreeNode
    '    Dim intCounter As Integer
    '    Dim intColumn As Integer
    '    Dim strLevel As String
    '    Dim strSpacer As String

    '    If Not myStartNode Is Nothing Then
    '        Me.treDataTree.SelectedNode = myStartNode
    '        intColumn = 1
    '        'ID
    '        xlSheet.Cells(intRow, intColumn) = myStartNode.Tag
    '        For intCounter = 0 To intLevel
    '            intColumn = intColumn + 1
    '        Next
    '        'Name
    '        xlSheet.Cells(intRow, intColumn) = myStartNode.Text
    '        For intCounter = intLevel To 8
    '            intColumn = intColumn + 1
    '        Next
    '        'Description
    '        xlSheet.Cells(intRow, intColumn) = funGetDescription(myStartNode.Tag)
    '        'Children
    '        For Each myNode In myStartNode.Nodes
    '            intRow = intRow + 1
    '            funExportIndentedHierarchyTree(xlSheet, myNode, intLevel + 1)
    '        Next
    '    End If

    'End Sub

    Private Sub funExportIndentedHierarchyText(ByVal sw As System.IO.StreamWriter, Optional ByVal myStartNode As TreeNode = Nothing)
        Dim myNode As TreeNode

        If Me.treDataTree.GetNodeCount(True) = 0 Then
            MsgBox("Please Create a Tree first.")
            Exit Sub
        End If
        intRow = 1
        'Header Row
        sw.WriteLine( _
            """ID""" & "," & _
            """Tier1""" & "," & _
            """Tier2""" & "," & _
            """Tier3""" & "," & _
            """Tier4""" & "," & _
            """Tier5""" & "," & _
            """Tier6""" & "," & _
            """Tier7""" & "," & _
            """Tier8""" & "," & _
            """Tier9""" & "," & _
            """Description""")
        intRow = intRow + 1
        sw.Flush()
        'Node
        If myStartNode Is Nothing Then
            'Entire Tree
            For Each myNode In Me.treDataTree.Nodes
                Me.funExportIndentedHierarchyTreeText(sw, myNode, 0)
            Next myNode
            Me.treDataTree.SelectedNode = Me.treDataTree.Nodes(0)
        Else
            'SubTree Only
            Me.funExportIndentedHierarchyTreeText(sw, myStartNode, 0)
            Me.treDataTree.SelectedNode = myStartNode
        End If

    End Sub

    Private Sub funExportIndentedHierarchyTreeText(ByVal sw As System.IO.StreamWriter, ByVal myStartNode As TreeNode, ByVal intLevel As Integer)
        Dim myNode As TreeNode
        Dim intCounter As Integer
        Dim intColumn As Integer
        'Dim strLevel As String
        'Dim strSpacer As String

        If Not myStartNode Is Nothing Then
            Me.treDataTree.SelectedNode = myStartNode
            intColumn = 1
            'ID
            sw.Write(myStartNode.Tag & "")
            For intCounter = 0 To intLevel
                intColumn = intColumn + 1
                sw.Write(",")
            Next
            'Name
            sw.Write(myStartNode.Text & "")
            For intCounter = intLevel To 8
                intColumn = intColumn + 1
                sw.Write(",")
            Next
            'Description
            sw.WriteLine("""" & funGetDescription(myStartNode.Tag) & """")
            'Children
            For Each myNode In myStartNode.Nodes
                intRow = intRow + 1
                funExportIndentedHierarchyTreeText(sw, myNode, intLevel + 1)
            Next
            sw.Flush()
        End If

    End Sub

    'Create Components
    Public Sub funCreateComponents()
        Dim myColumn As DataColumn
        Dim myTableStyle As New DataGridTableStyle
        'Dim strConfigurationFile As String

        'Reset Globals
        Me.funResetGlobals()
        'Create Data Adapter Table Mappings
        Me.tblChild.TableMappings.Clear()
        Me.tblChild.TableMappings.AddRange( _
            New System.Data.Common.DataTableMapping() { _
            New System.Data.Common.DataTableMapping("Table", myChildTable)})
        'Set Data Adapter Column Mappings Automatically
        Me.tblChild.MissingMappingAction = MissingMappingAction.Passthrough
        'Create Select Statement
        'This statement caused an error when Memo fields are used.
        'sqlSelect = _
        '        " SELECT DISTINCT [" & myChildTable & "].*"
        'sqlSelect = _
        '        " SELECT [" & myChildTable & "].*"
        sqlSelect = _
                " SELECT *"
        sqlFrom = _
            " FROM ([" & myChildTable & "]" & _
                " LEFT OUTER JOIN [" & myParentTable & "] AS ParentTable" & _
                " ON [" & myChildTable & "].[" & myChildTableChildID & "] = [ParentTable].[" & myParentTableChildID & "])"
        sqlWhere = _
            " WHERE [ParentTable].[" & myParentTableParentID & "] IS NULL"
        sqlOrder = _
            " ORDER BY [" & myChildTable & "].[" & myChildTableChildName & "]"
        If Me.myCriteriaWhere <> "" Then
            sqlWhere = _
                " WHERE [" & myChildTable & "].[" & myChildTableChildID & "] Not In (" & _
                    " SELECT [" & myParentTable & "].[" & myParentTableChildID & "]" & _
                    " FROM [" & myParentTable & "]" & _
                    " WHERE " & Me.myCriteriaWhere & ")"
        End If
        Me.tblChild.SelectCommand.CommandText = _
            sqlSelect & _
            sqlFrom & _
            sqlWhere & _
            sqlOrder
        'Create Update Statement
        sqlUpdate = _
            " UPDATE [" & myChildTable & "]"
        sqlSet = _
            " SET [" & myChildTableChildName & "] = 'New Item'"
        Me.tblChild.UpdateCommand.CommandText = _
            sqlUpdate & _
            sqlSet
        'Create Insert Statement
        sqlInsert = _
            " INSERT INTO [" & myChildTable & "]" & _
            "([" & myChildTableChildID & "], [" & myChildTableChildName & "], [" & myChildTableChildDescription & "])"
        sqlValues = _
            " VALUES (0, 'New Item', '')"
        Me.tblChild.InsertCommand.CommandText = _
            sqlInsert & _
            sqlValues
        'Create Delete Statement
        sqlDelete = _
            " DELETE"
        sqlFrom = _
            " FROM [" & myChildTable & "]"
        sqlWhere = _
            " WHERE [" & myChildTableChildID & "] = 0"
        Me.tblChild.DeleteCommand.CommandText = _
            sqlDelete & _
            sqlFrom & _
            sqlWhere
        'Fill Data Set
        Try
            'Create Tree
            Me.dsTree.Clear()
            Me.tblChild.Fill(Me.dsTree)
            Me.funCreateTree(Me.treDataTree, Me.dsTree.Tables(myChildTable), myChildTableChildID, myChildTableChildName)
            'Find ChildID Column Ordinal Position
            For Each myColumn In dsTree.Tables(myChildTable).Columns
                If myColumn.ColumnName = myChildTableChildID Then
                    intChildIDColumn = myColumn.Ordinal
                    Exit For
                End If
            Next
            'Enable Menu Items
            Me.funEnableFileMenuItems(True)
            Me.funEnableNodeMenuItems(True)
            'Status
            Me.lblSelected.Text = "Tree Created"
            Me.txtDescription.Text = "Next Step: View Nodes"
            'Data View to Disable AllowNew
            Me.dsChild.Clear()
            Me.tblChild.Fill(Me.dsChild)
            myDataTable = Me.dsChild.Tables(myChildTable)
            myDataView = New DataView(myDataTable)
            myDataView.AllowNew = False
            myDataView.AllowDelete = False
            '***** Set Column Order
            Me.myColumnList.Items.Clear()
            For Each myColumn In Me.dsTree.Tables(myChildTable).Columns
                Me.myColumnList.Items.Add(myColumn.ColumnName, True)
            Next
            Me.funSetColumnOrder(Me.grdChildren, Me.dsChild.Tables(myChildTable), Me.myColumnList, Me.mnuDataGrid)
            Me.funShowColumns()
            '***** Set Column Order
            'Set DataGridChildren to a formatted DataView of DataSetChild
            Me.grdChildren.DataSource = myDataView
            'Handle Data Table Changes
            AddHandler myDataTable.ColumnChanged, AddressOf OnColumnChanged
            Try
                'Select First Node
                Me.treDataTree.SelectedNode = Me.treDataTree.Nodes(0)
            Catch ex As Exception

            End Try

        Catch ex As OleDb.OleDbException
            'Error Message
            If ex.ErrorCode = -2147467259 Then
                If ex.Message = "Type mismatch in expression." Then
                    'Type Mismatch
                    MsgBox( _
                        "Type mismatch error." & vbCr & vbCr & _
                        "The ChildID and ParentID are not the same datatype in both tables." & "  " & _
                        "Make sure these tables are correct and then select from the available list of Table Structures or use Custom to create your own.", , _
                        "Create Tree Error")
                Else
                    'Tables Open Exclusively
                    MsgBox( _
                        "One or more Tables are open exclusively by another user." & vbCr & vbCr & _
                        "The database may be open in exclusive mode or the tables may be in design view." & "  " & _
                        "Make sure these tables are available and then select from the available list of Table Structures or use Custom to create your own.", , _
                        "Create Tree Error")
                End If
                Me.lblSelected.Text = "Table(s) Unavailable"
                Me.txtDescription.Text = "Next Step: Set Table Structure"
            ElseIf ex.ErrorCode = -2147217904 Then
                'Field Name Misspelled
                MsgBox( _
                    "Can not find the selected Field(s).  Verify the Field Names are spelled correctly." & vbCr & vbCr & _
                    "Select from the available list of Table Structures or use Custom to create your own.", , _
                    "Create Tree Error")
                Me.lblSelected.Text = "Field(s) Not Found"
                Me.txtDescription.Text = "Next Step: Set Table Structure"
            ElseIf ex.ErrorCode = -2147217865 Then
                'If Me.mnuCreateTree.Enabled Then
                'Tables Not Found
                'MsgBox( _
                '"Can not find the selected Table(s).  Verify the Table Names are spelled correctly." & vbCr & vbCr & _
                '    "Select from the available list of Table Structures or use Custom to create your own.", , _
                '    "Create Tree Error")
                Me.lblSelected.Text = "Database Connected > Table(s) Not Found"
                Me.txtDescription.Text = "Next Step: Set Table Structure"
                'End If
            End If
            'Clear Tree
            Me.dsTree.Clear()
            Me.dsChild.Clear()
            Me.treDataTree.Nodes.Clear()
            Me.treDataTree.SelectedNode = Nothing
            Me.grdChildren.DataSource = Nothing
            'Enable File Menu
            Me.funEnableFileMenuItems(True)
            'Disable Menu Items
            Me.funEnableEditMenuItems(False)
            Me.funEnableNodeMenuItems(False)
        Catch ex As Exception
            MsgBox("Create Components: " & ex.Message)
        End Try
        Me.funSetFormName(myConfigurationFile)
        Me.treDataTree.Focus()

    End Sub

    Private Sub funHelpContents()
        MsgBox( _
            "This form will display Hierarchical data from any database in a Tree View, " & _
            "but the tables must be in a specific format with specific relationships." & vbCr & vbCr & _
                vbTab & "ChildTable:" & vbTab & "Child" & vbCr & _
                vbTab & "ChildID:" & vbTab & vbTab & "ID" & vbTab & vbTab & _
                "(Note: ChildTable.ChildID MUST be part of the primary key for ParentTable)" & vbCr & _
                vbTab & "ChildName:" & vbTab & "Name" & vbCr & _
                vbTab & "ChildDescription:" & vbTab & "Description" & vbCr & _
                vbTab & "ParentTable:" & vbTab & "Parent" & vbCr & _
                vbTab & "ParentID:" & vbTab & vbTab & "ParentID" & vbCr & vbCr & _
                vbTab & "Child.ID = Parent.ID (one-to-many)" & _
                vbTab & "Child.ID = Parent.ParentID (one-to-many)" & vbCr & vbCr & _
            "Perform the following steps under the File Menu to Create a Data Tree View?" & vbCr & vbCr & vbTab & _
            "1) Open Database Connection" & vbCr & vbTab & _
            "2) Set Table Structure" & vbCr & vbTab & _
            "3) Create Tree", , "Help Contents")

    End Sub

    Private Function funConvertDataType(ByVal myValue As String, ByVal myField As String) As String
        Dim myColumn As DataColumn

        If myValue Is Nothing Then
            funConvertDataType = "Null"
            Exit Function
        End If
        'Get Data Type
        myColumn = Me.dsTree.Tables(myChildTable).Columns(myField)
        If myColumn.DataType Is System.Type.GetType("System.String") Then
            'String
            funConvertDataType = "'" & Me.funReplaceSingleQuote(myValue) & "'"
        ElseIf myColumn.DataType Is System.Type.GetType("System.DateTime") Then
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

    Private Sub funSearch(Optional ByVal myStartNode As TreeNode = Nothing)
        Dim myNode As TreeNode
        Dim mySearch As String

        If Me.treDataTree.GetNodeCount(True) = 0 Then
            MsgBox("Please Create a Tree first.")
            Exit Sub
        End If

        myFoundNode = Nothing
        'Reset Tree
        For Each myNode In Me.treDataTree.Nodes
            Me.funResetSearchTree(myNode)
        Next myNode
        'Find Node
        If myStartNode Is Nothing Then
            'Entire Tree
            mySearch = InputBox("Search Entire Tree For:", "Find")
            If mySearch = "" Then
                Exit Sub
            End If
            For Each myNode In Me.treDataTree.Nodes
                Me.funSearchTree(mySearch, myNode)
            Next myNode
            Me.treDataTree.SelectedNode = Me.treDataTree.Nodes(0)
        Else
            'SubTree Only
            mySearch = InputBox("Search In " & myStartNode.FullPath & " For:", "Find")
            If mySearch = "" Then
                Exit Sub
            End If
            Me.funSearchTree(mySearch, myStartNode)
            Me.treDataTree.SelectedNode = myStartNode
        End If
        If myFoundNode Is Nothing Then
            MsgBox("Node Not Found")
        End If

    End Sub

    Private Sub funSaveChanges()
        Try
            'Save Changes
            If Not mySelectedNode Is Nothing Then
                grdChildren.EndEdit(Me.grdChildren.TableStyles(0).GridColumnStyles(Me.grdChildren.CurrentCell.ColumnNumber), Me.grdChildren.CurrentRowIndex, False)
            End If
        Catch ex As Exception
        End Try

    End Sub

    Private Sub funSetColumnOrder(ByVal myDataGrid As DataGrid, ByVal myDataTableSource As DataTable, ByVal myColumnList As CheckedListBox, ByVal myMenu As Menu)
        'Dim myDataRow As DataRow
        Dim myColumn As DataColumn
        Dim myTableStyle As New DataGridTableStyle
        Dim intCounter As Integer

        'Remove Menu Items
        While myMenu.MenuItems.Count > 4
            myMenu.MenuItems.Remove(myMenu.MenuItems.Item(4))
        End While

        'Create DataGrid
        Try
            'myDataGrid.TableStyles.Clear()
            For intCounter = 0 To myColumnList.Items.Count - 1
                myColumn = myDataTableSource.Columns(myColumnList.Items(intCounter))
                'Set ChildID DataGrid Column Position (May be different than table)
                If myColumn.ColumnName = myChildTableChildID Then
                    intChildIDColumn = intCounter
                End If
                'Create Columns
                If myColumn.DataType Is System.Type.GetType("System.Decimal") Then
                    Dim myColumnStyle As New DataGridTextBoxColumn
                    myColumnStyle.Format() = "0.00"
                    myColumnStyle.MappingName = myColumn.ColumnName
                    myColumnStyle.HeaderText = myColumn.ColumnName
                    myColumnStyle.NullText() = ""
                    Try
                        myColumnStyle.Width = Me.grdChildren.TableStyles(0).GridColumnStyles(myColumn.ColumnName).Width
                    Catch ex As Exception
                    End Try
                    myTableStyle.GridColumnStyles.Add(myColumnStyle)
                ElseIf myColumn.DataType Is System.Type.GetType("System.Boolean") Then
                    Dim myColumnStyle As New DataGridBoolColumn
                    myColumnStyle.MappingName = myColumn.ColumnName
                    myColumnStyle.HeaderText = myColumn.ColumnName
                    myColumnStyle.NullText() = ""
                    Try
                        myColumnStyle.Width = Me.grdChildren.TableStyles(0).GridColumnStyles(myColumn.ColumnName).Width
                    Catch ex As Exception
                    End Try
                    myTableStyle.GridColumnStyles.Add(myColumnStyle)
                ElseIf myColumn.DataType Is System.Type.GetType("System.String") Then
                    If Me.mnuWordWrap.Checked = True Then
                        'Wrap Text
                        Dim myColumnStyle As New DataGridTextBoxColumnWrap
                        myColumnStyle.MappingName = myColumn.ColumnName
                        myColumnStyle.HeaderText = myColumn.ColumnName
                        myColumnStyle.NullText() = ""
                        Try
                            myColumnStyle.Width = Me.grdChildren.TableStyles(0).GridColumnStyles(myColumn.ColumnName).Width
                        Catch ex As Exception
                        End Try
                        myTableStyle.GridColumnStyles.Add(myColumnStyle)
                    Else
                        'No Wrap Text
                        Dim myColumnStyle As New DataGridTextBoxColumn
                        myColumnStyle.MappingName = myColumn.ColumnName
                        myColumnStyle.HeaderText = myColumn.ColumnName
                        myColumnStyle.NullText() = ""
                        Try
                            myColumnStyle.Width = Me.grdChildren.TableStyles(0).GridColumnStyles(myColumn.ColumnName).Width
                        Catch ex As Exception
                        End Try
                        myTableStyle.GridColumnStyles.Add(myColumnStyle)

                    End If
                Else
                    Dim myColumnStyle As New DataGridTextBoxColumn
                    myColumnStyle.MappingName = myColumn.ColumnName
                    myColumnStyle.HeaderText = myColumn.ColumnName
                    myColumnStyle.NullText() = ""
                    Try
                        myColumnStyle.Width = Me.grdChildren.TableStyles(0).GridColumnStyles(myColumn.ColumnName).Width
                    Catch ex As Exception
                    End Try
                    myTableStyle.GridColumnStyles.Add(myColumnStyle)
                End If
                'Create Menu Items
                myMenu.MenuItems.Add(myColumn.ColumnName, New EventHandler(AddressOf OnDataGridMenuClick))
                myMenu.MenuItems(myMenu.MenuItems.Count - 1).Checked = True
            Next
            myTableStyle.MappingName = myChildTable
            myTableStyle.PreferredRowHeight = intRowHeight
            Me.grdChildren.TableStyles.Clear()
            Me.grdChildren.TableStyles.Add(myTableStyle)
        Catch ex As Exception
            'MsgBox("Set Column Order: " & ex.Message)
        End Try

    End Sub

    Private Sub funShowColumns()
        Dim myDataGridForm As New frmDataGrid
        Dim myMenuItem As MenuItem
        Dim myColumn As DataGridColumnStyle

        For Each myMenuItem In Me.mnuDataGrid.MenuItems
            If myMenuItem.Index > 3 Then
                myColumn = Me.grdChildren.TableStyles(myChildTable).GridColumnStyles(myMenuItem.Text)
                If Me.myColumnList.CheckedItems.Contains(myMenuItem.Text) Then
                    myMenuItem.Checked = True
                    If myColumn.Width = 0 Then
                        myColumn.Width = 75
                    End If
                Else
                    myMenuItem.Checked = False
                    myColumn.Width = 0
                End If
            End If
        Next
        myDataGridForm.Dispose()

    End Sub

    Private Sub funResetGlobals()
        'Default Table Structure
        'SQL Statements
        sqlSelect = ""
        sqlFrom = ""
        sqlWhere = ""
        sqlOrder = ""
        sqlUpdate = ""
        sqlSet = ""
        sqlInsert = ""
        sqlValues = ""
        sqlDelete = ""
        'Nodes
        mySelectedNode = Nothing
        myFoundNode = Nothing
        myCutNode = Nothing
        myCopyNode = Nothing
        bolIsCollapsing = False
        'Drag and Drop
        myDragNode = Nothing
        mySourceNode = Nothing
        myDestinationNode = Nothing
        myHoverNode = Nothing
        'Constants
        'DataGrid
        myDataTable = Nothing
        myDataView = Nothing
        bolMouseOverColumnHeader = False
        bolMouseOverRowHeader = False

    End Sub

    Private Sub funResetTree()
        'If Me.myConfigurationFile = "" Then
        '    Me.Text = "Data Tree"
        'Else
        '    Me.Text = Me.myConfigurationFile
        'End If
        Me.funSetFormName(Me.myConfigurationFile)
        Me.treDataTree.Nodes.Clear()
        Me.lblSelected.Text = "Data Tree"
        Me.txtDescription.Text = ""
        Me.grdChildren.DataSource = Nothing
        Me.grdChildren.Refresh()

    End Sub

    Private Sub funSetFormName(ByVal strConfigurationFile As String)
        strConfigurationFile = myConfigurationFile
        While InStr(strConfigurationFile, "\") > 0
            strConfigurationFile = Mid(strConfigurationFile, InStr(strConfigurationFile, "\") + 1)
        End While
        If strConfigurationFile = "" Then
            Me.Text = strDatabase & " - Data Tree"
        Else
            Me.Text = strConfigurationFile & ": " & strDatabase & " - Data Tree"
        End If
        If myChildTable <> "" Then
            Me.Text = strConfigurationFile & ": " & strDatabase & " (" & myChildTable & " > " & myParentTable & ")" & " - Data Tree"
        End If

    End Sub

#End Region

#Region "Events"

    Private Sub OnColumnChanged(ByVal sender As Object, ByVal e As DataColumnChangeEventArgs)
        Dim strNewValue As String
        Dim myDataColumn As DataColumn
        Dim myNode As TreeNode

        If IsDBNull(e.ProposedValue) Then
            strNewValue = Nothing
        Else
            strNewValue = e.ProposedValue
        End If
        'Modify Update Statement
        sqlUpdate = _
            " UPDATE [" & myChildTable & "]"
        sqlSet = _
            " SET [" & e.Column.ColumnName & "] = " & Me.funConvertDataType(strNewValue, e.Column.ColumnName)
        sqlWhere = _
            " WHERE [" & myChildTable & "].[" & myChildTableChildID & "] = " & Me.funConvertDataType(mySelectedNode.Tag, myChildTableChildID)
        'sqlWhere = _
        '    " WHERE [" & myChildTable & "].[" & myChildTableChildID & "] = " & Me.funConvertDataType(e.Row.Item(myChildTableChildID), myChildTableChildID)
        'Modification Date
        For Each myDataColumn In myDataTable.Columns
            If e.Column.ColumnName <> "DateModified" And myDataColumn.ColumnName = "DateModified" Or _
                e.Column.ColumnName <> "ModificationDate" And myDataColumn.ColumnName = "ModificationDate" _
            Then
                sqlSet = sqlSet & ", " & _
                    "[" & myDataColumn.ColumnName & "] = " & Me.funConvertDataType(Now(), myDataColumn.ColumnName)
                'Data Grid Modification Date
                'Me.grdChildren.Item(Me.grdChildren.CurrentRowIndex, myDataColumn.Ordinal) = Now()
                'Me.dsChild.Tables(myChildTable).Rows(Me.grdChildren.CurrentRowIndex).Item(myDataColumn.ColumnName) = Now()
            End If
        Next
        'Update Selected Child
        Me.tblChild.UpdateCommand.CommandText = _
            sqlUpdate & _
            sqlSet & _
            sqlWhere
        Me.dbConnection.Open()
        Try
            Me.tblChild.UpdateCommand.ExecuteNonQuery()
            'Change Node ID
            If e.Column.ColumnName = myChildTableChildID Then
                'Reset Tree 
                For Each myNode In Me.treDataTree.Nodes
                    Me.funResetSearchTree(myNode)
                Next myNode
                'Find Node 
                For Each myNode In Me.treDataTree.Nodes
                    Me.funSearchTree(mySelectedNode.Tag, myNode, "ReplaceNodeID", strNewValue)
                Next myNode
            End If
            'Change Node Name
            If e.Column.ColumnName = myChildTableChildName Then
                'Reset Tree 
                For Each myNode In Me.treDataTree.Nodes
                    Me.funResetSearchTree(myNode)
                Next myNode
                'Find Node 
                For Each myNode In Me.treDataTree.Nodes
                    Me.funSearchTree(mySelectedNode.Tag, myNode, "ReplaceNode", strNewValue)
                Next myNode
            End If
            'Change Node Description
            If e.Column.ColumnName = myChildTableChildDescription Then
                Me.txtDescription.Text = strNewValue
            End If
        Catch ex As System.Data.OleDb.OleDbException
            MsgBox(ex.Message, , "Update Database Error")
            e.Row.CancelEdit()
        End Try
        Me.dbConnection.Close()

    End Sub

    Private Sub OnDataGridMenuClick(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim myMenuItem As System.Windows.Forms.MenuItem
        Dim myColumn As DataGridColumnStyle

        myMenuItem = sender
        myMenuItem.Checked = Not myMenuItem.Checked

        For Each myColumn In Me.grdChildren.TableStyles(myChildTable).GridColumnStyles
            If myColumn.MappingName = myMenuItem.Text Then
                If myMenuItem.Checked Then
                    Me.myColumnList.SetItemCheckState(Me.myColumnList.Items.IndexOf(myMenuItem.Text), CheckState.Checked)
                    If myColumn.Width = 0 Then
                        myColumn.Width = 75
                    End If
                Else
                    Me.myColumnList.SetItemCheckState(Me.myColumnList.Items.IndexOf(myMenuItem.Text), CheckState.Unchecked)
                    myColumn.Width = 0
                End If
                Me.bolSaveConfigurationFile = True
            End If
        Next

    End Sub

    Private Sub mnuWordWrap_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuWordWrap.Click
        If mnuWordWrap.Checked = True Then
            mnuWordWrap.Checked = False
        Else
            mnuWordWrap.Checked = True
        End If
        Me.funCreateComponents()
    End Sub

#End Region

End Class