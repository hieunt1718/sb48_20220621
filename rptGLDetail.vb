Imports SmartBooks.General
Imports VBReport
Imports System.Data.SqlClient
Imports Microsoft.ApplicationBlocks.Data
Imports Microsoft.Office.Interop
Imports System.IO
Imports OfficeOpenXml
Imports OfficeOpenXml.Style

Public Class rptGLDetail
    Inherits System.Windows.Forms.Form
    Dim tbResult As DataTable
    Private ExportToExcel As SmartBooks.General.ClsExportToExcel
    Friend WithEvents tbnPreview As Janus.Windows.EditControls.UIButton
    Friend WithEvents OtherOption As Janus.Windows.EditControls.UIGroupBox
    Friend WithEvents cboOptionOther As Janus.Windows.GridEX.EditControls.MultiColumnCombo
    Friend WithEvents lbOtherOption As Label
    Friend WithEvents cbOption As CheckBox
    Dim urlTemplate As String

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
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents UiGroupBox1 As Janus.Windows.EditControls.UIGroupBox
    Friend WithEvents btnClose As Janus.Windows.EditControls.UIButton
    Friend WithEvents btnPrintPreview As Janus.Windows.EditControls.UIButton
    Friend WithEvents UiGroupBox2 As Janus.Windows.EditControls.UIGroupBox
    Friend WithEvents rbtnAll As Janus.Windows.EditControls.UIRadioButton
    Friend WithEvents rbtnAcct As Janus.Windows.EditControls.UIRadioButton
    Friend WithEvents txtAcct As Janus.Windows.GridEX.EditControls.EditBox
    Friend WithEvents UiGroupBox3 As Janus.Windows.EditControls.UIGroupBox
    Friend WithEvents imgToolBars As System.Windows.Forms.ImageList
    Friend WithEvents cboCuryID As Janus.Windows.GridEX.EditControls.MultiColumnCombo
    Friend WithEvents lblCuryID As System.Windows.Forms.Label
    Friend WithEvents FocusHighlighter1 As SmartBooks.Ex.EWinform.FocusHighlighter
    Friend WithEvents cboToDate As SmartBooks.Ex.EWinform.ValidText
    Friend WithEvents cboFromDate As SmartBooks.Ex.EWinform.ValidText
    Friend WithEvents btnExportExcel As Janus.Windows.EditControls.UIButton
    Friend WithEvents cbDisplayVNDUSD As System.Windows.Forms.CheckBox
    Friend WithEvents errProvider As System.Windows.Forms.ErrorProvider
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim cboOptionOther_DesignTimeLayout As Janus.Windows.GridEX.GridEXLayout = New Janus.Windows.GridEX.GridEXLayout()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(rptGLDetail))
        Dim cboCuryID_DesignTimeLayout As Janus.Windows.GridEX.GridEXLayout = New Janus.Windows.GridEX.GridEXLayout()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.cbOption = New System.Windows.Forms.CheckBox()
        Me.OtherOption = New Janus.Windows.EditControls.UIGroupBox()
        Me.cboOptionOther = New Janus.Windows.GridEX.EditControls.MultiColumnCombo()
        Me.lbOtherOption = New System.Windows.Forms.Label()
        Me.cbDisplayVNDUSD = New System.Windows.Forms.CheckBox()
        Me.cboToDate = New SmartBooks.Ex.EWinform.ValidText()
        Me.cboFromDate = New SmartBooks.Ex.EWinform.ValidText()
        Me.UiGroupBox3 = New Janus.Windows.EditControls.UIGroupBox()
        Me.cboCuryID = New Janus.Windows.GridEX.EditControls.MultiColumnCombo()
        Me.lblCuryID = New System.Windows.Forms.Label()
        Me.UiGroupBox2 = New Janus.Windows.EditControls.UIGroupBox()
        Me.txtAcct = New Janus.Windows.GridEX.EditControls.EditBox()
        Me.rbtnAcct = New Janus.Windows.EditControls.UIRadioButton()
        Me.rbtnAll = New Janus.Windows.EditControls.UIRadioButton()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.UiGroupBox1 = New Janus.Windows.EditControls.UIGroupBox()
        Me.tbnPreview = New Janus.Windows.EditControls.UIButton()
        Me.imgToolBars = New System.Windows.Forms.ImageList(Me.components)
        Me.btnExportExcel = New Janus.Windows.EditControls.UIButton()
        Me.btnClose = New Janus.Windows.EditControls.UIButton()
        Me.btnPrintPreview = New Janus.Windows.EditControls.UIButton()
        Me.FocusHighlighter1 = New SmartBooks.Ex.EWinform.FocusHighlighter(Me.components)
        Me.errProvider = New System.Windows.Forms.ErrorProvider(Me.components)
        Me.Panel1.SuspendLayout()
        CType(Me.OtherOption, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.OtherOption.SuspendLayout()
        CType(Me.cboOptionOther, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UiGroupBox3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UiGroupBox3.SuspendLayout()
        CType(Me.cboCuryID, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UiGroupBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UiGroupBox2.SuspendLayout()
        CType(Me.UiGroupBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UiGroupBox1.SuspendLayout()
        CType(Me.errProvider, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.cbOption)
        Me.Panel1.Controls.Add(Me.OtherOption)
        Me.Panel1.Controls.Add(Me.cbDisplayVNDUSD)
        Me.Panel1.Controls.Add(Me.cboToDate)
        Me.Panel1.Controls.Add(Me.cboFromDate)
        Me.Panel1.Controls.Add(Me.UiGroupBox3)
        Me.Panel1.Controls.Add(Me.UiGroupBox2)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.FocusHighlighter1.SetHighlight(Me.Panel1, False)
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(542, 316)
        Me.Panel1.TabIndex = 0
        '
        'cbOption
        '
        Me.FocusHighlighter1.SetHighlight(Me.cbOption, False)
        Me.cbOption.Location = New System.Drawing.Point(16, 195)
        Me.cbOption.Name = "cbOption"
        Me.cbOption.Size = New System.Drawing.Size(244, 24)
        Me.cbOption.TabIndex = 25
        Me.cbOption.Text = "Other Option Export Excel"
        '
        'OtherOption
        '
        Me.OtherOption.BackColor = System.Drawing.Color.Transparent
        Me.OtherOption.Controls.Add(Me.cboOptionOther)
        Me.OtherOption.Controls.Add(Me.lbOtherOption)
        Me.FocusHighlighter1.SetHighlight(Me.OtherOption, False)
        Me.OtherOption.Location = New System.Drawing.Point(16, 222)
        Me.OtherOption.Name = "OtherOption"
        Me.OtherOption.Size = New System.Drawing.Size(244, 88)
        Me.OtherOption.TabIndex = 24
        Me.OtherOption.Text = "Other Option"
        Me.OtherOption.VisualStyle = Janus.Windows.UI.Dock.PanelVisualStyle.Office2003
        '
        'cboOptionOther
        '
        Me.cboOptionOther.ComboStyle = Janus.Windows.GridEX.ComboStyle.DropDownList
        cboOptionOther_DesignTimeLayout.LayoutString = resources.GetString("cboOptionOther_DesignTimeLayout.LayoutString")
        Me.cboOptionOther.DesignTimeLayout = cboOptionOther_DesignTimeLayout
        Me.FocusHighlighter1.SetHighlight(Me.cboOptionOther, False)
        Me.cboOptionOther.Location = New System.Drawing.Point(92, 25)
        Me.cboOptionOther.Name = "cboOptionOther"
        Me.cboOptionOther.SelectedIndex = -1
        Me.cboOptionOther.SelectedItem = Nothing
        Me.cboOptionOther.Size = New System.Drawing.Size(130, 21)
        Me.cboOptionOther.TabIndex = 8
        Me.cboOptionOther.TextAlignment = Janus.Windows.GridEX.TextAlignment.Near
        Me.cboOptionOther.VisualStyle = Janus.Windows.GridEX.VisualStyle.Office2003
        '
        'lbOtherOption
        '
        Me.lbOtherOption.AutoSize = True
        Me.FocusHighlighter1.SetHighlight(Me.lbOtherOption, False)
        Me.lbOtherOption.Location = New System.Drawing.Point(16, 28)
        Me.lbOtherOption.Name = "lbOtherOption"
        Me.lbOtherOption.Size = New System.Drawing.Size(70, 13)
        Me.lbOtherOption.TabIndex = 23
        Me.lbOtherOption.Text = "Other Option"
        Me.lbOtherOption.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'cbDisplayVNDUSD
        '
        Me.FocusHighlighter1.SetHighlight(Me.cbDisplayVNDUSD, False)
        Me.cbDisplayVNDUSD.Location = New System.Drawing.Point(227, 20)
        Me.cbDisplayVNDUSD.Name = "cbDisplayVNDUSD"
        Me.cbDisplayVNDUSD.Size = New System.Drawing.Size(104, 24)
        Me.cbDisplayVNDUSD.TabIndex = 13
        Me.cbDisplayVNDUSD.Text = "VND-USD"
        '
        'cboToDate
        '
        Me.cboToDate.ClearTime = False
        Me.cboToDate.DateDelimiter = "/"
        Me.cboToDate.DateInputFormat = SmartBooks.Ex.EUtilities.DateFormat.DDMMYYYY
        Me.cboToDate.DateOutputFormat = SmartBooks.Ex.EUtilities.DateFormat.DDMMYYYY
        Me.cboToDate.FieldReference = Nothing
        Me.FocusHighlighter1.SetHighlight(Me.cboToDate, True)
        Me.cboToDate.Location = New System.Drawing.Point(84, 44)
        Me.cboToDate.MaskEdit = ""
        Me.cboToDate.MessageLanguage = SmartBooks.Ex.EWinform.ValidText.MessageLanguages.Vietnamese
        Me.cboToDate.miForm = Nothing
        Me.cboToDate.Name = "cboToDate"
        Me.cboToDate.RegExPattern = SmartBooks.Ex.EWinform.ValidText.RegularExpressionModes.AutomatedDate
        Me.cboToDate.Required = False
        Me.cboToDate.ShowErrorIcon = True
        Me.cboToDate.Size = New System.Drawing.Size(130, 21)
        Me.cboToDate.TabIndex = 2
        Me.cboToDate.TextAlignment = Janus.Windows.GridEX.TextAlignment.Near
        Me.cboToDate.ValidationMode = SmartBooks.Ex.EWinform.ValidText.ValidationModes.ValidCharacters
        Me.cboToDate.ValidText = "0123456789/"
        Me.cboToDate.VisualStyle = Janus.Windows.GridEX.VisualStyle.Office2003
        Me.cboToDate.YearPrefix = CType(20, Short)
        '
        'cboFromDate
        '
        Me.cboFromDate.ClearTime = False
        Me.cboFromDate.DateDelimiter = "/"
        Me.cboFromDate.DateInputFormat = SmartBooks.Ex.EUtilities.DateFormat.DDMMYYYY
        Me.cboFromDate.DateOutputFormat = SmartBooks.Ex.EUtilities.DateFormat.DDMMYYYY
        Me.cboFromDate.FieldReference = Nothing
        Me.FocusHighlighter1.SetHighlight(Me.cboFromDate, True)
        Me.cboFromDate.Location = New System.Drawing.Point(84, 20)
        Me.cboFromDate.MaskEdit = ""
        Me.cboFromDate.MessageLanguage = SmartBooks.Ex.EWinform.ValidText.MessageLanguages.Vietnamese
        Me.cboFromDate.miForm = Nothing
        Me.cboFromDate.Name = "cboFromDate"
        Me.cboFromDate.RegExPattern = SmartBooks.Ex.EWinform.ValidText.RegularExpressionModes.AutomatedDate
        Me.cboFromDate.Required = False
        Me.cboFromDate.ShowErrorIcon = True
        Me.cboFromDate.Size = New System.Drawing.Size(130, 21)
        Me.cboFromDate.TabIndex = 1
        Me.cboFromDate.TextAlignment = Janus.Windows.GridEX.TextAlignment.Near
        Me.cboFromDate.ValidationMode = SmartBooks.Ex.EWinform.ValidText.ValidationModes.ValidCharacters
        Me.cboFromDate.ValidText = "0123456789/"
        Me.cboFromDate.VisualStyle = Janus.Windows.GridEX.VisualStyle.Office2003
        Me.cboFromDate.YearPrefix = CType(20, Short)
        '
        'UiGroupBox3
        '
        Me.UiGroupBox3.BackColor = System.Drawing.Color.Transparent
        Me.UiGroupBox3.Controls.Add(Me.cboCuryID)
        Me.UiGroupBox3.Controls.Add(Me.lblCuryID)
        Me.FocusHighlighter1.SetHighlight(Me.UiGroupBox3, False)
        Me.UiGroupBox3.Location = New System.Drawing.Point(280, 80)
        Me.UiGroupBox3.Name = "UiGroupBox3"
        Me.UiGroupBox3.Size = New System.Drawing.Size(244, 88)
        Me.UiGroupBox3.TabIndex = 7
        Me.UiGroupBox3.Text = "Select currency"
        Me.UiGroupBox3.VisualStyle = Janus.Windows.UI.Dock.PanelVisualStyle.Office2003
        '
        'cboCuryID
        '
        Me.cboCuryID.ComboStyle = Janus.Windows.GridEX.ComboStyle.DropDownList
        cboCuryID_DesignTimeLayout.LayoutString = resources.GetString("cboCuryID_DesignTimeLayout.LayoutString")
        Me.cboCuryID.DesignTimeLayout = cboCuryID_DesignTimeLayout
        Me.FocusHighlighter1.SetHighlight(Me.cboCuryID, False)
        Me.cboCuryID.Location = New System.Drawing.Point(92, 25)
        Me.cboCuryID.Name = "cboCuryID"
        Me.cboCuryID.SelectedIndex = -1
        Me.cboCuryID.SelectedItem = Nothing
        Me.cboCuryID.Size = New System.Drawing.Size(130, 21)
        Me.cboCuryID.TabIndex = 8
        Me.cboCuryID.TextAlignment = Janus.Windows.GridEX.TextAlignment.Near
        Me.cboCuryID.VisualStyle = Janus.Windows.GridEX.VisualStyle.Office2003
        '
        'lblCuryID
        '
        Me.lblCuryID.AutoSize = True
        Me.FocusHighlighter1.SetHighlight(Me.lblCuryID, False)
        Me.lblCuryID.Location = New System.Drawing.Point(16, 28)
        Me.lblCuryID.Name = "lblCuryID"
        Me.lblCuryID.Size = New System.Drawing.Size(65, 13)
        Me.lblCuryID.TabIndex = 23
        Me.lblCuryID.Text = "Currency ID"
        Me.lblCuryID.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'UiGroupBox2
        '
        Me.UiGroupBox2.BackColor = System.Drawing.Color.Transparent
        Me.UiGroupBox2.Controls.Add(Me.txtAcct)
        Me.UiGroupBox2.Controls.Add(Me.rbtnAcct)
        Me.UiGroupBox2.Controls.Add(Me.rbtnAll)
        Me.FocusHighlighter1.SetHighlight(Me.UiGroupBox2, False)
        Me.UiGroupBox2.Location = New System.Drawing.Point(16, 80)
        Me.UiGroupBox2.Name = "UiGroupBox2"
        Me.UiGroupBox2.Size = New System.Drawing.Size(244, 109)
        Me.UiGroupBox2.TabIndex = 3
        Me.UiGroupBox2.Text = "Option"
        Me.UiGroupBox2.VisualStyle = Janus.Windows.UI.Dock.PanelVisualStyle.Office2003
        '
        'txtAcct
        '
        Me.FocusHighlighter1.SetHighlight(Me.txtAcct, False)
        Me.txtAcct.Location = New System.Drawing.Point(104, 52)
        Me.txtAcct.Name = "txtAcct"
        Me.txtAcct.Size = New System.Drawing.Size(130, 21)
        Me.txtAcct.TabIndex = 6
        Me.txtAcct.TextAlignment = Janus.Windows.GridEX.TextAlignment.Near
        Me.txtAcct.VisualStyle = Janus.Windows.GridEX.VisualStyle.Office2003
        '
        'rbtnAcct
        '
        Me.rbtnAcct.BackColor = System.Drawing.Color.Transparent
        Me.FocusHighlighter1.SetHighlight(Me.rbtnAcct, False)
        Me.rbtnAcct.Location = New System.Drawing.Point(16, 52)
        Me.rbtnAcct.Name = "rbtnAcct"
        Me.rbtnAcct.Size = New System.Drawing.Size(80, 23)
        Me.rbtnAcct.TabIndex = 5
        Me.rbtnAcct.Text = "Account"
        Me.rbtnAcct.VisualStyle = Janus.Windows.UI.VisualStyle.Office2003
        '
        'rbtnAll
        '
        Me.rbtnAll.BackColor = System.Drawing.Color.Transparent
        Me.rbtnAll.Checked = True
        Me.FocusHighlighter1.SetHighlight(Me.rbtnAll, False)
        Me.rbtnAll.Location = New System.Drawing.Point(16, 24)
        Me.rbtnAll.Name = "rbtnAll"
        Me.rbtnAll.Size = New System.Drawing.Size(80, 23)
        Me.rbtnAll.TabIndex = 4
        Me.rbtnAll.TabStop = True
        Me.rbtnAll.Text = "All"
        Me.rbtnAll.VisualStyle = Janus.Windows.UI.VisualStyle.Office2003
        '
        'Label2
        '
        Me.FocusHighlighter1.SetHighlight(Me.Label2, False)
        Me.Label2.Location = New System.Drawing.Point(16, 44)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(76, 23)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "To date"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label1
        '
        Me.FocusHighlighter1.SetHighlight(Me.Label1, False)
        Me.Label1.Location = New System.Drawing.Point(16, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(76, 23)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "From date"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'UiGroupBox1
        '
        Me.UiGroupBox1.Controls.Add(Me.tbnPreview)
        Me.UiGroupBox1.Controls.Add(Me.btnExportExcel)
        Me.UiGroupBox1.Controls.Add(Me.btnClose)
        Me.UiGroupBox1.Controls.Add(Me.btnPrintPreview)
        Me.UiGroupBox1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.UiGroupBox1.FrameStyle = Janus.Windows.EditControls.FrameStyle.Top
        Me.FocusHighlighter1.SetHighlight(Me.UiGroupBox1, False)
        Me.UiGroupBox1.Location = New System.Drawing.Point(0, 316)
        Me.UiGroupBox1.Name = "UiGroupBox1"
        Me.UiGroupBox1.Size = New System.Drawing.Size(542, 60)
        Me.UiGroupBox1.TabIndex = 9
        Me.UiGroupBox1.VisualStyle = Janus.Windows.UI.Dock.PanelVisualStyle.Office2003
        '
        'tbnPreview
        '
        Me.tbnPreview.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.FocusHighlighter1.SetHighlight(Me.tbnPreview, False)
        Me.tbnPreview.ImageHorizontalAlignment = Janus.Windows.EditControls.ImageHorizontalAlignment.Near
        Me.tbnPreview.ImageIndex = 59
        Me.tbnPreview.ImageList = Me.imgToolBars
        Me.tbnPreview.Location = New System.Drawing.Point(227, 20)
        Me.tbnPreview.Name = "tbnPreview"
        Me.tbnPreview.Size = New System.Drawing.Size(100, 25)
        Me.tbnPreview.TabIndex = 13
        Me.tbnPreview.Text = "&Preview"
        Me.tbnPreview.Visible = False
        '
        'imgToolBars
        '
        Me.imgToolBars.ImageStream = CType(resources.GetObject("imgToolBars.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.imgToolBars.TransparentColor = System.Drawing.Color.Transparent
        Me.imgToolBars.Images.SetKeyName(0, "")
        Me.imgToolBars.Images.SetKeyName(1, "")
        Me.imgToolBars.Images.SetKeyName(2, "")
        Me.imgToolBars.Images.SetKeyName(3, "")
        Me.imgToolBars.Images.SetKeyName(4, "")
        Me.imgToolBars.Images.SetKeyName(5, "")
        Me.imgToolBars.Images.SetKeyName(6, "")
        Me.imgToolBars.Images.SetKeyName(7, "")
        Me.imgToolBars.Images.SetKeyName(8, "")
        Me.imgToolBars.Images.SetKeyName(9, "")
        Me.imgToolBars.Images.SetKeyName(10, "")
        Me.imgToolBars.Images.SetKeyName(11, "")
        Me.imgToolBars.Images.SetKeyName(12, "")
        Me.imgToolBars.Images.SetKeyName(13, "")
        Me.imgToolBars.Images.SetKeyName(14, "")
        Me.imgToolBars.Images.SetKeyName(15, "")
        Me.imgToolBars.Images.SetKeyName(16, "")
        Me.imgToolBars.Images.SetKeyName(17, "")
        Me.imgToolBars.Images.SetKeyName(18, "")
        Me.imgToolBars.Images.SetKeyName(19, "")
        Me.imgToolBars.Images.SetKeyName(20, "")
        Me.imgToolBars.Images.SetKeyName(21, "")
        Me.imgToolBars.Images.SetKeyName(22, "")
        Me.imgToolBars.Images.SetKeyName(23, "")
        Me.imgToolBars.Images.SetKeyName(24, "")
        Me.imgToolBars.Images.SetKeyName(25, "")
        Me.imgToolBars.Images.SetKeyName(26, "")
        Me.imgToolBars.Images.SetKeyName(27, "")
        Me.imgToolBars.Images.SetKeyName(28, "")
        Me.imgToolBars.Images.SetKeyName(29, "")
        Me.imgToolBars.Images.SetKeyName(30, "")
        Me.imgToolBars.Images.SetKeyName(31, "")
        Me.imgToolBars.Images.SetKeyName(32, "")
        Me.imgToolBars.Images.SetKeyName(33, "")
        Me.imgToolBars.Images.SetKeyName(34, "")
        Me.imgToolBars.Images.SetKeyName(35, "")
        Me.imgToolBars.Images.SetKeyName(36, "")
        Me.imgToolBars.Images.SetKeyName(37, "")
        Me.imgToolBars.Images.SetKeyName(38, "")
        Me.imgToolBars.Images.SetKeyName(39, "")
        Me.imgToolBars.Images.SetKeyName(40, "")
        Me.imgToolBars.Images.SetKeyName(41, "")
        Me.imgToolBars.Images.SetKeyName(42, "")
        Me.imgToolBars.Images.SetKeyName(43, "")
        Me.imgToolBars.Images.SetKeyName(44, "")
        Me.imgToolBars.Images.SetKeyName(45, "")
        Me.imgToolBars.Images.SetKeyName(46, "")
        Me.imgToolBars.Images.SetKeyName(47, "")
        Me.imgToolBars.Images.SetKeyName(48, "")
        Me.imgToolBars.Images.SetKeyName(49, "")
        Me.imgToolBars.Images.SetKeyName(50, "")
        Me.imgToolBars.Images.SetKeyName(51, "")
        Me.imgToolBars.Images.SetKeyName(52, "")
        Me.imgToolBars.Images.SetKeyName(53, "")
        Me.imgToolBars.Images.SetKeyName(54, "")
        Me.imgToolBars.Images.SetKeyName(55, "")
        Me.imgToolBars.Images.SetKeyName(56, "")
        Me.imgToolBars.Images.SetKeyName(57, "")
        Me.imgToolBars.Images.SetKeyName(58, "")
        Me.imgToolBars.Images.SetKeyName(59, "")
        Me.imgToolBars.Images.SetKeyName(60, "")
        Me.imgToolBars.Images.SetKeyName(61, "")
        Me.imgToolBars.Images.SetKeyName(62, "")
        Me.imgToolBars.Images.SetKeyName(63, "")
        Me.imgToolBars.Images.SetKeyName(64, "")
        Me.imgToolBars.Images.SetKeyName(65, "")
        Me.imgToolBars.Images.SetKeyName(66, "")
        Me.imgToolBars.Images.SetKeyName(67, "")
        Me.imgToolBars.Images.SetKeyName(68, "")
        Me.imgToolBars.Images.SetKeyName(69, "")
        Me.imgToolBars.Images.SetKeyName(70, "")
        Me.imgToolBars.Images.SetKeyName(71, "")
        Me.imgToolBars.Images.SetKeyName(72, "")
        Me.imgToolBars.Images.SetKeyName(73, "")
        Me.imgToolBars.Images.SetKeyName(74, "")
        Me.imgToolBars.Images.SetKeyName(75, "")
        Me.imgToolBars.Images.SetKeyName(76, "")
        Me.imgToolBars.Images.SetKeyName(77, "")
        Me.imgToolBars.Images.SetKeyName(78, "")
        Me.imgToolBars.Images.SetKeyName(79, "")
        Me.imgToolBars.Images.SetKeyName(80, "")
        Me.imgToolBars.Images.SetKeyName(81, "")
        Me.imgToolBars.Images.SetKeyName(82, "")
        Me.imgToolBars.Images.SetKeyName(83, "")
        Me.imgToolBars.Images.SetKeyName(84, "")
        Me.imgToolBars.Images.SetKeyName(85, "")
        Me.imgToolBars.Images.SetKeyName(86, "")
        '
        'btnExportExcel
        '
        Me.btnExportExcel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.FocusHighlighter1.SetHighlight(Me.btnExportExcel, False)
        Me.btnExportExcel.Icon = CType(resources.GetObject("btnExportExcel.Icon"), System.Drawing.Icon)
        Me.btnExportExcel.ImageHorizontalAlignment = Janus.Windows.EditControls.ImageHorizontalAlignment.Near
        Me.btnExportExcel.ImageIndex = 0
        Me.btnExportExcel.ImageList = Me.imgToolBars
        Me.btnExportExcel.Location = New System.Drawing.Point(4, 20)
        Me.btnExportExcel.Name = "btnExportExcel"
        Me.btnExportExcel.Size = New System.Drawing.Size(88, 25)
        Me.btnExportExcel.TabIndex = 12
        Me.btnExportExcel.Text = "&Export Excel"
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.FocusHighlighter1.SetHighlight(Me.btnClose, False)
        Me.btnClose.ImageHorizontalAlignment = Janus.Windows.EditControls.ImageHorizontalAlignment.Near
        Me.btnClose.ImageIndex = 76
        Me.btnClose.ImageList = Me.imgToolBars
        Me.btnClose.Location = New System.Drawing.Point(446, 20)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(80, 25)
        Me.btnClose.TabIndex = 11
        Me.btnClose.Text = "Close"
        '
        'btnPrintPreview
        '
        Me.btnPrintPreview.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.FocusHighlighter1.SetHighlight(Me.btnPrintPreview, False)
        Me.btnPrintPreview.ImageHorizontalAlignment = Janus.Windows.EditControls.ImageHorizontalAlignment.Near
        Me.btnPrintPreview.ImageIndex = 68
        Me.btnPrintPreview.ImageList = Me.imgToolBars
        Me.btnPrintPreview.Location = New System.Drawing.Point(338, 20)
        Me.btnPrintPreview.Name = "btnPrintPreview"
        Me.btnPrintPreview.Size = New System.Drawing.Size(100, 25)
        Me.btnPrintPreview.TabIndex = 10
        Me.btnPrintPreview.Text = "&Print Preview"
        '
        'FocusHighlighter1
        '
        Me.FocusHighlighter1.ChangeColors = True
        Me.FocusHighlighter1.HighlightBackColor = System.Drawing.Color.Yellow
        Me.FocusHighlighter1.HighlightForeColor = System.Drawing.Color.Black
        Me.FocusHighlighter1.LabelOrientation = SmartBooks.Ex.EWinform.FocusHighlighter.LabelOrientations.UpAbove
        Me.FocusHighlighter1.MakeBold = True
        '
        'errProvider
        '
        Me.errProvider.ContainerControl = Me
        '
        'rptGLDetail
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 14)
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(218, Byte), Integer), CType(CType(234, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(542, 376)
        Me.Controls.Add(Me.UiGroupBox1)
        Me.Controls.Add(Me.Panel1)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.FocusHighlighter1.SetHighlight(Me, False)
        Me.KeyPreview = True
        Me.Name = "rptGLDetail"
        Me.Text = "Ledger Account Listing Detail Report"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.OtherOption, System.ComponentModel.ISupportInitialize).EndInit()
        Me.OtherOption.ResumeLayout(False)
        Me.OtherOption.PerformLayout()
        CType(Me.cboOptionOther, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UiGroupBox3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UiGroupBox3.ResumeLayout(False)
        Me.UiGroupBox3.PerformLayout()
        CType(Me.cboCuryID, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UiGroupBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UiGroupBox2.ResumeLayout(False)
        Me.UiGroupBox2.PerformLayout()
        CType(Me.UiGroupBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UiGroupBox1.ResumeLayout(False)
        CType(Me.errProvider, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private cury As clsCurrency
    Private gltran As clsGLTran
    Dim tbAccount As New ArrayList

#Region "Private Methods"
    Private Sub Init()
        cury = New clsCurrency
        gltran = New clsGLTran
        'cboFromDate.Text = DateTime.Now.ToString("dd/MM/yyyy")
        'cboToDate.Text = DateTime.Now.ToString("dd/MM/yyyy")
        ExportToExcel = New SmartBooks.General.ClsExportToExcel
        txtAcct.Visible = False
        BindData4Cury()
        cboCuryID.Value = "VND"

        BindData4OtherOption()
        cboOptionOther.Value = "Option"
        cbOption.Checked = False
        OtherOption.Visible = False

        If DBConnection.Language = "VN" Then
            cbOption.Text = "Lựa chọn báo cáo khác"
        ElseIf DBConnection.Language = "EN" Then
            cbOption.Text = "Other option for Export Excel"
        ElseIf DBConnection.Language = "KR" Then
            cbOption.Text = "기타 옵션"
        End If

    End Sub
    Private Sub BindData4OtherOption()
        Dim mTable As DataTable
        Dim mRow As DataRow

        mTable = New DataTable
        mTable.Columns.Add(New DataColumn("ID", GetType(String)))
        mTable.Columns.Add(New DataColumn("Descr", GetType(String)))

        mRow = mTable.NewRow
        mRow("ID") = "Option"
        Select Case DBConnection.Language
            Case "VN"
                mRow("Descr") = "Lựa chọn"
            Case "EN"
                mRow("Descr") = "Option"
            Case "KR"
                mRow("Descr") = "선택"
        End Select
        mTable.Rows.Add(mRow)

        mRow = mTable.NewRow
        mRow("ID") = "FullReport"
        Select Case DBConnection.Language
            Case "VN"
                mRow("Descr") = "Báo cáo đầy đủ (Sổ chi tiết TK, CĐKT, CĐPS, KQKD)"
            Case "EN"
                mRow("Descr") = "Full Report (GL Detail, BS, TB, P&L)"
            Case "KR"
                mRow("Descr") = "전체 보고서 (GL Detail, BS, TB, P&L)"
        End Select
        mTable.Rows.Add(mRow)

        mRow = mTable.NewRow
        mRow("ID") = "MainAccount"
        Select Case DBConnection.Language
            Case "VN"
                mRow("Descr") = "Sổ chi tiết cấp 1"
            Case "EN"
                mRow("Descr") = "Main Account"
            Case "KR"
                mRow("Descr") = "메인 계정"
        End Select
        mTable.Rows.Add(mRow)

        mRow = mTable.NewRow
        mRow("ID") = "FinancialStatement"
        Select Case DBConnection.Language
            Case "VN"
                mRow("Descr") = "Báo cáo tài chính (CĐKT, CĐPS, KQKD)"
            Case "EN"
                mRow("Descr") = "Financial Statements (BS, TB, P&L)"
            Case "KR"
                mRow("Descr") = "재무제표"
        End Select
        mTable.Rows.Add(mRow)

        cboOptionOther.DataSource = mTable
        cboOptionOther.DropDownList.Columns("Descr").Width = 300 ''cboOptionOther.Width
        cboOptionOther.DropDownList.Columns("ID").Width = 110
        cboOptionOther.DisplayMember = "Descr"
        cboOptionOther.ValueMember = "ID"
    End Sub

    Private Sub BindData4Cury()
        cboCuryID.DataSource = cury.GetAll
        cboCuryID.DropDownList.Columns("CuryID").Width = cboCuryID.Width
        cboCuryID.DisplayMember = "CuryID"
        cboCuryID.ValueMember = "CuryID"
    End Sub
    Public Sub startInit()
        Init()
    End Sub
    Private Sub PreviewReport()
        If rbtnAll.Checked = True Then
            PreviewReportAllAccount()
        Else
            PreviewReportByAcct()
        End If
        'Dim rptViewer As New ReportViewer
        'Dim strReportFile As String
        '' gltran.GLPostedBegBal(cboFromDate.Value.Date, cboToDate.Value.Date)
        'gltran.GLPosted(cboFromDate.Value.Date, cboToDate.Value.Date)

        'If rbtnAll.Checked = True Then
        '    'If DBConnection.Language = "VN" Then
        '    '    If rbtnVND.Checked = True Then
        '    '        strReportFile = Application.StartupPath & "\" & "GLReports\rptGLDetail.rpt"
        '    '    Else
        '    '        strReportFile = Application.StartupPath & "\" & "GLReports\rptGLDetailUSD.rpt"
        '    '    End If
        '    'ElseIf DBConnection.Language = "EN" Then
        '    '    If rbtnVND.Checked = True Then
        '    '        strReportFile = Application.StartupPath & "\" & "GLReports\rptGLDetail_EN.rpt"
        '    '    Else
        '    '        strReportFile = Application.StartupPath & "\" & "GLReports\rptGLDetail_ENUSD.rpt"
        '    '    End If
        '    'Else
        '    '    If rbtnVND.Checked = True Then
        '    '        strReportFile = Application.StartupPath & "\" & "GLReports\rptGLDetail_KR.rpt"
        '    '    Else
        '    '        strReportFile = Application.StartupPath & "\" & "GLReports\rptGLDetail_KRUSD.rpt"
        '    '    End If
        '    'End If

        '    'Dim pValuesAll() As DateTime = {cboFromDate.Value.Date, cboToDate.Value.Date}
        '    'rptViewer.ShowReport(strReportFile, pValuesAll)
        'Else
        '    'If DBConnection.Language = "VN" Then
        '    '    'If rbtnVND.Checked = True Then
        '    '    '    strReportFile = Application.StartupPath & "\" & "GLReports\rptGLDetailByAcct.rpt"
        '    '    'Else
        '    '    '    strReportFile = Application.StartupPath & "\" & "GLReports\rptGLDetailByAcctUSD.rpt"
        '    '    'End If
        '    'Else
        '    '    'If rbtnVND.Checked = True Then
        '    '    '    strReportFile = Application.StartupPath & "\" & "GLReports\rptGLDetailByAcct_EN.rpt"
        '    '    'Else
        '    '    '    strReportFile = Application.StartupPath & "\" & "GLReports\rptGLDetailByAcct_ENUSD.rpt"
        '    '    'End If
        '    'End If

        '    'Dim pValuesAcct() As Object = {cboFromDate.Value.Date, cboToDate.Value.Date, txtAcct.Text.Trim}
        '    'rptViewer.ShowReport(strReportFile, pValuesAcct)
        'End If

        'rptViewer.Owner = Me
        'rptViewer.Show()
    End Sub

    Private Sub PreviewReportAllAccount()
        Dim tran As System.Data.SqlClient.SqlTransaction = DBConnection.Connection.BeginTransaction
        Try
            Dim rptViewer As New ReportViewer
            Dim strReportFile As String

            gltran.GLPostedBegBal(tran, StringToDate(cboFromDate.Text), StringToDate(cboToDate.Text))
            gltran.GLPosted(tran, StringToDate(cboFromDate.Text), StringToDate(cboToDate.Text))
            tran.Commit()
            If DBConnection.Language = "VN" Then
                If cboCuryID.Value = "VND" Then
                    If cbDisplayVNDUSD.Checked = True Then
                        strReportFile = Application.StartupPath & "\" & "GLReports\rptGLDetail_NT.rpt"
                    Else
                        strReportFile = Application.StartupPath & "\" & "GLReports\rptGLDetail.rpt"
                    End If

                Else
                    strReportFile = Application.StartupPath & "\" & "GLReports\rptGLDetailFC.rpt"
                End If
            Else
                '  If DBConnection.Language = "EN" Then
                If cboCuryID.Value = "VND" Then
                    strReportFile = Application.StartupPath & "\" & "GLReports\rptGLDetail_EN.rpt"
                Else
                    strReportFile = Application.StartupPath & "\" & "GLReports\rptGLDetailFC_EN.rpt"
                End If
                ' End If
            End If
            If cboCuryID.Value = "VND" Then
                Dim pValuesAll() As DateTime = {StringToDate(cboFromDate.Text), StringToDate(cboToDate.Text)}
                rptViewer.ShowReport(strReportFile, pValuesAll)
            Else
                Dim pValuesAll() As Object = {StringToDate(cboFromDate.Text), StringToDate(cboToDate.Text), cboCuryID.Value}
                rptViewer.ShowReport(strReportFile, pValuesAll)
            End If

            rptViewer.MdiParent = Me.MdiParent
            rptViewer.Show()
        Catch ex As Exception
            MsgBox(ex.Message)
            tran.Rollback()
        End Try

    End Sub

    Public Sub PreviewReportByAcct()
        Dim tran As System.Data.SqlClient.SqlTransaction = DBConnection.Connection.BeginTransaction
        Try
            Dim rptViewer As New ReportViewer
            Dim strReportFile As String
            gltran.GLPostedBegBal(tran, StringToDate(cboFromDate.Text), StringToDate(cboToDate.Text))
            gltran.GLPosted(tran, StringToDate(cboFromDate.Text), StringToDate(cboToDate.Text))
            tran.Commit()

            If DBConnection.Language = "VN" Then
                If cboCuryID.Value = "VND" Then
                    If cbDisplayVNDUSD.Checked = True Then
                        strReportFile = Application.StartupPath & "\" & "GLReports\rptGLDetailByAcct_NT.rpt"
                    Else
                        strReportFile = Application.StartupPath & "\" & "GLReports\rptGLDetailByAcct.rpt"
                    End If
                Else

                    strReportFile = Application.StartupPath & "\" & "GLReports\rptGLDetailByAcctFC.rpt"

                End If
                'If rbtnVND.Checked = True Then
                '    strReportFile = Application.StartupPath & "\" & "GLReports\rptGLDetailByAcct.rpt"
                'Else
                '    strReportFile = Application.StartupPath & "\" & "GLReports\rptGLDetailByAcctUSD.rpt"
                'End If
            Else
                If cboCuryID.Value = "VND" Then
                    strReportFile = Application.StartupPath & "\" & "GLReports\rptGLDetailByAcct_EN.rpt"
                Else
                    strReportFile = Application.StartupPath & "\" & "GLReports\rptGLDetailByAcctFC_EN.rpt"
                End If
                'If rbtnVND.Checked = True Then
                '    strReportFile = Application.StartupPath & "\" & "GLReports\rptGLDetailByAcct_EN.rpt"
                'Else
                '    strReportFile = Application.StartupPath & "\" & "GLReports\rptGLDetailByAcct_ENUSD.rpt"
                'End If
            End If
            If cboCuryID.Value = "VND" Then
                Dim pValuesAcct() As Object = {StringToDate(cboFromDate.Text), StringToDate(cboToDate.Text), txtAcct.Text.Trim}
                rptViewer.ShowReport(strReportFile, pValuesAcct)
            Else
                Dim pValuesAcct() As Object = {StringToDate(cboFromDate.Text), StringToDate(cboToDate.Text), txtAcct.Text.Trim, cboCuryID.Value}
                rptViewer.ShowReport(strReportFile, pValuesAcct)
            End If
            rptViewer.MdiParent = Me.MdiParent
            rptViewer.Show()
        Catch ex As Exception
            MsgBox(ex.Message)
            tran.Rollback()
        End Try

    End Sub

    Private Function CheckCurrencyIsMain() As Boolean
        Dim result As Boolean = False
        Return result
    End Function

    Public Sub ChangeLanguage()
        Dim ds As New DataSet

        Select Case DBConnection.Language
            Case "VN"
                'Me.Text = "Sổ cái chi tiết"
                ds.ReadXml(Application.StartupPath & "\AppLanguageVN.xml")

            Case "EN"
                'Me.Text = "Ledger Account Listing Detail Report"
                ds.ReadXml(Application.StartupPath & "\AppLanguageEN.xml")

            Case "KR"
                'Me.Text = "Ledger Account Listing Detail Report"
                ds.ReadXml(Application.StartupPath & "\AppLanguageKR.xml")
        End Select

        With ds.Tables("GLDetail").Rows(0)
            Me.btnClose.Text = .Item("close")
            Me.btnPrintPreview.Text = .Item("PrintPreview")

            Me.lblCuryID.Text = .Item("CurentcyID")
            Me.Label1.Text = .Item("Fromdate")
            Me.Label2.Text = .Item("ToDate")
            Me.rbtnAcct.Text = .Item("Account")
            Me.rbtnAll.Text = .Item("All")
            Me.UiGroupBox2.Text = .Item("Option")
            Me.UiGroupBox3.Text = .Item("selectCurency")
        End With
    End Sub
#End Region


    Private Sub rptGLDetail_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Init()
        ChangeLanguage()

    End Sub

    Private Sub btnClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Function IsValidate() As Boolean
        If (IsValidate_CheckFromDateToDate(cboFromDate, cboToDate, errProvider) = False) Then
            Return False
        End If
        Return True
    End Function
    Private Sub btnPrintPreview_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPrintPreview.Click
        If (IsValidate() = False) Then
            Exit Sub
        End If

        PreviewReport()
    End Sub

    Private Sub rbtnAll_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rbtnAll.CheckedChanged
        txtAcct.Visible = False
    End Sub

    Private Sub rbtnAcct_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rbtnAcct.CheckedChanged
        txtAcct.Visible = True
    End Sub

    Private Sub txtAcct_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtAcct.KeyDown
        If e.KeyCode = Keys.F3 Then
            Dim mFind As New FINDNET.FindProcess(New SqlClient.SqlConnection(DBConnection.strConnection))
            Dim strHeaders() As String = {"Account No", "Description (VN)", "Description (EN)"}
            Dim strColumns() As String = {"Acct", "DescrVN", "DescrEN"}
            Dim mRow As DataRow = mFind.Find(strHeaders, "Account", strColumns)

            If Not IsDBNull(mRow.Item(0)) Then
                txtAcct.Text = mRow.Item("Acct")
            End If
        End If
    End Sub

    Private Sub rptGLDetail_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        If e.Control Then
            Select Case e.KeyCode
                Case Keys.P
                    btnPrintPreview_Click(sender, e)

                    'Case Keys.Q
                    '    btnClose_Click(sender, e)
            End Select
        End If
        If e.Alt Then
            Select Case e.KeyCode
                Case Keys.E
                    btnExportExcel_Click(sender, e)
            End Select
        End If

        If e.KeyData = Keys.F5 Then
            Me.Init()
        End If
    End Sub

#Region "Overrides"
    Protected Overrides Function ProcessCmdKey(ByRef msg As System.Windows.Forms.Message, ByVal keyData As System.Windows.Forms.Keys) As Boolean
        Dim keyCode As Keys = CType(msg.WParam.ToInt32, Keys) And Keys.KeyCode
        Try
            Const WM_KEYDOWN As Integer = &H100
            Const WM_KEYUP As Integer = &H101
            If msg.Msg = WM_KEYDOWN Then
                Select Case keyCode
                    Case Keys.Enter

                        If ActiveControl.Name = "btnPrintPreview" Then
                            btnPrintPreview_Click(Nothing, Nothing)
                            Return True
                        End If

                        If ActiveControl.Name = "btnClose" Then
                            Me.Close()
                            Return True
                        End If
                        '---------------------- cac enter khach----------------------
                        If Not TypeOf ActiveControl Is Button And ActiveControl.Name <> "" Then
                            Me.SelectNextControl(ActiveControl, True, True, True, True)
                            Return True
                        End If

                    Case Keys.Escape

                        Me.SelectNextControl(ActiveControl, False, True, True, True)

                        Return True
                End Select
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        Return MyBase.ProcessKeyEventArgs(msg)
    End Function
#End Region

    Private Sub btnExportExcel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExportExcel.Click
        If (IsValidate() = False) Then
            Exit Sub
        End If

        If cbOption.Checked = True Then
            If cboOptionOther.Value = "FullReport" Then
                ExportReportfull()
            ElseIf cboOptionOther.Value = "MainAccount" Then
                ExportReport_MainAccount()
            ElseIf cboOptionOther.Value = "FinancialStatement" Then
                ExportReport_FinancialStatement()
            Else
                MessageBox.Show("Please choose option for export excel", "Warning", MessageBoxButtons.OK)
            End If
        Else
            If rbtnAll.Checked = True Then
                ExportReport()
            Else
                rptGLdetailReport()
            End If
        End If

    End Sub
    Private Sub ExportReport()

        Dim tran As System.Data.SqlClient.SqlTransaction = DBConnection.Connection.BeginTransaction


        Try

            gltran.GLPostedBegBal(tran, StringToDate(cboFromDate.Text), StringToDate(cboToDate.Text))
            gltran.GLPosted(tran, StringToDate(cboFromDate.Text), StringToDate(cboToDate.Text))
            tran.Commit()

            Dim tbCompany As DataTable
            tbCompany = ExportToExcel.GetCompanyInformation()


            If cbDisplayVNDUSD.Checked = False Then
                If cboCuryID.Value = "VND" Then
                    tbResult = ExportToExcel.ExcuteGLDetailVNDReport(StringToDate(cboFromDate.Text), StringToDate(cboToDate.Text))
                Else
                    tbResult = ExportToExcel.ExcuteGLDetailAllReportFC(StringToDate(cboFromDate.Text), StringToDate(cboToDate.Text), cboCuryID.Value)
                End If

            Else
                tbResult = ExportToExcel.ExcuteGLDetailVNDUSD_2(StringToDate(cboFromDate.Text), StringToDate(cboToDate.Text))

            End If

            Dim dtExport, dtDK As DataTable
            Dim _row As DataRow
            Dim dtrow As DataRowView
            If tbResult.Rows.Count = 0 Then
                MsgBox("Nothing to Export")
                Exit Sub
            End If


            Dim fileChooser As SaveFileDialog = New SaveFileDialog
            fileChooser.Filter = "Excel File|*.xlsx"
            fileChooser.FileName = "rptGLDetail.xlsx"
            Dim result As DialogResult = fileChooser.ShowDialog()
            fileChooser.CheckFileExists = False

            If result = DialogResult.OK Then
                If ExportToExcel.ExcelOpen(fileChooser.FileName.ToString) Then
                    MessageBox.Show("The file is open, Please close the file or rename a new file ", "Warning", MessageBoxButtons.OK)
                    Exit Try
                    Exit Sub
                End If
                Cursor = Cursors.WaitCursor
                'Show Progress Bar
                btnExportExcel.Enabled = False
                Dim f As New frmShow
                f.Show()
                f.UltraProgressBar1.Maximum = tbResult.Rows.Count

                'Lay DanhSachTaiKhoan
                Dim tbAccount As New ArrayList

                For Each dr As DataRow In tbResult.Rows
                    If (tbAccount.Contains(dr("Acct")) = False) Then
                        tbAccount.Add(dr("Acct"))
                    End If
                Next

                Dim AccountStr As String = ""

                'open file exel with control XlsReport
                Dim urlTemplate As String
                If cbDisplayVNDUSD.Checked = True Then
                    urlTemplate = Application.StartupPath() + "\GLReports\RptGLDetail_EN_VNDUSD.xlsx"
                Else

                    If cboCuryID.Value = "VND" Then
                        If DBConnection.Language = "VN" Then
                            urlTemplate = Application.StartupPath() + "\GLReports\RptGLDetailAll.xlsx"
                        Else
                            urlTemplate = Application.StartupPath() + "\GLReports\RptGLDetailAll_EN.xlsx"
                        End If
                    Else

                        If DBConnection.Language = "VN" Then
                            urlTemplate = Application.StartupPath() + "\GLReports\RptGLDetailAllFC.xlsx"
                        Else
                            urlTemplate = Application.StartupPath() + "\GLReports\RptGLDetailAll_ENFC.xlsx"
                        End If
                    End If
                End If

                Dim urlOut As String = fileChooser.FileName()
                Dim outputFile As FileInfo = New FileInfo(fileChooser.FileName)
                Using excelPackage As ExcelPackage = New ExcelPackage(New FileInfo(urlTemplate))

                    Dim worksheet_Index As ExcelWorksheet = excelPackage.Workbook.Worksheets.Copy("Index", "DocsMap")
                    worksheet_Index = excelPackage.Workbook.Worksheets("DocsMap")

                    '' Lay thong tin tai khoan
                    Dim clsacct As New clsAccount
                    Dim tbaccountname As New ArrayList
                    Dim AcctName As DataTable
                    Dim arrList As New ArrayList
                    Dim row_b As DataRow
                    Dim row_a As DataRow
                    AcctName = clsacct.GetAll
                    Dim AccountStrName As String = ""
                    tbAccount.Sort()
                    'Dim row_a As DataRow
                    'For k As Integer = 0 To tbAccount.Count - 1
                    '    For Each row_a In AcctName.Select("acct = '" & tbAccount(k) & "'")
                    '        If DBConnection.Language = "VN" Then
                    '            tbaccountname.Add(row_a("DescrVN"))
                    '        ElseIf DBConnection.Language = "EN" Then
                    '            tbaccountname.Add(row_a("DescrEN"))
                    '        Else
                    '            tbaccountname.Add(row_a("DescrKR"))

                    '        End If
                    '    Next
                    'Next
                    Dim textErrAcc As String = ""
                    For Each row_b In AcctName.Rows
                        arrList.Add(row_b("Acct"))
                    Next
                    For k As Integer = 0 To tbAccount.Count - 1
                        If Not arrList.Contains(tbAccount(k)) Then
                            textErrAcc = textErrAcc + " " + tbAccount(k)
                        End If
                        For Each row_a In AcctName.Select("acct = '" & tbAccount(k) & "'")
                            If DBConnection.Language = "VN" Then ' default 0-VN 1-EN 2-KR
                                tbaccountname.Add(row_a("DescrVN"))
                            ElseIf DBConnection.Language = "KR" Then ' default 0-VN 1-EN 2-KR
                                tbaccountname.Add(row_a("DescrKR"))
                            Else
                                tbaccountname.Add(row_a("DescrEN"))
                            End If
                        Next
                    Next

                    'Check phat sinh nhung tai khoan khong co trong danh muc tai khoan
                    If textErrAcc <> "" Then
                        MessageBox.Show("Accounts:" + textErrAcc + " do not belong to Account List!", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        f.Close()
                        Exit Sub
                    End If


                    Dim index_d As Integer = 2

                    For j As Integer = 0 To tbAccount.Count - 1
                        AccountStr = tbAccount(j)
                        AccountStrName = tbaccountname(j)
                        worksheet_Index.Cells(2, 1, 2, 2).Copy(worksheet_Index.Cells(index_d, 1, index_d, 2))
                        worksheet_Index.Cells("A" + index_d.ToString()).Value = AccountStr
                        worksheet_Index.Cells("B" + index_d.ToString()).Value = AccountStrName
                        index_d = index_d + 1
                    Next


                    Dim beginIndex As Integer = 2

                    For i As Integer = 0 To tbAccount.Count - 1
                        worksheet_Index.Cells(i + beginIndex, 1).Formula = String.Format("=HYPERLINK(""#'{0}'!A1"",""{0}"")", tbAccount(i))
                    Next

                    If cbDisplayVNDUSD.Checked = False Then
                        For i As Integer = 0 To tbAccount.Count - 1
                            AccountStr = tbAccount(i)
                            Dim worksheet As ExcelWorksheet = excelPackage.Workbook.Worksheets.Copy("template", AccountStr)

                            If cboCuryID.Value = "VND" Then
                                CreateExcelSheet(worksheet, tbCompany, f, AccountStr)
                            Else
                                CreateExcelSheetFC(worksheet, tbCompany, f, AccountStr)
                            End If
                        Next
                    Else


                        For i As Integer = 0 To tbAccount.Count - 1
                            AccountStr = tbAccount(i)
                            Dim worksheet As ExcelWorksheet = excelPackage.Workbook.Worksheets.Copy("template", AccountStr)
                            CreateExcelSheetALL(worksheet, tbCompany, f, AccountStr)
                        Next
                    End If
                    excelPackage.Workbook.Worksheets.Delete("template")
                    excelPackage.Workbook.Worksheets.Delete("Index")
                    excelPackage.SaveAs(outputFile)

                End Using

                Dim ps As New ProcessStartInfo


                Dim D_result As DialogResult = MessageBox.Show("Do you want to open excel file ?", Application.ProductName, MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                If D_result = DialogResult.Yes Then
                    ps.UseShellExecute = True
                    ps.FileName = urlOut
                    Process.Start(ps)

                End If


                f.Close()
                btnExportExcel.Enabled = True
                Cursor = Cursors.Default

            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            tran.Rollback()
        End Try
    End Sub
    Private Sub ExportReport_MainAccount()

        Dim tran As System.Data.SqlClient.SqlTransaction = DBConnection.Connection.BeginTransaction


        Try

            gltran.GLPostedBegBal(tran, StringToDate(cboFromDate.Text), StringToDate(cboToDate.Text))
            gltran.GLPosted(tran, StringToDate(cboFromDate.Text), StringToDate(cboToDate.Text))
            tran.Commit()

            Dim tbCompany As DataTable
            tbCompany = ExportToExcel.GetCompanyInformation()


            tbResult = ExportToExcel.ExcuteGLDetailMainAcct(StringToDate(cboFromDate.Text), StringToDate(cboToDate.Text))


            Dim dtExport, dtDK As DataTable
            Dim _row As DataRow
            Dim dtrow As DataRowView
            If tbResult.Rows.Count = 0 Then
                MsgBox("Nothing to Export")
                Exit Sub
            End If


            Dim fileChooser As SaveFileDialog = New SaveFileDialog
            fileChooser.Filter = "Excel File|*.xlsx"
            fileChooser.FileName = "rptGLDetailMainAccount.xlsx"
            Dim result As DialogResult = fileChooser.ShowDialog()
            fileChooser.CheckFileExists = False

            If result = DialogResult.OK Then
                If ExportToExcel.ExcelOpen(fileChooser.FileName.ToString) Then
                    MessageBox.Show("The file is open, Please close the file or rename a new file ", "Warning", MessageBoxButtons.OK)
                    Exit Try
                    Exit Sub
                End If
                Cursor = Cursors.WaitCursor
                'Show Progress Bar
                btnExportExcel.Enabled = False
                Dim f As New frmShow
                f.Show()
                f.UltraProgressBar1.Maximum = tbResult.Rows.Count

                tbResult.DefaultView.Sort = "Acct ASC"
                'Lay DanhSachTaiKhoan
                Dim tbAccount As New ArrayList

                For Each dr As DataRow In tbResult.Rows
                    If (tbAccount.Contains(dr("Acct")) = False) Then
                        tbAccount.Add(dr("Acct"))
                    End If
                Next

                Dim AccountStr As String = ""

                'open file exel with control XlsReport
                Dim urlTemplate As String

                urlTemplate = Application.StartupPath() + "\GLReports\RptGLDetail_EN_VNDUSD.xlsx"

                Dim urlOut As String = fileChooser.FileName()
                Dim outputFile As FileInfo = New FileInfo(fileChooser.FileName)
                Using excelPackage As ExcelPackage = New ExcelPackage(New FileInfo(urlTemplate))

                    Dim worksheet_Index As ExcelWorksheet = excelPackage.Workbook.Worksheets.Copy("Index", "DocsMap")
                    worksheet_Index = excelPackage.Workbook.Worksheets("DocsMap")

                    '' Lay thong tin tai khoan
                    Dim clsacct As New clsAccount
                    Dim tbaccountname As New ArrayList
                    Dim AcctName As DataTable
                    AcctName = clsacct.GetAll
                    Dim AccountStrName As String = ""
                    Dim soluong As Integer
                    soluong = 0

                    Dim row_a As DataRow
                    For k As Integer = 0 To tbAccount.Count - 1
                        For Each row_a In AcctName.Select("acct = '" & tbAccount(k) & "'")
                            If DBConnection.Language = "VN" Then
                                tbaccountname.Add(row_a("DescrVN"))
                            ElseIf DBConnection.Language = "EN" Then
                                tbaccountname.Add(row_a("DescrEN"))
                            Else
                                tbaccountname.Add(row_a("DescrKR"))

                            End If
                            soluong = soluong + 1
                        Next
                    Next
                    Dim index_d As Integer = 2

                    For j As Integer = 0 To tbAccount.Count - 1
                        AccountStr = tbAccount(j)
                        AccountStrName = tbaccountname(j)
                        worksheet_Index.Cells(2, 1, 2, 2).Copy(worksheet_Index.Cells(index_d, 1, index_d, 2))
                        worksheet_Index.Cells("A" + index_d.ToString()).Value = AccountStr
                        worksheet_Index.Cells("B" + index_d.ToString()).Value = AccountStrName
                        index_d = index_d + 1
                    Next

                    'Dim j As Integer
                    'j = 1
                    'Do While j < tbAccount.Count
                    '    AccountStr = tbAccount(j)
                    '    AccountStrName = tbaccountname(j)
                    '    worksheet_Index.Cells(2, 1, 2, 2).Copy(worksheet_Index.Cells(index_d, 1, index_d, 2))
                    '    worksheet_Index.Cells("A" + index_d.ToString()).Value = AccountStr
                    '    worksheet_Index.Cells("B" + index_d.ToString()).Value = AccountStrName
                    '    index_d = index_d + 1
                    'Loop


                    Dim beginIndex As Integer = 2

                    For i As Integer = 0 To tbAccount.Count - 1
                        worksheet_Index.Cells(i + beginIndex, 1).Formula = String.Format("=HYPERLINK(""#'{0}'!A1"",""{0}"")", tbAccount(i))
                    Next

                    If cbDisplayVNDUSD.Checked = False Then
                        For i As Integer = 0 To tbAccount.Count - 1
                            AccountStr = tbAccount(i)
                            Dim worksheet As ExcelWorksheet = excelPackage.Workbook.Worksheets.Copy("template", AccountStr)

                            If cboCuryID.Value = "VND" Then
                                CreateExcelSheet(worksheet, tbCompany, f, AccountStr)
                            Else
                                CreateExcelSheetFC(worksheet, tbCompany, f, AccountStr)
                            End If
                        Next
                    Else


                        For i As Integer = 0 To tbAccount.Count - 1
                            AccountStr = tbAccount(i)
                            Dim worksheet As ExcelWorksheet = excelPackage.Workbook.Worksheets.Copy("template", AccountStr)
                            CreateExcelSheetALL(worksheet, tbCompany, f, AccountStr)
                        Next
                    End If
                    excelPackage.Workbook.Worksheets.Delete("template")
                    excelPackage.Workbook.Worksheets.Delete("Index")
                    excelPackage.SaveAs(outputFile)

                End Using

                Dim ps As New ProcessStartInfo


                Dim D_result As DialogResult = MessageBox.Show("Do you want to open excel file ?", Application.ProductName, MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                If D_result = DialogResult.Yes Then
                    ps.UseShellExecute = True
                    ps.FileName = urlOut
                    Process.Start(ps)

                End If

                ''Open File
                'Dim ps As New ProcessStartInfo
                'ps.UseShellExecute = True
                'ps.FileName = urlOut
                'Process.Start(ps)

                f.Close()
                btnExportExcel.Enabled = True
                Cursor = Cursors.Default

            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            tran.Rollback()
        End Try
    End Sub



    Private Sub CreateExcelSheet(ByRef worksheet As ExcelWorksheet, ByVal tbCompany As DataTable, ByVal f As frmShow, ByVal account As String)

        'end file exel with control XlsReport

        worksheet.Cells("B7").Value = "Từ ngày:  " + cboFromDate.Text + " Đến ngày:     " + cboToDate.Text
        worksheet.Cells("B8").Value = "Tài khoản:    " + account
        'Pulmuone
        'worksheet.Cells("B9").Value = account
        'worksheet.Cells("B11").Value = account

        'process data for exel 
        Dim index As Integer = 14

        Dim row As DataRow
        Dim datetime As String


        'Show Progress Bar
        Dim i As Integer
        Dim PSNo, PSCo, PSNoNT, PSCoNT As Double ' So phat sinh No va Co trong ky
        PSNo = 0
        PSCo = 0
        PSNoNT = 0
        PSCoNT = 0
        Dim SDDK, SDDK_NT As Double ' So du dau ky
        SDDK = 0
        SDDK_NT = 0
        Dim TotalDr, TotalCr, TotalDr_NT, TotalCr_NT As Double  ' Tong phat sinh bao gom ca so du dau ky dung de tinh gia tri ton cuoi ky No hay Co
        TotalDr = 0  ' Tong Phat Sinh No bao gom so du dau ky
        TotalCr = 0 ' Tong Phat sinh Co bao gom so du dau ky
        TotalDr_NT = 0  ' Tong Phat Sinh No bao gom so du dau ky NT
        TotalCr_NT = 0 ' Tong Phat sinh Co bao gom so du dau ky NT

        ' For Each row In tbResult.Rows
        For Each row In tbResult.Select("Acct = '" & account & "' AND Module <>''", "TranDate ASC")

            'PRocess Bar
            If f.UltraProgressBar1.Value < f.UltraProgressBar1.Maximum Then
                f.UltraProgressBar1.Value += 1
                f.UltraProgressBar1.Refresh()
                f.Activate()
            End If


            index = index + 1
            worksheet.Cells(14, 1, 14, 12).Copy(worksheet.Cells(index, 1, index, 12))

            worksheet.Cells("B" + index.ToString()).Value = row("TranDate") ' Ngay chung Tu
            worksheet.Cells("C" + index.ToString()).Value = row("RefNbr") ' So Chung Tu
            worksheet.Cells("D" + index.ToString()).Value = row("Module") ' Phan He
            If DBConnection.Language = "KR" Then
                worksheet.Cells("E" + index.ToString()).Value = row("DescrEN") ' Dien Giai tiếng anh
            ElseIf DBConnection.Language = "EN" Then
                worksheet.Cells("E" + index.ToString()).Value = row("DescrEN") ' Dien Giai tiếng anh
            Else
                worksheet.Cells("E" + index.ToString()).Value = row("TranDescr") ' Dien Giai
            End If
            'Xls.Cells("E" + index.ToString()).Value = row("TranDescr") ' Dien Giai
            worksheet.Cells("F" + index.ToString()).Value = row("ID") ' So Hoa Don
            worksheet.Cells("G" + index.ToString()).Value = row("AcctRef") ' Tai Khoan Doi Ung
            worksheet.Cells("H" + index.ToString()).Value = row("DrAmt") ' So Tien Phat Sinh No
            worksheet.Cells("I" + index.ToString()).Value = row("CrAmt") ' So Tien Phat Sinh Co
            'worksheet.Cells("J" + index.ToString()).Value = row("CuryDrAmt") ' So Tien Phat Sinh No Ngoai Te
            'worksheet.Cells("K" + index.ToString()).Value = row("CuryCrAmt") ' So Tien Phat Sinh Co Ngoai Te

            PSNo += row("DrAmt")
            PSCo += row("CrAmt")
            'PSNoNT += row("CuryDrAmt")
            'PSCoNT += row("CuryCrAmt")

        Next

        '' Xu ly so du dau ky



        'For Each row In tbResult.Select("Module =''")
        '    SDDK += row("BegAmt")
        'Next
        '' thanhln sua lai ngay 11/07/2017

        For Each row In tbResult.Select("Acct = '" & account & "' AND AcctRef =''")
            SDDK += row("BegAmt")
            'SDDK_NT += row("CuryBegAmt")
        Next
        If SDDK >= 0 Then
            worksheet.Cells("H13").Value = SDDK
            worksheet.Cells("I13").Value = 0
            TotalDr = PSNo + SDDK
            TotalCr = PSCo

        Else
            worksheet.Cells("H13").Value = 0
            worksheet.Cells("I13").Value = (-1) * SDDK
            TotalDr = PSNo
            TotalCr = PSCo + (-1) * SDDK


        End If

        If SDDK_NT >= 0 Then
            worksheet.Cells("J13").Value = SDDK_NT
            worksheet.Cells("K13").Value = 0
            TotalDr_NT = PSNoNT + SDDK_NT
            TotalCr_NT = PSCoNT

        Else
            worksheet.Cells("J13").Value = 0
            worksheet.Cells("K13").Value = (-1) * SDDK_NT
            TotalDr_NT = PSNoNT
            TotalCr_NT = PSCoNT + (-1) * SDDK_NT
        End If

        ' worksheet.Cells("D" + (index + 1).ToString()).Formula = "=Sum(D14:D" + index.ToString() + ")"
        worksheet.Cells("H" + (index + 1).ToString()).Value = PSNo
        worksheet.Cells("I" + (index + 1).ToString()).Value = PSCo
        worksheet.Cells("E" + (index + 1).ToString()).Value = "Cộng Phát Sinh"
        'worksheet.Cells("J" + (index + 1).ToString()).Value = PSNoNT
        'worksheet.Cells("K" + (index + 1).ToString()).Value = PSCoNT

        '' end so du dau ky

        '' Xu ly so du cuoi ky

        If TotalDr > TotalCr Then
            worksheet.Cells("H" + (index + 2).ToString()).Value = TotalDr - TotalCr
            worksheet.Cells("I" + (index + 2).ToString()).Value = 0

        Else
            worksheet.Cells("H" + (index + 2).ToString()).Value = 0
            worksheet.Cells("I" + (index + 2).ToString()).Value = TotalCr - TotalDr

        End If

        ''' End xu ly so du cuoi ky
        'If TotalDr_NT > TotalCr_NT Then
        '    worksheet.Cells("J" + (index + 2).ToString()).Value = TotalDr_NT - TotalCr_NT
        '    worksheet.Cells("K" + (index + 2).ToString()).Value = 0

        'Else
        '    worksheet.Cells("J" + (index + 2).ToString()).Value = 0
        '    worksheet.Cells("K" + (index + 2).ToString()).Value = TotalCr_NT - TotalDr_NT

        'End If


        worksheet.Cells("B2").Value = tbCompany.Rows(0).Item("CpNyName")
        worksheet.Cells("B3").Value = tbCompany.Rows(0).Item("Address")

        worksheet.Cells("B10").Formula = String.Format("=HYPERLINK(""#'DOCSMAP'!A1"",""DOCSMAP"")")

        worksheet.Cells("B10").Style.Font.Bold = True
        worksheet.Cells("B10").Style.Font.Italic = True
        Dim cellFormat As String = String.Format("B{0}:K{0}", (index + 2).ToString())
        worksheet.Cells(String.Format("B{0}:K{0}", (index + 1).ToString())).Style.Fill.PatternType = ExcelFillStyle.Solid
        worksheet.Cells(String.Format("B{0}:K{0}", (index + 1).ToString())).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#FFF2CC"))
        worksheet.Cells(String.Format("B{0}:K{0}", (index + 2).ToString())).Style.Fill.PatternType = ExcelFillStyle.Solid
        worksheet.Cells(String.Format("B{0}:K{0}", (index + 2).ToString())).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#DDEBF7"))
        worksheet.Cells(String.Format("B{0}:K{0}", (index + 1).ToString())).Style.Font.Italic = True
        worksheet.Cells(String.Format("B{0}:K{0}", (index + 1).ToString())).Style.Font.Bold = True
        worksheet.Cells(String.Format("B{0}:K{0}", (index + 2).ToString())).Style.Font.Bold = True
        worksheet.Cells(cellFormat).Style.Border.Top.Style = ExcelBorderStyle.Hair
        worksheet.Cells(cellFormat).Style.Border.Bottom.Style = ExcelBorderStyle.Thin

        worksheet.Column(8).AutoFit()
        worksheet.Column(9).AutoFit()
        worksheet.DeleteRow(14)
        'end process data for exel
        'End operation with file exel
    End Sub

    Private Sub rptGLdetailReport()
        Dim tran As System.Data.SqlClient.SqlTransaction = DBConnection.Connection.BeginTransaction
        Try

            gltran.GLPostedBegBal(tran, StringToDate(cboFromDate.Text), StringToDate(cboToDate.Text))
            gltran.GLPosted(tran, StringToDate(cboFromDate.Text), StringToDate(cboToDate.Text))
            tran.Commit()
            Dim tbCompany As DataTable
            tbCompany = ExportToExcel.GetCompanyInformation()
            If cbDisplayVNDUSD.Checked = False Then
                If cboCuryID.Value = "VND" Then
                    tbResult = ExportToExcel.ExcuteGLDetailReport(StringToDate(cboFromDate.Text), StringToDate(cboToDate.Text), txtAcct.Text)
                Else
                    tbResult = ExportToExcel.ExcuteGLDetailReportFC(StringToDate(cboFromDate.Text), StringToDate(cboToDate.Text), txtAcct.Text, cboCuryID.Value)
                End If
            Else
                tbResult = ExportToExcel.ExcuteGLDetailReport_ByAcct(StringToDate(cboFromDate.Text), StringToDate(cboToDate.Text), txtAcct.Text)

            End If

            Dim dtExport, dtDK As DataTable
            Dim _row As DataRow
            Dim dtrow As DataRowView
            Dim filter As String
            filter = ""
            If tbResult.Rows.Count = 0 Then
                MsgBox("Nothing to Export")
                Exit Sub
            End If


            Dim fileChooser As SaveFileDialog = New SaveFileDialog
            fileChooser.Filter = "Excel File|*.xlsx"
            If cbDisplayVNDUSD.Checked = False Then
                fileChooser.FileName = "rptGLDetailByAcct_" + txtAcct.Text + ".xlsx"
            Else
                fileChooser.FileName = "rptGLDetailCuryByAcct_" + txtAcct.Text + ".xlsx"
            End If

            Dim result As DialogResult = fileChooser.ShowDialog()
            fileChooser.CheckFileExists = False

            If result = DialogResult.OK Then
                If ExportToExcel.ExcelOpen(fileChooser.FileName.ToString) Then
                    MessageBox.Show("The file is open, Please close the file or rename a new file ", "Warning", MessageBoxButtons.OK)
                    Exit Try
                    Exit Sub
                End If
                Cursor = Cursors.WaitCursor
                'Show Progress Bar
                btnExportExcel.Enabled = False
                Dim f As New frmShow
                f.Show()

                Dim index As Integer = 14
                'open file exel with control XlsReport
                Dim urlTemplate As String

                If cbDisplayVNDUSD.Checked = False Then
                    If DBConnection.Language = "KR" Then
                        urlTemplate = Application.StartupPath() + "\GLReports\RptGLDetailByAcct_EN.xlsx"
                    ElseIf DBConnection.Language = "EN" Then
                        urlTemplate = Application.StartupPath() + "\GLReports\RptGLDetailByAcct_EN.xlsx"
                    Else
                        urlTemplate = Application.StartupPath() + "\GLReports\RptGLDetailByAcct.xlsx"
                    End If
                Else
                    urlTemplate = Application.StartupPath() + "\GLReports\RptGLDetail_EN_VNDUSD_ByAcct.xlsx"
                End If

                Dim outPutFile As FileInfo = New FileInfo(fileChooser.FileName)
                Dim urlOut As String = fileChooser.FileName()

                Using excelPackage As ExcelPackage = New ExcelPackage(New FileInfo(urlTemplate))
                    Dim worksheet As ExcelWorksheet = excelPackage.Workbook.Worksheets(1)

                    worksheet.Cells("B7").Value = "Từ ngày:  " + cboFromDate.Text + " Đến ngày:     " + cboToDate.Text
                    worksheet.Cells("B8").Value = "Tài khoản:    " + txtAcct.Text


                    'process data for exel 

                    Dim count As Integer = 1
                    Dim row As DataRow
                    Dim datetime As String


                    'Show Progress Bar
                    f.UltraProgressBar1.Maximum = tbResult.Rows.Count
                    Dim i As Integer
                    Dim PSNo, PSCo, PSNoNT, PSCoNT As Double   ' So phat sinh No va Co trong ky
                    PSNo = 0
                    PSCo = 0
                    PSNoNT = 0
                    PSCoNT = 0

                    Dim SDDK, SDDK_NT As Double  ' So du dau ky
                    SDDK = 0
                    SDDK_NT = 0
                    Dim TotalDr, TotalCr, TotalDr_NT, TotalCr_NT As Double   ' Tong phat sinh bao gom ca so du dau ky dung de tinh gia tri ton cuoi ky No hay Co
                    TotalDr = 0  ' Tong Phat Sinh No bao gom so du dau ky
                    TotalCr = 0 ' Tong Phat sinh Co bao gom so du dau ky
                    TotalDr_NT = 0  ' Tong Phat Sinh No bao gom so du dau ky NT
                    TotalCr_NT = 0 ' Tong Phat sinh Co bao gom so du dau ky NT

                    ' For Each row In tbResult.Rows
                    If cbDisplayVNDUSD.Checked = False Then
                        For Each row In tbResult.Select("Module <>'" + filter + "'", "TranDate ASC")

                            'PRocess Bar
                            If f.UltraProgressBar1.Value < f.UltraProgressBar1.Maximum Then
                                f.UltraProgressBar1.Value += 1
                                f.UltraProgressBar1.Refresh()
                                f.Activate()
                            End If

                            index = index + 1
                            worksheet.Cells(14, 1, 14, 11).Copy(worksheet.Cells(index, 1, index, 11))

                            worksheet.Cells("B" + index.ToString()).Value = row("TranDate") ' Ngay chung Tu
                            worksheet.Cells("C" + index.ToString()).Value = row("RefNbr") ' So Chung Tu
                            worksheet.Cells("D" + index.ToString()).Value = row("Module") ' Phan He
                            If DBConnection.Language = "KR" Then
                                worksheet.Cells("E" + index.ToString()).Value = row("DescrEN") ' Dien Giai tiếng anh
                            ElseIf DBConnection.Language = "EN" Then
                                worksheet.Cells("E" + index.ToString()).Value = row("DescrEN") ' Dien Giai tiếng anh
                            Else
                                worksheet.Cells("E" + index.ToString()).Value = row("TranDescr") ' Dien Giai
                            End If
                            'Xls.Cells("E" + index.ToString()).Value = row("TranDescr") ' Dien Giai
                            worksheet.Cells("F" + index.ToString()).Value = row("ID") ' So Hoa Don
                            worksheet.Cells("G" + index.ToString()).Value = row("AcctRef") ' Tai Khoan Doi Ung
                            worksheet.Cells("H" + index.ToString()).Value = row("DrAmt") ' So Tien Phat Sinh No
                            worksheet.Cells("I" + index.ToString()).Value = row("CrAmt") ' So Tien Phat Sinh Co


                            count = count + 1
                            PSNo += row("DrAmt")
                            PSCo += row("CrAmt")

                        Next
                    Else  '' Export ca USD va VND
                        For Each row In tbResult.Select("Module <>'" + filter + "'", "TranDate ASC")

                            'PRocess Bar
                            If f.UltraProgressBar1.Value < f.UltraProgressBar1.Maximum Then
                                f.UltraProgressBar1.Value += 1
                                f.UltraProgressBar1.Refresh()
                                f.Activate()
                            End If

                            index = index + 1
                            worksheet.Cells(14, 1, 14, 11).Copy(worksheet.Cells(index, 1, index, 11))

                            worksheet.Cells("B" + index.ToString()).Value = row("TranDate") ' Ngay chung Tu
                            worksheet.Cells("C" + index.ToString()).Value = row("RefNbr") ' So Chung Tu
                            worksheet.Cells("D" + index.ToString()).Value = row("Module") ' Phan He
                            If DBConnection.Language = "KR" Then
                                worksheet.Cells("E" + index.ToString()).Value = row("DescrEN") ' Dien Giai tiếng anh
                            ElseIf DBConnection.Language = "EN" Then
                                worksheet.Cells("E" + index.ToString()).Value = row("DescrEN") ' Dien Giai tiếng anh
                            Else
                                worksheet.Cells("E" + index.ToString()).Value = row("TranDescr") ' Dien Giai
                            End If
                            'Xls.Cells("E" + index.ToString()).Value = row("TranDescr") ' Dien Giai
                            worksheet.Cells("F" + index.ToString()).Value = row("ID") ' So Hoa Don
                            worksheet.Cells("G" + index.ToString()).Value = row("AcctRef") ' Tai Khoan Doi Ung
                            worksheet.Cells("H" + index.ToString()).Value = row("DrAmt") ' So Tien Phat Sinh No
                            worksheet.Cells("I" + index.ToString()).Value = row("CrAmt") ' So Tien Phat Sinh Co
                            worksheet.Cells("J" + index.ToString()).Value = row("CuryDrAmt") ' So Tien Phat Sinh No Ngoai Te
                            worksheet.Cells("K" + index.ToString()).Value = row("CuryCrAmt") ' So Tien Phat Sinh Co Ngoai Te
                            count = count + 1
                            PSNo += row("DrAmt")
                            PSCo += row("CrAmt")
                            PSNoNT += row("CuryDrAmt")
                            PSCoNT += row("CuryCrAmt")

                        Next
                    End If


                    '' Xu ly dau ky
                    For Each row In tbResult.Select("Module ='" + filter + "'", "TranDate ASC")
                        If cbDisplayVNDUSD.Checked = True Then
                            SDDK += row("BegAmt")
                            SDDK_NT += row("CuryBegAmt")
                        Else
                            SDDK += row("BegAmt")
                        End If

                    Next
                    If cbDisplayVNDUSD.Checked = True Then
                        If SDDK >= 0 Then
                            worksheet.Cells("H13").Value = SDDK
                            worksheet.Cells("I13").Value = 0
                            TotalDr = PSNo + SDDK
                            TotalCr = PSCo
                        Else
                            worksheet.Cells("H13").Value = 0
                            worksheet.Cells("I13").Value = (-1) * SDDK
                            TotalDr = PSNo
                            TotalCr = PSCo + (-1) * SDDK
                        End If
                        If SDDK_NT >= 0 Then
                            worksheet.Cells("J13").Value = SDDK_NT
                            worksheet.Cells("K13").Value = 0
                            TotalDr_NT = PSNoNT + SDDK_NT
                            TotalCr_NT = PSCoNT
                        Else
                            worksheet.Cells("J13").Value = 0
                            worksheet.Cells("K13").Value = (-1) * SDDK_NT
                            TotalDr_NT = PSNoNT
                            TotalCr_NT = PSCoNT + (-1) * SDDK_NT
                        End If
                    Else
                        If SDDK >= 0 Then
                            worksheet.Cells("H13").Value = SDDK
                            worksheet.Cells("I13").Value = 0
                            TotalDr = PSNo + SDDK
                            TotalCr = PSCo
                        Else
                            worksheet.Cells("H13").Value = 0
                            worksheet.Cells("I13").Value = (-1) * SDDK
                            TotalDr = PSNo
                            TotalCr = PSCo + (-1) * SDDK
                        End If
                    End If


                    '' End xu ly dau ky
                    worksheet.Cells("H" + (index + 1).ToString()).Value = PSNo
                    worksheet.Cells("I" + (index + 1).ToString()).Value = PSCo
                    worksheet.Cells("E" + (index + 1).ToString()).Value = "Cộng Phát Sinh"
                    If cbDisplayVNDUSD.Checked = False Then
                        worksheet.Cells("H" + (index + 1).ToString()).Value = PSNo
                        worksheet.Cells("I" + (index + 1).ToString()).Value = PSCo
                        worksheet.Cells("E" + (index + 1).ToString()).Value = "Cộng Phát Sinh"
                    Else

                        worksheet.Cells("H" + (index + 1).ToString()).Value = PSNo
                        worksheet.Cells("I" + (index + 1).ToString()).Value = PSCo
                        worksheet.Cells("E" + (index + 1).ToString()).Value = "Cộng Phát Sinh | Total"
                        worksheet.Cells("J" + (index + 1).ToString()).Value = PSNoNT
                        worksheet.Cells("K" + (index + 1).ToString()).Value = PSCoNT
                    End If

                    '' end so du dau ky

                    '' Xu ly so du cuoi ky

                    If TotalDr > TotalCr Then
                        worksheet.Cells("H" + (index + 2).ToString()).Value = TotalDr - TotalCr
                        worksheet.Cells("I" + (index + 2).ToString()).Value = 0

                    Else
                        worksheet.Cells("H" + (index + 2).ToString()).Value = 0
                        worksheet.Cells("I" + (index + 2).ToString()).Value = TotalCr - TotalDr

                    End If
                    If TotalDr_NT > TotalCr_NT Then
                        worksheet.Cells("J" + (index + 2).ToString()).Value = TotalDr_NT - TotalCr_NT
                        worksheet.Cells("K" + (index + 2).ToString()).Value = 0

                    Else
                        worksheet.Cells("J" + (index + 2).ToString()).Value = 0
                        worksheet.Cells("K" + (index + 2).ToString()).Value = TotalCr_NT - TotalDr_NT

                    End If


                    worksheet.Cells("B2").Value = tbCompany.Rows(0).Item("CpNyName")
                    worksheet.Cells("B3").Value = tbCompany.Rows(0).Item("Address")
                    If cbDisplayVNDUSD.Checked = False Then
                        Dim cellFormat As String = String.Format("B{0}:I{0}", (index + 2).ToString())
                        worksheet.Cells(String.Format("B{0}:I{0}", (index + 1).ToString())).Style.Fill.PatternType = ExcelFillStyle.Solid
                        worksheet.Cells(String.Format("B{0}:I{0}", (index + 1).ToString())).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#FFF2CC"))
                        worksheet.Cells(String.Format("B{0}:I{0}", (index + 2).ToString())).Style.Fill.PatternType = ExcelFillStyle.Solid
                        worksheet.Cells(String.Format("B{0}:I{0}", (index + 2).ToString())).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#DDEBF7"))
                        worksheet.Cells(String.Format("B{0}:I{0}", (index + 1).ToString())).Style.Font.Italic = True
                        worksheet.Cells(String.Format("B{0}:I{0}", (index + 1).ToString())).Style.Font.Bold = True
                        worksheet.Cells(String.Format("B{0}:I{0}", (index + 2).ToString())).Style.Font.Bold = True
                        worksheet.Cells(cellFormat).Style.Border.Top.Style = ExcelBorderStyle.Hair
                        worksheet.Cells(cellFormat).Style.Border.Bottom.Style = ExcelBorderStyle.Thin
                    Else
                        Dim cellFormat As String = String.Format("B{0}:K{0}", (index + 2).ToString())
                        worksheet.Cells(String.Format("B{0}:K{0}", (index + 1).ToString())).Style.Fill.PatternType = ExcelFillStyle.Solid
                        worksheet.Cells(String.Format("B{0}:K{0}", (index + 1).ToString())).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#FFF2CC"))
                        worksheet.Cells(String.Format("B{0}:K{0}", (index + 2).ToString())).Style.Fill.PatternType = ExcelFillStyle.Solid
                        worksheet.Cells(String.Format("B{0}:K{0}", (index + 2).ToString())).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#DDEBF7"))
                        worksheet.Cells(String.Format("B{0}:K{0}", (index + 1).ToString())).Style.Font.Italic = True
                        worksheet.Cells(String.Format("B{0}:K{0}", (index + 1).ToString())).Style.Font.Bold = True
                        worksheet.Cells(String.Format("B{0}:K{0}", (index + 2).ToString())).Style.Font.Bold = True
                        worksheet.Cells(cellFormat).Style.Border.Top.Style = ExcelBorderStyle.Hair
                        worksheet.Cells(cellFormat).Style.Border.Bottom.Style = ExcelBorderStyle.Thin
                    End If

                    'end process data for exel
                    'End operation with file exel
                    worksheet.Column(8).AutoFit()
                    worksheet.Column(9).AutoFit()
                    worksheet.DeleteRow(14)
                    excelPackage.SaveAs(outPutFile)
                End Using

                'Open File
                Dim ps As New ProcessStartInfo


                Dim D_result As DialogResult = MessageBox.Show("Do you want to open excel file ?", Application.ProductName, MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                If D_result = DialogResult.Yes Then
                    ps.UseShellExecute = True
                    ps.FileName = urlOut
                    Process.Start(ps)

                End If

                ''Open File
                'Dim ps As New ProcessStartInfo
                'ps.UseShellExecute = True
                'ps.FileName = urlOut
                'Process.Start(ps)

                f.Close()
                btnExportExcel.Enabled = True
                Cursor = Cursors.Default

            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            tran.Rollback()
        End Try
    End Sub
    Private Sub CreateExcelSheetALL(ByRef worksheet As ExcelWorksheet, ByVal tbCompany As DataTable, ByVal f As frmShow, ByVal account As String)
        Try


            'end file exel with control XlsReport

            worksheet.Cells("B7").Value = "Từ ngày:  " + cboFromDate.Text + " Đến ngày:     " + cboToDate.Text
            worksheet.Cells("B8").Value = "Tài khoản:    " + account


            'process data for exel 
            Dim index As Integer = 14

            Dim row As DataRow
            Dim datetime As String


            'Show Progress Bar
            Dim i As Integer
            Dim PSNo, PSCo, PSNo_NT, PSCo_NT As Double ' So phat sinh No va Co trong ky
            PSNo = 0
            PSCo = 0
            PSNo_NT = 0
            PSCo_NT = 0
            Dim SDDK, SDDK_NT As Double  ' So du dau ky
            SDDK = 0
            SDDK_NT = 0
            Dim TotalDr, TotalCr, TotalDr_NT, TotalCr_NT As Double  ' Tong phat sinh bao gom ca so du dau ky dung de tinh gia tri ton cuoi ky No hay Co
            TotalDr = 0  ' Tong Phat Sinh No bao gom so du dau ky
            TotalCr = 0 ' Tong Phat sinh Co bao gom so du dau ky
            TotalDr_NT = 0
            TotalCr_NT = 0
            ' For Each row In tbResult.Rows
            For Each row In tbResult.Select("Acct = '" & account & "' AND Module <>''", "TranDate ASC")

                'PRocess Bar

                If f.UltraProgressBar1.Value < f.UltraProgressBar1.Maximum Then
                    f.UltraProgressBar1.Value += 1
                    f.UltraProgressBar1.Refresh()
                    f.Activate()
                End If

                index = index + 1
                worksheet.Cells(14, 1, 14, 11).Copy(worksheet.Cells(index, 1, index, 11))

                worksheet.Cells("B" + index.ToString()).Value = row("TranDate") ' Ngay chung Tu
                worksheet.Cells("C" + index.ToString()).Value = row("RefNbr") ' So Chung Tu
                worksheet.Cells("D" + index.ToString()).Value = row("Module") ' Phan He
                If DBConnection.Language = "KR" Then
                    worksheet.Cells("E" + index.ToString()).Value = row("DescrEN") ' Dien Giai tiếng anh
                ElseIf DBConnection.Language = "EN" Then
                    worksheet.Cells("E" + index.ToString()).Value = row("DescrEN") ' Dien Giai tiếng anh
                Else
                    worksheet.Cells("E" + index.ToString()).Value = row("TranDescr") ' Dien Giai
                End If
                'Xls.Cells("E" + index.ToString()).Value = row("TranDescr") ' Dien Giai
                worksheet.Cells("F" + index.ToString()).Value = row("ID") ' So Hoa Don
                worksheet.Cells("G" + index.ToString()).Value = row("AcctRef") ' Tai Khoan Doi Ung
                worksheet.Cells("H" + index.ToString()).Value = row("DrAmt") ' So Tien Phat Sinh No
                worksheet.Cells("I" + index.ToString()).Value = row("CrAmt") ' So Tien Phat Sinh Co
                worksheet.Cells("J" + index.ToString()).Value = row("CuryDrAmt") ' So Tien Phat Sinh No
                worksheet.Cells("K" + index.ToString()).Value = row("CuryCrAmt") ' So Tien Phat Sinh Co


                PSNo += row("DrAmt")
                PSCo += row("CrAmt")
                PSNo_NT += row("CuryDrAmt")
                PSCo_NT += row("CuryCrAmt")

            Next

            '' Xu ly so du dau ky



            'For Each row In tbResult.Select("Module =''")
            '    SDDK += row("BegAmt")
            'Next
            '' thanhln sua lai ngay 11/07/2017

            For Each row In tbResult.Select("Acct = '" & account & "' AND AcctRef =''")
                SDDK += row("BegAmt")
                SDDK_NT += row("CuryBegAmt")
            Next
            If SDDK >= 0 Then
                worksheet.Cells("H13").Value = SDDK
                worksheet.Cells("I13").Value = 0
                TotalDr = PSNo + SDDK
                TotalCr = PSCo

            Else
                worksheet.Cells("H13").Value = 0
                worksheet.Cells("I13").Value = (-1) * SDDK
                TotalDr = PSNo
                TotalCr = PSCo + (-1) * SDDK


            End If

            If SDDK_NT >= 0 Then
                worksheet.Cells("J13").Value = SDDK_NT
                worksheet.Cells("K13").Value = 0
                TotalDr_NT = PSNo_NT + SDDK_NT
                TotalCr_NT = PSCo_NT

            Else
                worksheet.Cells("J13").Value = 0
                worksheet.Cells("K13").Value = (-1) * SDDK_NT
                TotalDr_NT = PSNo_NT
                TotalCr_NT = PSCo_NT + (-1) * SDDK_NT


            End If

            worksheet.Cells("H" + (index + 1).ToString()).Value = PSNo
            worksheet.Cells("I" + (index + 1).ToString()).Value = PSCo
            worksheet.Cells("E" + (index + 1).ToString()).Value = "Cộng Phát Sinh | Total"
            worksheet.Cells("J" + (index + 1).ToString()).Value = PSNo_NT
            worksheet.Cells("K" + (index + 1).ToString()).Value = PSCo_NT

            '' end so du dau ky

            '' Xu ly so du cuoi ky

            If TotalDr > TotalCr Then
                worksheet.Cells("H" + (index + 2).ToString()).Value = TotalDr - TotalCr
                worksheet.Cells("I" + (index + 2).ToString()).Value = 0

            Else
                worksheet.Cells("H" + (index + 2).ToString()).Value = 0
                worksheet.Cells("I" + (index + 2).ToString()).Value = TotalCr - TotalDr

            End If

            If TotalDr_NT > TotalCr_NT Then
                worksheet.Cells("J" + (index + 2).ToString()).Value = TotalDr_NT - TotalCr_NT
                worksheet.Cells("K" + (index + 2).ToString()).Value = 0

            Else
                worksheet.Cells("J" + (index + 2).ToString()).Value = 0
                worksheet.Cells("K" + (index + 2).ToString()).Value = TotalCr_NT - TotalDr_NT

            End If

            '' End xu ly so du cuoi ky

            'Next
            'Xls.RowDelete(5)

            worksheet.Cells("B10").Formula = String.Format("=HYPERLINK(""#'DOCSMAP'!A1"",""DOCSMAP"")")
            worksheet.Cells("B2").Value = tbCompany.Rows(0).Item("CpNyName")
            worksheet.Cells("B3").Value = tbCompany.Rows(0).Item("Address")
            Dim cellFormat As String = String.Format("B{0}:K{0}", (index + 2).ToString())
            worksheet.Cells(String.Format("B{0}:K{0}", (index + 1).ToString())).Style.Fill.PatternType = ExcelFillStyle.Solid
            worksheet.Cells(String.Format("B{0}:K{0}", (index + 1).ToString())).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#FFF2CC"))
            worksheet.Cells(String.Format("B{0}:K{0}", (index + 2).ToString())).Style.Fill.PatternType = ExcelFillStyle.Solid
            worksheet.Cells(String.Format("B{0}:K{0}", (index + 2).ToString())).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#DDEBF7"))
            worksheet.Cells(String.Format("B{0}:K{0}", (index + 1).ToString())).Style.Font.Italic = True
            worksheet.Cells(String.Format("B{0}:K{0}", (index + 1).ToString())).Style.Font.Bold = True
            worksheet.Cells(String.Format("B{0}:K{0}", (index + 2).ToString())).Style.Font.Bold = True
            worksheet.Cells(cellFormat).Style.Border.Top.Style = ExcelBorderStyle.Hair
            worksheet.Cells(cellFormat).Style.Border.Bottom.Style = ExcelBorderStyle.Thin
            worksheet.Column(8).AutoFit()
            worksheet.Column(9).AutoFit()
            worksheet.DeleteRow(14)



        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub
    Private Sub CreateExcelSheetFC(ByRef worksheet As ExcelWorksheet, ByVal tbCompany As DataTable, ByVal f As frmShow, ByVal account As String)

        'end file exel with control XlsReport

        worksheet.Cells("B7").Value = "Từ ngày:  " + cboFromDate.Text + " Đến ngày:     " + cboToDate.Text
        worksheet.Cells("B8").Value = "Tài khoản:    " + account


        'process data for exel 
        Dim index As Integer = 14

        Dim row As DataRow
        Dim datetime As String


        'Show Progress Bar
        Dim i As Integer
        Dim PSNo, PSCo As Double ' So phat sinh No va Co trong ky
        PSNo = 0
        PSCo = 0
        Dim SDDK As Double ' So du dau ky
        SDDK = 0
        Dim TotalDr, TotalCr As Double  ' Tong phat sinh bao gom ca so du dau ky dung de tinh gia tri ton cuoi ky No hay Co
        TotalDr = 0  ' Tong Phat Sinh No bao gom so du dau ky
        TotalCr = 0 ' Tong Phat sinh Co bao gom so du dau ky

        ' For Each row In tbResult.Rows
        For Each row In tbResult.Select("Acct = '" & account & "' AND Module <>''", "TranDate ASC")

            'PRocess Bar
            If f.UltraProgressBar1.Value < f.UltraProgressBar1.Maximum Then
                f.UltraProgressBar1.Value += 1
                f.UltraProgressBar1.Refresh()
                f.Activate()
            End If

            index = index + 1
            worksheet.Cells(14, 1, 14, 11).Copy(worksheet.Cells(index, 1, index, 11))

            worksheet.Cells("B" + index.ToString()).Value = row("TranDate") ' Ngay chung Tu
            worksheet.Cells("C" + index.ToString()).Value = row("RefNbr") ' So Chung Tu
            worksheet.Cells("D" + index.ToString()).Value = row("Module") ' Phan He
            If DBConnection.Language = "KR" Then
                worksheet.Cells("E" + index.ToString()).Value = row("DescrEN") ' Dien Giai tiếng anh
            ElseIf DBConnection.Language = "EN" Then
                worksheet.Cells("E" + index.ToString()).Value = row("DescrEN") ' Dien Giai tiếng anh
            Else
                worksheet.Cells("E" + index.ToString()).Value = row("TranDescr") ' Dien Giai
            End If
            'Xls.Cells("E" + index.ToString()).Value = row("TranDescr") ' Dien Giai
            worksheet.Cells("F" + index.ToString()).Value = row("ID") ' So Hoa Don
            worksheet.Cells("G" + index.ToString()).Value = row("AcctRef") ' Tai Khoan Doi Ung
            worksheet.Cells("H" + index.ToString()).Value = row("CuryDrAmt") ' So Tien Phat Sinh No
            worksheet.Cells("I" + index.ToString()).Value = row("CuryCrAmt") ' So Tien Phat Sinh Co


            PSNo += row("CuryDrAmt")
            PSCo += row("CuryCrAmt")

        Next

        '' Xu ly so du dau ky



        'For Each row In tbResult.Select("Module =''")
        '    SDDK += row("BegAmt")
        'Next
        '' thanhln sua lai ngay 11/07/2017

        For Each row In tbResult.Select("Acct = '" & account & "' AND AcctRef =''")
            SDDK += row("CuryBegAmt")
        Next
        If SDDK >= 0 Then
            worksheet.Cells("H13").Value = SDDK
            worksheet.Cells("I13").Value = 0
            TotalDr = PSNo + SDDK
            TotalCr = PSCo

        Else
            worksheet.Cells("H13").Value = 0
            worksheet.Cells("I13").Value = (-1) * SDDK
            TotalDr = PSNo
            TotalCr = PSCo + (-1) * SDDK


        End If

        worksheet.Cells("H" + (index + 1).ToString()).Value = PSNo
        worksheet.Cells("I" + (index + 1).ToString()).Value = PSCo
        worksheet.Cells("E" + (index + 1).ToString()).Value = "Cộng Phát Sinh"

        '' end so du dau ky

        '' Xu ly so du cuoi ky

        If TotalDr > TotalCr Then
            worksheet.Cells("H" + (index + 2).ToString()).Value = TotalDr - TotalCr
            worksheet.Cells("I" + (index + 2).ToString()).Value = 0

        Else
            worksheet.Cells("H" + (index + 2).ToString()).Value = 0
            worksheet.Cells("I" + (index + 2).ToString()).Value = TotalCr - TotalDr

        End If

        '' End xu ly so du cuoi ky

        'Next
        'Xls.RowDelete(5)

        worksheet.Cells("B10").Formula = String.Format("=HYPERLINK(""#'DOCSMAP'!A1"",""DOCSMAP"")")

        worksheet.Cells("B2").Value = tbCompany.Rows(0).Item("CpNyName")
        worksheet.Cells("B3").Value = tbCompany.Rows(0).Item("Address")
        Dim cellFormat As String = String.Format("B{0}:I{0}", (index + 2).ToString())
        worksheet.Cells(String.Format("B{0}:K{0}", (index + 1).ToString())).Style.Fill.PatternType = ExcelFillStyle.Solid
        worksheet.Cells(String.Format("B{0}:K{0}", (index + 1).ToString())).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#FFF2CC"))
        worksheet.Cells(String.Format("B{0}:K{0}", (index + 2).ToString())).Style.Fill.PatternType = ExcelFillStyle.Solid
        worksheet.Cells(String.Format("B{0}:K{0}", (index + 2).ToString())).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#DDEBF7"))
        worksheet.Cells(String.Format("B{0}:K{0}", (index + 1).ToString())).Style.Font.Italic = True
        worksheet.Cells(String.Format("B{0}:K{0}", (index + 1).ToString())).Style.Font.Bold = True
        worksheet.Cells(String.Format("B{0}:K{0}", (index + 2).ToString())).Style.Font.Bold = True
        worksheet.Cells(cellFormat).Style.Border.Top.Style = ExcelBorderStyle.Hair
        worksheet.Cells(cellFormat).Style.Border.Bottom.Style = ExcelBorderStyle.Thin
        worksheet.Column(8).AutoFit()
        worksheet.Column(9).AutoFit()
        worksheet.DeleteRow(14)
    End Sub

    Private Sub ExportReportfull()

        Dim tran As System.Data.SqlClient.SqlTransaction = DBConnection.Connection.BeginTransaction

        Try

            gltran.GLPostedBegBal(tran, StringToDate(cboFromDate.Text), StringToDate(cboToDate.Text))
            gltran.GLPosted(tran, StringToDate(cboFromDate.Text), StringToDate(cboToDate.Text))
            gltran.GLPostedAcc(tran, StringToDate(cboFromDate.Text), StringToDate(cboToDate.Text))
            gltran.GLTrialBalance(tran, StringToDate(cboFromDate.Text), StringToDate(cboToDate.Text))
            gltran.GLBalanceSheet(tran, StringToDate(cboFromDate.Text), StringToDate(cboToDate.Text))
            gltran.GLTrialBalanceFC1(tran, StringToDate(cboFromDate.Text), StringToDate(cboToDate.Text), cboCuryID.Value)
            tran.Commit()
            gltran.IncomeBegin(StringToDate(cboFromDate.Text))
            gltran.IncomeCurrent(StringToDate(cboFromDate.Text), StringToDate(cboToDate.Text))



            Dim tbCompany As Data.DataTable
            tbCompany = ExportToExcel.GetCompanyInformation()

            Dim tbAccountList As Data.DataTable
            tbAccountList = ExportToExcel.GetAccountList()
            tbResult = ExportToExcel.ExcuteGLDetailReporALL(StringToDate(cboFromDate.Text), StringToDate(cboToDate.Text), txtAcct.Text, cboCuryID.Value)
            Dim dtExport, dtDK As Data.DataTable
            Dim _row As DataRow
            Dim dtrow As DataRowView
            If tbResult.Rows.Count = 0 Then
                MsgBox("Nothing to Export")
                Exit Sub
            End If

            Dim fileChooser As SaveFileDialog = New SaveFileDialog
            fileChooser.Filter = "Excel File|*.xlsx"
            fileChooser.FileName = "rptGLDetail_Docsmap.xlsx"
            Dim result As DialogResult = fileChooser.ShowDialog()
            fileChooser.CheckFileExists = False

            If result = DialogResult.OK Then
                If ExportToExcel.ExcelOpen(fileChooser.FileName.ToString) Then
                    MessageBox.Show("The file is open, Please close the file or rename a new file ", "Warning", MessageBoxButtons.OK)
                    Exit Try
                    Exit Sub
                End If
                Cursor = Cursors.WaitCursor
                btnExportExcel.Enabled = False
                Dim f As New frmShow
                f.Show()
                f.UltraProgressBar1.Maximum = tbResult.Rows.Count

                'Lay DanhSachTaiKhoan
                Dim tbAccount As New ArrayList
                Dim tbAcctName As New ArrayList

                For Each dr As DataRow In tbResult.Rows
                    If (tbAccount.Contains(dr("Acct")) = False) Then
                        tbAccount.Add(dr("Acct"))
                    End If
                Next

                tbAccount.Sort()
                'Truong: xuất tên TK theo ngôn ngữ
                Dim row_a As DataRow
                If DBConnection.Language = "KR" Then
                    For k As Integer = 0 To tbAccount.Count - 1
                        For Each row_a In tbAccountList.Select("Acct = '" & tbAccount(k) & "'")
                            tbAcctName.Add(row_a("DescrKR"))
                        Next
                    Next
                ElseIf DBConnection.Language = "EN" Then
                    For k As Integer = 0 To tbAccount.Count - 1
                        For Each row_a In tbAccountList.Select("Acct = '" & tbAccount(k) & "'")
                            tbAcctName.Add(row_a("DescrEN"))
                        Next
                    Next
                Else
                    For k As Integer = 0 To tbAccount.Count - 1
                        For Each row_a In tbAccountList.Select("Acct = '" & tbAccount(k) & "'")
                            tbAcctName.Add(row_a("DescrVN"))
                        Next
                    Next
                End If


                'tbAcctName.Sort()

                Dim AccountStr As String = ""
                Dim AcctNameStr As String = ""

                Dim index As Integer = 14
                Dim index_d As Integer = 2
                'open file exel with control XlsReport

                If DBConnection.Language = "KR" Then
                    urlTemplate = System.Windows.Forms.Application.StartupPath() + "\GLReports\RptGLDetail_EN_Full.xlsx"
                ElseIf DBConnection.Language = "EN" Then
                    urlTemplate = System.Windows.Forms.Application.StartupPath() + "\GLReports\RptGLDetail_EN_Full.xlsx"
                Else
                    urlTemplate = System.Windows.Forms.Application.StartupPath() + "\GLReports\RptGLDetail_Full.xlsx"
                End If

                Dim urlOut As String = fileChooser.FileName()
                Dim outputFile As FileInfo = New FileInfo(fileChooser.FileName)

                Using excelPackage As ExcelPackage = New ExcelPackage(New FileInfo(urlTemplate))

                    ''' Bat dau export data 

                    '1 - Thanhln add Trail Balance
                    Dim worksheet As ExcelWorksheet = excelPackage.Workbook.Worksheets(1)

                    Dim tbResultTB As DataTable
                    If cboCuryID.Value = "VND" Then
                        tbResultTB = gltran.TrialBalanceVND(StringToDate(cboFromDate.Text), StringToDate(cboToDate.Text))
                    Else
                        tbResultTB = gltran.TrialBalanceUSD(StringToDate(cboFromDate.Text), StringToDate(cboToDate.Text), cboCuryID.Value)
                    End If


                    Dim index_TB As Integer = 7

                    worksheet.Cells("E2").Value = "Từ ngày (From): " + cboFromDate.Text + " Đến ngày (To): " + cboToDate.Text

                    Dim row_TB As DataRow
                    For Each row_TB In tbResultTB.Rows

                        worksheet.Cells(6, 1, 6, 11).Copy(worksheet.Cells(index_TB, 1, index_TB, 11))
                        worksheet.Cells("B" + index_TB.ToString()).Value = row_TB("acctclass")

                        If DBConnection.Language = "KR" Then
                            worksheet.Cells("C" + index_TB.ToString()).Value = row_TB("acctnameKR")
                        ElseIf DBConnection.Language = "EN" Then
                            worksheet.Cells("C" + index_TB.ToString()).Value = row_TB("acctnameEN")
                        Else
                            worksheet.Cells("C" + index_TB.ToString()).Value = row_TB("acctname")
                        End If


                        If cboCuryID.Value = "VND" Then
                            worksheet.Cells("D" + index_TB.ToString()).Value = row_TB("DrBegAmt") 'Đầu kỳ Nợ
                            worksheet.Cells("E" + index_TB.ToString()).Value = row_TB("CrBegAmt") 'Đầu kỳ Có
                            worksheet.Cells("F" + index_TB.ToString()).Value = row_TB("DrAmt") 'Phát Sinh Nợ
                            worksheet.Cells("G" + index_TB.ToString()).Value = row_TB("CrAmt") 'Phát Sinh Có
                            worksheet.Cells("H" + index_TB.ToString()).Value = row_TB("DrAccAmt") ' Luỹ Kế Nợ
                            worksheet.Cells("I" + index_TB.ToString()).Value = row_TB("CrAccAmt") ' Lỹ Kế Có
                            worksheet.Cells("J" + index_TB.ToString()).Value = row_TB("DrEndAmt") 'Cuối kỳ Nợ
                            worksheet.Cells("K" + index_TB.ToString()).Value = row_TB("CrEndAmt") 'Cuối kỳ Có
                        Else
                            worksheet.Cells("D" + index_TB.ToString()).Value = row_TB("CuryDrBegAmt") 'Đầu kỳ Nợ
                            worksheet.Cells("E" + index_TB.ToString()).Value = row_TB("CuryCrBegAmt") 'Đầu kỳ Có
                            worksheet.Cells("F" + index_TB.ToString()).Value = row_TB("CuryDrAmt") 'Phát Sinh Nợ
                            worksheet.Cells("G" + index_TB.ToString()).Value = row_TB("CuryCrAmt") 'Phát Sinh Có
                            worksheet.Cells("H" + index_TB.ToString()).Value = row_TB("CuryDrAccAmt") ' Luỹ Kế Nợ
                            worksheet.Cells("I" + index_TB.ToString()).Value = row_TB("CuryCrAccAmt") ' Lỹ Kế Có
                            worksheet.Cells("J" + index_TB.ToString()).Value = row_TB("CuryDrEndAmt") 'Cuối kỳ Nợ
                            worksheet.Cells("K" + index_TB.ToString()).Value = row_TB("CuryCrEndAmt") 'Cuối kỳ Có
                        End If

                        If row_TB("Amt") = "GA" Or row_TB("Amt") = "SUM" Then
                            Dim rowformat As String = String.Format("B{0}:K{0}", (index_TB).ToString())
                            worksheet.Cells(rowformat).Style.Font.Bold = True
                        End If
                        index_TB = index_TB + 1

                    Next

                    worksheet.DeleteRow(6)
                    Dim cellFormatTB As String = String.Format("B{0}:K{0}", (index_TB - 2).ToString())
                    worksheet.Cells(cellFormatTB).Style.Border.Bottom.Style = ExcelBorderStyle.Thin
                    worksheet.Cells(cellFormatTB).Style.Fill.PatternType = ExcelFillStyle.Solid
                    worksheet.Cells(cellFormatTB).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#FFF2CC"))

                    ' End Thanhln Add trail Balance

                    ' 2 - Thanhln add BS Report '

                    worksheet = excelPackage.Workbook.Worksheets(2)

                    Dim tbResultBS As DataTable
                    tbResultBS = gltran.BlanceSheetVND()
                    Dim index_bs As Integer = 7
                    worksheet.Cells("B2").Value = "Tại ngày:  " + cboToDate.Text

                    Dim count As Integer = 1
                    Dim row_BS As DataRow
                    Dim datetime As String
                    For Each row_BS In tbResultBS.Rows

                        worksheet.Cells(6, 1, 6, 6).Copy(worksheet.Cells(index_bs, 1, index_bs, 6))

                        If DBConnection.Language = "KR" Then
                            worksheet.Cells("B" + index_bs.ToString()).Value = row_BS("DescrKR") 'Chỉ tiêu
                        ElseIf DBConnection.Language = "EN" Then
                            worksheet.Cells("B" + index_bs.ToString()).Value = row_BS("DescrEN") 'Chỉ tiêu
                        Else
                            worksheet.Cells("B" + index_bs.ToString()).Value = row_BS("DescrVN") 'Chỉ tiêu
                        End If

                        worksheet.Cells("C" + index_bs.ToString()).Value = row_BS("Code") 'Mã số
                        worksheet.Cells("D" + index_bs.ToString()).Value = row_BS("EndAmt") ' Cuối kỳ
                        worksheet.Cells("E" + index_bs.ToString()).Value = row_BS("BegAmt") ' Đầu kỳ

                        If row_BS("Total") = "1" Then
                            Dim rowformat As String = String.Format("B{0}:E{0}", (index_bs).ToString())
                            worksheet.Cells(rowformat).Style.Font.Bold = True
                        End If
                        count = count + 1
                        index_bs = index_bs + 1

                    Next
                    If (Roundss(tbResultBS.Rows(62).Item("BegAmt"), 0) <> Roundss(tbResultBS.Rows(112).Item("BegAmt"), 0)) Or (Roundss(tbResultBS.Rows(62).Item("EndAmt"), 0) <> Roundss(tbResultBS.Rows(112).Item("EndAmt"), 0)) Then
                        MessageBox.Show("BS NOT Balanced!", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    End If
                    worksheet.DeleteRow(6)
                    Dim cellFormatBS As String = String.Format("B{0}:E{0}", (index_bs - 2).ToString())
                    worksheet.Cells(cellFormatBS).Style.Border.Bottom.Style = ExcelBorderStyle.Thin
                    worksheet.Cells(cellFormatBS).Style.Fill.PatternType = ExcelFillStyle.Solid
                    worksheet.Cells(cellFormatBS).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#FFF2CC"))

                    ' End Thanhln Add BS Report '

                    '3 - Thanhln add P&L
                    worksheet = excelPackage.Workbook.Worksheets(3)

                    Dim tbResultPL As DataTable
                    tbResultPL = gltran.IncomeStatement(StringToDate(cboFromDate.Text), StringToDate(cboToDate.Text))
                    Dim indexPL As Integer = 7
                    worksheet.Cells("B2").Value = "Từ ngày:  " + cboFromDate.Text + " Đến ngày: " + cboToDate.Text
                    Dim row_PL As DataRow
                    For Each row_PL In tbResultPL.Rows

                        worksheet.Cells(6, 1, 6, 6).Copy(worksheet.Cells(indexPL, 1, indexPL, 6))
                        If DBConnection.Language = "KR" Then
                            worksheet.Cells("B" + indexPL.ToString()).Value = row_PL("DescrEN") 'Chỉ tiêu
                        ElseIf DBConnection.Language = "EN" Then
                            worksheet.Cells("B" + indexPL.ToString()).Value = row_PL("DescrEN") 'Chỉ tiêu
                        Else
                            worksheet.Cells("B" + indexPL.ToString()).Value = row_PL("DescrVN") 'Chỉ tiêu
                        End If

                        worksheet.Cells("C" + indexPL.ToString()).Value = row_PL("Code") 'Mã số

                        worksheet.Cells("D" + indexPL.ToString()).Value = row_PL("Comment") 'Thuyet Minh

                        worksheet.Cells("E" + indexPL.ToString()).Value = row_PL("Amount") ' Cuối kỳ

                        worksheet.Cells("F" + indexPL.ToString()).Value = row_PL("BegAmt") ' Đầu kỳ
                        indexPL = indexPL + 1


                    Next
                    worksheet.DeleteRow(6)
                    Dim cellFormatPL As String = String.Format("B{0}:F{0}", (indexPL - 2).ToString())
                    worksheet.Cells(cellFormatPL).Style.Border.Bottom.Style = ExcelBorderStyle.Thin

                    ' End Thanhln Add P&L

                    Dim row_d As DataRow
                    worksheet = excelPackage.Workbook.Worksheets(4)
                    For j As Integer = 0 To tbAccount.Count - 1
                        AccountStr = tbAccount(j)
                        AcctNameStr = tbAcctName(j)
                        worksheet.Cells(2, 1, 2, 2).Copy(worksheet.Cells(index_d, 1, index_d, 2))
                        worksheet.Cells("A" + index_d.ToString()).Value = AccountStr
                        worksheet.Cells("B" + index_d.ToString()).Value = AcctNameStr
                        index_d = index_d + 1
                    Next
                    Dim beginIndex As Integer = 2

                    For i As Integer = 0 To tbAccount.Count - 1
                        worksheet.Cells(i + beginIndex, 1).Formula = String.Format("=HYPERLINK(""#'{0}'!A1"",""{0}"")", tbAccount(i))
                        'Dim a = "=HYPERLINK(""#" & tbAccount(i) & "!A1"",""" & tbAccount(i) & """)"
                    Next


                    For i As Integer = 0 To tbAccount.Count - 1
                        AccountStr = tbAccount(i)
                        'Sheet account
                        worksheet = excelPackage.Workbook.Worksheets.Copy("template", AccountStr)
                        CreateExcelSheet(worksheet, tbCompany, f, AccountStr)
                    Next

                    excelPackage.Workbook.Worksheets.Delete("template")
                    excelPackage.SaveAs(outputFile)
                End Using

                'UpdateSheetLink(urlOut, tbAccount)
                Dim ps As New ProcessStartInfo


                Dim D_result As DialogResult = MessageBox.Show("Do you want to open excel file ?", Application.ProductName, MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                If D_result = DialogResult.Yes Then
                    ps.UseShellExecute = True
                    ps.FileName = urlOut
                    Process.Start(ps)

                End If

                ''Open File
                'Dim ps As New ProcessStartInfo
                'ps.UseShellExecute = True
                'ps.FileName = urlOut
                'Process.Start(ps)

                f.Close()
                btnExportExcel.Enabled = True
                Cursor = Cursors.Default

            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            tran.Rollback()
        End Try
    End Sub

    Private Sub ExportReport_FinancialStatement()

        Dim tran As System.Data.SqlClient.SqlTransaction = DBConnection.Connection.BeginTransaction

        Try

            gltran.GLPostedBegBal(tran, StringToDate(cboFromDate.Text), StringToDate(cboToDate.Text))
            gltran.GLPosted(tran, StringToDate(cboFromDate.Text), StringToDate(cboToDate.Text))
            gltran.GLPostedAcc(tran, StringToDate(cboFromDate.Text), StringToDate(cboToDate.Text))
            gltran.GLTrialBalance(tran, StringToDate(cboFromDate.Text), StringToDate(cboToDate.Text))
            gltran.GLBalanceSheet(tran, StringToDate(cboFromDate.Text), StringToDate(cboToDate.Text))
            gltran.GLTrialBalanceFC1(tran, StringToDate(cboFromDate.Text), StringToDate(cboToDate.Text), cboCuryID.Value)
            tran.Commit()
            gltran.IncomeBegin(StringToDate(cboFromDate.Text))
            gltran.IncomeCurrent(StringToDate(cboFromDate.Text), StringToDate(cboToDate.Text))



            Dim tbCompany As Data.DataTable
            tbCompany = ExportToExcel.GetCompanyInformation()

            Dim tbAccountList As Data.DataTable
            tbAccountList = ExportToExcel.GetAccountList()
            tbResult = ExportToExcel.ExcuteGLDetailReporALL(StringToDate(cboFromDate.Text), StringToDate(cboToDate.Text), txtAcct.Text, cboCuryID.Value)
            Dim dtExport, dtDK As Data.DataTable
            Dim _row As DataRow
            Dim dtrow As DataRowView
            If tbResult.Rows.Count = 0 Then
                MsgBox("Nothing to Export")
                Exit Sub
            End If

            Dim fileChooser As SaveFileDialog = New SaveFileDialog
            fileChooser.Filter = "Excel File|*.xlsx"
            fileChooser.FileName = "rptGLDetail_FinancialStatement.xlsx"
            Dim result As DialogResult = fileChooser.ShowDialog()
            fileChooser.CheckFileExists = False

            If result = DialogResult.OK Then
                If ExportToExcel.ExcelOpen(fileChooser.FileName.ToString) Then
                    MessageBox.Show("The file is open, Please close the file or rename a new file ", "Warning", MessageBoxButtons.OK)
                    Exit Try
                    Exit Sub
                End If
                Cursor = Cursors.WaitCursor
                btnExportExcel.Enabled = False
                Dim f As New frmShow
                f.Show()
                f.UltraProgressBar1.Maximum = tbResult.Rows.Count

                'Lay DanhSachTaiKhoan
                Dim tbAccount As New ArrayList
                Dim tbAcctName As New ArrayList

                For Each dr As DataRow In tbResult.Rows
                    If (tbAccount.Contains(dr("Acct")) = False) Then
                        tbAccount.Add(dr("Acct"))
                    End If
                Next

                tbAccount.Sort()
                'Truong: xuất tên TK theo ngôn ngữ
                Dim row_a As DataRow
                If DBConnection.Language = "KR" Then
                    For k As Integer = 0 To tbAccount.Count - 1
                        For Each row_a In tbAccountList.Select("Acct = '" & tbAccount(k) & "'")
                            tbAcctName.Add(row_a("DescrKR"))
                        Next
                    Next
                ElseIf DBConnection.Language = "EN" Then
                    For k As Integer = 0 To tbAccount.Count - 1
                        For Each row_a In tbAccountList.Select("Acct = '" & tbAccount(k) & "'")
                            tbAcctName.Add(row_a("DescrEN"))
                        Next
                    Next
                Else
                    For k As Integer = 0 To tbAccount.Count - 1
                        For Each row_a In tbAccountList.Select("Acct = '" & tbAccount(k) & "'")
                            tbAcctName.Add(row_a("DescrVN"))
                        Next
                    Next
                End If


                'tbAcctName.Sort()

                Dim AccountStr As String = ""
                Dim AcctNameStr As String = ""

                Dim index As Integer = 14
                Dim index_d As Integer = 2
                'open file exel with control XlsReport

                If DBConnection.Language = "KR" Then
                    urlTemplate = System.Windows.Forms.Application.StartupPath() + "\GLReports\RptGLDetail_EN_Financial.xlsx"
                ElseIf DBConnection.Language = "EN" Then
                    urlTemplate = System.Windows.Forms.Application.StartupPath() + "\GLReports\RptGLDetail_EN_Financial.xlsx"
                Else
                    urlTemplate = System.Windows.Forms.Application.StartupPath() + "\GLReports\RptGLDetail_Financial.xlsx"
                End If

                Dim urlOut As String = fileChooser.FileName()
                Dim outputFile As FileInfo = New FileInfo(fileChooser.FileName)

                Using excelPackage As ExcelPackage = New ExcelPackage(New FileInfo(urlTemplate))

                    ''' Bat dau export data 

                    '1 - Thanhln add Trial Balance
                    Dim worksheet As ExcelWorksheet = excelPackage.Workbook.Worksheets(1)

                    Dim tbResultTB As DataTable
                    If cboCuryID.Value = "VND" Then
                        tbResultTB = gltran.TrialBalanceVND(StringToDate(cboFromDate.Text), StringToDate(cboToDate.Text))
                    Else
                        tbResultTB = gltran.TrialBalanceUSD(StringToDate(cboFromDate.Text), StringToDate(cboToDate.Text), cboCuryID.Value)
                    End If


                    Dim index_TB As Integer = 7

                    worksheet.Cells("E2").Value = "Từ ngày (From): " + cboFromDate.Text + " Đến ngày (To): " + cboToDate.Text

                    Dim row_TB As DataRow
                    For Each row_TB In tbResultTB.Rows

                        worksheet.Cells(6, 1, 6, 11).Copy(worksheet.Cells(index_TB, 1, index_TB, 11))
                        worksheet.Cells("B" + index_TB.ToString()).Value = row_TB("acctclass")

                        If DBConnection.Language = "KR" Then
                            worksheet.Cells("C" + index_TB.ToString()).Value = row_TB("acctnameKR")
                        ElseIf DBConnection.Language = "EN" Then
                            worksheet.Cells("C" + index_TB.ToString()).Value = row_TB("acctnameEN")
                        Else
                            worksheet.Cells("C" + index_TB.ToString()).Value = row_TB("acctname")
                        End If


                        If cboCuryID.Value = "VND" Then
                            worksheet.Cells("D" + index_TB.ToString()).Value = row_TB("DrBegAmt") 'Đầu kỳ Nợ
                            worksheet.Cells("E" + index_TB.ToString()).Value = row_TB("CrBegAmt") 'Đầu kỳ Có
                            worksheet.Cells("F" + index_TB.ToString()).Value = row_TB("DrAmt") 'Phát Sinh Nợ
                            worksheet.Cells("G" + index_TB.ToString()).Value = row_TB("CrAmt") 'Phát Sinh Có
                            worksheet.Cells("H" + index_TB.ToString()).Value = row_TB("DrAccAmt") ' Luỹ Kế Nợ
                            worksheet.Cells("I" + index_TB.ToString()).Value = row_TB("CrAccAmt") ' Lỹ Kế Có
                            worksheet.Cells("J" + index_TB.ToString()).Value = row_TB("DrEndAmt") 'Cuối kỳ Nợ
                            worksheet.Cells("K" + index_TB.ToString()).Value = row_TB("CrEndAmt") 'Cuối kỳ Có
                        Else
                            worksheet.Cells("D" + index_TB.ToString()).Value = row_TB("CuryDrBegAmt") 'Đầu kỳ Nợ
                            worksheet.Cells("E" + index_TB.ToString()).Value = row_TB("CuryCrBegAmt") 'Đầu kỳ Có
                            worksheet.Cells("F" + index_TB.ToString()).Value = row_TB("CuryDrAmt") 'Phát Sinh Nợ
                            worksheet.Cells("G" + index_TB.ToString()).Value = row_TB("CuryCrAmt") 'Phát Sinh Có
                            worksheet.Cells("H" + index_TB.ToString()).Value = row_TB("CuryDrAccAmt") ' Luỹ Kế Nợ
                            worksheet.Cells("I" + index_TB.ToString()).Value = row_TB("CuryCrAccAmt") ' Lỹ Kế Có
                            worksheet.Cells("J" + index_TB.ToString()).Value = row_TB("CuryDrEndAmt") 'Cuối kỳ Nợ
                            worksheet.Cells("K" + index_TB.ToString()).Value = row_TB("CuryCrEndAmt") 'Cuối kỳ Có
                        End If

                        If row_TB("Amt") = "GA" Or row_TB("Amt") = "SUM" Then
                            Dim rowformat As String = String.Format("B{0}:K{0}", (index_TB).ToString())
                            worksheet.Cells(rowformat).Style.Font.Bold = True
                        End If
                        index_TB = index_TB + 1

                    Next

                    worksheet.DeleteRow(6)
                    Dim cellFormatTB As String = String.Format("B{0}:K{0}", (index_TB - 2).ToString())
                    worksheet.Cells(cellFormatTB).Style.Border.Bottom.Style = ExcelBorderStyle.Thin
                    worksheet.Cells(cellFormatTB).Style.Fill.PatternType = ExcelFillStyle.Solid
                    worksheet.Cells(cellFormatTB).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#FFF2CC"))

                    ' End Thanhln Add trail Balance

                    ' 2 - Thanhln add BS Report'

                    worksheet = excelPackage.Workbook.Worksheets(2)

                    Dim tbResultBS As DataTable
                    tbResultBS = gltran.BlanceSheetVND()
                    Dim index_bs As Integer = 7
                    worksheet.Cells("B2").Value = "Tại ngày:  " + cboToDate.Text

                    Dim count As Integer = 1
                    Dim row_BS As DataRow
                    Dim datetime As String
                    For Each row_BS In tbResultBS.Rows

                        worksheet.Cells(6, 1, 6, 6).Copy(worksheet.Cells(index_bs, 1, index_bs, 6))

                        If DBConnection.Language = "KR" Then
                            worksheet.Cells("B" + index_bs.ToString()).Value = row_BS("DescrKR") 'Chỉ tiêu
                        ElseIf DBConnection.Language = "EN" Then
                            worksheet.Cells("B" + index_bs.ToString()).Value = row_BS("DescrEN") 'Chỉ tiêu
                        Else
                            worksheet.Cells("B" + index_bs.ToString()).Value = row_BS("DescrVN") 'Chỉ tiêu
                        End If

                        worksheet.Cells("C" + index_bs.ToString()).Value = row_BS("Code") 'Mã số
                        worksheet.Cells("D" + index_bs.ToString()).Value = row_BS("EndAmt") ' Cuối kỳ
                        worksheet.Cells("E" + index_bs.ToString()).Value = row_BS("BegAmt") ' Đầu kỳ

                        If row_BS("Total") = "1" Then
                            Dim rowformat As String = String.Format("B{0}:E{0}", (index_bs).ToString())
                            worksheet.Cells(rowformat).Style.Font.Bold = True
                        End If
                        count = count + 1
                        index_bs = index_bs + 1

                    Next
                    If (Roundss(tbResultBS.Rows(62).Item("BegAmt"), 0) <> Roundss(tbResultBS.Rows(112).Item("BegAmt"), 0)) Or (Roundss(tbResultBS.Rows(62).Item("EndAmt"), 0) <> Roundss(tbResultBS.Rows(112).Item("EndAmt"), 0)) Then
                        MessageBox.Show("BS NOT Balanced!", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    End If
                    worksheet.DeleteRow(6)
                    Dim cellFormatBS As String = String.Format("B{0}:E{0}", (index_bs - 2).ToString())
                    worksheet.Cells(cellFormatBS).Style.Border.Bottom.Style = ExcelBorderStyle.Thin
                    worksheet.Cells(cellFormatBS).Style.Fill.PatternType = ExcelFillStyle.Solid
                    worksheet.Cells(cellFormatBS).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#FFF2CC"))

                    ' End Thanhln Add BS Report '

                    '3 - Thanhln add P&L
                    worksheet = excelPackage.Workbook.Worksheets(3)

                    Dim tbResultPL As DataTable
                    tbResultPL = gltran.IncomeStatement(StringToDate(cboFromDate.Text), StringToDate(cboToDate.Text))
                    Dim indexPL As Integer = 7
                    worksheet.Cells("B2").Value = "Từ ngày:  " + cboFromDate.Text + " Đến ngày: " + cboToDate.Text
                    Dim row_PL As DataRow
                    For Each row_PL In tbResultPL.Rows

                        worksheet.Cells(6, 1, 6, 6).Copy(worksheet.Cells(indexPL, 1, indexPL, 6))
                        If DBConnection.Language = "KR" Then
                            worksheet.Cells("B" + indexPL.ToString()).Value = row_PL("DescrEN") 'Chỉ tiêu
                        ElseIf DBConnection.Language = "EN" Then
                            worksheet.Cells("B" + indexPL.ToString()).Value = row_PL("DescrEN") 'Chỉ tiêu
                        Else
                            worksheet.Cells("B" + indexPL.ToString()).Value = row_PL("DescrVN") 'Chỉ tiêu
                        End If

                        worksheet.Cells("C" + indexPL.ToString()).Value = row_PL("Code") 'Mã số

                        worksheet.Cells("D" + indexPL.ToString()).Value = row_PL("Comment") 'Thuyet Minh

                        worksheet.Cells("E" + indexPL.ToString()).Value = row_PL("Amount") ' Cuối kỳ

                        worksheet.Cells("F" + indexPL.ToString()).Value = row_PL("BegAmt") ' Đầu kỳ
                        indexPL = indexPL + 1


                    Next

                    worksheet.DeleteRow(6)
                    Dim cellFormatPL As String = String.Format("B{0}:F{0}", (indexPL - 2).ToString())
                    worksheet.Cells(cellFormatPL).Style.Border.Bottom.Style = ExcelBorderStyle.Thin

                    ' End Thanhln Add P&L

                    'Dim row_d As DataRow
                    'worksheet = excelPackage.Workbook.Worksheets(4)
                    'For j As Integer = 0 To tbAccount.Count - 1
                    '    AccountStr = tbAccount(j)
                    '    AcctNameStr = tbAcctName(j)
                    '    worksheet.Cells(2, 1, 2, 2).Copy(worksheet.Cells(index_d, 1, index_d, 2))
                    '    worksheet.Cells("A" + index_d.ToString()).Value = AccountStr
                    '    worksheet.Cells("B" + index_d.ToString()).Value = AcctNameStr
                    '    index_d = index_d + 1
                    'Next
                    'Dim beginIndex As Integer = 2

                    'For i As Integer = 0 To tbAccount.Count - 1
                    '    worksheet.Cells(i + beginIndex, 1).Formula = String.Format("=HYPERLINK(""#'{0}'!A1"",""{0}"")", tbAccount(i))
                    '    'Dim a = "=HYPERLINK(""#" & tbAccount(i) & "!A1"",""" & tbAccount(i) & """)"
                    'Next


                    'For i As Integer = 0 To tbAccount.Count - 1
                    '    AccountStr = tbAccount(i)
                    '    'Sheet account
                    '    worksheet = excelPackage.Workbook.Worksheets.Copy("template", AccountStr)
                    '    CreateExcelSheet(worksheet, tbCompany, f, AccountStr)
                    'Next

                    'excelPackage.Workbook.Worksheets.Delete("template")
                    excelPackage.SaveAs(outputFile)
                End Using

                'UpdateSheetLink(urlOut, tbAccount)
                Dim ps As New ProcessStartInfo


                Dim D_result As DialogResult = MessageBox.Show("Do you want to open excel file ?", Application.ProductName, MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                If D_result = DialogResult.Yes Then
                    ps.UseShellExecute = True
                    ps.FileName = urlOut
                    Process.Start(ps)

                End If

                ''Open File
                'Dim ps As New ProcessStartInfo
                'ps.UseShellExecute = True
                'ps.FileName = urlOut
                'Process.Start(ps)

                f.Close()
                btnExportExcel.Enabled = True
                Cursor = Cursors.Default

            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            tran.Rollback()
        End Try
    End Sub
    Private Sub UpdateSheetLink(ByVal filePath As String, ByVal tbAccount As ArrayList)
        Dim excel As New Excel.Application, beginIndex = 2
        excel.DisplayAlerts = False
        Dim workbook As Excel.Workbook = excel.Workbooks.Open(filePath)
        Dim worksheet As Excel.Worksheet = workbook.Worksheets(4)

        Try
            If Not worksheet Is Nothing Then
                For i As Integer = 0 To tbAccount.Count - 1
                    worksheet.Range(String.Format("A{0}", i + beginIndex)).Formula = String.Format("=HYPERLINK(""#{0}!A1"",""{0}"")", tbAccount(i))
                Next
            End If

            workbook.Save()
            workbook.Close()
            ReleaseObject(worksheet)
            ReleaseObject(workbook)
            ReleaseObject(excel)
        Catch ex As Exception
            ReleaseObject(worksheet)
            ReleaseObject(workbook)
            ReleaseObject(excel)
        End Try
    End Sub
    Private Sub ReleaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Private Sub TbnPreview_Click(sender As Object, e As EventArgs) Handles tbnPreview.Click
        Try
            '' Show bao cao theo form

        Catch ex As Exception

        End Try
    End Sub

    Private Sub cbOption_CheckedChanged(sender As Object, e As EventArgs) Handles cbOption.CheckedChanged
        Try
            If cbOption.Checked = True Then
                OtherOption.Visible = True
            Else
                OtherOption.Visible = False
            End If
        Catch ex As Exception

        End Try
    End Sub

End Class
