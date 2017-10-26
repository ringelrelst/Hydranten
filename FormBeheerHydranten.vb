Option Explicit On 
Option Strict On

#Region " Imports namespaces "
Imports System.Drawing
Imports System.Runtime.InteropServices
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Carto.esriViewDrawPhase
Imports ESRI.ArcGIS.Editor
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Geometry
#End Region

''' -----------------------------------------------------------------------------
''' Project	 : Digipolis.Hydranten.BeheerHydranten
''' Class	 : Hydranten.BeheerHydranten.FormBeheerHydranten
''' 
''' -----------------------------------------------------------------------------
''' <summary>
'''      Form for managing hydrants.
''' </summary>
''' <remarks>
''' </remarks>
''' <history>
''' 	[Kristof Vydt]	16/09/2005	Created
''' 	[Kristof Vydt]	23/09/2005	Remove event handler(s) on close.
''' 	[Kristof Vydt]	05/10/2005	Change form layout.
'''                                 Replace ButtonConnect/Copy by CheckBoxConnect/Copy.
''' 	[Kristof Vydt]	10/10/2005	Type filter added.
''' 	                        	Update annotations moved until after storing all attributes.
''' 	[Kristof Vydt]	11/10/2005	Define color settings as private variables of the form.
''' 	[Kristof Vydt]	17/10/2005	Disable/enable filter combo boxes.
''' 	[Kristof Vydt]	21/10/2005	Adjust ButtonLabelAdd_Click to use the new FormAddAnnotation.
''' 	[Kristof Vydt]	24/10/2005	BrandweerID required &amp; unique when Actief &amp; Ondergronds.
''' 	                        	Deactivate listeners when initializing editing controls.
''' 	[Kristof Vydt]	27/10/2005	Set DropDownStyle of every ComboBox to List to force the user to select from the list.
'''  	[Kristof Vydt]	23/11/2005	Add optional zoomToFeature and resetCopyFromReference parameter to LoadFeature method.
'''                                 Support storing Null value for BrandweerID.
'''  	[Kristof Vydt]	17/07/2006	Refresh active view extent in StoreAttributeChanges' finally.
'''                                 Support storing Null value for attribute Diameter.
'''                                 Make sure attribute "Aanduiding" is filled, before adding a label.
''' 	[Kristof Vydt]	02/08/2006	Replace some global constant by a function that reads the value from config file.
''' 	[Kristof Vydt]	18/08/2006	Eliminate private marker element.
'''     [Kristof Vydt]  08/09/2006  EndDate default value and MinDate = yesterday instead of today.
'''                                 Use GetLayerWorkspace from ToolsLib instead of determining the workspace here.
'''     [Kristof Vydt]  22/02/2007  Eliminate on form legend code controls.
'''     [Koen Vermeer]  23/06/2008  Conversion to ArcGIS 9.2
'''     [Elton Manoku]  24/07/2008  See: RW: Added try catch statements in events. The form is not open in modal 
'''                                 so the global event that opened the form cannot catch exceptions. Storeattributechanges
'''                                 sets the eindatum equal to begindatum if smaller than the begindatum
''' </history>
''' -----------------------------------------------------------------------------
Public NotInheritable Class FormBeheerHydranten
    Inherits System.Windows.Forms.Form
    Implements IConnectFeature 'the form is using the <ConnectFeature> functionality

#Region " Private variables "

    Private m_application As IMxApplication 'hold current ArcMap application
    Private m_document As IMxDocument 'hold current ArcMap document
    Private m_layer As ILayer 'hydranten layer
    Private m_workspace As IWorkspace 'workspace of the hydranten
    'Private m_marker As IMarkerElement 'marker for current feature
    Private m_editing As Boolean 'indicated if form is ready for editing
    Private m_selectionSet As ISelectionSet
    Private m_selectionExtent As IEnvelope = Nothing
    Private m_enumOIDs As IEnumIDs 'enumeration of the feature IDs of the edit set
    Private m_OID As Integer 'the ObjectID of the current editable feature
    Private m_copyFrom As IFeature = Nothing 'when functionality "copy attributes from hydrant" is used

    'Private predefined colour settings.
    Dim EnabledEditControlForeColor As Color = Color.Black
    Dim EnabledEditControlBackColor As Color = Color.White
    Dim DisabledEditControlForeColor As Color = Color.Gray
    Dim DisabledEditControlBackColor As Color = Color.White
    Dim InitialEditControlForeColor As Color = Color.Gray
    Dim InitialEditControlBackColor As Color = SystemColors.Control
    Dim UnchangedLabelForeColor As Color = Color.Black
    Dim ChangedLabelForeColor As Color = Color.IndianRed
    Dim ActivatedToolbuttonForeColor As Color = Color.BlueViolet
    Dim DeactivatedToolbuttonForeColor As Color = Color.Black

#End Region

#Region " Windows Form Designer generated code "

    'Public Sub New()
    '    MyBase.New()

    '    'This call is required by the Windows Form Designer.
    '    InitializeComponent()

    '    'Add any initialization after the InitializeComponent() call

    'End Sub

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
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents RadioButtonMapSelection As System.Windows.Forms.RadioButton
    Friend WithEvents ButtonLoad As System.Windows.Forms.Button
    Friend WithEvents ButtonClose As System.Windows.Forms.Button
    Friend WithEvents ComboBoxStatusFilter As System.Windows.Forms.ComboBox
    Friend WithEvents RadioButtonStatusFilter As System.Windows.Forms.RadioButton
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents ButtonFirst As System.Windows.Forms.Button
    Friend WithEvents ButtonPrevious As System.Windows.Forms.Button
    Friend WithEvents ButtonNext As System.Windows.Forms.Button
    Friend WithEvents ButtonLast As System.Windows.Forms.Button
    Friend WithEvents LabelCounter As System.Windows.Forms.Label
    Friend WithEvents LabelTotal As System.Windows.Forms.Label
    Friend WithEvents LabelSeparator As System.Windows.Forms.Label
    Friend WithEvents ButtonSave As System.Windows.Forms.Button
    Friend WithEvents ComboBoxHydrantType As System.Windows.Forms.ComboBox
    Friend WithEvents LabelHydrantType As System.Windows.Forms.Label
    Friend WithEvents ComboBoxBron As System.Windows.Forms.ComboBox
    Friend WithEvents LabelBron As System.Windows.Forms.Label
    Friend WithEvents ComboBoxStatus As System.Windows.Forms.ComboBox
    Friend WithEvents LabelStatus As System.Windows.Forms.Label
    Friend WithEvents TextBoxPostcode As System.Windows.Forms.TextBox
    Friend WithEvents LabelPostcode As System.Windows.Forms.Label
    Friend WithEvents TextBoxStraatcode As System.Windows.Forms.TextBox
    Friend WithEvents LabelStraatcode As System.Windows.Forms.Label
    Friend WithEvents TextBoxStraatnaam As System.Windows.Forms.TextBox
    Friend WithEvents LabelStraatnaam As System.Windows.Forms.Label
    Friend WithEvents DatePickerEinddatum As System.Windows.Forms.DateTimePicker
    Friend WithEvents LabelEinddatum As System.Windows.Forms.Label
    Friend WithEvents DatePickerBegindatum As System.Windows.Forms.DateTimePicker
    Friend WithEvents LabelBegindatum As System.Windows.Forms.Label
    Friend WithEvents TextBoxLeidingID As System.Windows.Forms.TextBox
    Friend WithEvents LabelLeidingID As System.Windows.Forms.Label
    Friend WithEvents ComboBoxLeidingType As System.Windows.Forms.ComboBox
    Friend WithEvents LabelLeidingType As System.Windows.Forms.Label
    Friend WithEvents TextBoxLeverancierID As System.Windows.Forms.TextBox
    Friend WithEvents LabelLeverancierID As System.Windows.Forms.Label
    Friend WithEvents TextBoxYCoord As System.Windows.Forms.TextBox
    Friend WithEvents LabelYCoord As System.Windows.Forms.Label
    Friend WithEvents TextBoxXCoord As System.Windows.Forms.TextBox
    Friend WithEvents LabelXCoord As System.Windows.Forms.Label
    Friend WithEvents TextBoxBrandweerID As System.Windows.Forms.TextBox
    Friend WithEvents LabelBrandweerID As System.Windows.Forms.Label
    Friend WithEvents ComboBoxLigging As System.Windows.Forms.ComboBox
    Friend WithEvents LabelLigging As System.Windows.Forms.Label
    Friend WithEvents TextBoxDiameter As System.Windows.Forms.TextBox
    Friend WithEvents LabelDiameter As System.Windows.Forms.Label
    Friend WithEvents TextBoxAanduiding As System.Windows.Forms.TextBox
    Friend WithEvents LabelAanduiding As System.Windows.Forms.Label
    Friend WithEvents CheckBoxEinddatum As System.Windows.Forms.CheckBox
    Friend WithEvents ButtonLabelAdd As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ButtonLabelDel As System.Windows.Forms.Button
    Friend WithEvents CheckBoxConnect As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBoxCopy As System.Windows.Forms.CheckBox
    Friend WithEvents ComboBoxTypeFilter As System.Windows.Forms.ComboBox
    Friend WithEvents RadioButtonTypeFilter As System.Windows.Forms.RadioButton
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.ComboBoxTypeFilter = New System.Windows.Forms.ComboBox
        Me.RadioButtonTypeFilter = New System.Windows.Forms.RadioButton
        Me.ButtonLoad = New System.Windows.Forms.Button
        Me.ComboBoxStatusFilter = New System.Windows.Forms.ComboBox
        Me.RadioButtonMapSelection = New System.Windows.Forms.RadioButton
        Me.RadioButtonStatusFilter = New System.Windows.Forms.RadioButton
        Me.ButtonClose = New System.Windows.Forms.Button
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.CheckBoxCopy = New System.Windows.Forms.CheckBox
        Me.CheckBoxConnect = New System.Windows.Forms.CheckBox
        Me.ButtonLabelDel = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.CheckBoxEinddatum = New System.Windows.Forms.CheckBox
        Me.ButtonLabelAdd = New System.Windows.Forms.Button
        Me.ButtonSave = New System.Windows.Forms.Button
        Me.ComboBoxHydrantType = New System.Windows.Forms.ComboBox
        Me.LabelHydrantType = New System.Windows.Forms.Label
        Me.ComboBoxBron = New System.Windows.Forms.ComboBox
        Me.LabelBron = New System.Windows.Forms.Label
        Me.ComboBoxStatus = New System.Windows.Forms.ComboBox
        Me.LabelStatus = New System.Windows.Forms.Label
        Me.TextBoxPostcode = New System.Windows.Forms.TextBox
        Me.LabelPostcode = New System.Windows.Forms.Label
        Me.TextBoxStraatcode = New System.Windows.Forms.TextBox
        Me.LabelStraatcode = New System.Windows.Forms.Label
        Me.TextBoxStraatnaam = New System.Windows.Forms.TextBox
        Me.LabelStraatnaam = New System.Windows.Forms.Label
        Me.DatePickerEinddatum = New System.Windows.Forms.DateTimePicker
        Me.LabelEinddatum = New System.Windows.Forms.Label
        Me.DatePickerBegindatum = New System.Windows.Forms.DateTimePicker
        Me.LabelBegindatum = New System.Windows.Forms.Label
        Me.TextBoxLeidingID = New System.Windows.Forms.TextBox
        Me.LabelLeidingID = New System.Windows.Forms.Label
        Me.ComboBoxLeidingType = New System.Windows.Forms.ComboBox
        Me.LabelLeidingType = New System.Windows.Forms.Label
        Me.TextBoxLeverancierID = New System.Windows.Forms.TextBox
        Me.LabelLeverancierID = New System.Windows.Forms.Label
        Me.TextBoxYCoord = New System.Windows.Forms.TextBox
        Me.LabelYCoord = New System.Windows.Forms.Label
        Me.TextBoxXCoord = New System.Windows.Forms.TextBox
        Me.LabelXCoord = New System.Windows.Forms.Label
        Me.TextBoxBrandweerID = New System.Windows.Forms.TextBox
        Me.LabelBrandweerID = New System.Windows.Forms.Label
        Me.ComboBoxLigging = New System.Windows.Forms.ComboBox
        Me.LabelLigging = New System.Windows.Forms.Label
        Me.TextBoxDiameter = New System.Windows.Forms.TextBox
        Me.LabelDiameter = New System.Windows.Forms.Label
        Me.TextBoxAanduiding = New System.Windows.Forms.TextBox
        Me.LabelAanduiding = New System.Windows.Forms.Label
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.LabelCounter = New System.Windows.Forms.Label
        Me.LabelTotal = New System.Windows.Forms.Label
        Me.LabelSeparator = New System.Windows.Forms.Label
        Me.ButtonNext = New System.Windows.Forms.Button
        Me.ButtonLast = New System.Windows.Forms.Button
        Me.ButtonFirst = New System.Windows.Forms.Button
        Me.ButtonPrevious = New System.Windows.Forms.Button
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.ComboBoxTypeFilter)
        Me.GroupBox1.Controls.Add(Me.RadioButtonTypeFilter)
        Me.GroupBox1.Controls.Add(Me.ButtonLoad)
        Me.GroupBox1.Controls.Add(Me.ComboBoxStatusFilter)
        Me.GroupBox1.Controls.Add(Me.RadioButtonMapSelection)
        Me.GroupBox1.Controls.Add(Me.RadioButtonStatusFilter)
        Me.GroupBox1.Location = New System.Drawing.Point(8, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(272, 85)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Filter"
        '
        'ComboBoxTypeFilter
        '
        Me.ComboBoxTypeFilter.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBoxTypeFilter.Location = New System.Drawing.Point(112, 32)
        Me.ComboBoxTypeFilter.Name = "ComboBoxTypeFilter"
        Me.ComboBoxTypeFilter.Size = New System.Drawing.Size(152, 21)
        Me.ComboBoxTypeFilter.TabIndex = 3
        '
        'RadioButtonTypeFilter
        '
        Me.RadioButtonTypeFilter.BackColor = System.Drawing.Color.Transparent
        Me.RadioButtonTypeFilter.Location = New System.Drawing.Point(8, 32)
        Me.RadioButtonTypeFilter.Name = "RadioButtonTypeFilter"
        Me.RadioButtonTypeFilter.Size = New System.Drawing.Size(208, 24)
        Me.RadioButtonTypeFilter.TabIndex = 2
        Me.RadioButtonTypeFilter.Text = "Filter op type:"
        Me.RadioButtonTypeFilter.UseVisualStyleBackColor = False
        '
        'ButtonLoad
        '
        Me.ButtonLoad.Location = New System.Drawing.Point(208, 56)
        Me.ButtonLoad.Name = "ButtonLoad"
        Me.ButtonLoad.Size = New System.Drawing.Size(56, 24)
        Me.ButtonLoad.TabIndex = 5
        Me.ButtonLoad.Text = "Uitlezen"
        '
        'ComboBoxStatusFilter
        '
        Me.ComboBoxStatusFilter.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBoxStatusFilter.Location = New System.Drawing.Point(112, 8)
        Me.ComboBoxStatusFilter.Name = "ComboBoxStatusFilter"
        Me.ComboBoxStatusFilter.Size = New System.Drawing.Size(152, 21)
        Me.ComboBoxStatusFilter.TabIndex = 1
        '
        'RadioButtonMapSelection
        '
        Me.RadioButtonMapSelection.BackColor = System.Drawing.Color.Transparent
        Me.RadioButtonMapSelection.Checked = True
        Me.RadioButtonMapSelection.Location = New System.Drawing.Point(8, 56)
        Me.RadioButtonMapSelection.Name = "RadioButtonMapSelection"
        Me.RadioButtonMapSelection.Size = New System.Drawing.Size(208, 16)
        Me.RadioButtonMapSelection.TabIndex = 4
        Me.RadioButtonMapSelection.TabStop = True
        Me.RadioButtonMapSelection.Text = "Selectie op kaart"
        Me.RadioButtonMapSelection.UseVisualStyleBackColor = False
        '
        'RadioButtonStatusFilter
        '
        Me.RadioButtonStatusFilter.BackColor = System.Drawing.Color.Transparent
        Me.RadioButtonStatusFilter.Location = New System.Drawing.Point(8, 16)
        Me.RadioButtonStatusFilter.Name = "RadioButtonStatusFilter"
        Me.RadioButtonStatusFilter.Size = New System.Drawing.Size(208, 16)
        Me.RadioButtonStatusFilter.TabIndex = 0
        Me.RadioButtonStatusFilter.Text = "Filter op status:"
        Me.RadioButtonStatusFilter.UseVisualStyleBackColor = False
        '
        'ButtonClose
        '
        Me.ButtonClose.Location = New System.Drawing.Point(192, 608)
        Me.ButtonClose.Name = "ButtonClose"
        Me.ButtonClose.Size = New System.Drawing.Size(88, 24)
        Me.ButtonClose.TabIndex = 3
        Me.ButtonClose.Text = "Sluiten"
        '
        'GroupBox2
        '
        Me.GroupBox2.BackColor = System.Drawing.SystemColors.Control
        Me.GroupBox2.Controls.Add(Me.CheckBoxCopy)
        Me.GroupBox2.Controls.Add(Me.CheckBoxConnect)
        Me.GroupBox2.Controls.Add(Me.ButtonLabelDel)
        Me.GroupBox2.Controls.Add(Me.Label1)
        Me.GroupBox2.Controls.Add(Me.CheckBoxEinddatum)
        Me.GroupBox2.Controls.Add(Me.ButtonLabelAdd)
        Me.GroupBox2.Controls.Add(Me.ButtonSave)
        Me.GroupBox2.Controls.Add(Me.ComboBoxHydrantType)
        Me.GroupBox2.Controls.Add(Me.LabelHydrantType)
        Me.GroupBox2.Controls.Add(Me.ComboBoxBron)
        Me.GroupBox2.Controls.Add(Me.LabelBron)
        Me.GroupBox2.Controls.Add(Me.ComboBoxStatus)
        Me.GroupBox2.Controls.Add(Me.LabelStatus)
        Me.GroupBox2.Controls.Add(Me.TextBoxPostcode)
        Me.GroupBox2.Controls.Add(Me.LabelPostcode)
        Me.GroupBox2.Controls.Add(Me.TextBoxStraatcode)
        Me.GroupBox2.Controls.Add(Me.LabelStraatcode)
        Me.GroupBox2.Controls.Add(Me.TextBoxStraatnaam)
        Me.GroupBox2.Controls.Add(Me.LabelStraatnaam)
        Me.GroupBox2.Controls.Add(Me.DatePickerEinddatum)
        Me.GroupBox2.Controls.Add(Me.LabelEinddatum)
        Me.GroupBox2.Controls.Add(Me.DatePickerBegindatum)
        Me.GroupBox2.Controls.Add(Me.LabelBegindatum)
        Me.GroupBox2.Controls.Add(Me.TextBoxLeidingID)
        Me.GroupBox2.Controls.Add(Me.LabelLeidingID)
        Me.GroupBox2.Controls.Add(Me.ComboBoxLeidingType)
        Me.GroupBox2.Controls.Add(Me.LabelLeidingType)
        Me.GroupBox2.Controls.Add(Me.TextBoxLeverancierID)
        Me.GroupBox2.Controls.Add(Me.LabelLeverancierID)
        Me.GroupBox2.Controls.Add(Me.TextBoxYCoord)
        Me.GroupBox2.Controls.Add(Me.LabelYCoord)
        Me.GroupBox2.Controls.Add(Me.TextBoxXCoord)
        Me.GroupBox2.Controls.Add(Me.LabelXCoord)
        Me.GroupBox2.Controls.Add(Me.TextBoxBrandweerID)
        Me.GroupBox2.Controls.Add(Me.LabelBrandweerID)
        Me.GroupBox2.Controls.Add(Me.ComboBoxLigging)
        Me.GroupBox2.Controls.Add(Me.LabelLigging)
        Me.GroupBox2.Controls.Add(Me.TextBoxDiameter)
        Me.GroupBox2.Controls.Add(Me.LabelDiameter)
        Me.GroupBox2.Controls.Add(Me.TextBoxAanduiding)
        Me.GroupBox2.Controls.Add(Me.LabelAanduiding)
        Me.GroupBox2.Location = New System.Drawing.Point(8, 136)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(272, 464)
        Me.GroupBox2.TabIndex = 2
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Feature"
        '
        'CheckBoxCopy
        '
        Me.CheckBoxCopy.Appearance = System.Windows.Forms.Appearance.Button
        Me.CheckBoxCopy.Enabled = False
        Me.CheckBoxCopy.Location = New System.Drawing.Point(139, 400)
        Me.CheckBoxCopy.Name = "CheckBoxCopy"
        Me.CheckBoxCopy.Size = New System.Drawing.Size(125, 24)
        Me.CheckBoxCopy.TabIndex = 21
        Me.CheckBoxCopy.Text = "Attributen overnemen"
        Me.CheckBoxCopy.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'CheckBoxConnect
        '
        Me.CheckBoxConnect.Appearance = System.Windows.Forms.Appearance.Button
        Me.CheckBoxConnect.Enabled = False
        Me.CheckBoxConnect.Location = New System.Drawing.Point(8, 400)
        Me.CheckBoxConnect.Name = "CheckBoxConnect"
        Me.CheckBoxConnect.Size = New System.Drawing.Size(125, 24)
        Me.CheckBoxConnect.TabIndex = 20
        Me.CheckBoxConnect.Text = "Connecteren"
        Me.CheckBoxConnect.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'ButtonLabelDel
        '
        Me.ButtonLabelDel.Enabled = False
        Me.ButtonLabelDel.Location = New System.Drawing.Point(139, 368)
        Me.ButtonLabelDel.Name = "ButtonLabelDel"
        Me.ButtonLabelDel.Size = New System.Drawing.Size(125, 24)
        Me.ButtonLabelDel.TabIndex = 19
        Me.ButtonLabelDel.Text = "Labels verwijderen"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(21, 99)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(8, 16)
        Me.Label1.TabIndex = 91
        Me.Label1.Text = "/"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'CheckBoxEinddatum
        '
        Me.CheckBoxEinddatum.Enabled = False
        Me.CheckBoxEinddatum.ForeColor = System.Drawing.SystemColors.Control
        Me.CheckBoxEinddatum.Location = New System.Drawing.Point(96, 219)
        Me.CheckBoxEinddatum.Name = "CheckBoxEinddatum"
        Me.CheckBoxEinddatum.Size = New System.Drawing.Size(16, 16)
        Me.CheckBoxEinddatum.TabIndex = 9
        Me.CheckBoxEinddatum.Text = "CheckBox1"
        '
        'ButtonLabelAdd
        '
        Me.ButtonLabelAdd.Enabled = False
        Me.ButtonLabelAdd.Location = New System.Drawing.Point(8, 368)
        Me.ButtonLabelAdd.Name = "ButtonLabelAdd"
        Me.ButtonLabelAdd.Size = New System.Drawing.Size(125, 24)
        Me.ButtonLabelAdd.TabIndex = 18
        Me.ButtonLabelAdd.Text = "Label plaatsen"
        '
        'ButtonSave
        '
        Me.ButtonSave.Enabled = False
        Me.ButtonSave.Location = New System.Drawing.Point(8, 432)
        Me.ButtonSave.Name = "ButtonSave"
        Me.ButtonSave.Size = New System.Drawing.Size(256, 24)
        Me.ButtonSave.TabIndex = 22
        Me.ButtonSave.Text = "Wijzigingen opslaan"
        '
        'ComboBoxHydrantType
        '
        Me.ComboBoxHydrantType.BackColor = System.Drawing.SystemColors.Control
        Me.ComboBoxHydrantType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBoxHydrantType.Enabled = False
        Me.ComboBoxHydrantType.Location = New System.Drawing.Point(96, 315)
        Me.ComboBoxHydrantType.Name = "ComboBoxHydrantType"
        Me.ComboBoxHydrantType.Size = New System.Drawing.Size(104, 21)
        Me.ComboBoxHydrantType.TabIndex = 15
        '
        'LabelHydrantType
        '
        Me.LabelHydrantType.Location = New System.Drawing.Point(8, 315)
        Me.LabelHydrantType.Name = "LabelHydrantType"
        Me.LabelHydrantType.Size = New System.Drawing.Size(88, 16)
        Me.LabelHydrantType.TabIndex = 81
        Me.LabelHydrantType.Text = "Hydranttype"
        Me.LabelHydrantType.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'ComboBoxBron
        '
        Me.ComboBoxBron.BackColor = System.Drawing.SystemColors.Control
        Me.ComboBoxBron.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBoxBron.Enabled = False
        Me.ComboBoxBron.Location = New System.Drawing.Point(96, 339)
        Me.ComboBoxBron.Name = "ComboBoxBron"
        Me.ComboBoxBron.Size = New System.Drawing.Size(104, 21)
        Me.ComboBoxBron.TabIndex = 16
        '
        'LabelBron
        '
        Me.LabelBron.Location = New System.Drawing.Point(8, 339)
        Me.LabelBron.Name = "LabelBron"
        Me.LabelBron.Size = New System.Drawing.Size(88, 16)
        Me.LabelBron.TabIndex = 79
        Me.LabelBron.Text = "Bron"
        Me.LabelBron.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'ComboBoxStatus
        '
        Me.ComboBoxStatus.BackColor = System.Drawing.SystemColors.Control
        Me.ComboBoxStatus.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBoxStatus.Enabled = False
        Me.ComboBoxStatus.Location = New System.Drawing.Point(96, 267)
        Me.ComboBoxStatus.Name = "ComboBoxStatus"
        Me.ComboBoxStatus.Size = New System.Drawing.Size(104, 21)
        Me.ComboBoxStatus.TabIndex = 12
        '
        'LabelStatus
        '
        Me.LabelStatus.Location = New System.Drawing.Point(8, 267)
        Me.LabelStatus.Name = "LabelStatus"
        Me.LabelStatus.Size = New System.Drawing.Size(88, 16)
        Me.LabelStatus.TabIndex = 77
        Me.LabelStatus.Text = "Status"
        Me.LabelStatus.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'TextBoxPostcode
        '
        Me.TextBoxPostcode.BackColor = System.Drawing.SystemColors.Control
        Me.TextBoxPostcode.Enabled = False
        Me.TextBoxPostcode.ForeColor = System.Drawing.SystemColors.InactiveBorder
        Me.TextBoxPostcode.Location = New System.Drawing.Point(208, 339)
        Me.TextBoxPostcode.Name = "TextBoxPostcode"
        Me.TextBoxPostcode.Size = New System.Drawing.Size(56, 20)
        Me.TextBoxPostcode.TabIndex = 17
        '
        'LabelPostcode
        '
        Me.LabelPostcode.Location = New System.Drawing.Point(208, 323)
        Me.LabelPostcode.Name = "LabelPostcode"
        Me.LabelPostcode.Size = New System.Drawing.Size(56, 16)
        Me.LabelPostcode.TabIndex = 75
        Me.LabelPostcode.Text = "Postcode"
        Me.LabelPostcode.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'TextBoxStraatcode
        '
        Me.TextBoxStraatcode.BackColor = System.Drawing.SystemColors.Control
        Me.TextBoxStraatcode.Enabled = False
        Me.TextBoxStraatcode.ForeColor = System.Drawing.SystemColors.InactiveBorder
        Me.TextBoxStraatcode.Location = New System.Drawing.Point(208, 291)
        Me.TextBoxStraatcode.Name = "TextBoxStraatcode"
        Me.TextBoxStraatcode.Size = New System.Drawing.Size(56, 20)
        Me.TextBoxStraatcode.TabIndex = 14
        '
        'LabelStraatcode
        '
        Me.LabelStraatcode.Location = New System.Drawing.Point(208, 275)
        Me.LabelStraatcode.Name = "LabelStraatcode"
        Me.LabelStraatcode.Size = New System.Drawing.Size(50, 16)
        Me.LabelStraatcode.TabIndex = 73
        Me.LabelStraatcode.Text = "StrCode"
        Me.LabelStraatcode.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'TextBoxStraatnaam
        '
        Me.TextBoxStraatnaam.BackColor = System.Drawing.SystemColors.Control
        Me.TextBoxStraatnaam.Enabled = False
        Me.TextBoxStraatnaam.ForeColor = System.Drawing.SystemColors.InactiveBorder
        Me.TextBoxStraatnaam.Location = New System.Drawing.Point(96, 243)
        Me.TextBoxStraatnaam.Name = "TextBoxStraatnaam"
        Me.TextBoxStraatnaam.Size = New System.Drawing.Size(168, 20)
        Me.TextBoxStraatnaam.TabIndex = 11
        Me.TextBoxStraatnaam.TabStop = False
        '
        'LabelStraatnaam
        '
        Me.LabelStraatnaam.Location = New System.Drawing.Point(8, 243)
        Me.LabelStraatnaam.Name = "LabelStraatnaam"
        Me.LabelStraatnaam.Size = New System.Drawing.Size(88, 16)
        Me.LabelStraatnaam.TabIndex = 71
        Me.LabelStraatnaam.Text = "Straatnaam"
        Me.LabelStraatnaam.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'DatePickerEinddatum
        '
        Me.DatePickerEinddatum.Enabled = False
        Me.DatePickerEinddatum.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.DatePickerEinddatum.Location = New System.Drawing.Point(112, 219)
        Me.DatePickerEinddatum.Name = "DatePickerEinddatum"
        Me.DatePickerEinddatum.Size = New System.Drawing.Size(88, 20)
        Me.DatePickerEinddatum.TabIndex = 10
        Me.DatePickerEinddatum.Value = New Date(2005, 6, 29, 0, 0, 0, 0)
        '
        'LabelEinddatum
        '
        Me.LabelEinddatum.Location = New System.Drawing.Point(8, 219)
        Me.LabelEinddatum.Name = "LabelEinddatum"
        Me.LabelEinddatum.Size = New System.Drawing.Size(88, 16)
        Me.LabelEinddatum.TabIndex = 69
        Me.LabelEinddatum.Text = "EindDatum"
        Me.LabelEinddatum.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'DatePickerBegindatum
        '
        Me.DatePickerBegindatum.Enabled = False
        Me.DatePickerBegindatum.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.DatePickerBegindatum.Location = New System.Drawing.Point(96, 195)
        Me.DatePickerBegindatum.Name = "DatePickerBegindatum"
        Me.DatePickerBegindatum.Size = New System.Drawing.Size(104, 20)
        Me.DatePickerBegindatum.TabIndex = 8
        Me.DatePickerBegindatum.TabStop = False
        Me.DatePickerBegindatum.Value = New Date(2005, 6, 29, 0, 0, 0, 0)
        '
        'LabelBegindatum
        '
        Me.LabelBegindatum.Location = New System.Drawing.Point(8, 195)
        Me.LabelBegindatum.Name = "LabelBegindatum"
        Me.LabelBegindatum.Size = New System.Drawing.Size(88, 16)
        Me.LabelBegindatum.TabIndex = 67
        Me.LabelBegindatum.Text = "BeginDatum"
        Me.LabelBegindatum.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'TextBoxLeidingID
        '
        Me.TextBoxLeidingID.BackColor = System.Drawing.SystemColors.Control
        Me.TextBoxLeidingID.Enabled = False
        Me.TextBoxLeidingID.ForeColor = System.Drawing.SystemColors.InactiveBorder
        Me.TextBoxLeidingID.Location = New System.Drawing.Point(96, 171)
        Me.TextBoxLeidingID.Name = "TextBoxLeidingID"
        Me.TextBoxLeidingID.Size = New System.Drawing.Size(104, 20)
        Me.TextBoxLeidingID.TabIndex = 7
        '
        'LabelLeidingID
        '
        Me.LabelLeidingID.Location = New System.Drawing.Point(8, 171)
        Me.LabelLeidingID.Name = "LabelLeidingID"
        Me.LabelLeidingID.Size = New System.Drawing.Size(88, 16)
        Me.LabelLeidingID.TabIndex = 65
        Me.LabelLeidingID.Text = "Leidingnr"
        Me.LabelLeidingID.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'ComboBoxLeidingType
        '
        Me.ComboBoxLeidingType.BackColor = System.Drawing.SystemColors.Control
        Me.ComboBoxLeidingType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBoxLeidingType.Enabled = False
        Me.ComboBoxLeidingType.Location = New System.Drawing.Point(96, 147)
        Me.ComboBoxLeidingType.Name = "ComboBoxLeidingType"
        Me.ComboBoxLeidingType.Size = New System.Drawing.Size(104, 21)
        Me.ComboBoxLeidingType.TabIndex = 6
        '
        'LabelLeidingType
        '
        Me.LabelLeidingType.Location = New System.Drawing.Point(8, 147)
        Me.LabelLeidingType.Name = "LabelLeidingType"
        Me.LabelLeidingType.Size = New System.Drawing.Size(88, 16)
        Me.LabelLeidingType.TabIndex = 63
        Me.LabelLeidingType.Text = "Leidingtype"
        Me.LabelLeidingType.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'TextBoxLeverancierID
        '
        Me.TextBoxLeverancierID.BackColor = System.Drawing.SystemColors.Control
        Me.TextBoxLeverancierID.Enabled = False
        Me.TextBoxLeverancierID.ForeColor = System.Drawing.SystemColors.InactiveBorder
        Me.TextBoxLeverancierID.Location = New System.Drawing.Point(96, 123)
        Me.TextBoxLeverancierID.Name = "TextBoxLeverancierID"
        Me.TextBoxLeverancierID.Size = New System.Drawing.Size(104, 20)
        Me.TextBoxLeverancierID.TabIndex = 5
        Me.TextBoxLeverancierID.TabStop = False
        '
        'LabelLeverancierID
        '
        Me.LabelLeverancierID.Location = New System.Drawing.Point(8, 123)
        Me.LabelLeverancierID.Name = "LabelLeverancierID"
        Me.LabelLeverancierID.Size = New System.Drawing.Size(88, 16)
        Me.LabelLeverancierID.TabIndex = 61
        Me.LabelLeverancierID.Text = "Leveranciernr"
        Me.LabelLeverancierID.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'TextBoxYCoord
        '
        Me.TextBoxYCoord.BackColor = System.Drawing.SystemColors.Control
        Me.TextBoxYCoord.Enabled = False
        Me.TextBoxYCoord.ForeColor = System.Drawing.SystemColors.InactiveBorder
        Me.TextBoxYCoord.Location = New System.Drawing.Point(184, 99)
        Me.TextBoxYCoord.Name = "TextBoxYCoord"
        Me.TextBoxYCoord.Size = New System.Drawing.Size(80, 20)
        Me.TextBoxYCoord.TabIndex = 4
        Me.TextBoxYCoord.TabStop = False
        '
        'LabelYCoord
        '
        Me.LabelYCoord.Location = New System.Drawing.Point(32, 99)
        Me.LabelYCoord.Name = "LabelYCoord"
        Me.LabelYCoord.Size = New System.Drawing.Size(16, 16)
        Me.LabelYCoord.TabIndex = 59
        Me.LabelYCoord.Text = "Y"
        Me.LabelYCoord.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'TextBoxXCoord
        '
        Me.TextBoxXCoord.BackColor = System.Drawing.SystemColors.Control
        Me.TextBoxXCoord.Enabled = False
        Me.TextBoxXCoord.ForeColor = System.Drawing.SystemColors.WindowText
        Me.TextBoxXCoord.Location = New System.Drawing.Point(96, 99)
        Me.TextBoxXCoord.Name = "TextBoxXCoord"
        Me.TextBoxXCoord.Size = New System.Drawing.Size(80, 20)
        Me.TextBoxXCoord.TabIndex = 3
        Me.TextBoxXCoord.TabStop = False
        '
        'LabelXCoord
        '
        Me.LabelXCoord.Location = New System.Drawing.Point(8, 99)
        Me.LabelXCoord.Name = "LabelXCoord"
        Me.LabelXCoord.Size = New System.Drawing.Size(16, 16)
        Me.LabelXCoord.TabIndex = 57
        Me.LabelXCoord.Text = "X"
        Me.LabelXCoord.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'TextBoxBrandweerID
        '
        Me.TextBoxBrandweerID.BackColor = System.Drawing.SystemColors.Control
        Me.TextBoxBrandweerID.Enabled = False
        Me.TextBoxBrandweerID.ForeColor = System.Drawing.SystemColors.WindowText
        Me.TextBoxBrandweerID.Location = New System.Drawing.Point(96, 16)
        Me.TextBoxBrandweerID.Name = "TextBoxBrandweerID"
        Me.TextBoxBrandweerID.Size = New System.Drawing.Size(104, 20)
        Me.TextBoxBrandweerID.TabIndex = 0
        '
        'LabelBrandweerID
        '
        Me.LabelBrandweerID.Location = New System.Drawing.Point(8, 16)
        Me.LabelBrandweerID.Name = "LabelBrandweerID"
        Me.LabelBrandweerID.Size = New System.Drawing.Size(88, 16)
        Me.LabelBrandweerID.TabIndex = 55
        Me.LabelBrandweerID.Text = "Brandweernr"
        Me.LabelBrandweerID.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'ComboBoxLigging
        '
        Me.ComboBoxLigging.BackColor = System.Drawing.SystemColors.Control
        Me.ComboBoxLigging.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBoxLigging.Enabled = False
        Me.ComboBoxLigging.Location = New System.Drawing.Point(96, 291)
        Me.ComboBoxLigging.Name = "ComboBoxLigging"
        Me.ComboBoxLigging.Size = New System.Drawing.Size(104, 21)
        Me.ComboBoxLigging.TabIndex = 13
        '
        'LabelLigging
        '
        Me.LabelLigging.Location = New System.Drawing.Point(8, 291)
        Me.LabelLigging.Name = "LabelLigging"
        Me.LabelLigging.Size = New System.Drawing.Size(88, 16)
        Me.LabelLigging.TabIndex = 53
        Me.LabelLigging.Text = "Ligging"
        Me.LabelLigging.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'TextBoxDiameter
        '
        Me.TextBoxDiameter.BackColor = System.Drawing.SystemColors.Control
        Me.TextBoxDiameter.Enabled = False
        Me.TextBoxDiameter.Location = New System.Drawing.Point(96, 75)
        Me.TextBoxDiameter.Name = "TextBoxDiameter"
        Me.TextBoxDiameter.Size = New System.Drawing.Size(104, 20)
        Me.TextBoxDiameter.TabIndex = 2
        '
        'LabelDiameter
        '
        Me.LabelDiameter.Location = New System.Drawing.Point(8, 75)
        Me.LabelDiameter.Name = "LabelDiameter"
        Me.LabelDiameter.Size = New System.Drawing.Size(88, 16)
        Me.LabelDiameter.TabIndex = 51
        Me.LabelDiameter.Text = "Diameter"
        Me.LabelDiameter.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'TextBoxAanduiding
        '
        Me.TextBoxAanduiding.BackColor = System.Drawing.SystemColors.Control
        Me.TextBoxAanduiding.Enabled = False
        Me.TextBoxAanduiding.Location = New System.Drawing.Point(96, 40)
        Me.TextBoxAanduiding.Multiline = True
        Me.TextBoxAanduiding.Name = "TextBoxAanduiding"
        Me.TextBoxAanduiding.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.TextBoxAanduiding.Size = New System.Drawing.Size(168, 32)
        Me.TextBoxAanduiding.TabIndex = 1
        '
        'LabelAanduiding
        '
        Me.LabelAanduiding.Location = New System.Drawing.Point(8, 40)
        Me.LabelAanduiding.Name = "LabelAanduiding"
        Me.LabelAanduiding.Size = New System.Drawing.Size(88, 16)
        Me.LabelAanduiding.TabIndex = 49
        Me.LabelAanduiding.Text = "Aanduiding"
        Me.LabelAanduiding.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.LabelCounter)
        Me.GroupBox3.Controls.Add(Me.LabelTotal)
        Me.GroupBox3.Controls.Add(Me.LabelSeparator)
        Me.GroupBox3.Controls.Add(Me.ButtonNext)
        Me.GroupBox3.Controls.Add(Me.ButtonLast)
        Me.GroupBox3.Controls.Add(Me.ButtonFirst)
        Me.GroupBox3.Controls.Add(Me.ButtonPrevious)
        Me.GroupBox3.Location = New System.Drawing.Point(8, 86)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(272, 48)
        Me.GroupBox3.TabIndex = 1
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Featureset"
        '
        'LabelCounter
        '
        Me.LabelCounter.Location = New System.Drawing.Point(84, 19)
        Me.LabelCounter.Name = "LabelCounter"
        Me.LabelCounter.Size = New System.Drawing.Size(40, 16)
        Me.LabelCounter.TabIndex = 11
        Me.LabelCounter.Text = "#"
        Me.LabelCounter.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'LabelTotal
        '
        Me.LabelTotal.Location = New System.Drawing.Point(148, 19)
        Me.LabelTotal.Name = "LabelTotal"
        Me.LabelTotal.Size = New System.Drawing.Size(40, 16)
        Me.LabelTotal.TabIndex = 10
        Me.LabelTotal.Text = "#"
        '
        'LabelSeparator
        '
        Me.LabelSeparator.Location = New System.Drawing.Point(132, 19)
        Me.LabelSeparator.Name = "LabelSeparator"
        Me.LabelSeparator.Size = New System.Drawing.Size(8, 16)
        Me.LabelSeparator.TabIndex = 9
        Me.LabelSeparator.Text = "/"
        '
        'ButtonNext
        '
        Me.ButtonNext.Location = New System.Drawing.Point(192, 16)
        Me.ButtonNext.Name = "ButtonNext"
        Me.ButtonNext.Size = New System.Drawing.Size(32, 24)
        Me.ButtonNext.TabIndex = 1
        Me.ButtonNext.Text = ">"
        '
        'ButtonLast
        '
        Me.ButtonLast.Location = New System.Drawing.Point(232, 16)
        Me.ButtonLast.Name = "ButtonLast"
        Me.ButtonLast.Size = New System.Drawing.Size(32, 24)
        Me.ButtonLast.TabIndex = 2
        Me.ButtonLast.Text = ">>"
        '
        'ButtonFirst
        '
        Me.ButtonFirst.Location = New System.Drawing.Point(8, 16)
        Me.ButtonFirst.Name = "ButtonFirst"
        Me.ButtonFirst.Size = New System.Drawing.Size(32, 24)
        Me.ButtonFirst.TabIndex = 3
        Me.ButtonFirst.Text = "<<"
        '
        'ButtonPrevious
        '
        Me.ButtonPrevious.Location = New System.Drawing.Point(48, 16)
        Me.ButtonPrevious.Name = "ButtonPrevious"
        Me.ButtonPrevious.Size = New System.Drawing.Size(32, 24)
        Me.ButtonPrevious.TabIndex = 4
        Me.ButtonPrevious.Text = "<"
        '
        'FormBeheerHydranten
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(285, 640)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.ButtonClose)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.GroupBox2)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.Name = "FormBeheerHydranten"
        Me.Text = "Beheer van hydranten"
        Me.TopMost = True
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region " Overloaded constructor "

    <CLSCompliant(False)> _
    Public Sub New(ByVal ArcMapApplication As IMxApplication)

        MyBase.New()

        'Initialise locals.
        m_application = ArcMapApplication
        m_document = CType(CType(m_application, IApplication).Document, IMxDocument)
        m_layer = GetFeatureLayer(m_document.FocusMap, GetLayerName("Hydrant"))
        m_editing = False
        m_workspace = Nothing
        'm_marker = Nothing
        m_enumOIDs = Nothing
        m_copyFrom = Nothing

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Get the workspace of the hydranten.
        m_workspace = GetLayerWorkspace(ArcMapApplication, CType(m_layer, IFeatureLayer))

        'Custom form initialization.
        InitializeForm()
    End Sub

#End Region

#Region " Implementation of IConnectFeature "

    Private Property Straatnaam() As String Implements IConnectFeature.Straatnaam
        Get
            Return TextBoxStraatnaam.Text
        End Get
        Set(ByVal Value As String)
            TextBoxStraatnaam.Text = Value
        End Set
    End Property

    Private Property Straatcode() As String Implements IConnectFeature.Straatcode
        Get
            Return TextBoxStraatcode.Text
        End Get
        Set(ByVal Value As String)
            TextBoxStraatcode.Text = Value
        End Set
    End Property

    Private Property Postcode() As String Implements IConnectFeature.Postcode
        Get
            Return TextBoxPostcode.Text
        End Get
        Set(ByVal Value As String)
            TextBoxPostcode.Text = Value
        End Set
    End Property

    Public ReadOnly Property Toolbutton() As System.Windows.Forms.CheckBox Implements IConnectFeature.Toolbutton
        Get
            Return Me.CheckBoxConnect
        End Get
    End Property

#End Region

#Region " Form controls events "

    Private Sub ButtonClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonClose.Click

        Me.Close() 'Close form.
        'The user will be able to store his changes to current features attributes,
        'before the form is closed, becauce of the OnClosing event of current form.

    End Sub

    Private Sub ButtonLoad_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonLoad.Click
        'RW:07-08/2008
        Try

            Dim pHydrantLayer As ILayer = Nothing
            m_selectionSet = Nothing

            'Check if attributes of currently loaded feature are modified.
            'If so, allow the user to store these changes before loading another set of features.
            If ModifiedAttribute() Then
                If MsgBox(c_Message_SaveChanges, MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    If ValidateAttributeChanges() Then
                        StoreAttributeChanges()
                    Else
                        Exit Sub
                    End If
                End If
            End If

            'Get the hydranten layer from the map.
            pHydrantLayer = m_layer

            'Determine which radiobutton is selected.
            If Me.RadioButtonMapSelection.Checked Then
                'Get the current map selection.
                m_selectionSet = CType(pHydrantLayer, IFeatureSelection).SelectionSet

            ElseIf Me.RadioButtonStatusFilter.Checked Then
                'Filter on status of hydrants.

                'Get selected status value.
                Dim StatusFilterValue As String
                StatusFilterValue = Me.ComboBoxStatusFilter.Text
                StatusFilterValue = Mid(StatusFilterValue, 1, InStr(StatusFilterValue, ":") - 1)
                'MsgBox(StatusFilterValue)

                'Get a selectionset with specified status value.
                Dim pTable As ITable = CType(pHydrantLayer, ITable)
                Dim pQueryFilter As IQueryFilter = New QueryFilter
                pQueryFilter.WhereClause = GetAttributeName("Hydrant", "Status") & " = '" & StatusFilterValue & "'"
                m_selectionSet = pTable.Select(pQueryFilter, esriSelectionType.esriSelectionTypeHybrid, esriSelectionOption.esriSelectionOptionNormal, m_workspace)

            ElseIf Me.RadioButtonTypeFilter.Checked Then
                'Filter on type of hydrants.

                'Get selected status value.
                Dim TypeFilterValue As String
                TypeFilterValue = Me.ComboBoxTypeFilter.Text
                TypeFilterValue = Mid(TypeFilterValue, 1, InStr(TypeFilterValue, ":") - 1)
                'MsgBox(TypeFilterValue)

                'Get a selectionset with specified type value.
                Dim pTable As ITable = CType(pHydrantLayer, ITable)
                Dim pQueryFilter As IQueryFilter = New QueryFilter
                pQueryFilter.WhereClause = GetAttributeName("Hydrant", "HydrantType") & " = '" & TypeFilterValue & "'"
                m_selectionSet = pTable.Select(pQueryFilter, esriSelectionType.esriSelectionTypeHybrid, esriSelectionOption.esriSelectionOptionNormal, m_workspace)

            End If

            'Load the selection set in the form.
            If Not m_selectionSet Is Nothing Then LoadSelectionSet()
        Catch ex As Exception
            ErrorHandler(ex)
        End Try

    End Sub

    Private Sub ButtonSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonSave.Click
        'RW:07-08/2008

        Try

            'Store attribute changes if modifications are registered.
            If ModifiedAttribute() Then
                If ValidateAttributeChanges() Then
                    StoreAttributeChanges()
                End If
            End If

        Catch ex As Exception
            ErrorHandler(ex)
        End Try
    End Sub

    Private Sub ButtonFirst_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonFirst.Click
        'RW:07-08/2008
        Try

            If ModifiedAttribute() Then
                If MsgBox(c_Message_SaveChanges, MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    If ValidateAttributeChanges() Then
                        StoreAttributeChanges()
                        LoadFirstFeature()
                    End If
                Else
                    LoadFirstFeature()
                End If
            Else
                LoadFirstFeature()
            End If
        Catch ex As Exception
            ErrorHandler(ex)
        End Try
    End Sub

    Private Sub ButtonPrevious_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonPrevious.Click
        'RW:07-08/2008
        Try
            If ModifiedAttribute() Then
                If MsgBox(c_Message_SaveChanges, MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    If ValidateAttributeChanges() Then
                        StoreAttributeChanges()
                        LoadPreviousFeature()
                    End If
                Else
                    LoadPreviousFeature()
                End If
            Else
                LoadPreviousFeature()
            End If
        Catch ex As Exception
            ErrorHandler(ex)
        End Try

    End Sub

    Private Sub ButtonNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonNext.Click
        'RW:07-08/2008
        Try
            If ModifiedAttribute() Then
                If MsgBox(c_Message_SaveChanges, MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    If ValidateAttributeChanges() Then
                        StoreAttributeChanges()
                        LoadNextFeature()
                    End If
                Else
                    LoadNextFeature()
                End If
            Else
                LoadNextFeature()
            End If

        Catch ex As Exception
            ErrorHandler(ex)
        End Try
    End Sub

    Private Sub ButtonLast_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonLast.Click
        'RW:07-08/2008
        Try

            If ModifiedAttribute() Then
                If MsgBox(c_Message_SaveChanges, MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    If ValidateAttributeChanges() Then
                        StoreAttributeChanges()
                        LoadLastFeature()
                    End If
                Else
                    LoadLastFeature()
                End If
            Else
                LoadLastFeature()
            End If
        Catch ex As Exception
            ErrorHandler(ex)
        End Try
    End Sub

    Private Sub ButtonLabelAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonLabelAdd.Click

        Try

            ' Make sure there is some text to use as annotation.
            If Len(Trim(Me.TextBoxAanduiding.Text)) = 0 Then
                MsgBox(c_Message_AanduidingIsEmpty, MsgBoxStyle.Exclamation, c_Title_BeheerHydranten)
                Exit Sub
            End If

            ' Get annotations layer.
            Dim pLayer As ILayer = GetFeatureLayer(m_document.FocusMap, GetLayerName("HydrantAnno"))
            If pLayer Is Nothing Then Throw New LayerNotFoundException(GetLayerName("HydrantAnno"))
            Dim pAnnoLayer As IAnnotationLayer = CType(pLayer, IAnnotationLayer)

            ' Add new annotation feature.
            Dim pMarkerElement As IMarkerElement = GetMarkerElement(c_MarkerTag, m_document)
            If Not pMarkerElement Is Nothing Then
                Dim pGeometry As IGeometry = CType(pMarkerElement, IElement).Geometry
                If TypeOf pGeometry Is IPoint Then
                    Dim frm As FormAddAnnotation = _
                        New FormAddAnnotation( _
                            pAnnoLayer, _
                            CType(pGeometry, IPoint), _
                            Me.TextBoxAanduiding.Text, _
                            GetAttributeName("HydrantAnno", "LinkID"), _
                            Me.TextBoxLeverancierID.Text)
                    frm.ShowDialog()
                    frm.Dispose()
                End If
            End If

            ' Partial refresh to display new annotation.
            m_document.ActivatedView.PartialRefresh(esriViewDrawPhase.esriViewGeography, pAnnoLayer, Nothing)

            ' Activate the Edit Annotation Tool command if edit session is started.
            Dim pEditor As IEditor2 = GetEditorReference(m_application)
            EditSessionStart(pEditor, GetFeatureLayer(m_document.FocusMap, GetLayerName("Hydrant")), True)
            If Not pEditor.EditWorkspace Is Nothing Then
                ' Make the annotation layer selectable.
                GetFeatureLayer(m_document.FocusMap, GetLayerName("HydrantAnno")).Selectable = True
                ' Activate the Edit Annotation tool.
                ActivateTool(CType(m_document, IDocument), "esriEditor.AnnoEditTool")
            End If

        Catch ex As Exception
            ' Throw ex
            'RW:07-08/2008
            ErrorHandler(ex)

        End Try

    End Sub

    Private Sub ButtonLabelDel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonLabelDel.Click
        Try
            Dim title As String = c_Title_DeleteAnno
            Dim message As String = c_Message_ConfirmDeleteAnno

            'Ask the user for a confirmation.
            If MsgBox(message, MsgBoxStyle.OKCancel, title) = MsgBoxResult.OK Then

                'Remove all annotations with same LinkID.
                Dim linkID As String = CStr(TextBoxLeverancierID.Text)
                Dim pAnnoLayer As IAnnotationLayer = CType(GetFeatureLayer(m_document.FocusMap, GetLayerName("HydrantAnno")), IAnnotationLayer)
                Dim annoCount As Integer = RemoveLinkedAnnotations(pAnnoLayer, GetAttributeName("HydrantAnno", "LinkID"), linkID)

                'Partial refresh to display new annotation.
                m_document.ActivatedView.PartialRefresh(esriViewDrawPhase.esriViewGeography, pAnnoLayer, Nothing)

                'Inform the user about the number of deleted annotations.
                message = Replace(c_Message_DeleteAnnoCount, "^0", CStr(annoCount))
                MsgBox(message, MsgBoxStyle.OKOnly, title)

            End If
        Catch ex As Exception
            'Throw ex
            'RW:07-08/2008
            ErrorHandler(ex)

        End Try
    End Sub

    Private Sub TextBoxAanduiding_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBoxAanduiding.TextChanged
        If m_editing Then MarkAsChanged(Me.LabelAanduiding)
    End Sub

    Private Sub TextBoxDiameter_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBoxDiameter.TextChanged
        If m_editing Then MarkAsChanged(Me.LabelDiameter)
    End Sub

    Private Sub TextBoxBrandweerID_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBoxBrandweerID.TextChanged
        If m_editing Then MarkAsChanged(Me.LabelBrandweerID)
    End Sub

    Private Sub ComboBoxLeidingType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBoxLeidingType.SelectedIndexChanged
        If m_editing Then MarkAsChanged(Me.LabelLeidingType)
    End Sub

    Private Sub TextBoxLeidingID_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBoxLeidingID.TextChanged
        If m_editing Then MarkAsChanged(Me.LabelLeidingID)
    End Sub

    Private Sub DatePickerEinddatum_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DatePickerEinddatum.ValueChanged
        If m_editing Then MarkAsChanged(Me.LabelEinddatum)
        Me.CheckBoxEinddatum.Checked = True
    End Sub

    Private Sub ComboBoxStatus_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBoxStatus.SelectedIndexChanged
        If m_editing Then MarkAsChanged(Me.LabelStatus)
    End Sub

    Private Sub ComboBoxLigging_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBoxLigging.SelectedIndexChanged
        If m_editing Then MarkAsChanged(Me.LabelLigging)
    End Sub

    Private Sub ComboBoxHydrantType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBoxHydrantType.SelectedIndexChanged
        If m_editing Then MarkAsChanged(Me.LabelHydrantType)
    End Sub

    Private Sub ComboBoxBron_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBoxBron.SelectedIndexChanged
        If m_editing Then MarkAsChanged(Me.LabelBron)
    End Sub

    Private Sub CheckBoxEinddatum_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxEinddatum.CheckedChanged
        If m_editing Then
            MarkAsChanged(Me.LabelEinddatum)
            Me.DatePickerEinddatum.Enabled = Me.CheckBoxEinddatum.Checked
        End If
    End Sub

    Private Sub RadioButtonStatusFilter_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButtonStatusFilter.CheckedChanged
        Me.ComboBoxTypeFilter.Enabled = False
        Me.ComboBoxStatusFilter.Enabled = True
        UpdateButtonLoadAvailability() 'Disable the load button.
        If Not m_editing Then Exit Sub 'Avoid running into loops.
        ResetFormControls() 'Finalize ongoing editing and reset editing controls and labels.
    End Sub

    Private Sub RadioButtonMapSelection_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButtonMapSelection.CheckedChanged
        Me.ComboBoxTypeFilter.Enabled = False
        Me.ComboBoxStatusFilter.Enabled = False
        UpdateButtonLoadAvailability() 'Enable the load button.
        If Not m_editing Then Exit Sub 'Avoid running into loops.
        ResetFormControls() 'Finalize ongoing editing and reset editing controls and labels.
    End Sub

    Private Sub RadioButtonTypeFilter_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButtonTypeFilter.CheckedChanged
        Me.ComboBoxStatusFilter.Enabled = False
        Me.ComboBoxTypeFilter.Enabled = True
        UpdateButtonLoadAvailability() 'Disable the load button.
        If Not m_editing Then Exit Sub 'Avoid running into loops.
        ResetFormControls() 'Finalize ongoing editing and reset editing controls and labels.
    End Sub

    Private Sub TextBoxStraatnaam_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBoxStraatnaam.TextChanged
        If m_editing Then MarkAsChanged(Me.LabelStraatnaam)
    End Sub

    Private Sub TextBoxStraatcode_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBoxStraatcode.TextChanged
        If m_editing Then MarkAsChanged(Me.LabelStraatcode)
    End Sub

    Private Sub TextBoxPostcode_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBoxPostcode.TextChanged
        If m_editing Then MarkAsChanged(Me.LabelPostcode)
    End Sub

    Private Sub ComboBoxStatusFilter_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBoxStatusFilter.SelectedIndexChanged
        UpdateButtonLoadAvailability() 'Enable the load button.
        ResetFormControls() 'Finalize ongoing editing and reset editing controls and labels.
    End Sub

    Private Sub ComboBoxTypeFilter_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBoxTypeFilter.SelectedIndexChanged
        UpdateButtonLoadAvailability() 'Enable the load button.
        ResetFormControls() 'Finalize ongoing editing and reset editing controls and labels.
    End Sub

    Private Sub CheckBoxConnect_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxConnect.CheckedChanged
        Try
            If Me.CheckBoxConnect.Checked Then
                'Deactivate CopyAttributesFunctionality
                CopyAttributesFunctionality_Deactivate()
                'Activate ConnectFeatureFunctionality
                ConnectFeatureFunctionality_Activate(m_document, Me)
                'Show text in another color for better perception.
                Me.CheckBoxConnect.ForeColor = ActivatedToolbuttonForeColor
            Else
                'Deactivate ConnectFeatureFunctionality
                ConnectFeatureFunctionality_Deactivate()
                'Restore text to the default color.
                Me.CheckBoxConnect.ForeColor = DeactivatedToolbuttonForeColor
            End If
        Catch ex As Exception
            'Throw ex
            'RW:07-08/2008
            ErrorHandler(ex)
        End Try
    End Sub

    Private Sub CheckBoxCopy_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxCopy.CheckedChanged
        Try
            If Me.CheckBoxCopy.Checked Then
                'Deactivate ConnectFeatureFunctionality
                ConnectFeatureFunctionality_Deactivate()
                'Activate CopyAttributesFunctionality
                CopyAttributesFunctionality_Activate(m_document, Me)
                'Show text in another color for better perception.
                Me.CheckBoxCopy.ForeColor = ActivatedToolbuttonForeColor
            Else
                'Deactivate CopyAttributesFunctionality
                CopyAttributesFunctionality_Deactivate()
                'Restore text to the default color.
                Me.CheckBoxCopy.ForeColor = DeactivatedToolbuttonForeColor
            End If
        Catch ex As Exception
            'Throw ex
            'RW:07-08/2008
            ErrorHandler(ex)
        End Try
    End Sub

#End Region

#Region " Overridden form events "

    Protected Overrides Sub OnClosing(ByVal e As System.ComponentModel.CancelEventArgs)

        Dim pActiveView As IActiveView
        Dim pElement As IElement
        Dim pGraphics As IGraphicsContainer
        Dim pMarker As IMarkerElement
        Dim pMxDocument As IMxDocument
        Dim pEditor As IEditor2 = Nothing 'editor of edit session

        Try

            ' Allow the user to store modifications to current feature,
            ' before continuing.
            If ModifiedAttribute() Then
                If MsgBox(c_Message_SaveChanges, MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    If ValidateAttributeChanges() Then
                        StoreAttributeChanges()
                    Else
                        ' Do not continue, because user wanted to save changes.
                        ' But changes are invalid. So corrections are required.
                        e.Cancel = True
                        Exit Sub
                    End If
                End If
            End If

            'Close active edit session before continuing.
            pEditor = GetEditorReference(m_application)
            If pEditor.EditState = esriEditState.esriStateEditing Then

                'Does the user wants to save changes while closing the edit session?
                If MsgBox(c_Message_SaveEdits, vbYesNo, c_Title_SaveEdits) = MsgBoxResult.Yes Then

                    'Close the active edit session and save changes.
                    EditSessionSave(pEditor)

                Else

                    'Close the active edit session without saving changes.
                    EditSessionAbort(pEditor)

                End If
            End If

            ' Remove marker if there is one.
            pMxDocument = m_document
            pMarker = GetMarkerElement(c_MarkerTag, pMxDocument)
            If Not pMarker Is Nothing Then
                pGraphics = CType(pMxDocument.FocusMap, IGraphicsContainer)
                pElement = CType(pMarker, IElement)
                pGraphics.DeleteElement(pElement)
                pActiveView = pMxDocument.ActivatedView
                pActiveView.PartialRefresh(esriViewGraphics, Nothing, Nothing)
            End If

        Catch ex As Exception
            e.Cancel = True
            Throw ex
        End Try

    End Sub

    Private Sub FormBeheerHydranten_Closed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Closed
        'Be sure to remove remaining eventhandler from the "Connect feature" functionality.
        ConnectFeatureFunctionality_Deactivate()
        'Be sure to remove remaining eventhandler from the "Copy attributes" functionality.
        CopyAttributesFunctionality_Deactivate()
    End Sub

#End Region

#Region " Utility procedures "

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Use the SelectionSet as the set of data editable with this form.
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    ''' 	[Kristof Vydt]	12/10/2005	Initialise editing controls &amp; labels in case of an empty featureset.
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub LoadSelectionSet()

        Try
            m_editing = False

            'Set form navigation controls.
            Dim SelectionSetCount As Integer
            SelectionSetCount = m_selectionSet.Count
            Me.LabelTotal.Text = CStr(SelectionSetCount)
            Me.LabelCounter.Text = CStr(0)

            'Enumeration of feature IDs.
            m_enumOIDs = CType(m_selectionSet.IDs, IEnumIDs)

            'Get selection extent
            Dim cursor As ICursor = Nothing
            m_selectionSet.Search(Nothing, True, cursor)
            Dim featCursor As IFeatureCursor = CType(cursor, IFeatureCursor)
            Dim feat As IFeature = featCursor.NextFeature
            If (Not feat Is Nothing) Then
                m_selectionExtent = New EnvelopeClass()
                Dim geomType As esriGeometryType = feat.Shape.GeometryType
                While Not feat Is Nothing
                    m_document.FocusMap.SelectFeature(m_layer, feat)
                    m_selectionExtent.Union(feat.Shape.Envelope)
                    feat = featCursor.NextFeature
                End While

                'Expand the selection extent
                If geomType = esriGeometryType.esriGeometryPoint Then
                    'buffer around points.
                    m_selectionExtent.Expand(c_ZoomPointBuffer, c_ZoomPointBuffer, False)
                Else 'buffer around polygon
                    m_selectionExtent.Expand(c_ZoomPolygonBuffer, c_ZoomPolygonBuffer, True)
                End If
            End If

            'Load first record from SelectionSet into the form controls.
            'Or disable all controls if selectionset is empty.
            If SelectionSetCount > 0 Then
                LoadNextFeature()
                EnableNavigationControls(True)
            Else
                MsgBox(c_Message_EmptyFeatureSet, MsgBoxStyle.Exclamation)
                InitializeEditingControls()
                InitializeEditingLabels()
                EnableNavigationControls(False)
            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Load the feature with the specified objectID, displaying its attributes in the form.
    ''' </summary>
    ''' <param name="OID"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    ''' 	[Kristof Vydt]	05/10/2005	Set MinDate for EndDate when loading StartDate.
    ''' 	[Kristof Vydt]	11/10/2005	Use InitializeEditingLabels &amp; InitializeEditingLabels instead of individual method calls.
    ''' 	[Kristof Vydt]	23/11/2005	Add optional parameter to avoid resetting reference to the CopyAttributes reference during reload.
    '''                                 Add optional parameter to avoid zoom-to during reload of feature.
    ''' 	[Kristof Vydt]	02/08/2006	Replace some global constant by a function that reads the value from config file.
    '''     [Kristof Vydt]  18/08/2006  MarkerElement is no longer a parameter of MarkAndZoomTo().
    '''     [Kristof Vydt]  08/09/2006  Assign yesterday as default value for EindDate.
    '''     [Kristof Vydt]  22/09/2006  Use today as default EindDatum if BeginDatum is today.
    '''     [Kristof Vydt]  22/02/2007  Eliminate on form legend code controls. Correct attribute reference casing.
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub LoadFeature( _
        ByVal OID As Integer, _
        Optional ByVal zoomToFeatures As Boolean = True, _
        Optional ByVal resetCopyFromReference As Boolean = False)

        Try
            'Disable the MarkAsChanged listener.
            m_editing = False

            'Clear the reference to the "CopyFrom" feature, in case of a reload.
            If resetCopyFromReference Then SetCopyFrom(Nothing)

            'Make sure there is an enumeration of IDs.
            If m_enumOIDs Is Nothing Then Exit Sub

            'Hold the current objectID as a form private variable.
            m_OID = OID

            'Get a cursor from the SelectionSet.
            Dim pTable As ITable
            Dim pQueryFilter As IQueryFilter = New QueryFilter
            Dim pCursor As ICursor
            pQueryFilter.WhereClause = "OBJECTID = " & OID
            pTable = CType(m_layer, ITable)
            pCursor = pTable.Search(pQueryFilter, True)
            Dim pRow As IRow = pCursor.NextRow

            'Zoom to the feature and mark it on the map.
            Dim pFeature As IFeature
            pFeature = CType(pRow, IFeature)
            If zoomToFeatures And Not m_selectionExtent Is Nothing Then
                m_document.ActiveView.Extent = m_selectionExtent
            End If
            MarkAndZoomTo(pFeature, m_document, False)

            'Initialize layout of form controls and
            'Show feature attributes in the form controls.
            InitializeEditingControls()
            InitializeEditingLabels()
            Dim FieldIndex As Integer
            '- Aanduiding
            FieldIndex = pRow.Fields.FindField(GetAttributeName("Hydrant", "Aanduiding"))
            SetEditBoxValue(Me.TextBoxAanduiding, pRow.Value(FieldIndex))
            '- Begindatum
            FieldIndex = pRow.Fields.FindField(GetAttributeName("Hydrant", "BeginDatum"))
            'Assign current date if it doesn't have a value yet.
            Dim StartDate As Object = pRow.Value(FieldIndex)
            If TypeOf StartDate Is System.DBNull Then
                SetEditBoxValue(Me.DatePickerBegindatum, Today)
                MarkAsChanged(Me.LabelBegindatum)
                Me.DatePickerEinddatum.MinDate = Today.AddDays(-1)
            Else
                SetEditBoxValue(Me.DatePickerBegindatum, StartDate)
                Me.DatePickerEinddatum.MinDate = CDate(StartDate)
            End If
            '- BrandweerID
            FieldIndex = pRow.Fields.FindField(GetAttributeName("Hydrant", "BrandweerNr"))
            SetEditBoxValue(Me.TextBoxBrandweerID, pRow.Value(FieldIndex))
            '- Bron
            FieldIndex = pRow.Fields.FindField(GetAttributeName("Hydrant", "Bron"))
            SetEditBoxValue(Me.ComboBoxBron, CStr(pRow.Value(FieldIndex)))
            '- Diameter
            FieldIndex = pRow.Fields.FindField(GetAttributeName("Hydrant", "Diameter"))
            SetEditBoxValue(Me.TextBoxDiameter, pRow.Value(FieldIndex))
            '- Einddatum
            FieldIndex = pRow.Fields.FindField(GetAttributeName("Hydrant", "EindDatum"))
            'Assign yesterday as default if it doesn't have a value yet.
            Dim EndDate As Object = pRow.Value(FieldIndex)
            If TypeOf EndDate Is System.DBNull Then
                Dim Yesterday As DateTime = Today.AddDays(-1)
                If Yesterday < Me.DatePickerEinddatum.MinDate Then
                    SetEditBoxValue(Me.DatePickerEinddatum, Me.DatePickerEinddatum.MinDate())
                Else
                    SetEditBoxValue(Me.DatePickerEinddatum, Today.AddDays(-1))
                End If
                Me.CheckBoxEinddatum.Checked = False
            Else
                SetEditBoxValue(Me.DatePickerEinddatum, EndDate)
                Me.CheckBoxEinddatum.Checked = True
            End If
            '- Hydranttype
            FieldIndex = pRow.Fields.FindField(GetAttributeName("Hydrant", "HydrantType"))
            SetEditBoxValue(Me.ComboBoxHydrantType, CStr(pRow.Value(FieldIndex)))
            '- LeidingNummer
            FieldIndex = pRow.Fields.FindField(GetAttributeName("Hydrant", "LeidingNr"))
            SetEditBoxValue(Me.TextBoxLeidingID, pRow.Value(FieldIndex))
            '- LeidingType
            FieldIndex = pRow.Fields.FindField(GetAttributeName("Hydrant", "LeidingType"))
            SetEditBoxValue(Me.ComboBoxLeidingType, CStr(pRow.Value(FieldIndex)))
            '- LeverancierNummer
            FieldIndex = pRow.Fields.FindField(GetAttributeName("Hydrant", "LeverancierNr"))
            SetEditBoxValue(Me.TextBoxLeverancierID, pRow.Value(FieldIndex))
            'Create a new LeverancierID if it doesn't have a value yet.
            If (Me.TextBoxLeverancierID.Text = "") Then
                SetEditBoxValue(Me.TextBoxLeverancierID, NewLerancierNr(m_document))
                MarkAsChanged(Me.LabelLeverancierID)
            End If
            '- Ligging
            FieldIndex = pRow.Fields.FindField(GetAttributeName("Hydrant", "Ligging"))
            SetEditBoxValue(Me.ComboBoxLigging, pRow.Value(FieldIndex))
            '- Postcode
            FieldIndex = pRow.Fields.FindField(GetAttributeName("Hydrant", "Postcode"))
            SetEditBoxValue(Me.TextBoxPostcode, pRow.Value(FieldIndex))
            '- Status
            FieldIndex = pRow.Fields.FindField(GetAttributeName("Hydrant", "Status"))
            SetEditBoxValue(Me.ComboBoxStatus, CStr(pRow.Value(FieldIndex)))
            '- Straatcode
            FieldIndex = pRow.Fields.FindField(GetAttributeName("Hydrant", "Straatcode"))
            SetEditBoxValue(Me.TextBoxStraatcode, pRow.Value(FieldIndex))
            '- Straatnaam
            FieldIndex = pRow.Fields.FindField(GetAttributeName("Hydrant", "Straatnaam"))
            SetEditBoxValue(Me.TextBoxStraatnaam, pRow.Value(FieldIndex))
            '- XCoord
            FieldIndex = pRow.Fields.FindField(GetAttributeName("Hydrant", "CoordX"))
            SetEditBoxValue(Me.TextBoxXCoord, pRow.Value(FieldIndex))
            'Capture and show X-coordinates if not already known.
            If (Me.TextBoxXCoord.Text = "0") Or (Me.TextBoxXCoord.Text = "") Then
                SetEditBoxValue(Me.TextBoxXCoord, CType(pFeature.Shape, IPoint).X)
                MarkAsChanged(Me.LabelXCoord)
            End If
            '- YCoord
            FieldIndex = pRow.Fields.FindField(GetAttributeName("Hydrant", "CoordY"))
            SetEditBoxValue(Me.TextBoxYCoord, pRow.Value(FieldIndex))
            'Capture and show Y-coordinates if not already known.
            If (Me.TextBoxYCoord.Text = "0") Or (Me.TextBoxYCoord.Text = "") Then
                SetEditBoxValue(Me.TextBoxYCoord, CType(pFeature.Shape, IPoint).Y)
                MarkAsChanged(Me.LabelYCoord)
            End If

            'Disable editing historic records.
            If GetComboBoxCodeValue(Me.ComboBoxStatus) = "3" Then
                EnableEditingControls(False)
            Else
                EnableEditingControls(True)
            End If

            'Enable the MarkAsChanged listener.
            m_editing = True

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Mark a control as "changed".
    ''' </summary>
    ''' <param name="SomeControl">
    '''     The label of the control that is changed.
    ''' </param>
    ''' <remarks>
    '''     The label of the modified control gets another color.
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    ''' 	[Kristof Vydt]	11/10/2005	Use color settings declared as private to the form.
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub MarkAsChanged(ByVal SomeControl As Windows.Forms.Control)
        If TypeOf SomeControl Is Windows.Forms.Label Then
            SomeControl.ForeColor = ChangedLabelForeColor
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Load the last feature into the form, from the current editing set.
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub LoadLastFeature()

        Dim Counter As Integer = CInt(Me.LabelCounter.Text)
        Dim Total As Integer = CInt(Me.LabelTotal.Text)
        Dim ObjID As Integer

        If Counter < Total Then
            While Counter < Total
                Counter = Counter + 1
                ObjID = m_enumOIDs.Next
            End While
            Me.LabelCounter.Text = CStr(Counter)
            LoadFeature(ObjID)
        End If

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Load the following feature into the form, from the current editing set.
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub LoadNextFeature()

        Dim Counter As Integer = CInt(Me.LabelCounter.Text)
        Dim Total As Integer = CInt(Me.LabelTotal.Text)

        If Counter < Total Then
            Me.LabelCounter.Text = CStr(Counter + 1)
            LoadFeature(m_enumOIDs.Next)
        End If

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Load the previous feature into the form, from the current editing set.
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub LoadPreviousFeature()

        Dim Counter As Integer = CInt(Me.LabelCounter.Text)
        Dim Total As Integer = CInt(Me.LabelTotal.Text)
        Dim ObjID As Integer 'requested object ID
        Dim RequestedIndex As Integer

        If Counter > 1 Then
            RequestedIndex = Counter - 1

            'Reset the enumeration.
            Counter = 0
            m_enumOIDs.Reset()

            'Loop until the requested feature is found in the enumeration.
            While Counter < RequestedIndex
                ObjID = m_enumOIDs.Next
                Counter = Counter + 1
            End While

            'Load into form.
            Me.LabelCounter.Text = CStr(Counter)
            LoadFeature(ObjID)
        End If

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Load the first feature into the form, from the current editing set.
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub LoadFirstFeature()
        m_enumOIDs.Reset()
        Me.LabelCounter.Text = CStr(0)
        LoadNextFeature()
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Set some form controls (labels, text boxes, combo boxes, ...) to initial values.
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    ''' 	[Kristof Vydt]	10/10/2005	Type filter added.
    '''     [Kristof Vydt]  11/11/2005  Call initialization of editing controls and labels.
    ''' 	[Kristof Vydt]	17/10/2005	Disable filter comboboxes on load.
    ''' 	[Kristof Vydt]	02/08/2006	Replace some global constant by a function that reads the value from config file.
    '''     [Kristof Vydt]  22/02/2007  Eliminate on form legend code controls. Correct domain name casing.
    ''' 	[Kristof Vydt]	22/03/2007	Use the new CodedValueDomainManager instead of the deprecated ModuleDomainAccess.
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub InitializeForm()

        'Hide some form controls if not debugging.
        TextBoxStraatcode.Visible = c_DebugStatus
        LabelStraatcode.Visible = c_DebugStatus
        TextBoxPostcode.Visible = c_DebugStatus
        LabelPostcode.Visible = c_DebugStatus

        Dim domainMgr As CodedValueDomainManager

        ' List domain code values in the statusfilter combo box.
        'PopulateCodes(m_workspace, GetDomainName("Status"), Me.ComboBoxStatusFilter)
        domainMgr = New CodedValueDomainManager(m_workspace, "Status")
        domainMgr.PopulateCodes(Me.ComboBoxStatusFilter)

        ' List domain code values in the typefilter combo box.
        'PopulateCodes(m_workspace, GetDomainName("HydrantType"), Me.ComboBoxTypeFilter)
        domainMgr = New CodedValueDomainManager(m_workspace, "HydrantType")
        domainMgr.PopulateCodes(Me.ComboBoxTypeFilter)

        ' List domain code values in the leidingtype combo box.
        'PopulateCodes(m_workspace, GetDomainName("LeidingType"), Me.ComboBoxLeidingType)
        domainMgr = New CodedValueDomainManager(m_workspace, "LeidingType")
        domainMgr.PopulateCodes(Me.ComboBoxLeidingType)

        ' List domain code values in the status combo box.
        'PopulateCodes(m_workspace, GetDomainName("Status"), Me.ComboBoxStatus)
        domainMgr = New CodedValueDomainManager(m_workspace, "Status")
        domainMgr.PopulateCodes(Me.ComboBoxStatus)

        ' List domain code values in the ligging combo box.
        'PopulateCodes(m_workspace, GetDomainName("Ligging"), Me.ComboBoxLigging)
        domainMgr = New CodedValueDomainManager(m_workspace, "Ligging")
        domainMgr.PopulateCodes(Me.ComboBoxLigging)

        ' List domain code values in the hydranttype combo box.
        'PopulateCodes(m_workspace, GetDomainName("HydrantType"), Me.ComboBoxHydrantType)
        domainMgr = New CodedValueDomainManager(m_workspace, "HydrantType")
        domainMgr.PopulateCodes(Me.ComboBoxHydrantType)

        ' List domain code values in the bron combo box.
        'PopulateCodes(m_workspace, GetDomainName("Bron"), Me.ComboBoxBron)
        domainMgr = New CodedValueDomainManager(m_workspace, "Bron")
        domainMgr.PopulateCodes(Me.ComboBoxBron)

        'Disable editing & navigation controls.
        EnableNavigationControls(False)
        InitializeEditingControls()
        InitializeEditingLabels()

        'Disable filter combo boxes.
        Me.ComboBoxStatusFilter.Enabled = False
        Me.ComboBoxTypeFilter.Enabled = False

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Finalize current editing and reset all edit controls and labels.
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	11/10/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub ResetFormControls()

        'Check if attributes of currently loaded feature are modified.
        'If so, allow the user to store these changes before continuing.
        If ModifiedAttribute() Then
            If MsgBox(c_Message_SaveChanges, MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                If ValidateAttributeChanges() Then
                    StoreAttributeChanges()
                Else
                    Exit Sub
                End If
            End If
        End If

        'Re-initialize the form controls.
        m_editing = False
        EnableNavigationControls(False)
        InitializeEditingControls()
        InitializeEditingLabels()

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Set enabled/visible attributes of form controls for recordset navigation.
    ''' </summary>
    ''' <param name="value"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub EnableNavigationControls(ByVal value As Boolean)

        'Feature navigation button controls.
        Me.ButtonFirst.Enabled = value
        Me.ButtonPrevious.Enabled = value
        Me.ButtonNext.Enabled = value
        Me.ButtonLast.Enabled = value

        'Feature navigation label controls.
        Me.LabelCounter.Visible = value
        Me.LabelSeparator.Visible = value
        Me.LabelTotal.Visible = value

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Make form controls for attribute editing (un)available.
    ''' </summary>
    ''' <param name="value">
    '''     Boolean requested availability.
    ''' </param>
    ''' <remarks>
    '''     Controls will be set enabled or disabled, not only depending on the param,
    '''     but also depending on the functional requirements. Some attributes are read-only.
    '''     Forecolor &amp; backcolor are also adopted to create a visually attractive form.
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    '''     [Kristof Vydt]	05/10/2005  Replace ButtonConnect/Copy by CheckBoxConnect/Copy.
    ''' 	[Kristof Vydt]	11/10/2005	Implement color settings as private form declarations.
    '''     [Kristof Vydt]  22/02/2007  Eliminate on form legend code controls.
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub EnableEditingControls(ByVal value As Boolean)

        Dim ForeColor As Color
        Dim BackColor As Color

        'Tools button controls.
        ButtonSave.Enabled = value
        ButtonLabelAdd.Enabled = value
        ButtonLabelDel.Enabled = value
        CheckBoxConnect.Enabled = value
        CheckBoxCopy.Enabled = value

        'Set a color depending on the value.
        If value Then
            BackColor = EnabledEditControlBackColor
            ForeColor = EnabledEditControlForeColor
        Else
            BackColor = DisabledEditControlBackColor
            ForeColor = DisabledEditControlForeColor
        End If

        'Attribute edditing controls.

        '- Aanduiding
        TextBoxAanduiding.Enabled = value
        TextBoxAanduiding.ForeColor = ForeColor
        TextBoxAanduiding.BackColor = BackColor
        '- Begindatum
        DatePickerBegindatum.Enabled = False
        '- BrandweerNummer
        TextBoxBrandweerID.Enabled = value
        TextBoxBrandweerID.ForeColor = ForeColor
        TextBoxBrandweerID.BackColor = BackColor
        '- Bron
        ComboBoxBron.Enabled = value
        ComboBoxBron.ForeColor = ForeColor
        ComboBoxBron.BackColor = BackColor
        '- Diameter
        TextBoxDiameter.Enabled = value
        TextBoxDiameter.ForeColor = ForeColor
        TextBoxDiameter.BackColor = BackColor
        '- Einddatum
        CheckBoxEinddatum.Enabled = value
        DatePickerEinddatum.Enabled = value And CheckBoxEinddatum.Checked
        '- HydrantType
        ComboBoxHydrantType.Enabled = value
        ComboBoxHydrantType.ForeColor = ForeColor
        ComboBoxHydrantType.BackColor = BackColor
        '- LeidingNummer
        TextBoxLeidingID.Enabled = value
        TextBoxLeidingID.ForeColor = ForeColor
        TextBoxLeidingID.BackColor = BackColor
        '- LeidingType
        ComboBoxLeidingType.Enabled = value
        ComboBoxLeidingType.ForeColor = ForeColor
        ComboBoxLeidingType.BackColor = BackColor
        '- LeverancierNummer
        TextBoxLeverancierID.Enabled = False
        TextBoxLeverancierID.ForeColor = DisabledEditControlForeColor
        TextBoxLeverancierID.BackColor = DisabledEditControlBackColor
        '- Ligging
        ComboBoxLigging.Enabled = value
        ComboBoxLigging.ForeColor = ForeColor
        ComboBoxLigging.BackColor = BackColor
        '- Postcode
        TextBoxPostcode.Enabled = False
        TextBoxPostcode.ForeColor = DisabledEditControlForeColor
        TextBoxPostcode.BackColor = DisabledEditControlBackColor
        '- Status
        ComboBoxStatus.Enabled = value
        ComboBoxStatus.ForeColor = ForeColor
        ComboBoxStatus.BackColor = BackColor
        '- Straatcode
        TextBoxStraatcode.Enabled = False
        TextBoxStraatcode.ForeColor = DisabledEditControlForeColor
        TextBoxStraatcode.BackColor = DisabledEditControlBackColor
        '- Straatnaam
        TextBoxStraatnaam.Enabled = False
        TextBoxStraatnaam.ForeColor = DisabledEditControlForeColor
        TextBoxStraatnaam.BackColor = DisabledEditControlBackColor
        '- XCoord
        TextBoxXCoord.Enabled = False
        TextBoxXCoord.ForeColor = DisabledEditControlForeColor
        TextBoxXCoord.BackColor = DisabledEditControlBackColor
        '- YCoord
        TextBoxYCoord.Enabled = False
        TextBoxYCoord.ForeColor = DisabledEditControlForeColor
        TextBoxYCoord.BackColor = DisabledEditControlBackColor

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Set all editing control labels to their initial (not-edited) state.
    ''' </summary>
    ''' <remarks>
    '''     The color of the labels is used to track changes to the editing controls.
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	11/10/2005	Created
    '''     [Kristof Vydt]  22/02/2007  Eliminate on form legend code controls.
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub InitializeEditingLabels()

        LabelAanduiding.ForeColor = UnchangedLabelForeColor
        LabelBegindatum.ForeColor = UnchangedLabelForeColor
        LabelBrandweerID.ForeColor = UnchangedLabelForeColor
        LabelBron.ForeColor = UnchangedLabelForeColor
        LabelDiameter.ForeColor = UnchangedLabelForeColor
        LabelEinddatum.ForeColor = UnchangedLabelForeColor
        LabelHydrantType.ForeColor = UnchangedLabelForeColor
        LabelLeidingID.ForeColor = UnchangedLabelForeColor
        LabelLeidingType.ForeColor = UnchangedLabelForeColor
        LabelLeverancierID.ForeColor = UnchangedLabelForeColor
        LabelLigging.ForeColor = UnchangedLabelForeColor
        LabelPostcode.ForeColor = UnchangedLabelForeColor
        LabelStatus.ForeColor = UnchangedLabelForeColor
        LabelStraatcode.ForeColor = UnchangedLabelForeColor
        LabelStraatnaam.ForeColor = UnchangedLabelForeColor
        LabelXCoord.ForeColor = UnchangedLabelForeColor
        LabelYCoord.ForeColor = UnchangedLabelForeColor

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Disable and clear editing controls on the form.
    ''' </summary>
    ''' <remarks>
    '''     Labels of editing controls are not altered.
    '''     Buttons for feature-based functionality are also disabled.
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	11/10/2005	Created
    ''' 	[Kristof Vydt]	24/10/2005	Deactivate listeners
    '''     [Kristof Vydt]  22/02/2007  Eliminate on form legend code controls.
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub InitializeEditingControls()

        'Deactivate listeners.
        ConnectFeatureFunctionality_Deactivate()
        CopyAttributesFunctionality_Deactivate()
        CopyAddressFunctionality_Deactivate()

        'Tools button controls.
        ButtonSave.Enabled = False
        ButtonLabelAdd.Enabled = False
        ButtonLabelDel.Enabled = False
        CheckBoxConnect.Enabled = False
        CheckBoxCopy.Enabled = False

        'Attribute edditing controls.

        '- Aanduiding
        TextBoxAanduiding.Text = ""
        TextBoxAanduiding.Enabled = False
        TextBoxAanduiding.ForeColor = InitialEditControlForeColor
        TextBoxAanduiding.BackColor = InitialEditControlBackColor
        '- Begindatum
        DatePickerBegindatum.Text = ""
        DatePickerBegindatum.Enabled = False
        '- BrandweerNummer
        TextBoxBrandweerID.Text = ""
        TextBoxBrandweerID.Enabled = False
        TextBoxBrandweerID.ForeColor = InitialEditControlForeColor
        TextBoxBrandweerID.BackColor = InitialEditControlBackColor
        '- Bron
        ComboBoxBron.SelectedIndex = -1
        ComboBoxBron.Enabled = False
        ComboBoxBron.ForeColor = InitialEditControlForeColor
        ComboBoxBron.BackColor = InitialEditControlBackColor
        '- Diameter
        TextBoxDiameter.Text = ""
        TextBoxDiameter.Enabled = False
        TextBoxDiameter.ForeColor = InitialEditControlForeColor
        TextBoxDiameter.BackColor = InitialEditControlBackColor
        '- Einddatum
        CheckBoxEinddatum.Checked = False
        DatePickerEinddatum.Text = ""
        CheckBoxEinddatum.Enabled = False
        DatePickerEinddatum.Enabled = False
        '- HydrantType
        ComboBoxHydrantType.SelectedIndex = -1
        ComboBoxHydrantType.Enabled = False
        ComboBoxHydrantType.ForeColor = InitialEditControlForeColor
        ComboBoxHydrantType.BackColor = InitialEditControlBackColor
        '- LeidingNummer
        TextBoxLeidingID.Text = ""
        TextBoxLeidingID.Enabled = False
        TextBoxLeidingID.ForeColor = InitialEditControlForeColor
        TextBoxLeidingID.BackColor = InitialEditControlBackColor
        '- LeidingType
        ComboBoxLeidingType.SelectedIndex = -1
        ComboBoxLeidingType.Enabled = False
        ComboBoxLeidingType.ForeColor = InitialEditControlForeColor
        ComboBoxLeidingType.BackColor = InitialEditControlBackColor
        '- LeverancierNummer
        TextBoxLeverancierID.Text = ""
        TextBoxLeverancierID.Enabled = False
        TextBoxLeverancierID.ForeColor = InitialEditControlForeColor
        TextBoxLeverancierID.BackColor = InitialEditControlBackColor
        '- Ligging
        ComboBoxLigging.SelectedIndex = -1
        ComboBoxLigging.Enabled = False
        ComboBoxLigging.ForeColor = InitialEditControlForeColor
        ComboBoxLigging.BackColor = InitialEditControlBackColor
        '- Postcode
        TextBoxPostcode.Text = ""
        TextBoxPostcode.Enabled = False
        TextBoxPostcode.ForeColor = InitialEditControlForeColor
        TextBoxPostcode.BackColor = InitialEditControlBackColor
        '- Status
        ComboBoxStatus.SelectedIndex = -1
        ComboBoxStatus.Enabled = False
        ComboBoxStatus.ForeColor = InitialEditControlForeColor
        ComboBoxStatus.BackColor = InitialEditControlBackColor
        '- Straatcode
        TextBoxStraatcode.Text = ""
        TextBoxStraatcode.Enabled = False
        TextBoxStraatcode.ForeColor = InitialEditControlForeColor
        TextBoxStraatcode.BackColor = InitialEditControlBackColor
        '- Straatnaam
        TextBoxStraatnaam.Text = ""
        TextBoxStraatnaam.Enabled = False
        TextBoxStraatnaam.ForeColor = InitialEditControlForeColor
        TextBoxStraatnaam.BackColor = InitialEditControlBackColor
        '- XCoord
        TextBoxXCoord.Text = ""
        TextBoxXCoord.Enabled = False
        TextBoxXCoord.ForeColor = InitialEditControlForeColor
        TextBoxXCoord.BackColor = InitialEditControlBackColor
        '- YCoord
        TextBoxYCoord.Text = ""
        TextBoxYCoord.Enabled = False
        TextBoxYCoord.ForeColor = InitialEditControlForeColor
        TextBoxYCoord.BackColor = InitialEditControlBackColor

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Return true if at least one of the attribute editing controls is modified.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    '''     The color of the attribute label control indicates modified value.
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    ''' 	[Kristof Vydt]	23/09/2005	Check LeidingID &amp; LeverancierID for modifications.
    ''' 	[Kristof Vydt]	11/10/2005	Use color settings from private variables of the form.
    '''     [Kristof Vydt]  22/02/2007  Eliminate on form legend code controls.
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Function ModifiedAttribute() As Boolean

        If Me.LabelAanduiding.ForeColor.Equals(ChangedLabelForeColor) Or _
           Me.LabelBrandweerID.ForeColor.Equals(ChangedLabelForeColor) Or _
           Me.LabelBron.ForeColor.Equals(ChangedLabelForeColor) Or _
           Me.LabelDiameter.ForeColor.Equals(ChangedLabelForeColor) Or _
           Me.LabelEinddatum.ForeColor.Equals(ChangedLabelForeColor) Or _
           Me.LabelHydrantType.ForeColor.Equals(ChangedLabelForeColor) Or _
           Me.LabelLeidingID.ForeColor.Equals(ChangedLabelForeColor) Or _
           Me.LabelLeidingType.ForeColor.Equals(ChangedLabelForeColor) Or _
           Me.LabelLeverancierID.ForeColor.Equals(ChangedLabelForeColor) Or _
           Me.LabelLigging.ForeColor.Equals(ChangedLabelForeColor) Or _
           Me.LabelPostcode.ForeColor.Equals(ChangedLabelForeColor) Or _
           Me.LabelStatus.ForeColor.Equals(ChangedLabelForeColor) Or _
           Me.LabelStraatcode.ForeColor.Equals(ChangedLabelForeColor) Or _
           Me.LabelStraatnaam.ForeColor.Equals(ChangedLabelForeColor) Or _
           Me.LabelXCoord.ForeColor.Equals(ChangedLabelForeColor) Or _
           Me.LabelYCoord.ForeColor.Equals(ChangedLabelForeColor) Then

            ModifiedAttribute = True
        Else

            ModifiedAttribute = False
        End If

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Make all required checkes, before modified attributes are going to be saved.
    ''' </summary>
    ''' <returns>
    '''     Validity as boolean.
    ''' </returns>
    ''' <remarks>
    '''     A list of all violations is presented to the user, so he knows what to change.
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    ''' 	[Kristof Vydt]	24/10/2005	BrandweerID required &amp; unique when Actief &amp; Ondergronds.
    ''' 	[Kristof Vydt]	23/11/2005	Add criteria ObjectID&lt;&gt;m_OID when checking for unique BrandweerID.
    ''' 	[Kristof Vydt]	02/08/2006	Replace some global constant by a function that reads the value from config file.
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Function ValidateAttributeChanges() As Boolean

        Dim NumberOfViolations As Integer
        Dim ViolationsMessages As String() 'array of message strings, each holding a violation description

        Try

            ReDim ViolationsMessages(0)
            NumberOfViolations = 0

            Select Case GetComboBoxCodeValue(Me.ComboBoxStatus)

                Case "1" 'Status Actief

                    'Einddatum moet null zijn.
                    If Me.CheckBoxEinddatum.Checked Then
                        NumberOfViolations = NumberOfViolations + 1
                        ReDim Preserve ViolationsMessages(NumberOfViolations - 1)
                        ViolationsMessages(NumberOfViolations - 1) = c_Message_EinddatumIsNotEmpty
                    End If

                    'Straatnaam moet ingevuld zijn.
                    If Len(Trim(Me.TextBoxStraatnaam.Text)) = 0 Then
                        NumberOfViolations = NumberOfViolations + 1
                        ReDim Preserve ViolationsMessages(NumberOfViolations - 1)
                        ViolationsMessages(NumberOfViolations - 1) = c_Message_HydrantNotConnected
                    End If

                    'X & Y moet ingevuld zijn.
                    If (Len(Trim(Me.TextBoxXCoord.Text)) = 0) Or _
                       (Len(Trim(Me.TextBoxYCoord.Text)) = 0) Or _
                       (Me.TextBoxXCoord.Text = "0") Or _
                       (Me.TextBoxYCoord.Text = "0") Then
                        NumberOfViolations = NumberOfViolations + 1
                        ReDim Preserve ViolationsMessages(NumberOfViolations - 1)
                        ViolationsMessages(NumberOfViolations - 1) = c_Message_InvalidCoords
                    End If

                    'Bron moet ingevuld zijn.
                    If GetComboBoxCodeValue(ComboBoxBron) = "" Then
                        NumberOfViolations = NumberOfViolations + 1
                        ReDim Preserve ViolationsMessages(NumberOfViolations - 1)
                        ViolationsMessages(NumberOfViolations - 1) = c_Message_BronIsEmpty
                    End If

                    'LeverancierNummer moet ingevuld zijn. Dit is altijd het geval, 
                    'want waarde wordt automatisch ingevuld indien leeg bij laden.

                    'Begindatum moet ingevuld zijn. Dit is altijd het geval, 
                    'want waarde wordt automatisch ingevuld indien leeg bij laden.

                    Select Case GetComboBoxCodeValue(ComboBoxHydrantType)
                        Case "1" 'Hydranttype Ondergronds

                            'Bij ondergrondse, actieve hydrant: Ligging moet ingevuld zijn.
                            If GetComboBoxCodeValue(ComboBoxLigging) = "" Then
                                NumberOfViolations = NumberOfViolations + 1
                                ReDim Preserve ViolationsMessages(NumberOfViolations - 1)
                                ViolationsMessages(NumberOfViolations - 1) = c_Message_LiggingIsEmpty
                            End If

                            'Bij ondergrondse, actieve hydrant: Type leiding moet ingevuld zijn.
                            If GetComboBoxCodeValue(ComboBoxLeidingType) = "" Then
                                NumberOfViolations = NumberOfViolations + 1
                                ReDim Preserve ViolationsMessages(NumberOfViolations - 1)
                                ViolationsMessages(NumberOfViolations - 1) = c_Message_LeidingtypeIsEmpty
                            End If

                            'Bij ondergrondse, actieve hydrant: Brandweernummer moet ingevuld & uniek zijn.
                            If Len(Trim(TextBoxBrandweerID.Text)) = 0 Then
                                NumberOfViolations = NumberOfViolations + 1
                                ReDim Preserve ViolationsMessages(NumberOfViolations - 1)
                                ViolationsMessages(NumberOfViolations - 1) = c_Message_BrandweerNrIsEmpty
                            Else
                                'Check if the value is unique.
                                Dim pLayer As IFeatureLayer = CType(m_layer, IFeatureLayer)
                                Dim pQueryFilter As IQueryFilter = New QueryFilter
                                pQueryFilter.WhereClause = _
                                    "(" & GetAttributeName("Hydrant", "BrandweerNr") & "=" & CInt(TextBoxBrandweerID.Text) & ") AND " & _
                                    "(" & GetAttributeName("Hydrant", "Status") & "='1') AND " & _
                                    "(" & GetAttributeName("Hydrant", "HydrantType") & "='1') AND " & _
                                    "(ObjectID<>" & m_OID & ")"
                                Dim pFeatureCursor As IFeatureCursor = pLayer.FeatureClass.Search(pQueryFilter, True)
                                Dim pFeature As IFeature = pFeatureCursor.NextFeature
                                If Not pFeature Is Nothing Then
                                    NumberOfViolations = NumberOfViolations + 1
                                    ReDim Preserve ViolationsMessages(NumberOfViolations - 1)
                                    ViolationsMessages(NumberOfViolations - 1) = c_Message_BrandweerNrIsAlreadyInUse
                                End If
                                If Not pFeature Is Nothing Then Marshal.ReleaseComObject(pFeature)
                                If Not pFeatureCursor Is Nothing Then Marshal.ReleaseComObject(pFeatureCursor)
                                If Not pQueryFilter Is Nothing Then Marshal.ReleaseComObject(pQueryFilter)
                            End If

                            'Bij ondergrondse, actieve hydrant: Aanduiding moet ingevuld zijn.
                            If Len(Trim(TextBoxAanduiding.Text)) = 0 Then
                                NumberOfViolations = NumberOfViolations + 1
                                ReDim Preserve ViolationsMessages(NumberOfViolations - 1)
                                ViolationsMessages(NumberOfViolations - 1) = c_Message_AanduidingIsEmpty
                            End If

                            'Bij ondergrondse, actieve hydrant: Diameter moet ingevuld zijn.
                            If Len(Trim(TextBoxDiameter.Text)) = 0 Then
                                NumberOfViolations = NumberOfViolations + 1
                                ReDim Preserve ViolationsMessages(NumberOfViolations - 1)
                                ViolationsMessages(NumberOfViolations - 1) = c_Message_DiameterIsEmpty
                            End If

                            'Bij ondergrondse, actieve hydrant: moet label hebben.
                            Dim pAnnoLayer As IAnnotationLayer = CType(GetFeatureLayer(m_document.FocusMap, GetLayerName("HydrantAnno")), IAnnotationLayer)
                            If GetLinkedAnnotations(pAnnoLayer, GetAttributeName("HydrantAnno", "LinkID"), TextBoxLeverancierID.Text).Count < 1 Then
                                NumberOfViolations = NumberOfViolations + 1
                                ReDim Preserve ViolationsMessages(NumberOfViolations - 1)
                                ViolationsMessages(NumberOfViolations - 1) = c_Message_HydrantNotLabelled
                            End If

                    End Select

                Case "2", "4", "5" 'Status Niet bruikbaar, Defect, In Ontwerp

                    'X & Y moet ingevuld zijn.
                    If (Len(Trim(TextBoxXCoord.Text)) = 0) Or _
                       (Len(Trim(TextBoxYCoord.Text)) = 0) Or _
                       (TextBoxXCoord.Text = "0") Or _
                       (TextBoxYCoord.Text = "0") Then
                        NumberOfViolations = NumberOfViolations + 1
                        ReDim Preserve ViolationsMessages(NumberOfViolations - 1)
                        ViolationsMessages(NumberOfViolations - 1) = c_Message_InvalidCoords
                    End If

                    'LeverancierNummer moet ingevuld zijn. Dit is altijd het geval, 
                    'want waarde wordt automatisch ingevuld indien leeg bij laden.

                    'Begindatum moet ingevuld zijn. Dit is altijd het geval, 
                    'want waarde wordt automatisch ingevuld indien leeg bij laden.

                    'Straatnaam moet ingevuld zijn.
                    If Len(Trim(TextBoxStraatnaam.Text)) = 0 Then
                        NumberOfViolations = NumberOfViolations + 1
                        ReDim Preserve ViolationsMessages(NumberOfViolations - 1)
                        ViolationsMessages(NumberOfViolations - 1) = c_Message_HydrantNotConnected
                    End If

                    'Bron moet ingevuld zijn.
                    If GetComboBoxCodeValue(ComboBoxBron) = "" Then
                        NumberOfViolations = NumberOfViolations + 1
                        ReDim Preserve ViolationsMessages(NumberOfViolations - 1)
                        ViolationsMessages(NumberOfViolations - 1) = c_Message_BronIsEmpty
                    End If

                    'Einddatum moet null zijn
                    If Me.CheckBoxEinddatum.Checked Then
                        NumberOfViolations = NumberOfViolations + 1
                        ReDim Preserve ViolationsMessages(NumberOfViolations - 1)
                        ViolationsMessages(NumberOfViolations - 1) = c_Message_EinddatumIsNotEmpty
                    End If

                Case "3" 'Status Historiek

                    'Einddatum moet ingevuld zijn.
                    If Not Me.CheckBoxEinddatum.Checked Then
                        NumberOfViolations = NumberOfViolations + 1
                        ReDim Preserve ViolationsMessages(NumberOfViolations - 1)
                        ViolationsMessages(NumberOfViolations - 1) = c_Message_HistoricWithoutEinddatum
                    End If

                    'Bij historische hydrant: mag geen label hebben.
                    Dim pAnnoLayer As IAnnotationLayer = CType(GetFeatureLayer(m_document.FocusMap, GetLayerName("HydrantAnno")), IAnnotationLayer)
                    If GetLinkedAnnotations(pAnnoLayer, GetAttributeName("HydrantAnno", "LinkID"), TextBoxLeverancierID.Text).Count > 0 Then
                        NumberOfViolations = NumberOfViolations + 1
                        ReDim Preserve ViolationsMessages(NumberOfViolations - 1)
                        ViolationsMessages(NumberOfViolations - 1) = c_Message_HistoricHydrantLabelled
                    End If

                Case "6", "7", "8", "9", "10" 'Tijdelijke statussen.
                    'Status Nieuw, Verwijderd, Nakijken_c, Nakijken_a, Nakijken_ac

                    'Einddatum moet null zijn
                    If Me.CheckBoxEinddatum.Checked Then
                        NumberOfViolations = NumberOfViolations + 1
                        ReDim Preserve ViolationsMessages(NumberOfViolations - 1)
                        ViolationsMessages(NumberOfViolations - 1) = c_Message_EinddatumIsNotEmpty
                    End If

            End Select

            'If validation was not successfull: list violations to the user and abort this method.
            If NumberOfViolations = 0 Then
                ValidateAttributeChanges = True
                Exit Function
            Else
                ReDim Preserve ViolationsMessages(NumberOfViolations)
                ViolationsMessages(NumberOfViolations) = c_Message_CorrectBeforeContinue
                MsgBox(Concat(ViolationsMessages, vbNewLine), MsgBoxStyle.OKOnly, "Violations of attribute conditions")
                ValidateAttributeChanges = False
                Exit Function
            End If

        Catch ex As Exception
            Throw ex
        End Try

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Store current feature attribute modifications.
    ''' </summary>
    ''' <remarks>
    '''     Active edit session will be closed. User is asked if changes must be saved.
    '''     If copy attributes was used: ask for updating the sourc hydrant.
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    ''' 	[Kristof Vydt]	23/09/2005	Store modified LeidingID &amp; LeverancierID.
    ''' 	[Kristof Vydt]	26/09/2005	Throw an appl exception if feature not found.
    ''' 	                            Finalize edit session before saving changes.
    ''' 	[Kristof Vydt]	10/10/2005	Update annotations moved until after storing all attributes.
    ''' 	[Kristof Vydt]	11/10/2005	Use color settings from private variables of the form.
    ''' 	[Kristof Vydt]	24/10/2005	Introducing c_Message_HydrantChangesNotStored.
    '''     [Kristof Vydt]  23/11/2005  Support storing Null value for attribute BrandweerID.
    '''     [Kristof Vydt]  17/07/2006  Use global parameters in MsgBox, use try...catch and ReleaseComObject.
    '''                                 Support storing Null value for attribute Diameter.
    ''' 	[Kristof Vydt]	02/08/2006	Replace some global constant by a function that reads the value from config file.
    '''     [Kristof Vydt]  31/08/2006  Truncate multiline "aanduiding" if too long.
    '''     [Kristof Vydt]  22/02/2007  Eliminate on form legend code controls. Use ModuleHydrant.UpdateLegendCode() instead.
    '''     [Elton Manoku]  28/11/2008  If the overnamen attributen is active, then the verwijderde hydrant might be chosen to be
    ''' historic status and the einddatum is automatically set. If the einddatum is by accident smaller than begindatum (fault) it will be
    ''' set to equal.
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub StoreAttributeChanges()

        Dim annoLinkID As String 'identifier of the link between hydrant feature and hydrant label
        Dim attrIndex As Integer 'index of a feature attribute
        Dim attrMaxLength As Integer 'max length for a feature string attribute
        Dim attrStrValue As String 'value of a feature string attribute
        Dim pAnnoFeatLayer As IFeatureLayer = Nothing 'hydrant labels feature layer
        Dim pAnnoLayer As IAnnotationLayer = Nothing 'hydrant labels annotation layer
        Dim pEditor As IEditor2 = Nothing 'editor of edit session
        Dim pFeature As IFeature = Nothing 'the hydrant that is currently being edited
        Dim pFeatureClass As IFeatureClass = Nothing
        Dim pFeatureCursor As IFeatureCursor = Nothing
        Dim pParams As Hashtable 'hashtable (list of key-value pairs)
        Dim pQueryFilter As IQueryFilter = Nothing
        Dim pSourceFeature As IFeature = Nothing 'the hydrant from which the attributes where copied

        Try

            'Get the feature that is modified.
            pQueryFilter = New QueryFilter
            pQueryFilter.WhereClause = "OBJECTID = " & m_OID
            pFeatureClass = CType(m_layer, IFeatureLayer2).FeatureClass()
            pFeatureCursor = pFeatureClass.Search(pQueryFilter, True)
            pFeature = pFeatureCursor.NextFeature

            'Make sure that there is at least one feature.
            If pFeature Is Nothing Then
                Throw New ApplicationException(Replace(c_Message_HydrantChangesNotStored, "^0", CStr(m_OID)))
                Exit Sub
            End If

            'Close active edit session before continuing.
            pEditor = GetEditorReference(m_application)
            If pEditor.EditState = esriEditState.esriStateEditing Then

                'Does the user wants to save changes while closing the edit session?
                If MsgBox(c_Message_SaveEdits, vbYesNo, c_Title_SaveEdits) = MsgBoxResult.Yes Then

                    'Close the active edit session and save changes.
                    EditSessionSave(pEditor)

                Else

                    'Close the active edit session without saving changes.
                    EditSessionAbort(pEditor)

                End If
            End If

            'Modify current feature attributes.
            '- Aanduiding
            If Me.LabelAanduiding.ForeColor.Equals(ChangedLabelForeColor) Then
                attrIndex = pFeature.Fields.FindField(GetAttributeName("Hydrant", "Aanduiding"))
                'Compare the length of the input value with field length.
                attrMaxLength = pFeature.Fields.Field(attrIndex).Length
                attrStrValue = CStr(Me.TextBoxAanduiding.Text)
                If attrMaxLength < attrStrValue.Length Then
                    'Truncate to max length.
                    attrStrValue = attrStrValue.Substring(0, attrMaxLength)
                End If
                'Store as new attribute value.
                pFeature.Value(attrIndex) = attrStrValue
            End If
            'Changes to "aanduiding" should result in update of annotations.
            'This is done at the end, to be sure that the value of LinkID is already stored.

            '- Begindatum
            If Me.LabelBegindatum.ForeColor.Equals(ChangedLabelForeColor) Then
                attrIndex = pFeature.Fields.FindField(GetAttributeName("Hydrant", "BeginDatum"))
                'If the checkbox is checked, set begindatum attribute to new date.
                pFeature.Value(attrIndex) = CDate(Me.DatePickerBegindatum.Text)
            End If

            '- BrandweerID
            If Me.LabelBrandweerID.ForeColor.Equals(ChangedLabelForeColor) Then
                attrIndex = pFeature.Fields.FindField(GetAttributeName("Hydrant", "BrandweerNr"))
                If IsNumeric(Me.TextBoxBrandweerID.Text) Then
                    'If a number was filled in.
                    pFeature.Value(attrIndex) = CInt(Me.TextBoxBrandweerID.Text)
                Else
                    'If nothing or text was filled in.
                    pFeature.Value(attrIndex) = System.DBNull.Value
                End If
            End If

            '- Bron
            If Me.LabelBron.ForeColor.Equals(ChangedLabelForeColor) Then _
                pFeature.Value(pFeature.Fields.FindField(GetAttributeName("Hydrant", "Bron"))) = GetComboBoxCodeValue(Me.ComboBoxBron)

            '- Diameter
            If Me.LabelDiameter.ForeColor.Equals(ChangedLabelForeColor) Then
                attrIndex = pFeature.Fields.FindField(GetAttributeName("Hydrant", "Diameter"))
                If IsNumeric(Me.TextBoxDiameter.Text) Then
                    'If a number was filled in.
                    pFeature.Value(attrIndex) = CInt(Me.TextBoxDiameter.Text)
                Else
                    'If nothing or text was filled in.
                    pFeature.Value(attrIndex) = System.DBNull.Value
                End If
            End If

            '- Einddatum
            If Me.LabelEinddatum.ForeColor.Equals(ChangedLabelForeColor) Then
                attrIndex = pFeature.Fields.FindField(GetAttributeName("Hydrant", "EindDatum"))
                If Me.CheckBoxEinddatum.Checked Then
                    'If the checkbox is checked, set einddatum attribute to new date.
                    pFeature.Value(attrIndex) = CDate(Me.DatePickerEinddatum.Text)
                Else
                    'If the checkbox is not checked, einddatum attribute should have value null.
                    pFeature.Value(attrIndex) = System.DBNull.Value
                End If
            End If

            '- HydrantType
            If Me.LabelHydrantType.ForeColor.Equals(ChangedLabelForeColor) Then _
                pFeature.Value(pFeature.Fields.FindField(GetAttributeName("Hydrant", "HydrantType"))) = GetComboBoxCodeValue(Me.ComboBoxHydrantType)

            '- LeidingNummer
            If Me.LabelLeidingID.ForeColor.Equals(ChangedLabelForeColor) Then _
                pFeature.Value(pFeature.Fields.FindField(GetAttributeName("Hydrant", "LeidingNr"))) = CStr(Me.TextBoxLeidingID.Text)

            '- LeidingType
            If Me.LabelLeidingType.ForeColor.Equals(ChangedLabelForeColor) Then _
                pFeature.Value(pFeature.Fields.FindField(GetAttributeName("Hydrant", "LeidingType"))) = GetComboBoxCodeValue(Me.ComboBoxLeidingType)

            '- LeverancierNummer
            If Me.LabelLeverancierID.ForeColor.Equals(ChangedLabelForeColor) Then _
                pFeature.Value(pFeature.Fields.FindField(GetAttributeName("Hydrant", "LeverancierNr"))) = CStr(Me.TextBoxLeverancierID.Text)

            '- Ligging
            If Me.LabelLigging.ForeColor.Equals(ChangedLabelForeColor) Then _
                pFeature.Value(pFeature.Fields.FindField(GetAttributeName("Hydrant", "Ligging"))) = GetComboBoxCodeValue(Me.ComboBoxLigging)

            '- Postcode
            If Me.LabelPostcode.ForeColor.Equals(ChangedLabelForeColor) Then _
                pFeature.Value(pFeature.Fields.FindField(GetAttributeName("Hydrant", "Postcode"))) = CStr(Me.TextBoxPostcode.Text)

            '- Status
            If Me.LabelStatus.ForeColor.Equals(ChangedLabelForeColor) Then
                Dim newStatus As String = GetComboBoxCodeValue(Me.ComboBoxStatus)
                pFeature.Value(pFeature.Fields.FindField(GetAttributeName("Hydrant", "Status"))) = newStatus
            End If

            '- Straatcode
            If Me.LabelStraatcode.ForeColor.Equals(ChangedLabelForeColor) Then _
                pFeature.Value(pFeature.Fields.FindField(GetAttributeName("Hydrant", "Straatcode"))) = CStr(Me.TextBoxStraatcode.Text)

            '- Straatnaam
            If Me.LabelStraatnaam.ForeColor.Equals(ChangedLabelForeColor) Then _
                pFeature.Value(pFeature.Fields.FindField(GetAttributeName("Hydrant", "Straatnaam"))) = CStr(Me.TextBoxStraatnaam.Text)

            '- XCoord
            If Me.LabelXCoord.ForeColor.Equals(ChangedLabelForeColor) Then _
                pFeature.Value(pFeature.Fields.FindField(GetAttributeName("Hydrant", "CoordX"))) = CStr(Me.TextBoxXCoord.Text)

            '- YCoord
            If Me.LabelYCoord.ForeColor.Equals(ChangedLabelForeColor) Then _
                pFeature.Value(pFeature.Fields.FindField(GetAttributeName("Hydrant", "CoordY"))) = CStr(Me.TextBoxYCoord.Text)

            '- Legend code
            Call ModuleHydrant.UpdateLegendCode(pFeature)

            '- Annotations
            If Me.LabelAanduiding.ForeColor.Equals(ChangedLabelForeColor) Then

                'Changes to "aanduiding" should result in update of annotations.
                'Modify existing related annotation features.
                annoLinkID = CStr(pFeature.Value(pFeature.Fields.FindField(GetAttributeName("Hydrant", "LinkID"))))
                pAnnoLayer = CType(GetFeatureLayer(m_document.FocusMap, GetLayerName("HydrantAnno")), IAnnotationLayer)
                If pAnnoLayer Is Nothing Then Throw New LayerNotFoundException(GetLayerName("HydrantAnno"))
                pParams = New Hashtable
                pParams.Add("TextString", Me.TextBoxAanduiding.Text)
                UpdateAnno(pAnnoLayer, pParams, GetAttributeName("HydrantAnno", "LinkID"), annoLinkID)

            End If

            'Commit changes.
            pFeature.Store()

            'Reload current feature into the form,
            'without resetting the CopyAttributes feature reference,
            'and without rezooming to the current feature.
            LoadFeature(m_OID, False, False)

            'In case the CopyAttributes functionality has been used:
            'Modify source feature after saving changes to destination feature.
            If Not m_copyFrom Is Nothing Then

                'Usk user if source hydrant must be updated.
                If MsgBox(c_Message_ModifySourceHydrant, MsgBoxStyle.YesNo, c_Title_CopyAttributes) = MsgBoxResult.Yes Then

                    'Transfer annotations from the source hydrant (copy from) to the destination hydrant (copy to).
                    pAnnoFeatLayer = GetFeatureLayer(m_document.FocusMap, GetLayerName("HydrantAnno"))
                    If pAnnoFeatLayer Is Nothing Then Throw New LayerNotFoundException(GetLayerName("HydrantAnno"))
                    Dim hydrantAnnoLayer As IAnnotationLayer = CType(pAnnoFeatLayer, IAnnotationLayer)
                    Dim fieldIndex As Integer = m_copyFrom.Fields.FindField(GetAttributeName("Hydrant", "LinkID"))
                    Dim oldLinkID As String = Trim(CStr(m_copyFrom.Value(fieldIndex)))
                    Dim newLinkID As String = Trim(TextBoxLeverancierID.Text)
                    RelinkAnnotations(hydrantAnnoLayer, oldLinkID, newLinkID)

                    'Change source hydrant that was used to copy from, to status "historiek".
                    pParams = New Hashtable
                    'RW:2008 The eid datum is checked if it is smaller than the begindatum. If Yes the einddatum is the same as the begindatum
                    fieldIndex = m_copyFrom.Fields.FindField(GetAttributeName("Hydrant", "BeginDatum"))
                    Dim beginDatum As Date = CDate(m_copyFrom.Value(fieldIndex))
                    '- Einddatum
                    Dim EindDatum As Date = CDate(DatePickerBegindatum.Text).AddDays(-1)
                    If beginDatum > EindDatum Then
                        EindDatum = beginDatum
                    End If

                    pParams.Add("EindDatum", EindDatum)
                    '- Status
                    pParams.Add("Status", CStr(3))   'status=historiek
                    '- LegendCode is updated by ModifyHydrantAttributes.
                    Dim success As Boolean = ModuleHydrant.ModifyHydrantAttributes(m_copyFrom, pParams)

                End If

                'Clear the reference to the "CopyFrom" feature.
                SetCopyFrom(Nothing)

            End If

        Catch ex As Exception

            Throw ex

        Finally

            'Refresh hydrants symbology.
            Dim pActiveView As IActiveView = m_document.ActivatedView()
            pActiveView.PartialRefresh(esriViewDrawPhase.esriViewGeography, m_layer, pActiveView.Extent)

            'Be sure to release as much objects as possible.
            If Not pQueryFilter Is Nothing Then Marshal.ReleaseComObject(pQueryFilter)
            If Not pFeatureClass Is Nothing Then Marshal.ReleaseComObject(pFeatureClass)
            If Not pFeatureCursor Is Nothing Then Marshal.ReleaseComObject(pFeatureCursor)
            If Not pFeature Is Nothing Then Marshal.ReleaseComObject(pFeature)
            If Not pEditor Is Nothing Then Marshal.ReleaseComObject(pEditor)
            If Not pSourceFeature Is Nothing Then Marshal.ReleaseComObject(pSourceFeature)
            If Not pAnnoFeatLayer Is Nothing Then Marshal.ReleaseComObject(pAnnoFeatLayer)
            If Not pAnnoLayer Is Nothing Then Marshal.ReleaseComObject(pAnnoLayer)

        End Try

    End Sub

    '''' -----------------------------------------------------------------------------
    '''' <summary>
    ''''     Update the hidden LegendCode attribute.
    '''' </summary>
    '''' <remarks>
    '''' </remarks>
    '''' <history>
    '''' 	[Kristof Vydt]	16/09/2005	Created
    ''''     [Kristof Vydt]  22/02/2007  deprecated. Elimination of on form legend code controls.
    '''' </history>
    '''' -----------------------------------------------------------------------------
    'Private Sub UpdateLegendCode()
    '    Try
    '        Dim status As String = GetComboBoxCodeValue(ComboBoxStatus)
    '        Dim hydrantType As String = GetComboBoxCodeValue(ComboBoxHydrantType)
    '        Dim ligging As String = GetComboBoxCodeValue(ComboBoxLigging)
    '        Dim diameter As Integer = CInt(IIf(TextBoxDiameter.Text = "", 0, TextBoxDiameter.Text))
    '        Dim legend As String = CStr(HydrantLegendCodeEx(status, hydrantType, ligging, diameter))

    '        MarkAsChanged(LabelLegende)
    '        TextBoxLegende.Text = legend
    '    Catch ex As Exception
    '        Throw ex
    '    End Try
    'End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Store a pointer to the feature that was copied from,
    '''     in order to use this reference when changes to current feature are saved.
    ''' </summary>
    ''' <param name="pFeature">
    '''     The feature to remember.
    ''' </param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    <CLSCompliant(False)> _
    Public Sub SetCopyFrom(ByVal pFeature As IFeature)
        Try
            m_copyFrom = pFeature
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Re-evaluate the enabled property of the control ButtonLoad
    '''     in function of selected filter type and filter value.
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	27/10/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub UpdateButtonLoadAvailability()
        Try
            If Me.RadioButtonStatusFilter.Checked Then Me.ButtonLoad.Enabled = (Me.ComboBoxStatusFilter.SelectedIndex > -1)
            If Me.RadioButtonTypeFilter.Checked Then Me.ButtonLoad.Enabled = (Me.ComboBoxTypeFilter.SelectedIndex > -1)
            If Me.RadioButtonMapSelection.Checked Then Me.ButtonLoad.Enabled = True
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

#End Region

End Class
