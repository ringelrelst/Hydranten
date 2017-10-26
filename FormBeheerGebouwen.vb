Option Explicit On 
Option Strict On

#Region " Imports namespaces "
Imports System.Drawing
Imports System.Drawing.Color
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
''' Class	 : Hydranten.BeheerHydranten.FormBeheerGebouwen
''' 
''' -----------------------------------------------------------------------------
''' <summary>
'''     Form for managing buildings.
''' </summary>
''' <remarks>
''' </remarks>
''' <history>
''' 	[Kristof Vydt]	16/09/2005	Created
''' 	[Kristof Vydt]	23/09/2005	Remove event handler(s) on close.
''' 	[Kristof Vydt]	05/10/2005	Replace ButtonConnect/Copy by CheckBoxConnect/Copy.
''' 	                        	Hide Straatcode.
''' 	[Kristof Vydt]	21/10/2005	Adjust ButtonLabelAdd_Click to use the new FormAddAnnotation.
''' 	[Kristof Vydt]	24/10/2005  Deactivate listeners when loading feature.
''' 	[Kristof Vydt]	27/10/2005	Set DropDownStyle of every ComboBox to List to force the user to select from the list.
''' 	                            FormBeheerGebouwen_Closed added.
'''  	[Kristof Vydt]	23/11/2005	Add optional zoomToFeature parameter to LoadFeature method.
'''                                 Start edit session before activating the Edit Annotation Tool command.
'''  	[Kristof Vydt]	17/07/2006	Close active edit session in StoreAttributeChanges before modifying feature attributes.
'''                                 Changes to "Volgnummer" should result in update of existing related annotation features.
''' 	[Kristof Vydt]	02/08/2006	Replace some global constant by a function that reads the value from config file.
''' 	[Kristof Vydt]	11/08/2006	Handle DBNull input values in StoreAttributeChanges.
''' 	[Kristof Vydt]	18/08/2006	Eliminate private marker element.
''' 	[Kristof Vydt]	22/02/2007	Adopt to XML configuration.
''' 	[Kristof Vydt]	14/03/2007	Add GebouwType attribute.
''' 	[Kristof Vydt]	22/03/2007	Use the new CodedValueDomainManager instead of the deprecated ModuleDomainAccess.
'''     [Koen Vermeer]  23/06/2008  Conversion to ArcGIS 9.2
''' </history>
''' -----------------------------------------------------------------------------
Public NotInheritable Class FormBeheerGebouwen
    Inherits System.Windows.Forms.Form

#Region " Private variables "

    Private m_application As IMxApplication 'hold current ArcMap application
    Private m_document As IMxDocument 'hold current ArcMap document
    'Private m_layer As ILayer 'hydranten layer
    'Private m_workspace As IWorkspace 'workspace of the hydranten
    'Private m_marker As IMarkerElement 'marker for current feature
    Private m_editing As Boolean 'indicated if form is ready for editing
    'Private m_selectionSet As ISelectionSet
    Private m_enumOIDs As IEnumIDs 'enumeration of the feature IDs of the edit set
    Private m_OID As Integer 'the ObjectID of the current editable feature
    'Private m_copyFrom As IFeature = Nothing 'when functionality "copy attributes from hydrant" is used

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
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents ButtonLabelDel As System.Windows.Forms.Button
    Friend WithEvents ButtonLabelAdd As System.Windows.Forms.Button
    Friend WithEvents ButtonSave As System.Windows.Forms.Button
    Friend WithEvents TextBoxStraatnaam As System.Windows.Forms.TextBox
    Friend WithEvents LabelStraatnaam As System.Windows.Forms.Label
    Friend WithEvents TextBoxAanduiding As System.Windows.Forms.TextBox
    Friend WithEvents LabelAanduiding As System.Windows.Forms.Label
    Friend WithEvents ButtonClose As System.Windows.Forms.Button
    Friend WithEvents LabelNaam As System.Windows.Forms.Label
    Friend WithEvents TextBoxVolgNr As System.Windows.Forms.TextBox
    Friend WithEvents LabelVolgNr As System.Windows.Forms.Label
    Friend WithEvents TextBoxInfo As System.Windows.Forms.TextBox
    Friend WithEvents GroupBoxDossier As System.Windows.Forms.GroupBox
    Friend WithEvents TextBoxDossierNrNieuw As System.Windows.Forms.TextBox
    Friend WithEvents LabelDossierNrNieuw As System.Windows.Forms.Label
    Friend WithEvents TextBoxDossierNrOud As System.Windows.Forms.TextBox
    Friend WithEvents LabelDossierNrOud As System.Windows.Forms.Label
    Friend WithEvents TextBoxDossierInterventie As System.Windows.Forms.TextBox
    Friend WithEvents LabelDossierInterventie As System.Windows.Forms.Label
    Friend WithEvents TextBoxHuisnr As System.Windows.Forms.TextBox
    Friend WithEvents LabelHuisnr As System.Windows.Forms.Label
    Friend WithEvents TextBoxPostcode As System.Windows.Forms.TextBox
    Friend WithEvents LabelPostcode As System.Windows.Forms.Label
    Friend WithEvents TextBoxNaam As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxStraatcode As System.Windows.Forms.TextBox
    Friend WithEvents LabelStraatcode As System.Windows.Forms.Label
    Friend WithEvents LabelInfo As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ButtonLoad As System.Windows.Forms.Button
    Friend WithEvents ComboBoxLayerFilter As System.Windows.Forms.ComboBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents LabelCounter As System.Windows.Forms.Label
    Friend WithEvents LabelTotal As System.Windows.Forms.Label
    Friend WithEvents LabelSeparator As System.Windows.Forms.Label
    Friend WithEvents ButtonNext As System.Windows.Forms.Button
    Friend WithEvents ButtonLast As System.Windows.Forms.Button
    Friend WithEvents ButtonFirst As System.Windows.Forms.Button
    Friend WithEvents ButtonPrevious As System.Windows.Forms.Button
    Friend WithEvents CheckBoxConnect As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBoxCopy As System.Windows.Forms.CheckBox
    Friend WithEvents ComboBoxGebouwType As System.Windows.Forms.ComboBox
    Friend WithEvents LabelGebouwType As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.CheckBoxCopy = New System.Windows.Forms.CheckBox
        Me.CheckBoxConnect = New System.Windows.Forms.CheckBox
        Me.LabelInfo = New System.Windows.Forms.Label
        Me.GroupBoxDossier = New System.Windows.Forms.GroupBox
        Me.TextBoxDossierNrNieuw = New System.Windows.Forms.TextBox
        Me.LabelDossierNrNieuw = New System.Windows.Forms.Label
        Me.TextBoxDossierNrOud = New System.Windows.Forms.TextBox
        Me.LabelDossierNrOud = New System.Windows.Forms.Label
        Me.TextBoxDossierInterventie = New System.Windows.Forms.TextBox
        Me.LabelDossierInterventie = New System.Windows.Forms.Label
        Me.TextBoxHuisnr = New System.Windows.Forms.TextBox
        Me.LabelHuisnr = New System.Windows.Forms.Label
        Me.TextBoxPostcode = New System.Windows.Forms.TextBox
        Me.LabelPostcode = New System.Windows.Forms.Label
        Me.ButtonLabelDel = New System.Windows.Forms.Button
        Me.ButtonLabelAdd = New System.Windows.Forms.Button
        Me.ButtonSave = New System.Windows.Forms.Button
        Me.TextBoxStraatnaam = New System.Windows.Forms.TextBox
        Me.LabelStraatnaam = New System.Windows.Forms.Label
        Me.TextBoxNaam = New System.Windows.Forms.TextBox
        Me.LabelNaam = New System.Windows.Forms.Label
        Me.TextBoxStraatcode = New System.Windows.Forms.TextBox
        Me.LabelStraatcode = New System.Windows.Forms.Label
        Me.TextBoxVolgNr = New System.Windows.Forms.TextBox
        Me.LabelVolgNr = New System.Windows.Forms.Label
        Me.TextBoxAanduiding = New System.Windows.Forms.TextBox
        Me.LabelAanduiding = New System.Windows.Forms.Label
        Me.TextBoxInfo = New System.Windows.Forms.TextBox
        Me.ButtonClose = New System.Windows.Forms.Button
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.ButtonLoad = New System.Windows.Forms.Button
        Me.ComboBoxLayerFilter = New System.Windows.Forms.ComboBox
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.LabelCounter = New System.Windows.Forms.Label
        Me.LabelTotal = New System.Windows.Forms.Label
        Me.LabelSeparator = New System.Windows.Forms.Label
        Me.ButtonNext = New System.Windows.Forms.Button
        Me.ButtonLast = New System.Windows.Forms.Button
        Me.ButtonFirst = New System.Windows.Forms.Button
        Me.ButtonPrevious = New System.Windows.Forms.Button
        Me.LabelGebouwType = New System.Windows.Forms.Label
        Me.ComboBoxGebouwType = New System.Windows.Forms.ComboBox
        Me.GroupBox2.SuspendLayout()
        Me.GroupBoxDossier.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox2
        '
        Me.GroupBox2.BackColor = System.Drawing.SystemColors.Control
        Me.GroupBox2.Controls.Add(Me.ComboBoxGebouwType)
        Me.GroupBox2.Controls.Add(Me.LabelGebouwType)
        Me.GroupBox2.Controls.Add(Me.CheckBoxCopy)
        Me.GroupBox2.Controls.Add(Me.CheckBoxConnect)
        Me.GroupBox2.Controls.Add(Me.LabelInfo)
        Me.GroupBox2.Controls.Add(Me.GroupBoxDossier)
        Me.GroupBox2.Controls.Add(Me.TextBoxHuisnr)
        Me.GroupBox2.Controls.Add(Me.LabelHuisnr)
        Me.GroupBox2.Controls.Add(Me.TextBoxPostcode)
        Me.GroupBox2.Controls.Add(Me.LabelPostcode)
        Me.GroupBox2.Controls.Add(Me.ButtonLabelDel)
        Me.GroupBox2.Controls.Add(Me.ButtonLabelAdd)
        Me.GroupBox2.Controls.Add(Me.ButtonSave)
        Me.GroupBox2.Controls.Add(Me.TextBoxStraatnaam)
        Me.GroupBox2.Controls.Add(Me.LabelStraatnaam)
        Me.GroupBox2.Controls.Add(Me.TextBoxNaam)
        Me.GroupBox2.Controls.Add(Me.LabelNaam)
        Me.GroupBox2.Controls.Add(Me.TextBoxStraatcode)
        Me.GroupBox2.Controls.Add(Me.LabelStraatcode)
        Me.GroupBox2.Controls.Add(Me.TextBoxVolgNr)
        Me.GroupBox2.Controls.Add(Me.LabelVolgNr)
        Me.GroupBox2.Controls.Add(Me.TextBoxAanduiding)
        Me.GroupBox2.Controls.Add(Me.LabelAanduiding)
        Me.GroupBox2.Controls.Add(Me.TextBoxInfo)
        Me.GroupBox2.Location = New System.Drawing.Point(8, 96)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(272, 432)
        Me.GroupBox2.TabIndex = 52
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Feature"
        '
        'CheckBoxCopy
        '
        Me.CheckBoxCopy.Appearance = System.Windows.Forms.Appearance.Button
        Me.CheckBoxCopy.Enabled = False
        Me.CheckBoxCopy.Location = New System.Drawing.Point(140, 368)
        Me.CheckBoxCopy.Name = "CheckBoxCopy"
        Me.CheckBoxCopy.Size = New System.Drawing.Size(124, 24)
        Me.CheckBoxCopy.TabIndex = 110
        Me.CheckBoxCopy.Text = "Adres overnemen"
        Me.CheckBoxCopy.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'CheckBoxConnect
        '
        Me.CheckBoxConnect.Appearance = System.Windows.Forms.Appearance.Button
        Me.CheckBoxConnect.Enabled = False
        Me.CheckBoxConnect.Location = New System.Drawing.Point(8, 368)
        Me.CheckBoxConnect.Name = "CheckBoxConnect"
        Me.CheckBoxConnect.Size = New System.Drawing.Size(125, 24)
        Me.CheckBoxConnect.TabIndex = 109
        Me.CheckBoxConnect.Text = "Connecteren"
        Me.CheckBoxConnect.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'LabelInfo
        '
        Me.LabelInfo.Location = New System.Drawing.Point(8, 272)
        Me.LabelInfo.Name = "LabelInfo"
        Me.LabelInfo.Size = New System.Drawing.Size(88, 16)
        Me.LabelInfo.TabIndex = 108
        Me.LabelInfo.Text = "Info"
        Me.LabelInfo.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'GroupBoxDossier
        '
        Me.GroupBoxDossier.Controls.Add(Me.TextBoxDossierNrNieuw)
        Me.GroupBoxDossier.Controls.Add(Me.LabelDossierNrNieuw)
        Me.GroupBoxDossier.Controls.Add(Me.TextBoxDossierNrOud)
        Me.GroupBoxDossier.Controls.Add(Me.LabelDossierNrOud)
        Me.GroupBoxDossier.Controls.Add(Me.TextBoxDossierInterventie)
        Me.GroupBoxDossier.Controls.Add(Me.LabelDossierInterventie)
        Me.GroupBoxDossier.Location = New System.Drawing.Point(8, 184)
        Me.GroupBoxDossier.Name = "GroupBoxDossier"
        Me.GroupBoxDossier.Size = New System.Drawing.Size(256, 88)
        Me.GroupBoxDossier.TabIndex = 107
        Me.GroupBoxDossier.TabStop = False
        Me.GroupBoxDossier.Text = "Dossier"
        '
        'TextBoxDossierNrNieuw
        '
        Me.TextBoxDossierNrNieuw.BackColor = System.Drawing.SystemColors.Control
        Me.TextBoxDossierNrNieuw.Enabled = False
        Me.TextBoxDossierNrNieuw.ForeColor = System.Drawing.SystemColors.InactiveBorder
        Me.TextBoxDossierNrNieuw.Location = New System.Drawing.Point(88, 40)
        Me.TextBoxDossierNrNieuw.Name = "TextBoxDossierNrNieuw"
        Me.TextBoxDossierNrNieuw.Size = New System.Drawing.Size(160, 20)
        Me.TextBoxDossierNrNieuw.TabIndex = 110
        Me.TextBoxDossierNrNieuw.Text = ""
        '
        'LabelDossierNrNieuw
        '
        Me.LabelDossierNrNieuw.Location = New System.Drawing.Point(8, 40)
        Me.LabelDossierNrNieuw.Name = "LabelDossierNrNieuw"
        Me.LabelDossierNrNieuw.Size = New System.Drawing.Size(80, 16)
        Me.LabelDossierNrNieuw.TabIndex = 109
        Me.LabelDossierNrNieuw.Text = "Nieuw nummer"
        Me.LabelDossierNrNieuw.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'TextBoxDossierNrOud
        '
        Me.TextBoxDossierNrOud.BackColor = System.Drawing.SystemColors.Control
        Me.TextBoxDossierNrOud.Enabled = False
        Me.TextBoxDossierNrOud.ForeColor = System.Drawing.SystemColors.InactiveBorder
        Me.TextBoxDossierNrOud.Location = New System.Drawing.Point(88, 16)
        Me.TextBoxDossierNrOud.Name = "TextBoxDossierNrOud"
        Me.TextBoxDossierNrOud.Size = New System.Drawing.Size(160, 20)
        Me.TextBoxDossierNrOud.TabIndex = 108
        Me.TextBoxDossierNrOud.Text = ""
        '
        'LabelDossierNrOud
        '
        Me.LabelDossierNrOud.Location = New System.Drawing.Point(8, 16)
        Me.LabelDossierNrOud.Name = "LabelDossierNrOud"
        Me.LabelDossierNrOud.Size = New System.Drawing.Size(80, 16)
        Me.LabelDossierNrOud.TabIndex = 107
        Me.LabelDossierNrOud.Text = "Oud nummer"
        Me.LabelDossierNrOud.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'TextBoxDossierInterventie
        '
        Me.TextBoxDossierInterventie.BackColor = System.Drawing.SystemColors.Control
        Me.TextBoxDossierInterventie.Enabled = False
        Me.TextBoxDossierInterventie.ForeColor = System.Drawing.SystemColors.InactiveBorder
        Me.TextBoxDossierInterventie.Location = New System.Drawing.Point(88, 64)
        Me.TextBoxDossierInterventie.Name = "TextBoxDossierInterventie"
        Me.TextBoxDossierInterventie.Size = New System.Drawing.Size(160, 20)
        Me.TextBoxDossierInterventie.TabIndex = 106
        Me.TextBoxDossierInterventie.Text = ""
        '
        'LabelDossierInterventie
        '
        Me.LabelDossierInterventie.Location = New System.Drawing.Point(8, 64)
        Me.LabelDossierInterventie.Name = "LabelDossierInterventie"
        Me.LabelDossierInterventie.Size = New System.Drawing.Size(80, 16)
        Me.LabelDossierInterventie.TabIndex = 105
        Me.LabelDossierInterventie.Text = "Interventie"
        Me.LabelDossierInterventie.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'TextBoxHuisnr
        '
        Me.TextBoxHuisnr.BackColor = System.Drawing.SystemColors.Control
        Me.TextBoxHuisnr.Enabled = False
        Me.TextBoxHuisnr.ForeColor = System.Drawing.SystemColors.InactiveBorder
        Me.TextBoxHuisnr.Location = New System.Drawing.Point(72, 112)
        Me.TextBoxHuisnr.Name = "TextBoxHuisnr"
        Me.TextBoxHuisnr.Size = New System.Drawing.Size(192, 20)
        Me.TextBoxHuisnr.TabIndex = 96
        Me.TextBoxHuisnr.Text = ""
        '
        'LabelHuisnr
        '
        Me.LabelHuisnr.Location = New System.Drawing.Point(8, 112)
        Me.LabelHuisnr.Name = "LabelHuisnr"
        Me.LabelHuisnr.Size = New System.Drawing.Size(40, 16)
        Me.LabelHuisnr.TabIndex = 95
        Me.LabelHuisnr.Text = "Huisnr"
        Me.LabelHuisnr.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'TextBoxPostcode
        '
        Me.TextBoxPostcode.BackColor = System.Drawing.SystemColors.Control
        Me.TextBoxPostcode.Enabled = False
        Me.TextBoxPostcode.ForeColor = System.Drawing.SystemColors.InactiveBorder
        Me.TextBoxPostcode.Location = New System.Drawing.Point(72, 136)
        Me.TextBoxPostcode.Name = "TextBoxPostcode"
        Me.TextBoxPostcode.Size = New System.Drawing.Size(64, 20)
        Me.TextBoxPostcode.TabIndex = 94
        Me.TextBoxPostcode.Text = ""
        '
        'LabelPostcode
        '
        Me.LabelPostcode.Location = New System.Drawing.Point(8, 136)
        Me.LabelPostcode.Name = "LabelPostcode"
        Me.LabelPostcode.Size = New System.Drawing.Size(64, 16)
        Me.LabelPostcode.TabIndex = 93
        Me.LabelPostcode.Text = "Postcode"
        Me.LabelPostcode.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'ButtonLabelDel
        '
        Me.ButtonLabelDel.Enabled = False
        Me.ButtonLabelDel.Location = New System.Drawing.Point(139, 336)
        Me.ButtonLabelDel.Name = "ButtonLabelDel"
        Me.ButtonLabelDel.Size = New System.Drawing.Size(125, 24)
        Me.ButtonLabelDel.TabIndex = 92
        Me.ButtonLabelDel.Text = "Labels verwijderen"
        '
        'ButtonLabelAdd
        '
        Me.ButtonLabelAdd.Enabled = False
        Me.ButtonLabelAdd.Location = New System.Drawing.Point(8, 336)
        Me.ButtonLabelAdd.Name = "ButtonLabelAdd"
        Me.ButtonLabelAdd.Size = New System.Drawing.Size(125, 24)
        Me.ButtonLabelAdd.TabIndex = 84
        Me.ButtonLabelAdd.Text = "Label plaatsen"
        '
        'ButtonSave
        '
        Me.ButtonSave.Location = New System.Drawing.Point(8, 400)
        Me.ButtonSave.Name = "ButtonSave"
        Me.ButtonSave.Size = New System.Drawing.Size(256, 24)
        Me.ButtonSave.TabIndex = 83
        Me.ButtonSave.Text = "Wijzigingen opslaan"
        '
        'TextBoxStraatnaam
        '
        Me.TextBoxStraatnaam.BackColor = System.Drawing.SystemColors.Control
        Me.TextBoxStraatnaam.Enabled = False
        Me.TextBoxStraatnaam.ForeColor = System.Drawing.SystemColors.InactiveBorder
        Me.TextBoxStraatnaam.Location = New System.Drawing.Point(72, 88)
        Me.TextBoxStraatnaam.Name = "TextBoxStraatnaam"
        Me.TextBoxStraatnaam.Size = New System.Drawing.Size(192, 20)
        Me.TextBoxStraatnaam.TabIndex = 72
        Me.TextBoxStraatnaam.Text = ""
        '
        'LabelStraatnaam
        '
        Me.LabelStraatnaam.Location = New System.Drawing.Point(8, 88)
        Me.LabelStraatnaam.Name = "LabelStraatnaam"
        Me.LabelStraatnaam.Size = New System.Drawing.Size(40, 16)
        Me.LabelStraatnaam.TabIndex = 71
        Me.LabelStraatnaam.Text = "Straat"
        Me.LabelStraatnaam.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'TextBoxNaam
        '
        Me.TextBoxNaam.BackColor = System.Drawing.SystemColors.Control
        Me.TextBoxNaam.Enabled = False
        Me.TextBoxNaam.ForeColor = System.Drawing.SystemColors.InactiveBorder
        Me.TextBoxNaam.Location = New System.Drawing.Point(8, 40)
        Me.TextBoxNaam.Name = "TextBoxNaam"
        Me.TextBoxNaam.Size = New System.Drawing.Size(256, 20)
        Me.TextBoxNaam.TabIndex = 66
        Me.TextBoxNaam.Text = ""
        '
        'LabelNaam
        '
        Me.LabelNaam.Location = New System.Drawing.Point(8, 24)
        Me.LabelNaam.Name = "LabelNaam"
        Me.LabelNaam.Size = New System.Drawing.Size(64, 16)
        Me.LabelNaam.TabIndex = 65
        Me.LabelNaam.Text = "Naam"
        Me.LabelNaam.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'TextBoxStraatcode
        '
        Me.TextBoxStraatcode.BackColor = System.Drawing.SystemColors.Control
        Me.TextBoxStraatcode.Enabled = False
        Me.TextBoxStraatcode.ForeColor = System.Drawing.SystemColors.InactiveBorder
        Me.TextBoxStraatcode.Location = New System.Drawing.Point(200, 136)
        Me.TextBoxStraatcode.Name = "TextBoxStraatcode"
        Me.TextBoxStraatcode.Size = New System.Drawing.Size(64, 20)
        Me.TextBoxStraatcode.TabIndex = 62
        Me.TextBoxStraatcode.Text = ""
        '
        'LabelStraatcode
        '
        Me.LabelStraatcode.Location = New System.Drawing.Point(136, 136)
        Me.LabelStraatcode.Name = "LabelStraatcode"
        Me.LabelStraatcode.Size = New System.Drawing.Size(64, 16)
        Me.LabelStraatcode.TabIndex = 61
        Me.LabelStraatcode.Text = "Straatcode"
        Me.LabelStraatcode.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'TextBoxVolgNr
        '
        Me.TextBoxVolgNr.BackColor = System.Drawing.SystemColors.Control
        Me.TextBoxVolgNr.Enabled = False
        Me.TextBoxVolgNr.Location = New System.Drawing.Point(216, 16)
        Me.TextBoxVolgNr.Name = "TextBoxVolgNr"
        Me.TextBoxVolgNr.Size = New System.Drawing.Size(48, 20)
        Me.TextBoxVolgNr.TabIndex = 52
        Me.TextBoxVolgNr.Text = ""
        '
        'LabelVolgNr
        '
        Me.LabelVolgNr.Location = New System.Drawing.Point(176, 16)
        Me.LabelVolgNr.Name = "LabelVolgNr"
        Me.LabelVolgNr.Size = New System.Drawing.Size(48, 16)
        Me.LabelVolgNr.TabIndex = 51
        Me.LabelVolgNr.Text = "Volgnr"
        Me.LabelVolgNr.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'TextBoxAanduiding
        '
        Me.TextBoxAanduiding.BackColor = System.Drawing.SystemColors.Control
        Me.TextBoxAanduiding.Enabled = False
        Me.TextBoxAanduiding.Location = New System.Drawing.Point(72, 64)
        Me.TextBoxAanduiding.Name = "TextBoxAanduiding"
        Me.TextBoxAanduiding.Size = New System.Drawing.Size(192, 20)
        Me.TextBoxAanduiding.TabIndex = 50
        Me.TextBoxAanduiding.Text = ""
        '
        'LabelAanduiding
        '
        Me.LabelAanduiding.Location = New System.Drawing.Point(8, 64)
        Me.LabelAanduiding.Name = "LabelAanduiding"
        Me.LabelAanduiding.Size = New System.Drawing.Size(64, 16)
        Me.LabelAanduiding.TabIndex = 49
        Me.LabelAanduiding.Text = "Aanduiding"
        Me.LabelAanduiding.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'TextBoxInfo
        '
        Me.TextBoxInfo.BackColor = System.Drawing.SystemColors.Control
        Me.TextBoxInfo.Enabled = False
        Me.TextBoxInfo.ForeColor = System.Drawing.SystemColors.InactiveBorder
        Me.TextBoxInfo.Location = New System.Drawing.Point(8, 288)
        Me.TextBoxInfo.Multiline = True
        Me.TextBoxInfo.Name = "TextBoxInfo"
        Me.TextBoxInfo.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.TextBoxInfo.Size = New System.Drawing.Size(256, 42)
        Me.TextBoxInfo.TabIndex = 103
        Me.TextBoxInfo.Text = ""
        '
        'ButtonClose
        '
        Me.ButtonClose.Location = New System.Drawing.Point(192, 536)
        Me.ButtonClose.Name = "ButtonClose"
        Me.ButtonClose.Size = New System.Drawing.Size(88, 24)
        Me.ButtonClose.TabIndex = 53
        Me.ButtonClose.Text = "Sluiten"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.ButtonLoad)
        Me.GroupBox1.Controls.Add(Me.ComboBoxLayerFilter)
        Me.GroupBox1.Location = New System.Drawing.Point(8, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(272, 48)
        Me.GroupBox1.TabIndex = 54
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Filter"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 20)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(64, 16)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "Selectie uit"
        '
        'ButtonLoad
        '
        Me.ButtonLoad.Location = New System.Drawing.Point(208, 14)
        Me.ButtonLoad.Name = "ButtonLoad"
        Me.ButtonLoad.Size = New System.Drawing.Size(56, 24)
        Me.ButtonLoad.TabIndex = 4
        Me.ButtonLoad.Text = "Uitlezen"
        '
        'ComboBoxLayerFilter
        '
        Me.ComboBoxLayerFilter.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBoxLayerFilter.Location = New System.Drawing.Point(72, 16)
        Me.ComboBoxLayerFilter.Name = "ComboBoxLayerFilter"
        Me.ComboBoxLayerFilter.Size = New System.Drawing.Size(128, 21)
        Me.ComboBoxLayerFilter.TabIndex = 3
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
        Me.GroupBox3.Location = New System.Drawing.Point(8, 48)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(272, 48)
        Me.GroupBox3.TabIndex = 55
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
        Me.ButtonNext.TabIndex = 7
        Me.ButtonNext.Text = ">"
        '
        'ButtonLast
        '
        Me.ButtonLast.Location = New System.Drawing.Point(232, 16)
        Me.ButtonLast.Name = "ButtonLast"
        Me.ButtonLast.Size = New System.Drawing.Size(32, 24)
        Me.ButtonLast.TabIndex = 6
        Me.ButtonLast.Text = ">>"
        '
        'ButtonFirst
        '
        Me.ButtonFirst.Location = New System.Drawing.Point(8, 16)
        Me.ButtonFirst.Name = "ButtonFirst"
        Me.ButtonFirst.Size = New System.Drawing.Size(32, 24)
        Me.ButtonFirst.TabIndex = 4
        Me.ButtonFirst.Text = "<<"
        '
        'ButtonPrevious
        '
        Me.ButtonPrevious.Location = New System.Drawing.Point(48, 16)
        Me.ButtonPrevious.Name = "ButtonPrevious"
        Me.ButtonPrevious.Size = New System.Drawing.Size(32, 24)
        Me.ButtonPrevious.TabIndex = 5
        Me.ButtonPrevious.Text = "<"
        '
        'LabelGebouwType
        '
        Me.LabelGebouwType.Location = New System.Drawing.Point(8, 160)
        Me.LabelGebouwType.Name = "LabelGebouwType"
        Me.LabelGebouwType.Size = New System.Drawing.Size(64, 16)
        Me.LabelGebouwType.TabIndex = 112
        Me.LabelGebouwType.Text = "Type"
        Me.LabelGebouwType.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'ComboBoxGebouwType
        '
        Me.ComboBoxGebouwType.BackColor = System.Drawing.SystemColors.Control
        Me.ComboBoxGebouwType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBoxGebouwType.Enabled = False
        Me.ComboBoxGebouwType.ForeColor = System.Drawing.SystemColors.ControlText
        Me.ComboBoxGebouwType.Location = New System.Drawing.Point(72, 160)
        Me.ComboBoxGebouwType.Name = "ComboBoxGebouwType"
        Me.ComboBoxGebouwType.Size = New System.Drawing.Size(192, 21)
        Me.ComboBoxGebouwType.TabIndex = 113
        '
        'FormBeheerGebouwen
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(284, 568)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.ButtonClose)
        Me.Controls.Add(Me.GroupBox2)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.Name = "FormBeheerGebouwen"
        Me.Text = "Beheer Speciale Gebouwen"
        Me.TopMost = True
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBoxDossier.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
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
        'm_layer = GetFeatureLayer(m_document.FocusMap, GetLayerName("Hydrant"))
        m_editing = False
        'm_workspace = Nothing
        'm_marker = Nothing
        m_enumOIDs = Nothing

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Custom form initialization in Form.OnLoad.

    End Sub

#End Region

#Region " Form controls events "

    Private Sub ButtonLoad_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonLoad.Click

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

        'Load the current map selection in the form.
        Dim pTargetLayer As IFeatureLayer = GetFeatureLayer(m_document.FocusMap, ComboBoxLayerFilter.Text)
        Dim pSelectionSet As ISelectionSet = CType(pTargetLayer, IFeatureSelection).SelectionSet
        If Not pSelectionSet Is Nothing Then LoadSelectionSet(pSelectionSet)

    End Sub

    Private Sub ButtonClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonClose.Click
        Me.Close() 'Close form.
        'The user will be able to store his changes to current features attributes,
        'before the form is closed, becauce of the OnClosing event of current form.
    End Sub

    Private Sub ButtonSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonSave.Click

        'Store attribute changes if modifications are registered.
        If ModifiedAttribute() Then
            If ValidateAttributeChanges() Then
                StoreAttributeChanges()
            End If
        End If

    End Sub

    Private Sub ButtonFirst_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonFirst.Click
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
    End Sub

    Private Sub ButtonPrevious_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonPrevious.Click
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
    End Sub

    Private Sub ButtonNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonNext.Click
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
    End Sub

    Private Sub ButtonLast_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonLast.Click
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
    End Sub

    Private Sub TextBoxNaam_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBoxNaam.TextChanged
        If m_editing Then MarkAsChanged(LabelNaam)
    End Sub

    Private Sub TextBoxAanduiding_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBoxAanduiding.TextChanged
        If m_editing Then MarkAsChanged(LabelAanduiding)
    End Sub

    Private Sub TextBoxStraatnaam_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBoxStraatnaam.TextChanged
        If m_editing Then MarkAsChanged(LabelStraatnaam)
    End Sub

    Private Sub TextBoxHuisnr_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBoxHuisnr.TextChanged
        If m_editing Then MarkAsChanged(LabelHuisnr)
    End Sub

    Private Sub TextBoxVolgNr_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBoxVolgNr.TextChanged
        If m_editing Then MarkAsChanged(LabelVolgNr)
    End Sub

    Private Sub TextBoxStraatcode_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBoxStraatcode.TextChanged
        If m_editing Then MarkAsChanged(LabelStraatcode)
    End Sub

    Private Sub TextBoxPostcode_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBoxPostcode.TextChanged
        If m_editing Then MarkAsChanged(LabelPostcode)
    End Sub

    Private Sub TextBoxDossierNrOud_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBoxDossierNrOud.TextChanged
        If m_editing Then MarkAsChanged(LabelDossierNrOud)
    End Sub

    Private Sub TextBoxDossierNrNieuw_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBoxDossierNrNieuw.TextChanged
        If m_editing Then MarkAsChanged(LabelDossierNrNieuw)
    End Sub

    Private Sub TextBoxDossierInterventie_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBoxDossierInterventie.TextChanged
        If m_editing Then MarkAsChanged(LabelDossierInterventie)
    End Sub

    Private Sub TextBoxInfo_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBoxInfo.TextChanged
        If m_editing Then MarkAsChanged(LabelInfo)
    End Sub

    Private Sub ComboBoxGebouwType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBoxGebouwType.SelectedIndexChanged
        If m_editing Then MarkAsChanged(LabelGebouwType)
    End Sub

    Private Sub ButtonLabelAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonLabelAdd.Click

        Try
            'Get annotations layer.
            Dim pLayer As ILayer = GetFeatureLayer(m_document.FocusMap, GetLayerName("SpeciaalGebouwAnno"))
            If pLayer Is Nothing Then Throw New LayerNotFoundException(GetLayerName("SpeciaalGebouwAnno"))
            Dim pAnnoLayer As IAnnotationLayer = CType(pLayer, IAnnotationLayer)

            'Add new annotation feature.
            Dim pMarkerElement As IMarkerElement = GetMarkerElement(c_MarkerTag, m_document)
            If Not pMarkerElement Is Nothing Then
                Dim pGeometry As IGeometry = CType(pMarkerElement, IElement).Geometry
                If TypeOf pGeometry Is IPoint Then
                    Dim frm As FormAddAnnotation = _
                        New FormAddAnnotation( _
                            pAnnoLayer, _
                            CType(pGeometry, IPoint), _
                            TextBoxVolgNr.Text, _
                            GetAttributeName("SpeciaalGebouwAnno", "LinkID"), _
                            TextBoxVolgNr.Text)
                    frm.ShowDialog()
                End If
            End If

            'Partial refresh to display new annotation.
            m_document.ActivatedView.PartialRefresh(esriViewDrawPhase.esriViewGeography, pAnnoLayer, Nothing)

            'Activate the Edit Annotation Tool command if edit session is started.
            Dim pEditor As IEditor2 = GetEditorReference(m_application)
            EditSessionStart(pEditor, GetFeatureLayer(m_document.FocusMap, GetLayerName("SpeciaalGebouw")), True)
            If Not pEditor.EditWorkspace Is Nothing Then
                'Make the annotation layer selectable.
                GetFeatureLayer(m_document.FocusMap, GetLayerName("SpeciaalGebouw")).Selectable = True
                'Activate the Edit Annotation tool.
                ActivateTool(CType(m_document, IDocument), "esriEditor.AnnoEditTool")
            End If

        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    Private Sub ButtonLabelDel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonLabelDel.Click
        Try
            Dim title As String = c_Title_DeleteAnno
            Dim message As String = c_Message_ConfirmDeleteAnno

            'Ask the user for a confirmation.
            If MsgBox(message, MsgBoxStyle.OKCancel, title) = MsgBoxResult.OK Then

                'Remove all annotations with same LinkID.
                Dim linkID As String = CStr(TextBoxVolgNr.Text)
                Dim pAnnoLayer As IAnnotationLayer = CType(GetFeatureLayer(m_document.FocusMap, GetLayerName("SpeciaalGebouwAnno")), IAnnotationLayer)
                Dim annoCount As Integer = RemoveLinkedAnnotations(pAnnoLayer, GetAttributeName("SpeciaalGebouwAnno", "LinkID"), linkID)

                'Partial refresh to display new annotation.
                m_document.ActivatedView.PartialRefresh(esriViewDrawPhase.esriViewGeography, pAnnoLayer, Nothing)

                'Inform the user about the number of deleted annotations.
                message = Replace(c_Message_DeleteAnnoCount, "^0", CStr(annoCount))
                MsgBox(message, MsgBoxStyle.OKOnly, title)

            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub CheckBoxCopy_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxCopy.CheckedChanged
        Try
            If Me.CheckBoxCopy.Checked Then
                'Activate CopyAddressFunctionality
                CopyAddressFunctionality_Activate(m_document, Me)
                'Show text in another color for better perception.
                Me.CheckBoxCopy.ForeColor = System.Drawing.Color.BlueViolet
            Else
                'Deactivate CopyAddressFunctionality
                CopyAddressFunctionality_Deactivate()
                'Restore text to the default color.
                Me.CheckBoxCopy.ForeColor = System.Drawing.Color.Black
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

#End Region

#Region " Overridden form events "

    Protected Overrides Sub OnLoad(ByVal e As System.EventArgs)

        ' Check availability of configuration.
        If Config Is Nothing Then Throw New ApplicationException("No configuration loaded.")

#If DEBUG Then
        ' Show some controls.
        With Me
            .TextBoxStraatcode.Visible = True
            .LabelStraatcode.Visible = True
        End With
#Else
        ' Hide some controls.
        With Me
            .TextBoxStraatcode.Visible = False
            .LabelStraatcode.Visible = False
        End With
#End If

        ' Get a list of related layers from configuration.
        Dim layerNames As Collection = Config.BuildingLayers

        ' If layer is available on map, add it to the list.
        Dim featureLayer As IFeatureLayer = Nothing
        For Each layerName As String In layerNames
            featureLayer = GetFeatureLayer(m_document.FocusMap, layerName)
            If Not featureLayer Is Nothing Then ComboBoxLayerFilter.Items.Add(layerName)
        Next

        ' List domain code values in the ligging combo box.
        If Not featureLayer Is Nothing Then
            Dim domainMgr As CodedValueDomainManager
            domainMgr = New CodedValueDomainManager(featureLayer, "GebouwType")
            domainMgr.PopulateCodes(Me.ComboBoxGebouwType)
        End If

        ' Select the first item in the list.
        ComboBoxLayerFilter.SelectedIndex = 0

        ' Simulate selection if only one item in list.
        If ComboBoxLayerFilter.Items.Count = 1 Then Call ButtonLoad_Click(Nothing, Nothing)

    End Sub

    Protected Overrides Sub Finalize()

        'Free the private property objects for garbage collection.
        m_application = Nothing
        m_document = Nothing
        'm_layer = Nothing
        m_editing = Nothing
        'm_workspace = Nothing
        'm_marker = Nothing
        m_enumOIDs = Nothing

        MyBase.Finalize()
    End Sub

    Protected Overrides Sub OnClosing(ByVal e As System.ComponentModel.CancelEventArgs)

        Dim pActiveView As IActiveView
        Dim pElement As IElement
        Dim pGraphics As IGraphicsContainer
        Dim pMarker As IMarkerElement
        Dim pMxDocument As IMxDocument

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

    Private Sub FormBeheerGebouwen_Closed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Closed
        'Be sure to remove remaining eventhandler from the "Copy address" functionality.
        CopyAddressFunctionality_Deactivate()
    End Sub

#End Region

#Region " Utility procedures "

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Use the SelectionSet as the set of data editable with this form.
    ''' </summary>
    ''' <param name="pSelectionSet"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub LoadSelectionSet(ByVal pSelectionSet As ISelectionSet)
        Try
            m_editing = False

            'Set form navigation controls.
            Dim SelectionSetCount As Integer = pSelectionSet.Count
            LabelTotal.Text = CStr(SelectionSetCount)
            LabelCounter.Text = CStr(0)

            'Enumeration of feature IDs.
            m_enumOIDs = CType(pSelectionSet.IDs, IEnumIDs)

            'Load first record from SelectionSet into the form controls.
            'Or disable all controls if selectionset is empty.
            If SelectionSetCount > 0 Then
                LoadNextFeature()
                EnableNavigationControls(True)
            Else
                MsgBox(c_Message_EmptyFeatureSet, MsgBoxStyle.Exclamation)
                EnableEditingControls(False)
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
    ''' 	[Kristof Vydt]	24/10/2005	Deactivate listeners
    ''' 	[Kristof Vydt]	23/11/2005	Add optional parameter to avoid zoom-to during reload of feature.
    ''' 	[Kristof Vydt]	02/08/2006	Replace some global constant by a function that reads the value from config file.
    '''     [Kristof Vydt]  18/08/2006  MarkerElement is no longer a parameter of MarkAndZoomTo().
    ''' 	[Kristof Vydt]	14/03/2007	Add GebouwType attribute.
    '''  	[Kristof Vydt]	04/04/2007	Correct LabelInfo to LabelGebouwType.
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub LoadFeature( _
        ByVal OID As Integer, _
        Optional ByVal zoomToFeature As Boolean = True)

        Try

            'Make sure there is an enumeration of IDs.
            If m_enumOIDs Is Nothing Then Exit Sub

            'Disable editing.
            EnableEditingControls(False)
            m_editing = False

            'Deactivate listeners.
            ConnectFeatureFunctionality_Deactivate()
            CopyAttributesFunctionality_Deactivate()
            CopyAddressFunctionality_Deactivate()

            'Hold the current objectID as a form private variable.
            m_OID = OID

            'Get the first feature from the SelectionSet.
            Dim pLayer As IFeatureLayer = GetFeatureLayer(m_document.FocusMap, GetLayerName("SpeciaalGebouw"))
            Dim pTable As ITable = CType(pLayer, ITable)
            Dim pQueryFilter As IQueryFilter = New QueryFilter
            pQueryFilter.WhereClause = "OBJECTID = " & OID
            Dim pCursor As ICursor = pTable.Search(pQueryFilter, True)
            Dim pRow As IRow = pCursor.NextRow

            'Zoom to the feature and mark it on the map.
            Dim pFeature As IFeature = CType(pRow, IFeature)
            MarkAndZoomTo(pFeature, m_document, False)

            'Initialize layout of form controls and
            'Show feature attributes in the form controls.
            Dim FieldIndex As Integer
            '- Aanduiding
            LabelAanduiding.ForeColor = Black
            FieldIndex = pRow.Fields.FindField(GetAttributeName("SpeciaalGebouw", "Aanduiding"))
            SetEditBoxValue(TextBoxAanduiding, pRow.Value(FieldIndex))
            '- Huisnummer
            LabelHuisnr.ForeColor = Black
            FieldIndex = pRow.Fields.FindField(GetAttributeName("SpeciaalGebouw", "Huisnr"))
            SetEditBoxValue(TextBoxHuisnr, pRow.Value(FieldIndex))
            '- Info
            LabelInfo.ForeColor = Black
            FieldIndex = pRow.Fields.FindField(GetAttributeName("SpeciaalGebouw", "Info"))
            SetEditBoxValue(TextBoxInfo, pRow.Value(FieldIndex))
            '- GebouwType
            LabelGebouwType.ForeColor = Black
            FieldIndex = pRow.Fields.FindField(GetAttributeName("SpeciaalGebouw", "GebouwType"))
            SetEditBoxValue(ComboBoxGebouwType, pRow.Value(FieldIndex))
            '- InterventieDossier
            LabelDossierInterventie.ForeColor = Black
            FieldIndex = pRow.Fields.FindField(GetAttributeName("SpeciaalGebouw", "DossierInterventie"))
            SetEditBoxValue(TextBoxDossierInterventie, pRow.Value(FieldIndex))
            '- Naam
            LabelNaam.ForeColor = Black
            FieldIndex = pRow.Fields.FindField(GetAttributeName("SpeciaalGebouw", "Naam"))
            SetEditBoxValue(TextBoxNaam, pRow.Value(FieldIndex))
            '- NieuwDossiernr
            LabelDossierNrNieuw.ForeColor = Black
            FieldIndex = pRow.Fields.FindField(GetAttributeName("SpeciaalGebouw", "DossierNrNieuw"))
            SetEditBoxValue(TextBoxDossierNrNieuw, pRow.Value(FieldIndex))
            '- OudDossiernr
            LabelDossierNrOud.ForeColor = Black
            FieldIndex = pRow.Fields.FindField(GetAttributeName("SpeciaalGebouw", "DossierNrOud"))
            SetEditBoxValue(TextBoxDossierNrOud, pRow.Value(FieldIndex))
            '- Postcode
            LabelPostcode.ForeColor = Black
            FieldIndex = pRow.Fields.FindField(GetAttributeName("SpeciaalGebouw", "Postcode"))
            SetEditBoxValue(TextBoxPostcode, pRow.Value(FieldIndex))
            '- Straatcode
            LabelStraatcode.ForeColor = Black
            FieldIndex = pRow.Fields.FindField(GetAttributeName("SpeciaalGebouw", "Straatcode"))
            SetEditBoxValue(TextBoxStraatcode, pRow.Value(FieldIndex))
            '- Straatnaam
            LabelStraatnaam.ForeColor = Black
            FieldIndex = pRow.Fields.FindField(GetAttributeName("SpeciaalGebouw", "Straatnaam"))
            SetEditBoxValue(TextBoxStraatnaam, pRow.Value(FieldIndex))
            '- Volgnummer
            LabelVolgNr.ForeColor = Black
            FieldIndex = pRow.Fields.FindField(GetAttributeName("SpeciaalGebouw", "Volgnr"))
            SetEditBoxValue(TextBoxVolgNr, pRow.Value(FieldIndex))
            'Create a new Volgnummer if it doesn't have a value yet.
            If (TextBoxVolgNr.Text = "") Or (TextBoxVolgNr.Text = "0") Then
                SetEditBoxValue(TextBoxVolgNr, NextUniqueAttributeValue(pLayer, GetAttributeName("SpeciaalGebouw", "Volgnr"), True))
                MarkAsChanged(LabelVolgNr)
            End If

            'Enable editing.
            EnableEditingControls(True)
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
    '''     The control that is changed.
    ''' </param>
    ''' <remarks>
    '''     The label of the modified control is displayed in (fixed) IndianRed.
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub MarkAsChanged( _
        ByVal SomeControl As Windows.Forms.Control)
        If TypeOf SomeControl Is Windows.Forms.Label Then
            SomeControl.ForeColor = IndianRed
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
    '''     Set enabled/visible attributes of form controls for attribute editing.
    ''' </summary>
    ''' <param name="value"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    '''     [Kristof Vydt]	05/10/2005  Replace ButtonConnect/Copy by CheckBoxConnect/Copy.
    ''' 	[Kristof Vydt]	14/03/2007	Add GebouwType attribute.
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub EnableEditingControls(ByVal value As Boolean)

        Dim ForeColorEnabled As Color = Black
        Dim BackColorEnabled As Color = White
        Dim ForeColorDisabled As Color = Gray
        Dim BackColorDisabled As Color = White
        Dim ForeColor As Color
        Dim BackColor As Color

        'Set a color depending on the value.
        If value Then
            BackColor = BackColorEnabled
            ForeColor = ForeColorEnabled
        Else
            BackColor = BackColorDisabled
            ForeColor = ForeColorDisabled
        End If

        'Status of TextBox controls.
        '- Aanduiding
        TextBoxAanduiding.Enabled = value
        TextBoxAanduiding.ForeColor = ForeColor
        TextBoxAanduiding.BackColor = BackColor
        '- DossierNr Oud
        TextBoxDossierNrOud.Enabled = value
        TextBoxDossierNrOud.ForeColor = ForeColor
        TextBoxDossierNrOud.BackColor = BackColor
        '- DossierNr Nieuw
        TextBoxDossierNrNieuw.Enabled = value
        TextBoxDossierNrNieuw.ForeColor = ForeColor
        TextBoxDossierNrNieuw.BackColor = BackColor
        '- Dossier Interventie
        TextBoxDossierInterventie.Enabled = value
        TextBoxDossierInterventie.ForeColor = ForeColor
        TextBoxDossierInterventie.BackColor = BackColor
        '- Huisnummer
        TextBoxHuisnr.Enabled = False 'read-only
        TextBoxHuisnr.ForeColor = ForeColorDisabled
        TextBoxHuisnr.BackColor = BackColorDisabled
        '- Info
        TextBoxInfo.Enabled = value
        TextBoxInfo.ForeColor = ForeColor
        TextBoxInfo.BackColor = BackColor
        '- GebouwType
        ComboBoxGebouwType.Enabled = value
        ComboBoxGebouwType.ForeColor = ForeColor
        ComboBoxGebouwType.BackColor = BackColor
        '- Naam
        TextBoxNaam.Enabled = value
        TextBoxNaam.ForeColor = ForeColor
        TextBoxNaam.BackColor = BackColor
        '- Postcode
        TextBoxPostcode.Enabled = False 'read-only
        TextBoxPostcode.ForeColor = ForeColorDisabled
        TextBoxPostcode.BackColor = BackColorDisabled
        '- Straatcode
        TextBoxStraatcode.Enabled = False 'read-only
        TextBoxStraatcode.BackColor = BackColorDisabled
        TextBoxStraatcode.ForeColor = ForeColorDisabled
        '- Straatnaam
        TextBoxStraatnaam.Enabled = False 'read-only
        TextBoxStraatnaam.ForeColor = ForeColorDisabled
        TextBoxStraatnaam.BackColor = BackColorDisabled
        '- Volgnummer
        TextBoxVolgNr.Enabled = value
        TextBoxVolgNr.ForeColor = ForeColor
        TextBoxVolgNr.BackColor = BackColor

        'Status of Button controls.
        ButtonLabelAdd.Enabled = value
        ButtonLabelDel.Enabled = value
        CheckBoxConnect.Enabled = False 'no such functionality for this form
        CheckBoxCopy.Enabled = value
        ButtonSave.Enabled = value
        ButtonClose.Enabled = True 'always available

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Return true if at least one of the attribute editing controls is modified.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    '''     The color of the attribute label control is supposed to be (fixed) IndianRed if modified.
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    ''' 	[Kristof Vydt]	14/03/2007	Add GebouwType attribute.
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Function ModifiedAttribute() As Boolean
        If LabelAanduiding.ForeColor.Equals(IndianRed) Or _
           LabelVolgNr.ForeColor.Equals(IndianRed) Or _
           LabelNaam.ForeColor.Equals(IndianRed) Or _
           LabelStraatnaam.ForeColor.Equals(IndianRed) Or _
           LabelHuisnr.ForeColor.Equals(IndianRed) Or _
           LabelStraatcode.ForeColor.Equals(IndianRed) Or _
           LabelPostcode.ForeColor.Equals(IndianRed) Or _
           LabelDossierNrOud.ForeColor.Equals(IndianRed) Or _
           LabelDossierNrNieuw.ForeColor.Equals(IndianRed) Or _
           LabelDossierInterventie.ForeColor.Equals(IndianRed) Or _
           LabelInfo.ForeColor.Equals(IndianRed) Or _
           LabelGebouwType.ForeColor.Equals(IndianRed) Then

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
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Function ValidateAttributeChanges() As Boolean
        'TODO:
        '- Aanduiding
        '- Huisnummer
        '- ID
        '- Info
        '- InterventieDossier
        '- Naam
        '- NieuwDossiernr
        '- OudDossiernr
        '- Postcode
        '- Straatcode
        '- Straatnaam
        '- Volgnummer
        Try
            Return True
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Store current feature attribute modifications.
    ''' </summary>
    ''' <remarks>
    '''     No modifications will be saved if there is an active edit session.
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    ''' 	[Kristof Vydt]	17/07/2006	Close active edit session before modifying feature attributes.
    '''                                 Changes to "Volgnummer" should result in update of existing related annotation features.
    ''' 	[Kristof Vydt]	02/08/2006	Replace some global constant by a function that reads the value from config file.
    ''' 	[Kristof Vydt]	11/08/2006	Check if DBNull before validating length.
    '''                                 Do no longer CStr the txtbox value when assiging to feature attribute value.
    '''     [Kristof Vydt]  18/08/2006  Check if LinkID is DBNull before updating existing linked annotations.
    ''' 	[Kristof Vydt]	14/03/2007	Add GebouwType attribute.
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub StoreAttributeChanges()

        Dim fldIndex As Integer
        Dim pFeatureLayer As IFeatureLayer = Nothing
        Dim pQueryFilter As IQueryFilter = Nothing
        Dim pFeatureCursor As IFeatureCursor = Nothing
        Dim pFeature As IFeature = Nothing
        Dim pEditor As IEditor2 = Nothing
        Dim objLinkID As Object 'identifier of the link between hydrant feature and hydrant label
        Dim strLinkID As String 'same identifier but as string
        Dim pAnnoLayer As IAnnotationLayer = Nothing 'hydrant labels annotation layer
        Dim pParams As Hashtable 'hashtable (list of key-value pairs)

        Try
            ' Get the feature that is modified.
            pFeatureLayer = GetFeatureLayer(m_document.FocusMap, ComboBoxLayerFilter.Text)
            If pFeatureLayer Is Nothing Then Throw New LayerNotFoundException(ComboBoxLayerFilter.Text)
            pQueryFilter = New QueryFilter
            pQueryFilter.WhereClause = "OBJECTID = " & m_OID
            pFeatureCursor = pFeatureLayer.FeatureClass.Search(pQueryFilter, True)
            pFeature = pFeatureCursor.NextFeature

            ' Make sure that there is at least one feature.
            If pFeature Is Nothing Then
                MsgBox("Cannot save changed because the feature with OBJECTID " & CStr(m_OID) & " could not be found in the feature class.", _
                    MsgBoxStyle.Exclamation, "Wijzigingen opslaan.")
                Exit Sub
            End If

            ' Close active edit session before continuing.
            pEditor = GetEditorReference(m_application)
            If pEditor.EditState = esriEditState.esriStateEditing Then
                ' Does the user wants to save changes while closing the edit session?
                If MsgBox(c_Message_SaveEdits, vbYesNo, c_Title_SaveEdits) = MsgBoxResult.Yes Then
                    ' Close the active edit session and save changes.
                    EditSessionSave(pEditor)
                Else
                    ' Close the active edit session without saving changes.
                    EditSessionAbort(pEditor)
                End If
            End If

            ' Modify current feature attributes.
            '- Aanduiding
            If LabelAanduiding.ForeColor.Equals(IndianRed) Then
                fldIndex = pFeature.Fields.FindField(GetAttributeName("SpeciaalGebouw", "Aanduiding"))
                If Not String.IsNullOrEmpty(TextBoxHuisnr.Text) Then
                    If Len(CStr(TextBoxAanduiding.Text)) > pFeature.Fields.Field(fldIndex).Length Then
                        Throw New AttributeSizeNotSufficientException(ComboBoxLayerFilter.Text, GetAttributeName("SpeciaalGebouw", "Aanduiding"))
                    End If
                End If
                pFeature.Value(fldIndex) = TextBoxAanduiding.Text
            End If
            '- Huisnummer
            If LabelHuisnr.ForeColor.Equals(IndianRed) Then
                fldIndex = pFeature.Fields.FindField(GetAttributeName("SpeciaalGebouw", "Huisnr"))
                If Not String.IsNullOrEmpty(TextBoxHuisnr.Text) Then
                    If Len(CStr(TextBoxHuisnr.Text)) > pFeature.Fields.Field(fldIndex).Length Then
                        Throw New AttributeSizeNotSufficientException(ComboBoxLayerFilter.Text, GetAttributeName("SpeciaalGebouw", "Huisnr"))
                    End If
                End If
                pFeature.Value(fldIndex) = TextBoxHuisnr.Text
            End If
            '- Info
            If LabelInfo.ForeColor.Equals(IndianRed) Then
                fldIndex = pFeature.Fields.FindField(GetAttributeName("SpeciaalGebouw", "Info"))
                If Not String.IsNullOrEmpty(TextBoxInfo.Text) Then
                    If Len(CStr(TextBoxInfo.Text)) > pFeature.Fields.Field(fldIndex).Length Then
                        Throw New AttributeSizeNotSufficientException(ComboBoxLayerFilter.Text, GetAttributeName("SpeciaalGebouw", "Info"))
                    End If
                End If
                pFeature.Value(fldIndex) = TextBoxInfo.Text
            End If
            '- GebouwType
            If LabelGebouwType.ForeColor.Equals(IndianRed) Then
                fldIndex = pFeature.Fields.FindField(GetAttributeName("SpeciaalGebouw", "GebouwType"))
                pFeature.Value(fldIndex) = GetComboBoxCodeValue(ComboBoxGebouwType)
            End If
            '- InterventieDossier
            If LabelDossierInterventie.ForeColor.Equals(IndianRed) Then
                fldIndex = pFeature.Fields.FindField(GetAttributeName("SpeciaalGebouw", "DossierInterventie"))
                If Not String.IsNullOrEmpty(TextBoxDossierInterventie.Text) Then
                    If Len(CStr(TextBoxDossierInterventie.Text)) > pFeature.Fields.Field(fldIndex).Length Then
                        Throw New AttributeSizeNotSufficientException(ComboBoxLayerFilter.Text, GetAttributeName("SpeciaalGebouw", "DossierInterventie"))
                    End If
                End If
                pFeature.Value(fldIndex) = TextBoxDossierInterventie.Text
            End If
            '- Naam
            If LabelNaam.ForeColor.Equals(IndianRed) Then
                fldIndex = pFeature.Fields.FindField(GetAttributeName("SpeciaalGebouw", "Naam"))
                If Not String.IsNullOrEmpty(TextBoxNaam.Text) Then
                    If Len(CStr(TextBoxNaam.Text)) > pFeature.Fields.Field(fldIndex).Length Then
                        Throw New AttributeSizeNotSufficientException(ComboBoxLayerFilter.Text, GetAttributeName("SpeciaalGebouw", "Naam"))
                    End If
                End If
                pFeature.Value(fldIndex) = TextBoxNaam.Text
            End If
            '- NieuwDossiernr
            If LabelDossierNrNieuw.ForeColor.Equals(IndianRed) Then
                fldIndex = pFeature.Fields.FindField(GetAttributeName("SpeciaalGebouw", "DossierNrNieuw"))
                If Not String.IsNullOrEmpty(TextBoxDossierNrNieuw.Text) Then
                    If Len(CStr(TextBoxDossierNrNieuw.Text)) > pFeature.Fields.Field(fldIndex).Length Then
                        Throw New AttributeSizeNotSufficientException(ComboBoxLayerFilter.Text, GetAttributeName("SpeciaalGebouw", "DossierNrNieuw"))
                    End If
                End If
                pFeature.Value(fldIndex) = TextBoxDossierNrNieuw.Text
            End If
            '- OudDossiernr
            If LabelDossierNrOud.ForeColor.Equals(IndianRed) Then
                fldIndex = pFeature.Fields.FindField(GetAttributeName("SpeciaalGebouw", "DossierNrOud"))
                If Not String.IsNullOrEmpty(TextBoxDossierNrOud.Text) Then
                    If Len(CStr(TextBoxDossierNrOud.Text)) > pFeature.Fields.Field(fldIndex).Length Then
                        Throw New AttributeSizeNotSufficientException(ComboBoxLayerFilter.Text, GetAttributeName("SpeciaalGebouw", "DossierNrOud"))
                    End If
                End If
                pFeature.Value(fldIndex) = TextBoxDossierNrOud.Text
            End If
            '- Postcode
            If LabelPostcode.ForeColor.Equals(IndianRed) Then
                fldIndex = pFeature.Fields.FindField(GetAttributeName("SpeciaalGebouw", "Postcode"))
                If Not String.IsNullOrEmpty(TextBoxPostcode.Text) Then
                    If Len(CStr(TextBoxPostcode.Text)) > pFeature.Fields.Field(fldIndex).Length Then
                        Throw New AttributeSizeNotSufficientException(ComboBoxLayerFilter.Text, GetAttributeName("SpeciaalGebouw", "Postcode"))
                    End If
                End If
                pFeature.Value(fldIndex) = TextBoxPostcode.Text
            End If
            '- Straatcode
            If LabelStraatcode.ForeColor.Equals(IndianRed) Then
                fldIndex = pFeature.Fields.FindField(GetAttributeName("SpeciaalGebouw", "Straatcode"))
                If Not String.IsNullOrEmpty(TextBoxStraatcode.Text) Then
                    If Len(CStr(TextBoxStraatcode.Text)) > pFeature.Fields.Field(fldIndex).Length Then
                        Throw New AttributeSizeNotSufficientException(ComboBoxLayerFilter.Text, GetAttributeName("SpeciaalGebouw", "Straatcode"))
                    End If
                End If
                pFeature.Value(fldIndex) = TextBoxStraatcode.Text
            End If
            '- Straatnaam
            If LabelStraatnaam.ForeColor.Equals(IndianRed) Then
                fldIndex = pFeature.Fields.FindField(GetAttributeName("SpeciaalGebouw", "Straatnaam"))
                If Not String.IsNullOrEmpty(TextBoxStraatnaam.Text) Then
                    If Len(CStr(TextBoxStraatnaam.Text)) > pFeature.Fields.Field(fldIndex).Length Then
                        Throw New AttributeSizeNotSufficientException(ComboBoxLayerFilter.Text, GetAttributeName("SpeciaalGebouw", "Straatnaam"))
                    End If
                End If
                pFeature.Value(fldIndex) = TextBoxStraatnaam.Text
            End If
            '- Volgnummer
            If LabelVolgNr.ForeColor.Equals(IndianRed) Then

                ' Changes to "Volgnummer" should result in update of existing related annotation features.
                objLinkID = pFeature.Value(pFeature.Fields.FindField(GetAttributeName("SpeciaalGebouw", "LinkID")))
                If Not TypeOf objLinkID Is System.DBNull Then
                    strLinkID = CStr(objLinkID)
                    pAnnoLayer = CType(GetFeatureLayer(m_document.FocusMap, GetLayerName("SpeciaalGebouwAnno")), IAnnotationLayer)
                    If pAnnoLayer Is Nothing Then Throw New LayerNotFoundException(GetLayerName("SpeciaalGebouwAnno"))
                    pParams = New Hashtable
                    pParams.Add("TextString", Me.TextBoxVolgNr.Text)
                    pParams.Add("LinkField", GetAttributeName("SpeciaalGebouwAnno", "LinkID"))
                    pParams.Add("LinkValue", Me.TextBoxVolgNr.Text)
                    UpdateAnno(pAnnoLayer, pParams, GetAttributeName("SpeciaalGebouwAnno", "LinkID"), strLinkID)
                End If

                'Update "Volgnummer" attribute of feature.
                fldIndex = pFeature.Fields.FindField(GetAttributeName("SpeciaalGebouw", "Volgnr"))
                If Len(CStr(TextBoxVolgNr.Text)) > pFeature.Fields.Field(fldIndex).Length Then _
                    Throw New AttributeSizeNotSufficientException(ComboBoxLayerFilter.Text, GetAttributeName("SpeciaalGebouw", "Volgnr"))
                pFeature.Value(fldIndex) = CStr(TextBoxVolgNr.Text)

            End If

            'Commit changes.
            pFeature.Store()

            'Reload current feature into the form.
            LoadFeature(m_OID)

        Catch ex As AttributeSizeNotSufficientException
            Dim title As String = Me.Text
            MsgBox(ex.Message, , title)

        Catch ex As Exception
            Throw ex

        Finally
            If Not pFeatureLayer Is Nothing Then Marshal.ReleaseComObject(pFeatureLayer)
            If Not pQueryFilter Is Nothing Then Marshal.ReleaseComObject(pQueryFilter)
            If Not pFeatureCursor Is Nothing Then Marshal.ReleaseComObject(pFeatureCursor)
            If Not pFeature Is Nothing Then Marshal.ReleaseComObject(pFeature)
            If Not pEditor Is Nothing Then Marshal.ReleaseComObject(pEditor)
            If Not pAnnoLayer Is Nothing Then Marshal.ReleaseComObject(pAnnoLayer)

        End Try
    End Sub

#End Region

End Class
