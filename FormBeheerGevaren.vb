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
'Imports ESRI.ArcGIS.Geometry
#End Region

''' -----------------------------------------------------------------------------
''' Project	 : Digipolis.Hydranten.BeheerHydranten
''' Class	 : Hydranten.BeheerHydranten.FormBeheerGevaren
''' 
''' -----------------------------------------------------------------------------
''' <summary>
'''     Form for managing danger objects.
''' </summary>
''' <remarks>
''' </remarks>
''' <history>
''' 	[Kristof Vydt]	16/09/2005	Created
''' 	[Kristof Vydt]	23/09/2005	Remove event handler(s) on close.
'''     [Kristof Vydt]	05/10/2005	Replace ButtonConnect/Copy by CheckBoxConnect/Copy.
''' 	[Kristof Vydt]	24/10/2005	Attribute validation on Straatnaam &amp; Aanduiding.
''' 	                        	Deactivate listeners when loading feature.
''' 	[Kristof Vydt]	27/10/2005	Set DropDownStyle of every ComboBox to List to force the user to select from the list.
''' 	                            FormBeheerGebouwen_Closed added.
'''  	[Kristof Vydt]	23/11/2005	Add optional zoomToFeature parameter to LoadFeature method.
'''  	[Kristof Vydt]	17/07/2006	Close active edit session in StoreAttributeChanges before modifying feature attributes.
''' 	[Kristof Vydt]	02/08/2006	Replace some global constant by a function that reads the value from config file.
''' 	[Kristof Vydt]	18/08/2006	Eliminate private marker element.
''' 	[Kristof Vydt]	22/02/2007	Adopt to XML configuration.
'''     [Koen Vermeer]  23/06/2008  Conversion to ArcGIS 9.2
''' </history>
''' -----------------------------------------------------------------------------
Public NotInheritable Class FormBeheerGevaren
    Inherits System.Windows.Forms.Form
    Implements IConnectFeature 'the form is using the <ConnectFeature> functionality

#Region " Private variables "

    Private m_application As IMxApplication 'hold current ArcMap application
    Private m_document As IMxDocument 'hold current ArcMap document
    'Private m_layer As ILayer 'hydranten layer
    'Private m_workspace As IWorkspace 'workspace of the hydranten
    'Private m_marker As IMarkerElement 'marker for current feature
    Private m_editing As Boolean 'indicated if form is ready for editing
    'Private m_selectionSet As ISelectionSet
    Private m_enumOIDs As IEnumIDs 'enumeration of the feature IDs of the edit set
    Private m_OID As Integer 'the feature/object ID of the current editable feature
    'Private m_copyFrom As IFeature = Nothing 'when functionality "copy attributes from hydrant" is used

#End Region

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
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents ButtonLoad As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents LabelCounter As System.Windows.Forms.Label
    Friend WithEvents LabelTotal As System.Windows.Forms.Label
    Friend WithEvents LabelSeparator As System.Windows.Forms.Label
    Friend WithEvents ButtonNext As System.Windows.Forms.Button
    Friend WithEvents ButtonLast As System.Windows.Forms.Button
    Friend WithEvents ButtonFirst As System.Windows.Forms.Button
    Friend WithEvents ButtonPrevious As System.Windows.Forms.Button
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents ButtonLabelDel As System.Windows.Forms.Button
    Friend WithEvents ButtonLabelAdd As System.Windows.Forms.Button
    Friend WithEvents ButtonSave As System.Windows.Forms.Button
    Friend WithEvents TextBoxPostcode As System.Windows.Forms.TextBox
    Friend WithEvents LabelPostcode As System.Windows.Forms.Label
    Friend WithEvents TextBoxStraatcode As System.Windows.Forms.TextBox
    Friend WithEvents LabelStraatcode As System.Windows.Forms.Label
    Friend WithEvents TextBoxStraatnaam As System.Windows.Forms.TextBox
    Friend WithEvents LabelStraatnaam As System.Windows.Forms.Label
    Friend WithEvents TextBoxAanduiding As System.Windows.Forms.TextBox
    Friend WithEvents LabelAanduiding As System.Windows.Forms.Label
    Friend WithEvents ComboBoxLayerFilter As System.Windows.Forms.ComboBox
    Friend WithEvents ButtonClose As System.Windows.Forms.Button
    Friend WithEvents CheckBoxConnect As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBoxCopy As System.Windows.Forms.CheckBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
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
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.CheckBoxCopy = New System.Windows.Forms.CheckBox
        Me.CheckBoxConnect = New System.Windows.Forms.CheckBox
        Me.ButtonLabelDel = New System.Windows.Forms.Button
        Me.ButtonLabelAdd = New System.Windows.Forms.Button
        Me.ButtonSave = New System.Windows.Forms.Button
        Me.TextBoxPostcode = New System.Windows.Forms.TextBox
        Me.LabelPostcode = New System.Windows.Forms.Label
        Me.TextBoxStraatcode = New System.Windows.Forms.TextBox
        Me.LabelStraatcode = New System.Windows.Forms.Label
        Me.TextBoxStraatnaam = New System.Windows.Forms.TextBox
        Me.LabelStraatnaam = New System.Windows.Forms.Label
        Me.TextBoxAanduiding = New System.Windows.Forms.TextBox
        Me.LabelAanduiding = New System.Windows.Forms.Label
        Me.ButtonClose = New System.Windows.Forms.Button
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.ButtonLoad)
        Me.GroupBox1.Controls.Add(Me.ComboBoxLayerFilter)
        Me.GroupBox1.Location = New System.Drawing.Point(8, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(272, 48)
        Me.GroupBox1.TabIndex = 2
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
        Me.GroupBox3.TabIndex = 51
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
        Me.LabelCounter.Visible = False
        '
        'LabelTotal
        '
        Me.LabelTotal.Location = New System.Drawing.Point(148, 19)
        Me.LabelTotal.Name = "LabelTotal"
        Me.LabelTotal.Size = New System.Drawing.Size(40, 16)
        Me.LabelTotal.TabIndex = 10
        Me.LabelTotal.Text = "#"
        Me.LabelTotal.Visible = False
        '
        'LabelSeparator
        '
        Me.LabelSeparator.Location = New System.Drawing.Point(132, 19)
        Me.LabelSeparator.Name = "LabelSeparator"
        Me.LabelSeparator.Size = New System.Drawing.Size(8, 16)
        Me.LabelSeparator.TabIndex = 9
        Me.LabelSeparator.Text = "/"
        Me.LabelSeparator.Visible = False
        '
        'ButtonNext
        '
        Me.ButtonNext.Enabled = False
        Me.ButtonNext.Location = New System.Drawing.Point(192, 16)
        Me.ButtonNext.Name = "ButtonNext"
        Me.ButtonNext.Size = New System.Drawing.Size(32, 24)
        Me.ButtonNext.TabIndex = 7
        Me.ButtonNext.Text = ">"
        '
        'ButtonLast
        '
        Me.ButtonLast.Enabled = False
        Me.ButtonLast.Location = New System.Drawing.Point(232, 16)
        Me.ButtonLast.Name = "ButtonLast"
        Me.ButtonLast.Size = New System.Drawing.Size(32, 24)
        Me.ButtonLast.TabIndex = 6
        Me.ButtonLast.Text = ">>"
        '
        'ButtonFirst
        '
        Me.ButtonFirst.Enabled = False
        Me.ButtonFirst.Location = New System.Drawing.Point(8, 16)
        Me.ButtonFirst.Name = "ButtonFirst"
        Me.ButtonFirst.Size = New System.Drawing.Size(32, 24)
        Me.ButtonFirst.TabIndex = 4
        Me.ButtonFirst.Text = "<<"
        '
        'ButtonPrevious
        '
        Me.ButtonPrevious.Enabled = False
        Me.ButtonPrevious.Location = New System.Drawing.Point(48, 16)
        Me.ButtonPrevious.Name = "ButtonPrevious"
        Me.ButtonPrevious.Size = New System.Drawing.Size(32, 24)
        Me.ButtonPrevious.TabIndex = 5
        Me.ButtonPrevious.Text = "<"
        '
        'GroupBox2
        '
        Me.GroupBox2.BackColor = System.Drawing.SystemColors.Control
        Me.GroupBox2.Controls.Add(Me.CheckBoxCopy)
        Me.GroupBox2.Controls.Add(Me.CheckBoxConnect)
        Me.GroupBox2.Controls.Add(Me.ButtonLabelDel)
        Me.GroupBox2.Controls.Add(Me.ButtonLabelAdd)
        Me.GroupBox2.Controls.Add(Me.ButtonSave)
        Me.GroupBox2.Controls.Add(Me.TextBoxPostcode)
        Me.GroupBox2.Controls.Add(Me.LabelPostcode)
        Me.GroupBox2.Controls.Add(Me.TextBoxStraatcode)
        Me.GroupBox2.Controls.Add(Me.LabelStraatcode)
        Me.GroupBox2.Controls.Add(Me.TextBoxStraatnaam)
        Me.GroupBox2.Controls.Add(Me.LabelStraatnaam)
        Me.GroupBox2.Controls.Add(Me.TextBoxAanduiding)
        Me.GroupBox2.Controls.Add(Me.LabelAanduiding)
        Me.GroupBox2.Location = New System.Drawing.Point(8, 96)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(272, 200)
        Me.GroupBox2.TabIndex = 52
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Feature"
        '
        'CheckBoxCopy
        '
        Me.CheckBoxCopy.Appearance = System.Windows.Forms.Appearance.Button
        Me.CheckBoxCopy.Enabled = False
        Me.CheckBoxCopy.Location = New System.Drawing.Point(139, 136)
        Me.CheckBoxCopy.Name = "CheckBoxCopy"
        Me.CheckBoxCopy.Size = New System.Drawing.Size(125, 24)
        Me.CheckBoxCopy.TabIndex = 94
        Me.CheckBoxCopy.Text = "Overnemen"
        Me.CheckBoxCopy.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'CheckBoxConnect
        '
        Me.CheckBoxConnect.Appearance = System.Windows.Forms.Appearance.Button
        Me.CheckBoxConnect.Enabled = False
        Me.CheckBoxConnect.Location = New System.Drawing.Point(8, 136)
        Me.CheckBoxConnect.Name = "CheckBoxConnect"
        Me.CheckBoxConnect.Size = New System.Drawing.Size(125, 24)
        Me.CheckBoxConnect.TabIndex = 93
        Me.CheckBoxConnect.Text = "Connecteren"
        Me.CheckBoxConnect.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'ButtonLabelDel
        '
        Me.ButtonLabelDel.Enabled = False
        Me.ButtonLabelDel.Location = New System.Drawing.Point(139, 104)
        Me.ButtonLabelDel.Name = "ButtonLabelDel"
        Me.ButtonLabelDel.Size = New System.Drawing.Size(125, 24)
        Me.ButtonLabelDel.TabIndex = 92
        Me.ButtonLabelDel.Text = "Labels verwijderen"
        '
        'ButtonLabelAdd
        '
        Me.ButtonLabelAdd.Enabled = False
        Me.ButtonLabelAdd.Location = New System.Drawing.Point(8, 104)
        Me.ButtonLabelAdd.Name = "ButtonLabelAdd"
        Me.ButtonLabelAdd.Size = New System.Drawing.Size(125, 24)
        Me.ButtonLabelAdd.TabIndex = 84
        Me.ButtonLabelAdd.Text = "Label plaatsen"
        '
        'ButtonSave
        '
        Me.ButtonSave.Enabled = False
        Me.ButtonSave.Location = New System.Drawing.Point(8, 168)
        Me.ButtonSave.Name = "ButtonSave"
        Me.ButtonSave.Size = New System.Drawing.Size(256, 24)
        Me.ButtonSave.TabIndex = 83
        Me.ButtonSave.Text = "Wijzigingen opslaan"
        '
        'TextBoxPostcode
        '
        Me.TextBoxPostcode.BackColor = System.Drawing.SystemColors.Control
        Me.TextBoxPostcode.Enabled = False
        Me.TextBoxPostcode.ForeColor = System.Drawing.SystemColors.InactiveBorder
        Me.TextBoxPostcode.Location = New System.Drawing.Point(200, 80)
        Me.TextBoxPostcode.Name = "TextBoxPostcode"
        Me.TextBoxPostcode.Size = New System.Drawing.Size(64, 20)
        Me.TextBoxPostcode.TabIndex = 76
        Me.TextBoxPostcode.Text = ""
        '
        'LabelPostcode
        '
        Me.LabelPostcode.Location = New System.Drawing.Point(144, 80)
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
        Me.TextBoxStraatcode.Location = New System.Drawing.Point(72, 80)
        Me.TextBoxStraatcode.Name = "TextBoxStraatcode"
        Me.TextBoxStraatcode.Size = New System.Drawing.Size(64, 20)
        Me.TextBoxStraatcode.TabIndex = 74
        Me.TextBoxStraatcode.Text = ""
        '
        'LabelStraatcode
        '
        Me.LabelStraatcode.Location = New System.Drawing.Point(8, 80)
        Me.LabelStraatcode.Name = "LabelStraatcode"
        Me.LabelStraatcode.Size = New System.Drawing.Size(64, 16)
        Me.LabelStraatcode.TabIndex = 73
        Me.LabelStraatcode.Text = "StraatCode"
        Me.LabelStraatcode.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'TextBoxStraatnaam
        '
        Me.TextBoxStraatnaam.BackColor = System.Drawing.SystemColors.Control
        Me.TextBoxStraatnaam.Enabled = False
        Me.TextBoxStraatnaam.ForeColor = System.Drawing.SystemColors.InactiveBorder
        Me.TextBoxStraatnaam.Location = New System.Drawing.Point(72, 56)
        Me.TextBoxStraatnaam.Name = "TextBoxStraatnaam"
        Me.TextBoxStraatnaam.Size = New System.Drawing.Size(192, 20)
        Me.TextBoxStraatnaam.TabIndex = 72
        Me.TextBoxStraatnaam.Text = ""
        '
        'LabelStraatnaam
        '
        Me.LabelStraatnaam.Location = New System.Drawing.Point(8, 56)
        Me.LabelStraatnaam.Name = "LabelStraatnaam"
        Me.LabelStraatnaam.Size = New System.Drawing.Size(64, 16)
        Me.LabelStraatnaam.TabIndex = 71
        Me.LabelStraatnaam.Text = "Straatnaam"
        Me.LabelStraatnaam.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'TextBoxAanduiding
        '
        Me.TextBoxAanduiding.BackColor = System.Drawing.SystemColors.Control
        Me.TextBoxAanduiding.Enabled = False
        Me.TextBoxAanduiding.Location = New System.Drawing.Point(8, 32)
        Me.TextBoxAanduiding.Name = "TextBoxAanduiding"
        Me.TextBoxAanduiding.Size = New System.Drawing.Size(256, 20)
        Me.TextBoxAanduiding.TabIndex = 50
        Me.TextBoxAanduiding.Text = ""
        '
        'LabelAanduiding
        '
        Me.LabelAanduiding.Location = New System.Drawing.Point(8, 16)
        Me.LabelAanduiding.Name = "LabelAanduiding"
        Me.LabelAanduiding.Size = New System.Drawing.Size(88, 16)
        Me.LabelAanduiding.TabIndex = 49
        Me.LabelAanduiding.Text = "Aanduiding"
        Me.LabelAanduiding.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'ButtonClose
        '
        Me.ButtonClose.Location = New System.Drawing.Point(192, 300)
        Me.ButtonClose.Name = "ButtonClose"
        Me.ButtonClose.Size = New System.Drawing.Size(88, 24)
        Me.ButtonClose.TabIndex = 54
        Me.ButtonClose.Text = "Sluiten"
        '
        'FormBeheerGevaren
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(286, 328)
        Me.Controls.Add(Me.ButtonClose)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.Name = "FormBeheerGevaren"
        Me.Text = "Beheer Gevarenthema's"
        Me.TopMost = True
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
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

        'Custom form initialization.
        'InitializeForm()

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

    Private Sub TextBoxAanduiding_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBoxAanduiding.TextChanged
        If m_editing Then MarkAsChanged(LabelAanduiding)
    End Sub

    Private Sub TextBoxStraatnaam_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBoxStraatnaam.TextChanged
        If m_editing Then MarkAsChanged(LabelStraatnaam)
    End Sub

    Private Sub TextBoxStraatcode_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBoxStraatcode.TextChanged
        If m_editing Then MarkAsChanged(LabelStraatcode)
    End Sub

    Private Sub TextBoxPostcode_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBoxPostcode.TextChanged
        If m_editing Then MarkAsChanged(LabelPostcode)
    End Sub

    Private Sub ButtonSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonSave.Click
        'Store attribute changes if modifications are registered.
        If ModifiedAttribute() Then
            If ValidateAttributeChanges() Then
                StoreAttributeChanges()
            End If
        End If
    End Sub

    Private Sub ButtonClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonClose.Click
        Me.Close()
    End Sub

    Private Sub CheckBoxConnect_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxConnect.CheckedChanged
        Try
            If Me.CheckBoxConnect.Checked Then
                'Activate ConnectFeatureFunctionality
                ConnectFeatureFunctionality_Activate(m_document, Me)
                'Show text in another color for better perception.
                Me.CheckBoxConnect.ForeColor = System.Drawing.Color.BlueViolet
            Else
                'Deactivate ConnectFeatureFunctionality
                ConnectFeatureFunctionality_Deactivate()
                'Restore text to the default color.
                Me.CheckBoxConnect.ForeColor = System.Drawing.Color.Black
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

#End Region

#Region " Overridden form events "

    Protected Overrides Sub OnLoad(ByVal e As System.EventArgs)
        Try

            ' Check availability of configuration.
            If Config Is Nothing Then Throw New ApplicationException("No configuration loaded.")

            ' Get a list of related layers from configuration.
            Dim layerNames As Collection = Config.DangerLayers

            ' If layer is available on map, add it to the list.
            For Each layerName As String In layerNames
                Dim featureLayer As IFeatureLayer = GetFeatureLayer(m_document.FocusMap, layerName)
                If Not featureLayer Is Nothing Then ComboBoxLayerFilter.Items.Add(layerName)
            Next

            ' Select the first item in the list.
            ComboBoxLayerFilter.SelectedIndex = 0

            ' Simulate selection if only one item in list.
            If ComboBoxLayerFilter.Items.Count = 1 Then Call ButtonLoad_Click(Nothing, Nothing)

        Catch ex As Exception
            Throw ex
        End Try

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

    Private Sub FormBeheerGevaren_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        'Be sure to remove remaining eventhandler from the "Connect feature" functionality.
        ConnectFeatureFunctionality_Deactivate()
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
                TextBoxAanduiding.Text = ""
                TextBoxPostcode.Text = ""
                TextBoxStraatcode.Text = ""
                TextBoxStraatnaam.Text = ""
                EnableEditingControls(False)
                EnableNavigationControls(False)
                MsgBox(c_Message_EmptyFeatureSet, MsgBoxStyle.Exclamation)
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
            Dim pLayer As IFeatureLayer = GetFeatureLayer(m_document.FocusMap, ComboBoxLayerFilter.Text)
            Dim pTable As ITable = CType(pLayer, ITable)
            Dim pQueryFilter As IQueryFilter = New QueryFilter
            pQueryFilter.WhereClause = "OBJECTID = " & OID
            Dim pCursor As ICursor = pTable.Search(pQueryFilter, True)
            Dim pRow As IRow = pCursor.NextRow

            'Zoom to the feature and mark it on the map.
            Dim pFeature As IFeature = CType(pRow, IFeature)
            If zoomToFeature Then MarkAndZoomTo(pFeature, m_document, False)

            'Initialize layout of form controls and
            'Show feature attributes in the form controls.
            Dim FieldIndex As Integer
            '- Aanduiding
            LabelAanduiding.ForeColor = Black
            FieldIndex = pRow.Fields.FindField(GetAttributeName("GevarenThema", "Aanduiding"))
            SetEditBoxValue(TextBoxAanduiding, pRow.Value(FieldIndex))
            '- Postcode
            LabelPostcode.ForeColor = Black
            FieldIndex = pRow.Fields.FindField(GetAttributeName("GevarenThema", "Postcode"))
            SetEditBoxValue(TextBoxPostcode, pRow.Value(FieldIndex))
            '- Straatcode
            LabelStraatcode.ForeColor = Black
            FieldIndex = pRow.Fields.FindField(GetAttributeName("GevarenThema", "Straatcode"))
            SetEditBoxValue(TextBoxStraatcode, pRow.Value(FieldIndex))
            '- Straatnaam
            LabelStraatnaam.ForeColor = Black
            FieldIndex = pRow.Fields.FindField(GetAttributeName("GevarenThema", "Straatnaam"))
            SetEditBoxValue(TextBoxStraatnaam, pRow.Value(FieldIndex))

            'Enable editing.
            EnableEditingControls(True)
            m_editing = True

        Catch ex As Exception
            Throw ex
        End Try
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
            LabelCounter.Text = CStr(Counter)
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
            LabelCounter.Text = CStr(Counter + 1)
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
        LabelCounter.Text = CStr(0)
        LoadNextFeature()
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
    Public Sub MarkAsChanged(ByVal SomeControl As Windows.Forms.Control)
        If TypeOf SomeControl Is Windows.Forms.Label Then
            SomeControl.ForeColor = IndianRed
        End If
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

        'Status of Button controls.
        ButtonLabelAdd.Enabled = False 'no such functionality for this form
        ButtonLabelDel.Enabled = False 'no such functionality for this form
        CheckBoxConnect.Enabled = value
        CheckBoxCopy.Enabled = False 'no such functionality for this form
        ButtonSave.Enabled = value
        'ButtonClose.Enabled = True 'always available

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
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Function ModifiedAttribute() As Boolean
        If LabelAanduiding.ForeColor.Equals(IndianRed) Or _
           LabelStraatnaam.ForeColor.Equals(IndianRed) Or _
           LabelStraatcode.ForeColor.Equals(IndianRed) Or _
           LabelPostcode.ForeColor.Equals(IndianRed) Then

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
    ''' 	[Kristof Vydt]	24/10/2005	Aanduiding &amp; Straatnaam must be filled.
    ''' 	[Kristof Vydt]	24/10/2005	Application exceptions using global constant messages.
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Function ValidateAttributeChanges() As Boolean
        Try
            '- Aanduiding
            If Len(Trim(Me.TextBoxAanduiding.Text)) = 0 Then _
                Throw New ApplicationException(c_Message_GevaarAanduidingIsEmpty)
            '- ID
            '- Postcode
            '- Straatcode
            '- Straatnaam
            If Len(Trim(Me.TextBoxStraatnaam.Text)) = 0 Then _
                Throw New ApplicationException(c_Message_GevaarStraatnaamIsEmpty)
            'Valid.
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.OKOnly, "Validatie van attributen")
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
    ''' 	[Kristof Vydt]	02/08/2006	Replace some global constant by a function that reads the value from config file.
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub StoreAttributeChanges()

        Dim pFeatureLayer As IFeatureLayer = Nothing
        Dim pQueryFilter As IQueryFilter = Nothing
        Dim pFeatureCursor As IFeatureCursor = Nothing
        Dim pFeature As IFeature = Nothing
        Dim pEditor As IEditor2 = Nothing

        Try
            'Get the feature that is modified.
            pFeatureLayer = GetFeatureLayer(m_document.FocusMap, ComboBoxLayerFilter.Text)
            If pFeatureLayer Is Nothing Then Throw New LayerNotFoundException(ComboBoxLayerFilter.Text)
            pQueryFilter = New QueryFilter
            pQueryFilter.WhereClause = "OBJECTID = " & m_OID
            pFeatureCursor = pFeatureLayer.FeatureClass.Search(pQueryFilter, True)
            pFeature = pFeatureCursor.NextFeature

            'Make sure that there is at least one feature.
            If pFeature Is Nothing Then
                MsgBox("Cannot save changed because the feature with OBJECTID " & CStr(m_OID) & " could not be found in the feature class.", _
                    MsgBoxStyle.Exclamation, "Wijzigingen opslaan.")
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
            Dim AttributeIndex As Integer
            '- Aanduiding
            If LabelAanduiding.ForeColor.Equals(IndianRed) Then
                AttributeIndex = pFeature.Fields.FindField(GetAttributeName("GevarenThema", "Aanduiding"))
                If Len(CStr(TextBoxAanduiding.Text)) > pFeature.Fields.Field(AttributeIndex).Length Then _
                    Throw New AttributeSizeNotSufficientException(ComboBoxLayerFilter.Text, GetAttributeName("GevarenThema", "Aanduiding"))
                pFeature.Value(AttributeIndex) = CStr(TextBoxAanduiding.Text)
            End If
            '- Postcode
            If LabelPostcode.ForeColor.Equals(IndianRed) Then
                AttributeIndex = pFeature.Fields.FindField(GetAttributeName("SpeciaalGebouw", "Postcode"))
                If Len(CStr(TextBoxPostcode.Text)) > pFeature.Fields.Field(AttributeIndex).Length Then _
                    Throw New AttributeSizeNotSufficientException(ComboBoxLayerFilter.Text, GetAttributeName("GevarenThema", "Postcode"))
                pFeature.Value(AttributeIndex) = CStr(TextBoxPostcode.Text)
            End If
            '- Straatcode
            If LabelStraatcode.ForeColor.Equals(IndianRed) Then
                AttributeIndex = pFeature.Fields.FindField(GetAttributeName("SpeciaalGebouw", "Straatcode"))
                If Len(CStr(TextBoxStraatcode.Text)) > pFeature.Fields.Field(AttributeIndex).Length Then _
                    Throw New AttributeSizeNotSufficientException(ComboBoxLayerFilter.Text, GetAttributeName("GevarenThema", "Straatcode"))
                pFeature.Value(AttributeIndex) = CStr(TextBoxStraatcode.Text)
            End If
            '- Straatnaam
            If LabelStraatnaam.ForeColor.Equals(IndianRed) Then
                AttributeIndex = pFeature.Fields.FindField(GetAttributeName("SpeciaalGebouw", "Straatnaam"))
                If Len(CStr(TextBoxStraatnaam.Text)) > pFeature.Fields.Field(AttributeIndex).Length Then _
                    Throw New AttributeSizeNotSufficientException(ComboBoxLayerFilter.Text, GetAttributeName("GevarenThema", "Straatnaam"))
                pFeature.Value(AttributeIndex) = CStr(TextBoxStraatnaam.Text)
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
        End Try
    End Sub

#End Region

End Class
