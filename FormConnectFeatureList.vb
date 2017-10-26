Option Explicit On 
Option Strict On

Imports System.Windows.Forms
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geodatabase

''' -----------------------------------------------------------------------------
''' Project	 : Digipolis.Hydranten.BeheerHydranten
''' Class	 : Hydranten.BeheerHydranten.FormConnectFeatureList
''' 
''' -----------------------------------------------------------------------------
''' <summary>
'''     Show a list of selected multiple features to connect to.
''' </summary>
''' <remarks>
''' </remarks>
''' <history>
''' 	[Kristof Vydt]	16/09/2005	Created
''' 	[Kristof Vydt]	26/09/2005	Form not resizable. Cancel button added.
''' 	[Kristof Vydt]	24/10/2005	Skip &lt;null&gt;-values when filling the selection list.
''' 	[Kristof Vydt]	02/08/2006	Replace some global constant by a function that reads the value from config file.
'''     [Koen Vermeer]  23/06/2008  Conversion to ArcGIS 9.2
''' </history>
''' -----------------------------------------------------------------------------
Public Class FormConnectFeatureList
    Inherits Form

    'A keyword refering to the layer that the user wants to connect to.
    ' Equals "straat" in case at least one street has been selected.
    ' Equals "dok" in case no streets but at least one dock has been selected.
    ' Equals "park" in case no streets and no docks but at least one park has been selected.
    'This keyword determines what attributes are displayed to the user in case of multiple selection.
    'This keyword determines what attributes are copied to the management form that called this connect feature functionality.
    Private m_ConnectionType As String

    'Store the selectionset for determining the selected feature.
    Private m_SelectionSet As ISelectionSet

    'The ArcMap document.
    Private m_MxDocument As IMxDocument

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
    Friend WithEvents ListBoxFeatures As System.Windows.Forms.ListBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ButtonOK As System.Windows.Forms.Button
    Friend WithEvents ButtonAnnuleren As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.ListBoxFeatures = New System.Windows.Forms.ListBox
        Me.ButtonOK = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.ButtonAnnuleren = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'ListBoxFeatures
        '
        Me.ListBoxFeatures.Location = New System.Drawing.Point(8, 24)
        Me.ListBoxFeatures.Name = "ListBoxFeatures"
        Me.ListBoxFeatures.Size = New System.Drawing.Size(208, 160)
        Me.ListBoxFeatures.TabIndex = 0
        '
        'ButtonOK
        '
        Me.ButtonOK.Enabled = False
        Me.ButtonOK.Location = New System.Drawing.Point(96, 192)
        Me.ButtonOK.Name = "ButtonOK"
        Me.ButtonOK.Size = New System.Drawing.Size(48, 24)
        Me.ButtonOK.TabIndex = 1
        Me.ButtonOK.Text = "OK"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(272, 16)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Selecteer één referentie uit volgende lijst:"
        '
        'ButtonAnnuleren
        '
        Me.ButtonAnnuleren.Location = New System.Drawing.Point(152, 192)
        Me.ButtonAnnuleren.Name = "ButtonAnnuleren"
        Me.ButtonAnnuleren.Size = New System.Drawing.Size(64, 24)
        Me.ButtonAnnuleren.TabIndex = 3
        Me.ButtonAnnuleren.Text = "Annuleren"
        '
        'FormConnectFeatureList
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(224, 222)
        Me.ControlBox = False
        Me.Controls.Add(Me.ButtonAnnuleren)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.ButtonOK)
        Me.Controls.Add(Me.ListBoxFeatures)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.Name = "FormConnectFeatureList"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.Text = "Connecteer Feature"
        Me.TopMost = True
        Me.ResumeLayout(False)

    End Sub

#End Region

    <CLSCompliant(False)> _
    Public Sub New( _
            ByVal pSelectionSet As ISelectionSet, _
            ByVal FeatureType As String, _
            ByRef pMxDocument As IMxDocument _
            )

        'Required by the Windows Form Designer.
        MyBase.New()
        InitializeComponent()

        Dim pCursor As ICursor = Nothing
        Dim pFCursor As IFeatureCursor
        Dim pFeature As IFeature

        'Remember the passed arguments for later use.
        m_ConnectionType = FeatureType
        m_SelectionSet = pSelectionSet
        m_MxDocument = pMxDocument

        'Convert the selectionset into a featurecursor that can be looped.
        pSelectionSet.Search(Nothing, False, pCursor)
        pFCursor = CType(pCursor, IFeatureCursor)

        'List each feature of the selectionset in the listbox.
        pFeature = pFCursor.NextFeature
        Me.ListBoxFeatures.ClearSelected()
        While Not pFeature Is Nothing
            AddFeatureRef(pFeature)
            pFeature = pFCursor.NextFeature
        End While

    End Sub

    'Add the features reference data to the listbox.
    Private Sub AddFeatureRef(ByVal pFeature As IFeature)

        Dim pFields As IFields
        Dim FieldIndex As Integer
        Dim ItemText As String

        'Combine the requested attributes into a reference string.
        pFields = pFeature.Fields
        ItemText = ""
        Select Case m_ConnectionType
            Case "straat"
                FieldIndex = pFields.FindField(GetAttributeName("Straatassen", "Straatnaam"))
                If FieldIndex < 0 Then Throw New AttributeNotFoundException(GetLayerName("Straatassen"), GetAttributeName("Straatassen", "Straatnaam"))
                If Not TypeOf pFeature.Value(FieldIndex) Is System.DBNull Then ItemText = CStr(pFeature.Value(FieldIndex))
                FieldIndex = pFields.FindField(GetAttributeName("Straatassen", "Postcode"))
                If FieldIndex < 0 Then Throw New AttributeNotFoundException(GetLayerName("Straatassen"), GetAttributeName("Straatassen", "Postcode"))
                If Not TypeOf pFeature.Value(FieldIndex) Is System.DBNull Then ItemText = ItemText & " - " & CStr(pFeature.Value(FieldIndex))
            Case "dok"
                FieldIndex = pFields.FindField(GetAttributeName("Water", "Naam"))
                If FieldIndex < 0 Then Throw New AttributeNotFoundException(GetLayerName("Water"), GetAttributeName("Water", "Naam"))
                If Not TypeOf pFeature.Value(FieldIndex) Is System.DBNull Then ItemText = CStr(pFeature.Value(FieldIndex))
            Case "park"
                FieldIndex = pFields.FindField(GetAttributeName("Park", "Naam"))
                If FieldIndex < 0 Then Throw New AttributeNotFoundException(GetLayerName("Park"), GetAttributeName("Park", "Naam"))
                If Not TypeOf pFeature.Value(FieldIndex) Is System.DBNull Then ItemText = CStr(pFeature.Value(FieldIndex))
        End Select

        'Add this reference string to the listbox.
        If Len(ItemText) > 0 Then Me.ListBoxFeatures.Items.Add(ItemText)

    End Sub

    Private Sub ButtonOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonOK.Click

        'Determine the selected feature.
        Dim pFeature As IFeature = GetSelectedFeature(m_SelectionSet, Me.ListBoxFeatures.SelectedIndex)

        'Return the selected feature to the management form.
        If Not pFeature Is Nothing Then ReturnFeature(pFeature)

        'Close current list form.
        Me.Close()

    End Sub

    Private Sub ButtonAnnuleren_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonAnnuleren.Click

        'Deactivate functionality.
        ConnectFeatureFunctionality_Deactivate()

        'Close current list form.
        Me.Close()

    End Sub

    Private Sub ListBoxFeatures_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBoxFeatures.SelectedIndexChanged

        'Enable OK button.
        Me.ButtonOK.Enabled = True

        'Flash the selected feature.
        Dim pFeature As IFeature = GetSelectedFeature(m_SelectionSet, Me.ListBoxFeatures.SelectedIndex)
        If Not pFeature Is Nothing Then FlashFeature(pFeature, m_MxDocument)

    End Sub

End Class
