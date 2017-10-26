Option Explicit On 
Option Strict On

#Region " Imports namespaces "
Imports ADODB
Imports System.IO
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.Framework

#End Region

''' -----------------------------------------------------------------------------
''' Project	 : Digipolis.Hydranten.BeheerHydranten
''' Class	 : Hydranten.BeheerHydranten.FormIndexGebouwen
''' 
''' -----------------------------------------------------------------------------
''' <summary>
'''     GUI for generating building index pages.
''' </summary>
''' <remarks>
''' </remarks>
''' <history>
''' 	[Kristof Vydt]	16/09/2005	Created
''' 	[Kristof Vydt]	26/09/2005	Open *.dot as template, and not the document itself.
'''     [Kristof Vydt]  13/07/2006  Update the name of some variables/constants.
''' 	[Kristof Vydt]	02/08/2006	Replace some global constant by a function that reads the value from config file.
''' 	[Kristof Vydt]	21/03/2007	Add type ordered index.
''' 	[Kristof Vydt]	28/03/2007	Group types into categories for the type ordered index.
'''     [Koen Vermeer]  23/06/2008  Conversion to ArcGIS 9.2
''' </history>
''' -----------------------------------------------------------------------------
Public Class FormIndexGebouwen
    Inherits System.Windows.Forms.Form

#Region " Private variables & constants "

    'Constants.
    Private Const c_IdFieldName As String = "VOLGNUMMER" 'name of field with building id
    Private Const c_NameFieldName As String = "NAAM" 'name of field with building name
    Private Const c_RefListFieldName As String = "REFERENTIE" 'name of field with list of kwadrant references
    Private Const c_StreetFieldName As String = "STRAATNAAM" 'name of field with streetname
    Private Const c_DescriptionFieldName As String = "AANDUIDING" 'name of field with building position description
    Private Const c_TypeFieldName As String = "TYPE" 'name of field with building type
    Private Const c_CategoryFieldName As String = "CATEGORY" 'name of field with building category
    Private Const c_LookupTableName As String = "straatlijst" 'name of table in personal geodatabase for street lookup info

    'Locals.
    Private m_application As IMxApplication 'set by constructor
    Private m_document As IMxDocument 'set by constructor
    Private m_IdMaxLength As Integer 'maximum number of characters available for a building id
    Private m_NameMaxLength As Integer 'maximum number of characters available for a building name
    Private m_RefListMaxLength As Integer 'maximum number of characters available for the building reference list (a ";"-separated list of kwadrants, each consisting of 2 char raster page + 1 char kwadrant letter)
    Private m_StreetMaxLength As Integer 'maximum number of characters available for a building streetname
    Private m_DescMaxLength As Integer 'maximum number of characters available for a building description
    'TODO: waarde m_typeMaxLength & m_CategoryMaxLength uitlezen met code ?
    Private m_TypeMaxLength As Integer = 50 'maximum number of characters available for a building type
    Private m_CategoryMaxLength As Integer = 50 'maximum number of characters available for a building type

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
    Friend WithEvents ButtonCancel As System.Windows.Forms.Button
    Friend WithEvents ButtonOK As System.Windows.Forms.Button
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents ProgressBar2 As System.Windows.Forms.ProgressBar
    Friend WithEvents LabelProgressMessage As System.Windows.Forms.Label
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox2 As System.Windows.Forms.CheckBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents CheckBox3 As System.Windows.Forms.CheckBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.CheckBox3 = New System.Windows.Forms.CheckBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.CheckBox2 = New System.Windows.Forms.CheckBox
        Me.CheckBox1 = New System.Windows.Forms.CheckBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.ButtonCancel = New System.Windows.Forms.Button
        Me.ButtonOK = New System.Windows.Forms.Button
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.ProgressBar2 = New System.Windows.Forms.ProgressBar
        Me.LabelProgressMessage = New System.Windows.Forms.Label
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.CheckBox3)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.CheckBox2)
        Me.GroupBox1.Controls.Add(Me.CheckBox1)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Location = New System.Drawing.Point(8, 8)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(280, 88)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Sortering"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(56, 64)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(176, 16)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "Op type van het gebouw."
        '
        'CheckBox3
        '
        Me.CheckBox3.Checked = True
        Me.CheckBox3.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox3.Location = New System.Drawing.Point(40, 64)
        Me.CheckBox3.Name = "CheckBox3"
        Me.CheckBox3.Size = New System.Drawing.Size(16, 16)
        Me.CheckBox3.TabIndex = 6
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(56, 40)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(176, 16)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Op naam van het gebouw."
        '
        'CheckBox2
        '
        Me.CheckBox2.Checked = True
        Me.CheckBox2.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox2.Location = New System.Drawing.Point(40, 40)
        Me.CheckBox2.Name = "CheckBox2"
        Me.CheckBox2.Size = New System.Drawing.Size(16, 16)
        Me.CheckBox2.TabIndex = 3
        '
        'CheckBox1
        '
        Me.CheckBox1.Checked = True
        Me.CheckBox1.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox1.Location = New System.Drawing.Point(40, 16)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(16, 16)
        Me.CheckBox1.TabIndex = 2
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(56, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(176, 16)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Op volgnummer van het gebouw."
        '
        'ButtonCancel
        '
        Me.ButtonCancel.Location = New System.Drawing.Point(208, 184)
        Me.ButtonCancel.Name = "ButtonCancel"
        Me.ButtonCancel.TabIndex = 4
        Me.ButtonCancel.Text = "Annuleren"
        '
        'ButtonOK
        '
        Me.ButtonOK.Location = New System.Drawing.Point(120, 184)
        Me.ButtonOK.Name = "ButtonOK"
        Me.ButtonOK.TabIndex = 3
        Me.ButtonOK.Text = "OK"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.ProgressBar2)
        Me.GroupBox2.Controls.Add(Me.LabelProgressMessage)
        Me.GroupBox2.Controls.Add(Me.ProgressBar1)
        Me.GroupBox2.Location = New System.Drawing.Point(8, 104)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(280, 72)
        Me.GroupBox2.TabIndex = 9
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Vooruitgang"
        '
        'ProgressBar2
        '
        Me.ProgressBar2.Enabled = False
        Me.ProgressBar2.Location = New System.Drawing.Point(4, 51)
        Me.ProgressBar2.Name = "ProgressBar2"
        Me.ProgressBar2.Size = New System.Drawing.Size(272, 16)
        Me.ProgressBar2.TabIndex = 9
        '
        'LabelProgressMessage
        '
        Me.LabelProgressMessage.Location = New System.Drawing.Point(8, 16)
        Me.LabelProgressMessage.Name = "LabelProgressMessage"
        Me.LabelProgressMessage.Size = New System.Drawing.Size(264, 16)
        Me.LabelProgressMessage.TabIndex = 8
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Enabled = False
        Me.ProgressBar1.Location = New System.Drawing.Point(4, 32)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(272, 16)
        Me.ProgressBar1.TabIndex = 7
        '
        'FormIndexGebouwen
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(293, 216)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.ButtonCancel)
        Me.Controls.Add(Me.ButtonOK)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.Name = "FormIndexGebouwen"
        Me.Text = "Index Gebouwen"
        Me.TopMost = True
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region " Overloaded constructor "
    <CLSCompliant(False)> _
    Public Sub New(ByRef pMxApplication As IMxApplication)
        MyBase.New()

        'Initialize locals.
        m_application = pMxApplication
        m_document = CType(CType(m_application, IApplication).Document, IMxDocument)

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        'InitializeForm()

    End Sub
#End Region

#Region " Form controls events "

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        'Enable OK button only if at least one checkbox is checked.
        Me.ButtonOK.Enabled = (Me.CheckBox1.Checked Or Me.CheckBox2.Checked Or Me.CheckBox3.Checked)
    End Sub

    Private Sub CheckBox2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox2.CheckedChanged
        'Enable OK button only if at least one checkbox is checked.
        Me.ButtonOK.Enabled = (Me.CheckBox1.Checked Or Me.CheckBox2.Checked Or Me.CheckBox3.Checked)
    End Sub

    Private Sub CheckBox3_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckBox3.CheckedChanged
        'Enable OK button only if at least one checkbox is checked.
        Me.ButtonOK.Enabled = (Me.CheckBox1.Checked Or Me.CheckBox2.Checked Or Me.CheckBox3.Checked)
    End Sub

    Private Sub ButtonOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonOK.Click

        'Generating index files.
        Dim lookupRS As ADODB.Recordset = Nothing
        Try
            ' Delete existing export file(s).
            If File.Exists(c_FilePath_IndexGebouwNummer) Then File.Delete(c_FilePath_IndexGebouwNummer)
            If File.Exists(c_FilePath_IndexGebouwNaam) Then File.Delete(c_FilePath_IndexGebouwNaam)
            If File.Exists(c_FilePath_IndexGebouwType) Then File.Delete(c_FilePath_IndexGebouwType)

            'Maak een lege ADODB recordset met 2 velden van gepaste grootte.
            lookupRS = CreateLookupRecordset()

            'Vul de recordset met straatnamen en kwadranten.
            UpdateLookupInfo(lookupRS, AddressOf OnShowProgress1, AddressOf OnSetMaxProgress1)

            If c_DebugStatus Then MsgBox(lookupRS.RecordCount & " records in lookup recordset", , "Gebouwen Index")

            ' Compose index data ordered by number and export to txt file.
            If Me.CheckBox1.Checked Then
                lookupRS.Sort = c_IdFieldName
                ExportRecordset(lookupRS, _
                    AddressOf OnShowProgress2, _
                    AddressOf OnSetMaxProgress2, _
                    "nksa", c_FilePath_IndexGebouwNummer)
            End If

            ' Compose index data ordered by name and export to txt file.
            If Me.CheckBox2.Checked Then
                lookupRS.Sort = c_NameFieldName
                ExportRecordset(lookupRS, _
                    AddressOf OnShowProgress2, _
                    AddressOf OnSetMaxProgress2, _
                    "snka", c_FilePath_IndexGebouwNaam)
            End If

            ' Compose index data ordered by type and export to txt file.
            If Me.CheckBox3.Checked Then
                lookupRS.Sort = c_CategoryFieldName & "," & c_NameFieldName
                ExportRecordset(lookupRS, _
                    AddressOf OnShowProgress2, _
                    AddressOf OnSetMaxProgress2, _
                    "cskna", c_FilePath_IndexGebouwType)
            End If

            'Progress monitor info
            OnShowProgress2(-2, "Export succesvol beëindigd")

        Catch ex As Exception
            'Progress monitor info
            OnShowProgress2(-2, "Export niet succesvol beëindigd")
            Throw ex
        Finally
            If Not lookupRS Is Nothing Then lookupRS = Nothing
        End Try

        Dim oWord As Word.ApplicationClass = Nothing
        Dim dotPath As String

        ' Open Word template and run macro.
        Try
            oWord = New Word.ApplicationClass
            dotPath = oWord.Application.Options.DefaultFilePath(Word.WdDefaultFilePath.wdWorkgroupTemplatesPath)
            dotPath &= "\" & c_FileName_WordTemplateIndexGebouwen
            'oWord.Documents.Open(CType(dotPath, System.Object)) 'to open as a regular document
            oWord.Documents.Add(CType(dotPath, System.Object)) 'to make a new document based on template
            oWord.Run(c_MacroName_IndexGebouwen)
            oWord.Visible = True

        Catch ex As Exception
            Throw ex

        Finally
            If Not oWord Is Nothing Then oWord = Nothing

        End Try

        ' Close form if export finished successfully.
        Me.Close()

    End Sub

    Private Sub ButtonCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonCancel.Click
        Close()
    End Sub

#End Region

#Region " Form events "

#End Region

#Region " Progress information management "

    Public Delegate Sub ShowProgress(ByVal progress As Integer, ByVal message As String)

    Public Delegate Sub SetMaxProgress(ByVal max As Integer)

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     OnShowProgress1:
    '''     OnShowProgress2:
    '''     Set progress information
    ''' </summary>
    ''' <param name="progress">
    '''     The value the progress bar is set to.
    ''' </param>
    ''' <param name="message">
    '''     Some text info that is displayed together with the progress bar.
    ''' </param>
    ''' <remarks>
    '''     A progress value of -1 is interpreted as the minimum value. 
    '''     A progress value of -2 is interpreted as the maximum value. 
    '''     A progress value higher than the maximum is reduced to the maximum value.
    '''     A progress value lower than the minimum is reduced to the minimum value.
    '''     Use an empty string as message to clear the text info.
    '''     Use Nothing as message to leave the text info unchanged.
    '''     There is a delegate : ShowProgress.
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub OnShowProgress1(ByVal progress As Integer, ByVal message As String)
        Select Case progress
            Case -1
                ProgressBar1.Value = ProgressBar1.Minimum
            Case -2
                ProgressBar1.Value = ProgressBar1.Maximum
            Case Else
                If progress > ProgressBar1.Maximum Then
                    ProgressBar1.Value = ProgressBar1.Maximum
                ElseIf progress < ProgressBar1.Minimum Then
                    ProgressBar1.Value = ProgressBar1.Minimum
                Else
                    ProgressBar1.Value = progress
                End If
        End Select
        ProgressBar1.Refresh()
        If Not message Is Nothing Then
            LabelProgressMessage.Text = message
            If c_DebugStatus Then LabelProgressMessage.Text &= " [" & CStr(ProgressBar1.Value) & "/" & CStr(ProgressBar1.Maximum) & "]"
            LabelProgressMessage.Refresh()
        End If
    End Sub

    Public Sub OnShowProgress2(ByVal progress As Integer, ByVal message As String)
        Select Case progress
            Case -1
                ProgressBar2.Value = ProgressBar2.Minimum
            Case -2
                ProgressBar2.Value = ProgressBar2.Maximum
            Case Else
                If progress > ProgressBar2.Maximum Then
                    ProgressBar2.Value = ProgressBar2.Maximum
                ElseIf progress < ProgressBar2.Minimum Then
                    ProgressBar2.Value = ProgressBar2.Minimum
                Else
                    ProgressBar2.Value = progress
                End If
        End Select
        ProgressBar2.Refresh()
        If Not message Is Nothing Then
            LabelProgressMessage.Text = message
            If c_DebugStatus Then LabelProgressMessage.Text &= " [" & CStr(ProgressBar2.Value) & "/" & CStr(ProgressBar2.Maximum) & "]"
            LabelProgressMessage.Refresh()
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     OnSetMaxProgress1:
    '''     OnSetMaxProgress2:
    '''     Set minimum and maximum value of the progress bar control.
    ''' </summary>
    ''' <param name="max">
    '''     Maximum value for the progress bar.
    ''' </param>
    ''' <remarks>
    '''     There is a delegate : SetMaxProgress.
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub OnSetMaxProgress1(ByVal max As Integer)
        ProgressBar1.Minimum = 0
        ProgressBar1.Maximum = max
        ProgressBar1.Refresh()
    End Sub

    Public Sub OnSetMaxProgress2(ByVal max As Integer)
        ProgressBar2.Minimum = 0
        ProgressBar2.Maximum = max
        ProgressBar2.Refresh()
    End Sub

#End Region

#Region " Utility procedures "

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Return a recordset to hold links between buildings to kwadrants.
    ''' </summary>
    ''' <returns>
    '''     An empty recordset with the required structure.
    ''' </returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    '''     [Kristof Vydt]  13/07/2006  Update the name of some variables/constants.
    '''                                 Determine the field sizes by reading the corresponding table field.
    ''' 	[Kristof Vydt]	02/08/2006	Replace some global constant by a function that reads the value from config file.
    ''' 	[Kristof Vydt]	21/03/2007	Add type field.
    ''' 	[Kristof Vydt]	28/03/2007	Add category field.
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Function CreateLookupRecordset() As ADODB.Recordset

        Dim rs As New ADODB.Recordset
        Dim pStandaloneTable As IStandaloneTable
        Dim pTableFields As ITableFields

        Try

            'Read available field lengths from the street lookup table.
            pStandaloneTable = New StandaloneTable
            pStandaloneTable.Table = GetTable(c_lookupTableName, m_document.FocusMap)
            pTableFields = CType(pStandaloneTable, ITableFields)
            With pTableFields
                If .FindField(c_refListFieldName) < 0 Then Throw New AttributeNotFoundException(c_lookupTableName, c_refListFieldName)
                m_refListMaxLength = .Field(.FindField(c_refListFieldName)).Length
            End With

            'Read available field lengths from the building feature class.
            Dim pFeatureLayer As IFeatureLayer
            pFeatureLayer = GetFeatureLayer(m_document.FocusMap, GetLayerName("SpeciaalGebouw"))
            If pFeatureLayer Is Nothing Then Throw New LayerNotFoundException(GetLayerName("SpeciaalGebouw"))
            With pFeatureLayer.FeatureClass.Fields()
                If .FindField(GetAttributeName("SpeciaalGebouw", "Volgnr")) < 0 Then Throw New AttributeNotFoundException(GetLayerName("SpeciaalGebouw"), GetAttributeName("SpeciaalGebouw", "Volgnr"))
                m_idMaxLength = .Field(.FindField(GetAttributeName("SpeciaalGebouw", "Volgnr"))).Length
                If .FindField(GetAttributeName("SpeciaalGebouw", "Naam")) < 0 Then Throw New AttributeNotFoundException(GetLayerName("SpeciaalGebouw"), GetAttributeName("SpeciaalGebouw", "Naam"))
                m_nameMaxLength = .Field(.FindField(GetAttributeName("SpeciaalGebouw", "Naam"))).Length
                If .FindField(GetAttributeName("SpeciaalGebouw", "Straatnaam")) < 0 Then Throw New AttributeNotFoundException(GetLayerName("SpeciaalGebouw"), GetAttributeName("SpeciaalGebouw", "Straatnaam"))
                m_streetMaxLength = .Field(.FindField(GetAttributeName("SpeciaalGebouw", "Straatnaam"))).Length
                If .FindField(GetAttributeName("SpeciaalGebouw", "Aanduiding")) < 0 Then Throw New AttributeNotFoundException(GetLayerName("SpeciaalGebouw"), GetAttributeName("SpeciaalGebouw", "Aanduiding"))
                m_descMaxLength = .Field(.FindField(GetAttributeName("SpeciaalGebouw", "Aanduiding"))).Length
            End With

            'Append each individual field.
            rs.Fields.Append(c_IdFieldName, DataTypeEnum.adInteger, m_idMaxLength)
            rs.Fields.Append(c_NameFieldName, DataTypeEnum.adChar, m_nameMaxLength)
            rs.Fields.Append(c_RefListFieldName, DataTypeEnum.adChar, m_refListMaxLength)
            rs.Fields.Append(c_StreetFieldName, DataTypeEnum.adChar, m_streetMaxLength)
            rs.Fields.Append(c_DescriptionFieldName, DataTypeEnum.adChar, m_descMaxLength)
            rs.Fields.Append(c_TypeFieldName, DataTypeEnum.adChar, m_typeMaxLength)
            rs.Fields.Append(c_CategoryFieldName, DataTypeEnum.adChar, m_CategoryMaxLength)

            'Open the recordset and return it.
            rs.Open()
            Return rs

        Catch ex As Exception
            Throw ex

        Finally
            If Not rs Is Nothing Then rs = Nothing

        End Try

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Fill recordset with buildings and their corresponding kwadrants.
    ''' </summary>
    ''' <param name="rs">
    '''     Recordset to fill.
    ''' </param>
    ''' <param name="ProgressDel">
    '''     Delegate for progress bar update procedure.
    ''' </param>
    ''' <param name="MaxProgressDel">
    '''     Delegate for progress bar maximum procedure.
    ''' </param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    ''' 	[Kristof Vydt]	02/08/2006	Replace some global constant by a function that reads the value from config file.
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub UpdateLookupInfo( _
        ByRef rs As ADODB.Recordset, _
        ByVal ProgressDel As ShowProgress, _
        ByVal MaxProgressDel As SetMaxProgress)

        'Variables.
        Dim progress As Integer
        Dim sectorCode As String = ""
        Dim rasterLayer As IFeatureLayer = Nothing
        Dim queryFilter As IQueryFilter = Nothing
        Dim rasterCursor As IFeatureCursor = Nothing
        Dim rasterFeature As IFeature = Nothing
        Dim arrayKwadranten As IGeometry() = Nothing
        Dim i As Integer 'loop index
        Dim kwadrantReference As String = ""
        Dim pageFieldIndex As Integer
        Dim spatialFilter As ISpatialFilter = Nothing
        Dim gebouwLayer As IFeatureLayer = Nothing
        Dim gebouwCursor As IFeatureCursor = Nothing

        Try

            'Progress monitor
            progress = -1 'empty progress bar
            ProgressDel(progress, "Reference update begint ...")

            'Determine current sector.
            sectorCode = GetSectorCode(m_document)

            'Get raster layer.
            rasterLayer = GetFeatureLayer(m_document.FocusMap, GetLayerName("Raster"))
            If rasterLayer Is Nothing Then Throw New LayerNotFoundException(GetLayerName("Raster"))
            pageFieldIndex = rasterLayer.FeatureClass.Fields.FindField("BLZ_" & sectorCode)
            If pageFieldIndex < 0 Then Throw New AttributeNotFoundException(GetLayerName("Raster"), "BLZ_" & sectorCode)

            'Get building layer.
            gebouwLayer = GetFeatureLayer(m_document.FocusMap, GetLayerName("SpeciaalGebouw"))
            If gebouwLayer Is Nothing Then Throw New LayerNotFoundException(GetLayerName("SpeciaalGebouw"))

            'Build a cursor of all raster features for current sector.
            queryFilter = New queryFilter
            queryFilter.WhereClause = "BLZ_" & sectorCode & " <> ''"
            rasterCursor = CType(rasterLayer, IFeatureLayer2).FeatureClass.Search(queryFilter, True)
            rasterFeature = rasterCursor.NextFeature()
            If rasterFeature Is Nothing Then Exit Sub

            'Set maximum progress value
            i = 0
            While Not rasterFeature Is Nothing
                i += 5
                rasterFeature = rasterCursor.NextFeature
            End While
            MaxProgressDel(i)

            'Loop through raster cursor.
            rasterCursor = CType(rasterLayer, IFeatureLayer2).FeatureClass.Search(queryFilter, True)
            rasterFeature = rasterCursor.NextFeature()
            While Not rasterFeature Is Nothing

                'Progress monitor
                progress += 1
                ProgressDel(progress, "Reference update: " & CStr(rasterFeature.Value(pageFieldIndex)))

                'Split each raster in 4 kwadrants.
                SplitFeatureEnvelopeIntoKwadrants(rasterFeature, Nothing, arrayKwadranten)
                For i = 0 To arrayKwadranten.Length - 1

                    'Get reference of current kwadrant.
                    Select Case i
                        Case 0 'Kwadrant A
                            kwadrantReference = CStr(rasterFeature.Value(pageFieldIndex)) & "A" 'keep 0 at the beginning
                            'kwadrantReference = CStr(CInt(rasterFeature.Value(pageFieldIndex))) & "A" 'remove 0 at the beginning
                        Case 1 'Kwadrant B
                            kwadrantReference = CStr(rasterFeature.Value(pageFieldIndex)) & "B" 'keep 0 at the beginning
                            'kwadrantReference = CStr(CInt(rasterFeature.Value(pageFieldIndex))) & "B" 'remove 0 at the beginning
                        Case 2 'Kwadrant C
                            kwadrantReference = CStr(rasterFeature.Value(pageFieldIndex)) & "C" 'keep 0 at the beginning
                            'kwadrantReference = CStr(CInt(rasterFeature.Value(pageFieldIndex))) & "C" 'remove 0 at the beginning
                        Case 3 'Kwadrant D
                            kwadrantReference = CStr(rasterFeature.Value(pageFieldIndex)) & "D" 'keep 0 at the beginning
                            'kwadrantReference = CStr(CInt(rasterFeature.Value(pageFieldIndex))) & "D" 'remove 0 at the beginning
                    End Select

                    'Progress monitor
                    progress += 1
                    ProgressDel(progress, "Reference update: " & kwadrantReference)

                    'Get all buildings for current kwadrant.
                    spatialFilter = New spatialFilter
                    spatialFilter.Geometry = arrayKwadranten(i)
                    spatialFilter.SpatialRel = esriSpatialRelEnum.esriSpatialRelIntersects
                    spatialFilter.WhereClause = ""
                    gebouwCursor = gebouwLayer.Search(spatialFilter, Nothing)

                    'Add each selected building to the recordset.
                    If Not gebouwCursor Is Nothing Then _
                        AddCursorToLookupRecordset(gebouwCursor, rs, kwadrantReference)

                Next 'Loop to next kwadrant of current raster feature.

                'Loop to next raster feature.
                rasterFeature = rasterCursor.NextFeature()
            End While

        Catch ex As Exception
            Throw ex
        Finally
            If Not rasterLayer Is Nothing Then Marshal.ReleaseComObject(rasterLayer)
            If Not queryFilter Is Nothing Then Marshal.ReleaseComObject(queryFilter)
            If Not rasterCursor Is Nothing Then Marshal.ReleaseComObject(rasterCursor)
            If Not rasterFeature Is Nothing Then Marshal.ReleaseComObject(rasterFeature)
            If Not spatialFilter Is Nothing Then Marshal.ReleaseComObject(spatialFilter)
            If Not gebouwLayer Is Nothing Then Marshal.ReleaseComObject(gebouwLayer)
            If Not gebouwCursor Is Nothing Then Marshal.ReleaseComObject(gebouwCursor)
        End Try
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Register a set of buildings with a kwadrant reference, in the recordset.
    ''' </summary>
    ''' <param name="pCursor">
    '''     A cursor with building features.
    ''' </param>
    ''' <param name="pRecordset">
    '''     The recordset to add to.
    ''' </param>
    ''' <param name="kwadrantReference">
    '''     The single kwadrant reference (raster page &amp; singel kwadrant letter).
    ''' </param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    ''' 	[Kristof Vydt]	02/08/2006	Replace some global constant by a function that reads the value from config file.
    ''' 	[Kristof Vydt]	21/03/2007	Also copy building type.
    ''' 	[Kristof Vydt]	28/03/2007	Determine building category.
    ''' 	[Kristof Vydt]	19/04/2007	Initialise category to empty string. Add "Hoogbouw".
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub AddCursorToLookupRecordset( _
            ByVal pCursor As IFeatureCursor, _
            ByRef pRecordset As ADODB.Recordset, _
            ByVal kwadrantReference As String)

        Try
            Dim pFeature As IFeature
            Dim fieldIndex1 As Integer 'Volgnummer
            Dim fieldValue1 As Integer
            Dim fieldIndex2 As Integer 'Naam
            Dim fieldValue2 As String
            Dim fieldIndex4 As Integer 'Straatnaam
            Dim fieldValue4 As String
            Dim fieldIndex5 As Integer 'Aanduiding
            Dim fieldValue5 As String
            Dim fieldIndex6 As Integer 'Type
            Dim fieldValue6 As String
            Dim fieldValue7 As String 'Category
            Dim searchCondition As String

            'Get all field indices.
            pFeature = pCursor.NextFeature
            If Not pFeature Is Nothing Then
                fieldIndex1 = pFeature.Fields.FindField(GetAttributeName("SpeciaalGebouw", "Volgnr"))
                If fieldIndex1 < 0 Then Throw New AttributeNotFoundException(GetLayerName("SpeciaalGebouw"), GetAttributeName("SpeciaalGebouw", "Volgnr"))
                fieldIndex2 = pFeature.Fields.FindField(GetAttributeName("SpeciaalGebouw", "Naam"))
                If fieldIndex2 < 0 Then Throw New AttributeNotFoundException(GetLayerName("SpeciaalGebouw"), GetAttributeName("SpeciaalGebouw", "Naam"))
                fieldIndex4 = pFeature.Fields.FindField(GetAttributeName("SpeciaalGebouw", "Straatnaam"))
                If fieldIndex4 < 0 Then Throw New AttributeNotFoundException(GetLayerName("SpeciaalGebouw"), GetAttributeName("SpeciaalGebouw", "Straatnaam"))
                fieldIndex5 = pFeature.Fields.FindField(GetAttributeName("SpeciaalGebouw", "Aanduiding"))
                If fieldIndex5 < 0 Then Throw New AttributeNotFoundException(GetLayerName("SpeciaalGebouw"), GetAttributeName("SpeciaalGebouw", "Aanduiding"))
                fieldIndex6 = pFeature.Fields.FindField(GetAttributeName("SpeciaalGebouw", "GebouwType"))
                If fieldIndex6 < 0 Then Throw New AttributeNotFoundException(GetLayerName("SpeciaalGebouw"), GetAttributeName("SpeciaalGebouw", "GebouwType"))
            End If

            'Loop through the complete cursor.
            While Not pFeature Is Nothing

                'Read all attributes that need to be stores in the recordset.
                fieldValue1 = 0  'Volgnummer
                fieldValue2 = String.Empty 'Naam
                fieldValue4 = String.Empty 'Straatnaam
                fieldValue5 = String.Empty 'Aanduiding
                fieldValue6 = String.Empty 'Type
                fieldValue7 = String.Empty 'Categorie
                If Not TypeOf pFeature.Value(fieldIndex1) Is System.DBNull Then fieldValue1 = CInt(pFeature.Value(fieldIndex1))
                If Not TypeOf pFeature.Value(fieldIndex2) Is System.DBNull Then fieldValue2 = CStr(pFeature.Value(fieldIndex2))
                If Not TypeOf pFeature.Value(fieldIndex4) Is System.DBNull Then fieldValue4 = CStr(pFeature.Value(fieldIndex4))
                If Not TypeOf pFeature.Value(fieldIndex5) Is System.DBNull Then fieldValue5 = CStr(pFeature.Value(fieldIndex5))
                If Not TypeOf pFeature.Value(fieldIndex6) Is System.DBNull Then

                    ' Derive building type label from type code.
                    Dim codeValue As String = CStr(pFeature.Value(fieldIndex6))
                    Dim pDomain As New CodedValueDomainManager(pFeature, "GebouwType")
                    fieldValue6 = pDomain.CodeName(codeValue)

                    ' Derive building category label from type code.
                    Select Case codeValue
                        Case "1", "9"
                            fieldValue7 = "Bank - kantoorgebouw"
                        Case "2"
                            fieldValue7 = "Bibliotheek"
                        Case "3"
                            fieldValue7 = "Consulaat"
                        Case "4", "5", "18", "32"
                            fieldValue7 = "Cultuur - feestzaal - muziek - theater"
                        Case "6"
                            fieldValue7 = "Hoogbouw"
                        Case "7"
                            fieldValue7 = "Hotel"
                        Case "8"
                            fieldValue7 = "Industriële instelling"
                        Case "10", "12", "14", "17", "32"
                            fieldValue7 = "Kapel - kerk - klooster - moskee - synagoge"
                        Case "11", "23"
                            fieldValue7 = "Kazerne - politiedienst"
                        Case "13"
                            fieldValue7 = "Kinderkribbe"
                        Case "15"
                            fieldValue7 = "KMO"
                        Case "16"
                            fieldValue7 = "Markt"
                        Case "18"
                            fieldValue7 = "Museum"
                        Case "20", "22", "29"
                            fieldValue7 = "Park - plantsoen - Stadsdienst"
                        Case "21"
                            fieldValue7 = "Petrochemische instelling"
                        Case "24"
                            fieldValue7 = "Rustoord"
                        Case "25"
                            fieldValue7 = "School"
                        Case "26", "31"
                            fieldValue7 = "Spoorinfrastructuur – station"
                        Case "27", "28"
                            fieldValue7 = "Sportinfrastructuur - stadion"
                        Case "30"
                            fieldValue7 = "Stadsmagazijn"
                        Case "34"
                            fieldValue7 = "Tunnel"
                        Case "35"
                            fieldValue7 = "Ziekenhuis"
                    End Select

                End If

                ' -----------------------------------------------------------------------------
                ' Register building attributes and kwadrantreference into the recordset,
                ' by adding new building registrations or adding a new single  
                ' kwadrantreference to an already registered building. 
                ' -----------------------------------------------------------------------------

                'Look for this building in the recordset.
                'If Trim(fieldValue1) = "" Then Exit While
                If Not pRecordset.EOF Then pRecordset.MoveFirst()
                searchCondition = c_IdFieldName & "=" & CStr(fieldValue1)
                pRecordset.Find(searchCondition)

                If pRecordset.EOF Then
                    ' Building not yet in the recordset.

                    ' Add this building as a new record to the recordset.
                    pRecordset.AddNew()
                    pRecordset(c_IdFieldName).Value = fieldValue1
                    pRecordset(c_NameFieldName).Value = fieldValue2
                    pRecordset(c_RefListFieldName).Value = kwadrantReference
                    pRecordset(c_StreetFieldName).Value = fieldValue4
                    pRecordset(c_DescriptionFieldName).Value = fieldValue5
                    pRecordset(c_TypeFieldName).Value = fieldValue6
                    pRecordset(c_CategoryFieldName).Value = fieldValue7

                Else
                    ' Building is already in the recordset.

                    ' Get all kwadrants already in recordset for current building.
                    Dim listKwadrant As String
                    Dim arrayKwadrant As String()
                    listKwadrant = CStr(pRecordset(c_RefListFieldName).Value)
                    arrayKwadrant = Split2(listKwadrant, c_ListSeparator, True)
                    If Array.IndexOf(arrayKwadrant, kwadrantReference) = -1 Then

                        ' Add ref only if kwadrant is not yet added in recordset.
                        ReDim Preserve arrayKwadrant(arrayKwadrant.Length)
                        arrayKwadrant.SetValue(kwadrantReference, arrayKwadrant.Length - 1)
                        Array.Sort(arrayKwadrant) 'sort references
                        listKwadrant = Concat(arrayKwadrant)
                        If Len(listKwadrant) > m_refListMaxLength Then
                            Throw New RecordsetFieldSizeNotSufficientException(c_RefListFieldName)
                        Else
                            pRecordset(c_RefListFieldName).Value = listKwadrant
                        End If

                    End If
                End If

                'Next feature from cursor. 
                pFeature = pCursor.NextFeature
            End While

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Export recordset data to text file on local disk.
    ''' </summary>
    ''' <param name="Recordset">
    '''     Recordset with building attributes and list of individual kwadrant references.
    ''' </param>
    ''' <param name="ProgressDelegate">
    '''     Delegate to progress bar update procedure.
    ''' </param>
    ''' <param name="MaxProgressDelegate">
    '''     Delegate to progress bar maximum procedure.
    ''' </param>
    ''' <param name="ComponentOrder">
    '''     The order in which the building attributes must be written to the text file.
    '''     Available lettercodes:
    '''     - n = gebouwnummer
    '''     - s = gebouwnaam
    '''     - k = kwadrants
    '''     - a = aanduiding
    '''     - t = type
    ''' </param>
    ''' <param name="ExportFilePath">
    '''     The full local path of the export text file.
    ''' </param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    '''     [Kristof Vydt]  13/07/2006  Update the name of some variables/constants.
    ''' 	[Kristof Vydt]	21/03/2007	Add building type.
    ''' 	[Kristof Vydt]	22/03/2007	Export with unicode encoding.
    ''' 	[Kristof Vydt]	28/03/2007	Add building category.
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub ExportRecordset( _
            ByVal Recordset As ADODB.Recordset, _
            ByRef ProgressDelegate As ShowProgress, _
            ByRef MaxProgressDelegate As SetMaxProgress, _
            ByVal ComponentOrder As String, _
            ByVal ExportFilePath As String)

        Dim StreamWriter As StreamWriter 'for writing to the export file
        Dim i, rec As Integer 'index
        Dim KwadrantArray As String() 'array of raster-kwadrant references
        Dim ThisKwadrant As String 'one singel raster-kwadrant reference (pagenumber & 1 kwadrant char)

        'Progress monitor information.
        MaxProgressDelegate(Recordset.RecordCount)
        ProgressDelegate(-1, "Exportbestand wordt voorbereid...")

        'Prepare export file.
        Try
            If File.Exists(ExportFilePath) Then
                ' Append to existing export file.
                StreamWriter = New StreamWriter(ExportFilePath, False, System.Text.Encoding.Unicode)
            Else
                ' Create new unicode export file.
                StreamWriter = New StreamWriter(ExportFilePath, True, System.Text.Encoding.Unicode)
                ' Add an empty line at the begin of file, to avoid problem
                ' with hidden character(s) as a result of unicode encoding.
                StreamWriter.WriteLine("")
            End If
        Catch ex As Exception
            Throw New RecreateExportFileException(ExportFilePath)
        End Try

        Try
            rec = 1
            Recordset.MoveFirst()
            While Not Recordset.EOF

                ' Update progress monitor information.
                ProgressDelegate(rec, Nothing)

                For i = 0 To ComponentOrder.Length - 1
                    Select Case ComponentOrder.Substring(i, 1)
                        Case "n" 'VOLGNUMMER
                            '-- N --
                            StreamWriter.WriteLine("n," & Trim(CStr(Recordset(c_idFieldName).Value)))
                        Case "s" 'NAAM
                            '-- S --
                            StreamWriter.WriteLine("s," & Trim(CStr(Recordset(c_nameFieldName).Value)))
                        Case "k" 'KWADRANTEN
                            '-- K --
                            KwadrantArray = Split2(Trim(CStr(Recordset(c_refListFieldName).Value)), c_ListSeparator, True)
                            Array.Sort(KwadrantArray)
                            For Each ThisKwadrant In KwadrantArray
                                While ThisKwadrant.Substring(0, 1) = "0"
                                    ThisKwadrant = ThisKwadrant.Substring(1) 'Verwijder nullen aan het begin.
                                End While
                                StreamWriter.WriteLine("k," & ThisKwadrant)
                            Next
                        Case "a" 'AANDUIDING
                            '-- A --
                            StreamWriter.WriteLine("a," & Trim(CStr(Recordset(c_DescriptionFieldName).Value)))
                        Case "t" 'TYPE
                            '-- T --
                            StreamWriter.WriteLine("t," & Trim(CStr(Recordset(c_TypeFieldName).Value)))
                        Case "c" 'CATEGORIE
                            '-- C --
                            StreamWriter.WriteLine("c," & Trim(CStr(Recordset(c_CategoryFieldName).Value)))
                    End Select
                    StreamWriter.Flush()
                Next 'Next component of current record.
                Recordset.MoveNext()
                rec = rec + 1
            End While 'Next record in recordset.
        Catch ex As Exception
            Throw ex
        End Try

        'Sluit het txt export bestand.
        StreamWriter.Close()
        ProgressDelegate(-2, "Exportbestand gesloten")

        'Release objects.
        If Not StreamWriter Is Nothing Then StreamWriter = Nothing

    End Sub

#End Region

End Class
