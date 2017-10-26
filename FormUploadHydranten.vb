Option Explicit On 
Option Strict On

#Region " Imports namespaces "
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.Marshal
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Editor
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Geometry
#End Region

''' -----------------------------------------------------------------------------
''' Project	 : Digipolis.Hydranten.BeheerHydranten
''' Class	 : Hydranten.BeheerHydranten.FormUploadHydranten
''' 
''' -----------------------------------------------------------------------------
''' <summary>
'''     Upload hydrants from external file into the geodatabase.
''' </summary>
''' <remarks>
''' </remarks>
''' <history>
''' 	[Kristof Vydt]	??/??/2005	Created
''' 	[Kristof Vydt]	02/08/2006	Replace some global constant by a function that reads the value from config file.
'''     [Kristof Vydt]  08/09/2006  Start of other implementation approach ...
'''     [Kristof Vydt]  15/09/2006  ... continue this
'''     [Kristof Vydt]  21/09/2006  ... continue this
'''     [Kristof Vydt]  28/09/2006  ... continue this
''' 	[Kristof Vydt]	22/03/2007	Use the new CodedValueDomainManager instead of the deprecated ModuleDomainAccess.
'''     [Koen Vermeer]  23/06/2008  Conversion to ArcGIS 9.2
''' </history>
''' -----------------------------------------------------------------------------
Public Class FormUploadHydranten
    Inherits System.Windows.Forms.Form

#Region " Local variables "

    'Locals.
    Private m_application As IMxApplication 'set by constructor
    Private m_document As IMxDocument 'set by constructor
    Private m_blockUpload As Boolean = False
    Private ds As New DataSet 'will contain tables UPLOAD, PROBLEMS

    'Schema file parameters
    Private Globals As Hashtable = New Hashtable 'global settings
    Private Columns As Hashtable = New Hashtable 'column mapping
    Private IgnoreFieldValues As Hashtable 'ignore records with field values
    Private ErrorFieldValues As Hashtable 'error records with field values
    Private Sectors As Hashtable 'label and code of Sector import data
    Private HydrantTypeMapping As Hashtable 'label and code of HydrantTypes import data
    Private LeidingTypeMapping As Hashtable 'label and code of LeidingType import data

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
    Friend WithEvents ButtonClose As System.Windows.Forms.Button
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents LabelProgressMessage As System.Windows.Forms.Label
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents ButtonUpload As System.Windows.Forms.Button
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents TextBoxCountStatus0 As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxCountStatus10 As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxCountStatus8 As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxCountStatus9 As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxCountStatus7 As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxCountStatus6 As System.Windows.Forms.TextBox
    Friend WithEvents OpenFileDialogDataFile As System.Windows.Forms.OpenFileDialog
    Friend WithEvents ButtonOpenExcel As System.Windows.Forms.Button
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents TextBoxCountStatus1 As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents RadioButtonIntegral As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButtonPartial As System.Windows.Forms.RadioButton
    Friend WithEvents TextBoxDataFile As System.Windows.Forms.TextBox
    Friend WithEvents ButtonBrowseDataFile As System.Windows.Forms.Button
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents ComboBoxProvider As System.Windows.Forms.ComboBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents TextBoxSectorName As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxSectorCode As System.Windows.Forms.TextBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents ButtonErrorReport As System.Windows.Forms.Button
    Friend WithEvents SaveFileDialog1 As System.Windows.Forms.SaveFileDialog
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.ButtonClose = New System.Windows.Forms.Button
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.TextBoxCountStatus1 = New System.Windows.Forms.TextBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.TextBoxCountStatus10 = New System.Windows.Forms.TextBox
        Me.TextBoxCountStatus8 = New System.Windows.Forms.TextBox
        Me.TextBoxCountStatus9 = New System.Windows.Forms.TextBox
        Me.TextBoxCountStatus7 = New System.Windows.Forms.TextBox
        Me.TextBoxCountStatus6 = New System.Windows.Forms.TextBox
        Me.TextBoxCountStatus0 = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.LabelProgressMessage = New System.Windows.Forms.Label
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar
        Me.ButtonUpload = New System.Windows.Forms.Button
        Me.OpenFileDialogDataFile = New System.Windows.Forms.OpenFileDialog
        Me.ButtonOpenExcel = New System.Windows.Forms.Button
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.ComboBoxProvider = New System.Windows.Forms.ComboBox
        Me.TextBoxDataFile = New System.Windows.Forms.TextBox
        Me.ButtonBrowseDataFile = New System.Windows.Forms.Button
        Me.RadioButtonPartial = New System.Windows.Forms.RadioButton
        Me.RadioButtonIntegral = New System.Windows.Forms.RadioButton
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.TextBoxSectorCode = New System.Windows.Forms.TextBox
        Me.TextBoxSectorName = New System.Windows.Forms.TextBox
        Me.ButtonErrorReport = New System.Windows.Forms.Button
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.SuspendLayout()
        '
        'ButtonClose
        '
        Me.ButtonClose.Location = New System.Drawing.Point(240, 208)
        Me.ButtonClose.Name = "ButtonClose"
        Me.ButtonClose.TabIndex = 0
        Me.ButtonClose.Text = "Sluiten"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.TextBoxCountStatus1)
        Me.GroupBox2.Controls.Add(Me.Label10)
        Me.GroupBox2.Controls.Add(Me.TextBoxCountStatus10)
        Me.GroupBox2.Controls.Add(Me.TextBoxCountStatus8)
        Me.GroupBox2.Controls.Add(Me.TextBoxCountStatus9)
        Me.GroupBox2.Controls.Add(Me.TextBoxCountStatus7)
        Me.GroupBox2.Controls.Add(Me.TextBoxCountStatus6)
        Me.GroupBox2.Controls.Add(Me.TextBoxCountStatus0)
        Me.GroupBox2.Controls.Add(Me.Label9)
        Me.GroupBox2.Controls.Add(Me.Label7)
        Me.GroupBox2.Controls.Add(Me.Label6)
        Me.GroupBox2.Controls.Add(Me.Label8)
        Me.GroupBox2.Controls.Add(Me.Label5)
        Me.GroupBox2.Controls.Add(Me.Label4)
        Me.GroupBox2.Controls.Add(Me.LabelProgressMessage)
        Me.GroupBox2.Controls.Add(Me.ProgressBar1)
        Me.GroupBox2.Location = New System.Drawing.Point(328, 8)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(280, 232)
        Me.GroupBox2.TabIndex = 9
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Vooruitgang"
        '
        'TextBoxCountStatus1
        '
        Me.TextBoxCountStatus1.Enabled = False
        Me.TextBoxCountStatus1.Location = New System.Drawing.Point(16, 56)
        Me.TextBoxCountStatus1.Name = "TextBoxCountStatus1"
        Me.TextBoxCountStatus1.Size = New System.Drawing.Size(40, 20)
        Me.TextBoxCountStatus1.TabIndex = 26
        Me.TextBoxCountStatus1.Text = ""
        Me.TextBoxCountStatus1.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label10
        '
        Me.Label10.Location = New System.Drawing.Point(64, 64)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(208, 15)
        Me.Label10.TabIndex = 25
        Me.Label10.Text = "Status ""OK"""
        '
        'TextBoxCountStatus10
        '
        Me.TextBoxCountStatus10.Enabled = False
        Me.TextBoxCountStatus10.Location = New System.Drawing.Point(16, 176)
        Me.TextBoxCountStatus10.Name = "TextBoxCountStatus10"
        Me.TextBoxCountStatus10.Size = New System.Drawing.Size(40, 20)
        Me.TextBoxCountStatus10.TabIndex = 24
        Me.TextBoxCountStatus10.Text = ""
        Me.TextBoxCountStatus10.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TextBoxCountStatus8
        '
        Me.TextBoxCountStatus8.Enabled = False
        Me.TextBoxCountStatus8.Location = New System.Drawing.Point(16, 152)
        Me.TextBoxCountStatus8.Name = "TextBoxCountStatus8"
        Me.TextBoxCountStatus8.Size = New System.Drawing.Size(40, 20)
        Me.TextBoxCountStatus8.TabIndex = 23
        Me.TextBoxCountStatus8.Text = ""
        Me.TextBoxCountStatus8.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TextBoxCountStatus9
        '
        Me.TextBoxCountStatus9.Enabled = False
        Me.TextBoxCountStatus9.Location = New System.Drawing.Point(16, 128)
        Me.TextBoxCountStatus9.Name = "TextBoxCountStatus9"
        Me.TextBoxCountStatus9.Size = New System.Drawing.Size(40, 20)
        Me.TextBoxCountStatus9.TabIndex = 22
        Me.TextBoxCountStatus9.Text = ""
        Me.TextBoxCountStatus9.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TextBoxCountStatus7
        '
        Me.TextBoxCountStatus7.Enabled = False
        Me.TextBoxCountStatus7.Location = New System.Drawing.Point(16, 104)
        Me.TextBoxCountStatus7.Name = "TextBoxCountStatus7"
        Me.TextBoxCountStatus7.Size = New System.Drawing.Size(40, 20)
        Me.TextBoxCountStatus7.TabIndex = 21
        Me.TextBoxCountStatus7.Text = ""
        Me.TextBoxCountStatus7.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TextBoxCountStatus6
        '
        Me.TextBoxCountStatus6.Enabled = False
        Me.TextBoxCountStatus6.Location = New System.Drawing.Point(16, 80)
        Me.TextBoxCountStatus6.Name = "TextBoxCountStatus6"
        Me.TextBoxCountStatus6.Size = New System.Drawing.Size(40, 20)
        Me.TextBoxCountStatus6.TabIndex = 20
        Me.TextBoxCountStatus6.Text = ""
        Me.TextBoxCountStatus6.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TextBoxCountStatus0
        '
        Me.TextBoxCountStatus0.Enabled = False
        Me.TextBoxCountStatus0.Location = New System.Drawing.Point(16, 200)
        Me.TextBoxCountStatus0.Name = "TextBoxCountStatus0"
        Me.TextBoxCountStatus0.Size = New System.Drawing.Size(40, 20)
        Me.TextBoxCountStatus0.TabIndex = 19
        Me.TextBoxCountStatus0.Text = ""
        Me.TextBoxCountStatus0.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label9
        '
        Me.Label9.Location = New System.Drawing.Point(64, 208)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(208, 16)
        Me.Label9.TabIndex = 18
        Me.Label9.Text = "Niet-ingeladen records met fouten"
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(64, 160)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(208, 15)
        Me.Label7.TabIndex = 16
        Me.Label7.Text = "Status ""nakijken coördinaten"""
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(64, 136)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(208, 15)
        Me.Label6.TabIndex = 15
        Me.Label6.Text = "Status ""nakijken attributen"""
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(64, 184)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(212, 15)
        Me.Label8.TabIndex = 14
        Me.Label8.Text = "Status ""nakijken attributen && coördinaten"""
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(64, 112)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(208, 15)
        Me.Label5.TabIndex = 11
        Me.Label5.Text = "Status ""verwijderd"""
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(64, 88)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(208, 15)
        Me.Label4.TabIndex = 10
        Me.Label4.Text = "Status ""nieuw"""
        '
        'LabelProgressMessage
        '
        Me.LabelProgressMessage.Location = New System.Drawing.Point(8, 32)
        Me.LabelProgressMessage.Name = "LabelProgressMessage"
        Me.LabelProgressMessage.Size = New System.Drawing.Size(264, 16)
        Me.LabelProgressMessage.TabIndex = 8
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Enabled = False
        Me.ProgressBar1.Location = New System.Drawing.Point(8, 16)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(264, 16)
        Me.ProgressBar1.TabIndex = 7
        '
        'ButtonUpload
        '
        Me.ButtonUpload.Enabled = False
        Me.ButtonUpload.Location = New System.Drawing.Point(16, 208)
        Me.ButtonUpload.Name = "ButtonUpload"
        Me.ButtonUpload.TabIndex = 15
        Me.ButtonUpload.Text = "Opladen"
        '
        'OpenFileDialogDataFile
        '
        Me.OpenFileDialogDataFile.Filter = "Excell bestand|*.xls"
        Me.OpenFileDialogDataFile.Title = "Import gegevens"
        '
        'ButtonOpenExcel
        '
        Me.ButtonOpenExcel.Enabled = False
        Me.ButtonOpenExcel.Location = New System.Drawing.Point(232, 48)
        Me.ButtonOpenExcel.Name = "ButtonOpenExcel"
        Me.ButtonOpenExcel.Size = New System.Drawing.Size(72, 23)
        Me.ButtonOpenExcel.TabIndex = 26
        Me.ButtonOpenExcel.Text = "Open Excel"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label11)
        Me.GroupBox1.Controls.Add(Me.ComboBoxProvider)
        Me.GroupBox1.Controls.Add(Me.TextBoxDataFile)
        Me.GroupBox1.Controls.Add(Me.ButtonBrowseDataFile)
        Me.GroupBox1.Controls.Add(Me.RadioButtonPartial)
        Me.GroupBox1.Controls.Add(Me.RadioButtonIntegral)
        Me.GroupBox1.Controls.Add(Me.ButtonOpenExcel)
        Me.GroupBox1.Location = New System.Drawing.Point(8, 64)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(312, 128)
        Me.GroupBox1.TabIndex = 28
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Gegevensbestand"
        '
        'Label11
        '
        Me.Label11.Location = New System.Drawing.Point(8, 56)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(64, 16)
        Me.Label11.TabIndex = 32
        Me.Label11.Text = "Leverancier"
        '
        'ComboBoxProvider
        '
        Me.ComboBoxProvider.Location = New System.Drawing.Point(80, 48)
        Me.ComboBoxProvider.Name = "ComboBoxProvider"
        Me.ComboBoxProvider.Size = New System.Drawing.Size(144, 21)
        Me.ComboBoxProvider.TabIndex = 31
        '
        'TextBoxDataFile
        '
        Me.TextBoxDataFile.Enabled = False
        Me.TextBoxDataFile.Location = New System.Drawing.Point(8, 16)
        Me.TextBoxDataFile.Name = "TextBoxDataFile"
        Me.TextBoxDataFile.Size = New System.Drawing.Size(272, 20)
        Me.TextBoxDataFile.TabIndex = 14
        Me.TextBoxDataFile.Text = ""
        '
        'ButtonBrowseDataFile
        '
        Me.ButtonBrowseDataFile.Location = New System.Drawing.Point(280, 16)
        Me.ButtonBrowseDataFile.Name = "ButtonBrowseDataFile"
        Me.ButtonBrowseDataFile.Size = New System.Drawing.Size(24, 23)
        Me.ButtonBrowseDataFile.TabIndex = 15
        Me.ButtonBrowseDataFile.Text = "..."
        '
        'RadioButtonPartial
        '
        Me.RadioButtonPartial.Location = New System.Drawing.Point(24, 104)
        Me.RadioButtonPartial.Name = "RadioButtonPartial"
        Me.RadioButtonPartial.Size = New System.Drawing.Size(240, 16)
        Me.RadioButtonPartial.TabIndex = 1
        Me.RadioButtonPartial.Text = "Partieel (enkel toevoegen en wijzigen)"
        '
        'RadioButtonIntegral
        '
        Me.RadioButtonIntegral.Location = New System.Drawing.Point(24, 80)
        Me.RadioButtonIntegral.Name = "RadioButtonIntegral"
        Me.RadioButtonIntegral.Size = New System.Drawing.Size(240, 16)
        Me.RadioButtonIntegral.TabIndex = 0
        Me.RadioButtonIntegral.Text = "Integraal (incl. status op verwijderd zetten)"
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.Panel1)
        Me.GroupBox3.Controls.Add(Me.TextBoxSectorCode)
        Me.GroupBox3.Controls.Add(Me.TextBoxSectorName)
        Me.GroupBox3.Location = New System.Drawing.Point(8, 8)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(312, 48)
        Me.GroupBox3.TabIndex = 29
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Sector"
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Location = New System.Drawing.Point(280, 16)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(24, 24)
        Me.Panel1.TabIndex = 29
        '
        'TextBoxSectorCode
        '
        Me.TextBoxSectorCode.Enabled = False
        Me.TextBoxSectorCode.Location = New System.Drawing.Point(8, 16)
        Me.TextBoxSectorCode.Name = "TextBoxSectorCode"
        Me.TextBoxSectorCode.ReadOnly = True
        Me.TextBoxSectorCode.Size = New System.Drawing.Size(32, 20)
        Me.TextBoxSectorCode.TabIndex = 28
        Me.TextBoxSectorCode.Text = "XXX"
        Me.TextBoxSectorCode.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBoxSectorName
        '
        Me.TextBoxSectorName.Enabled = False
        Me.TextBoxSectorName.Location = New System.Drawing.Point(48, 16)
        Me.TextBoxSectorName.Name = "TextBoxSectorName"
        Me.TextBoxSectorName.ReadOnly = True
        Me.TextBoxSectorName.Size = New System.Drawing.Size(224, 20)
        Me.TextBoxSectorName.TabIndex = 13
        Me.TextBoxSectorName.Text = "Sectornaam"
        '
        'ButtonErrorReport
        '
        Me.ButtonErrorReport.AccessibleRole = System.Windows.Forms.AccessibleRole.PageTabList
        Me.ButtonErrorReport.Enabled = False
        Me.ButtonErrorReport.Location = New System.Drawing.Point(128, 208)
        Me.ButtonErrorReport.Name = "ButtonErrorReport"
        Me.ButtonErrorReport.TabIndex = 30
        Me.ButtonErrorReport.Text = "Problemen"
        '
        'FormUploadHydranten
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(610, 248)
        Me.ControlBox = False
        Me.Controls.Add(Me.ButtonErrorReport)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.ButtonUpload)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.ButtonClose)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.Name = "FormUploadHydranten"
        Me.Text = "Hydranten opladen en integreren"
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
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
        InitializeForm()

    End Sub
#End Region

#Region " Initialization procedures "

    'Initial state of the form when loading.
    Private Sub InitializeForm()

        ' Set labels and text boxes.
        Me.TextBoxSectorName.Text = GetSectorName(m_document)
        Me.TextBoxSectorCode.Text = GetSectorCode(m_document)
        Me.TextBoxCountStatus0.Text = ""
        Me.TextBoxCountStatus1.Text = ""
        Me.TextBoxCountStatus6.Text = ""
        Me.TextBoxCountStatus7.Text = ""
        Me.TextBoxCountStatus8.Text = ""
        Me.TextBoxCountStatus9.Text = ""
        Me.TextBoxCountStatus10.Text = ""

        ' Manipulate some colors.
        Dim Color As System.drawing.Color = New System.Drawing.Color
        With TextBoxDataFile
            .Enabled = False
            .BackColor = Color.White
            .ForeColor = Color.Black
        End With
        With TextBoxSectorName
            .Enabled = False
            .BackColor = Color.White
            .ForeColor = Color.Black
        End With

        Dim pFeature As IFeature
        Dim pFeatureCursor As IFeatureCursor
        Dim pFeatureLayer As IFeatureLayer = Nothing
        Dim pQueryFilter As IQueryFilter = New QueryFilter
        Dim counter As Integer

        ' Make sure that the current sector has no temporary status hydrants.
        Try

            pFeatureLayer = GetFeatureLayer(m_document.FocusMap, GetLayerName("Hydrant"))
            If pFeatureLayer Is Nothing Then Throw New LayerNotFoundException(GetLayerName("Hydrant"))
            pQueryFilter.WhereClause = _
                "(" & GetAttributeName("Hydrant", "Status") & "='6' ) OR " & _
                "(" & GetAttributeName("Hydrant", "Status") & "='7' ) OR " & _
                "(" & GetAttributeName("Hydrant", "Status") & "='8' ) OR " & _
                "(" & GetAttributeName("Hydrant", "Status") & "='9' ) OR " & _
                "(" & GetAttributeName("Hydrant", "Status") & "='10')"
            pFeatureCursor = pFeatureLayer.Search(pQueryFilter, False)
            pFeature = pFeatureCursor.NextFeature
            If Not pFeature Is Nothing Then
                While Not pFeature Is Nothing
                    counter += 1
                    pFeature = pFeatureCursor.NextFeature
                End While
                Me.ButtonUpload.Enabled = False
                Throw New HydrantsWithTemporaryStatus(counter, Me.TextBoxSectorName.Text)
            End If
            Me.Panel1.BackColor = Color.Green

        Catch ex As Exception
            m_blockUpload = True
            Me.Panel1.BackColor = Color.Red
            MsgBox(ex.Message)

        End Try

        ' Fill data providers list.
        'Dim pWorkspace As IWorkspace = GetLayerWorkspace(m_application, pFeatureLayer)
        'Dim domainName As String = GetDomainName("Bron")
        'PopulateCodes(pWorkspace, domainName, Me.ComboBoxProvider)
        'SetEditBoxValue(Me.ComboBoxProvider, GetDomainCodeValue(pWorkspace, domainName, "AWW"))
        Dim domainMgr As CodedValueDomainManager
        domainMgr = New CodedValueDomainManager(pFeatureLayer, "Bron")
        domainMgr.PopulateCodes(Me.ComboBoxProvider)
        SetEditBoxValue(Me.ComboBoxProvider, domainMgr.CodeValue("AWW"))

        ' Set buttons.
        Me.ButtonBrowseDataFile.Enabled = True
        Me.ButtonOpenExcel.Enabled = False
        Me.ButtonUpload.Enabled = False
        Me.ButtonClose.Enabled = True

    End Sub

    ' Make Upload button available only if all conditions are fullfilled.
    Private Sub UpdateButtonUploadEnabledState()

        If (Len(Me.TextBoxDataFile.Text) > 0) _
                    And (Me.RadioButtonIntegral.Checked Or Me.RadioButtonPartial.Checked) _
                    And (Not m_blockUpload) Then
            Me.ButtonUpload.Enabled = True
        Else
            Me.ButtonUpload.Enabled = False
        End If

    End Sub

#End Region

#Region " Form controls events "

    Private Sub ButtonBrowseDataFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonBrowseDataFile.Click

        Dim path As String
        Dim target As String = "\/|"
        Dim anyOf As Char() = target.ToCharArray()

        ' Browse and retrieve xls file name.
        With OpenFileDialogDataFile
            .Title = "Import gegevens"
            .Filter = "Excell bestand (*.xls)|*.xls|dBase bestand (*.dbf)|*.dbf"
            .ShowDialog()
            path = .FileName
        End With
        TextBoxDataFile.Text = path.Substring(1 + path.LastIndexOfAny(anyOf))

    End Sub

    Private Sub TextBoxDataFile_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBoxDataFile.TextChanged

        ' Button Open XLS
        If (Len(Me.TextBoxDataFile.Text) > 0) Then
            Me.ButtonOpenExcel.Enabled = True
        Else
            Me.ButtonOpenExcel.Enabled = False
        End If

        ' Button Upload
        UpdateButtonUploadEnabledState()

    End Sub

    Private Sub RadioButtonPartial_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButtonPartial.CheckedChanged
        UpdateButtonUploadEnabledState()
    End Sub

    Private Sub RadioButtonIntegral_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButtonIntegral.CheckedChanged
        UpdateButtonUploadEnabledState()
    End Sub

    Private Sub ButtonClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonClose.Click
        Close()
    End Sub

    Private Sub ButtonOpenExcel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonOpenExcel.Click
        ShowXLS(OpenFileDialogDataFile.FileName)
    End Sub

    Private Sub ButtonUpload_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonUpload.Click

        ' Disable controls to avoid user interaction while processing.
        Me.ButtonBrowseDataFile.Enabled = False
        Me.ButtonUpload.Enabled = False
        Me.ButtonOpenExcel.Enabled = False
        Me.ComboBoxProvider.Enabled = False

        ' Initialise status counters.
        Me.TextBoxCountStatus0.Text = "0"
        Me.TextBoxCountStatus1.Text = "0"
        Me.TextBoxCountStatus6.Text = "0"
        Me.TextBoxCountStatus7.Text = "0"
        Me.TextBoxCountStatus8.Text = "0"
        Me.TextBoxCountStatus9.Text = "0"
        Me.TextBoxCountStatus10.Text = "0"

        ' Update progress monitor.
        OnShowProgress1(0, "Voorbereiding ...")

        ' Read some input values from form.
        Dim sectorCode As String = Me.TextBoxSectorCode.Text
        Dim providerCode As String = GetComboBoxCodeValue(Me.ComboBoxProvider)

        ' Read first sheet name.
        Dim firstSheetName As String = GetXLSSheetName(Me.OpenFileDialogDataFile.FileName, 1)

        ' Integration of hydrants from a file.
        If Me.RadioButtonIntegral.Checked Or Me.RadioButtonPartial.Checked Then
            IntegrateDataFile(Me.OpenFileDialogDataFile.FileName, firstSheetName, _
                sectorCode, providerCode)
        End If

        ' Flag hydrants that are not mentioned in the file.
        If Me.RadioButtonIntegral.Checked Then
            If MsgBox(c_Message_DepricateHydrants, vbYesNo, c_Title_OpladenHydranten) = MsgBoxResult.Yes Then _
                FlagRemovedHydrants(sectorCode, providerCode)
        End If

        ' Manipulate form to allow user interaction.
        Me.ButtonOpenExcel.Enabled = True
        Me.ButtonClose.Enabled = True
        Me.ButtonErrorReport.Enabled = True
        ' ButtonUpload must remain disabled because checking for  
        ' temporary status hydrants is only done when the form is loaded.

        ' Inform the user that it's finished.
        OnShowProgress1(0, c_Message_UploadFinished)
        'MsgBox(c_Message_UploadFinished, MsgBoxStyle.OKOnly, c_Title_OpladenHydranten)

    End Sub

    Private Sub ButtonErrorReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonErrorReport.Click
        Call ShowProblemRows()
    End Sub

#End Region

#Region " Progress information management "

    Public Delegate Sub ShowProgress(ByVal progress As Integer, ByVal message As String)

    Public Delegate Sub SetMaxProgress(ByVal max As Integer)

    ''' -----------------------------------------------------------------------------
    ''' <summary>
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
    ''' 	[Kristof Vydt]	15/09/2006  Second progress bar control deleted.
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
            'If c_DebugStatus Then LabelProgressMessage.Text &= " [" & CStr(ProgressBar1.Value) & "/" & CStr(ProgressBar1.Maximum) & "]"
            Me.LabelProgressMessage.Refresh()
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
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
    ''' 	[Kristof Vydt]	15/09/2006  Second progress bar control deleted.
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub OnSetMaxProgress1(ByVal max As Integer)
        ProgressBar1.Minimum = 0
        ProgressBar1.Maximum = max
        ProgressBar1.Refresh()
    End Sub

#End Region

#Region " Utility procedures "

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Upload hydrants from an Excel data file.
    ''' </summary>
    ''' <param name="dataFile">The Excel file path.</param>
    ''' <param name="dataSheet">The name of the Excel worksheet.</param>
    ''' <param name="sectorCode">The code of the current sector.</param>
    ''' <param name="providerCode">The code of the provider of the data file.</param>
    ''' <remarks>
    '''     This method contains complex logic regarding which hydrants
    '''     are evaluated and how data rows are translated to hydrant feature changes.
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	22/09/2006	Created
    ''' 	[Kristof Vydt]	28/09/2006	Release COM objects.
    ''' 	[Kristof Vydt]	14/12/2006	Correct check for multiple features with same LeverancierId.
    ''' 	[Kristof Vydt]	22/03/2007	Use the new CodedValueDomainManager instead of the deprecated ModuleDomainAccess.
    '''     [RW / Elton Manoku] 22/07/2008 See below for more explanation under RW:07-08/2008.
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub IntegrateDataFile( _
            ByVal dataFile As String, _
            ByVal dataSheet As String, _
            ByVal sectorCode As String, _
            ByVal providerCode As String)

        ' ----- General structure -----
        ' Connect to data file.
        ' Read data from file into a dataset.
        ' > Use WHERE-clause to retrieve only rows with ID and current sector code.
        ' Determine number of rows to be processed.
        ' Initialize progress monitor.
        ' Evaluate and process first row in dataset.
        ' > Try to find a matching registered feature ...
        '   ... insert new feature based on the dataset row.
        ' > Determine if matching feature has similar attributes ...
        '   ... insert new feature partly based on the dataset row and partly based on the matched feature.
        '   ... change the state attribute of the matching feature.
        '   ... change just a single attribute of the matching feature.
        ' > When something went wrong, duplicate the dataset row to the problems dataset.
        ' Continu with next row in recordset.
        ' Export the problems dataset.
        ' -----------------------------

        ' Read data from file into a new datatable.
        ds.Tables.Clear() 'Clears the collection of all DataTable objects.
        'RW:07-08/2008 Records from the excel file are filtered by the condition: 
        ' must have a leveranciernummer 
        ' and must have been evaluated OK from the schoon programma (Excel macro)
        ' and it is inside the envelope of the sector selected. It uses first the sector envelope 
        ' to filter out most of the records. Later it filters the remaining records again 
        ' by using a spatial filter with the sector shape.
        Try

            Dim queryCondition As String = _
                "(LeverancierNummer IS NOT NULL) AND Evaluatie ='OK' AND " ' and (Sector=" & CStrSql(sectorCode) & ")"

            'Get an array with values for minx, miny, maxx, maxy of the envelope of the sector
            Dim EnvelopeOfSectorCoords As Double() = GetSectorEnvelope(sectorCode)

            'RW:2008 We need to define a culture to convert the number with dot decimal seperation
            'instead of comma
            Dim cultureToFormatDouble As System.Globalization.CultureInfo = System.Globalization.CultureInfo.CreateSpecificCulture("en-US")

            'finalize the where condition for the excel records
            queryCondition = queryCondition _
                & String.Format(" CoordX >= {0} and CoordX <= {2} and CoordY >= {1} and CoordY<={3}", _
                    New Object() {EnvelopeOfSectorCoords(0).ToString("G", cultureToFormatDouble), EnvelopeOfSectorCoords(1).ToString("G", cultureToFormatDouble), _
                    EnvelopeOfSectorCoords(2).ToString("G", cultureToFormatDouble), EnvelopeOfSectorCoords(3).ToString("G", cultureToFormatDouble)})
            ImportXLSWorksheet(ds, "UPLOAD", dataFile, dataSheet, queryCondition)
        Catch ex As Exception
            Throw New ApplicationException("Er is een fout tijdens de import van de excel bestaand:" + ex.Message)
        End Try

        'RW:07-08/2008 Delete rows that are not found inside the sector
        Try

            Dim dtRow As DataRow
            For Each dtRow In ds.Tables.Item("UPLOAD").Rows
                'RW:07-08/2008 Find the sector of the point. If no sector found don't handle the row
                Dim rowCoordX As Double = Math.Round(CType(dtRow.Item("CoordX"), Double), 3)
                Dim rowCoordY As Double = Math.Round(CType(dtRow.Item("CoordY"), Double), 3)
                Dim sectorCodeOfPointInRow As String = GetSectorCodeOfPoint(rowCoordX, rowCoordY)

                If sectorCodeOfPointInRow.Equals(sectorCode) Then
                    'dtRow.Item("Sector") = sectorCodeOfPointInRow
                Else
                    dtRow.Delete()
                End If
            Next
            ds.Tables.Item("UPLOAD").AcceptChanges()
        Catch ex As Exception
            Throw New ApplicationException("Er is een fout tijdens de sector van de hydranten te vinden:" + ex.Message)
        End Try

        ' Check if LeverancierNummer is unique in the whole table.
        Try
            ds.Tables.Item("UPLOAD").Columns.Item("LeverancierNummer").Unique = True
        Catch ex As InvalidConstraintException
            Throw New ApplicationException(String.Format(c_Message_NonUniqueColumn, "LeverancierNummer"))
        End Try

        'DataGrid1.DataSource = ds.Tables(0)
        'MsgBox(ds.Tables(0).Rows.Count)

        ' Set progress monitor max value.
        OnSetMaxProgress1(ds.Tables(0).Rows.Count)

        ' Prepare a table for problem records. 
        ' Each upload record that gives us problems, will be added to this table.
        Dim dt_ErrRec As DataTable = ds.Tables.Item("UPLOAD").Clone
        dt_ErrRec.TableName = "PROBLEMS"
        ds.Tables.Add(dt_ErrRec)

        '== Loop through each row of the upload table. ==

        ' Common part of the query that is used to look for matching features.
        Dim pFeatureLayer As IFeatureLayer = GetFeatureLayer(m_document.FocusMap, GetLayerName("Hydrant"))
        'Dim pWorkspace As IWorkspace = GetLayerWorkspace(m_application, pFeatureLayer)
        'Dim strDomainName As String = GetDomainName("Status")
        'Dim strBaseWhereClause As String = _
        '    "(" & GetAttributeName("Hydrant", "Status") & " IN " & _
        '        "(" & CStrSql(CStr(GetDomainCodeValue(pWorkspace, strDomainName, "actief"))) & _
        '        "," & CStrSql(CStr(GetDomainCodeValue(pWorkspace, strDomainName, "niet bruikbaar"))) & ")) and " & _
        '    "(" & GetAttributeName("Hydrant", "EindDatum") & " IS NULL)"
        Dim pDomainMgr As New CodedValueDomainManager(pFeatureLayer, "Status")
        Dim strBaseWhereClause As String = _
            "(" & GetAttributeName("Hydrant", "Status") & " IN " & _
                "(" & CStrSql(pDomainMgr.CodeValue("actief")) & _
                "," & CStrSql(pDomainMgr.CodeValue("niet bruikbaar")) & ")) and " & _
            "(" & GetAttributeName("Hydrant", "EindDatum") & " IS NULL)"

        ' Prepare an editor and start edit session on hydrants feature layer.
        Dim pEditor As IEditor2 = GetEditorReference(m_application)

        'The edit session starts without an edit operation because that operation has to start for any row in the datatable
        EditSessionStart(pEditor, GetFeatureLayer(m_document.FocusMap, GetLayerName("Hydrant")), False)

        Try

            Dim row As DataRow
            For Each row In ds.Tables.Item("UPLOAD").Rows

                Dim pQueryFilter As QueryFilter = Nothing
                Dim pFeatureCursor As IFeatureCursor = Nothing
                Dim pFeature As IFeature = Nothing

                Try

                    ' Start edit operation.
                    ' This allows rollback in case of exceptions.
                    pEditor.StartOperation()

                    'RW:07-08/2008 Get the coordinates of the hydrant in the input table 
                    Dim rowCoordX As Double = Math.Round(CType(row.Item("CoordX"), Double), 3)
                    Dim rowCoordY As Double = Math.Round(CType(row.Item("CoordY"), Double), 3)
                    'Dim rowCoordX As Double = CType(row.Item("CoordX"), Double)
                    'Dim rowCoordY As Double = CType(row.Item("CoordY"), Double)

                    'RW:07-08/2008 Find the sector of the point. If no sector found don't handle the row
                    'This piece of code below will be active if the whole file of excel will be treated as one
                    'Dim sectorCodeOfPointInRow As String = GetSectorCodeOfPoint(rowCoordX, rowCoordY)
                    'If (Not sectorCode.Equals(sectorCodeOfPointInRow)) Then
                    'Continue For
                    'Throw New ApplicationException("De coordinaten staan niet in de sector: " + sectorCode)
                    'End If


                    ' Determine ID value.
                    Dim objLevNr As Object = row.Item("LeverancierNummer")
                    Dim strLevNr As String = CType(objLevNr, String)

                    ' Update progress monitor.
                    OnShowProgress1(Me.ProgressBar1.Value + 1, "Opladen " & strLevNr)

                    ' Look for matching feature:
                    ' - same active Sector code
                    ' - same LeverancierNr
                    ' - status "Actief"/"Niet bruikbaar"
                    ' - no EindDatum defined
                    pQueryFilter = New QueryFilter
                    pQueryFilter.WhereClause = strBaseWhereClause & _
                        " and (" & GetAttributeName("Hydrant", "LeverancierNr") & "='" & strLevNr & "')"
                    pFeatureCursor = pFeatureLayer.Search(pQueryFilter, Nothing)
                    pFeature = pFeatureCursor.NextFeature
                    'RW:07-08/2008 It maintains the status of the new feature. -1 means no value assigned
                    'Based in the conditions below status of the new feature can be 
                    'NIEUW, ACTIEF, NAKIJKEN ATTRIBUTEN, NAKIJKEN COORDINATEN, NAKIJKEN ATTRIBUTEN EN COORDINATEN

                    Dim statusOfNewFeature As Integer = -1
                    If pFeature Is Nothing Then

                        ' If no matching registered feature ...
                        '   ... insert new feature based on the dataset row.

                        '==================================================================================
                        '== Action: Add new feature with status NIEUW.                                   ==
                        '==================================================================================

                        statusOfNewFeature = CType(pDomainMgr.CodeValue("nieuw"), Integer)
                        ' Create new feature based on upload row only.
                        Call CreateNewFeature(pFeatureLayer, row, Nothing, providerCode, statusOfNewFeature.ToString()) '"6"

                        ' Increase status "nieuw" counter.
                        Call IncrementStatusCounter(statusOfNewFeature)

                    Else

                        ' Refuse data row if >1 matching feature was found.
                        Dim count As Integer = 1
                        Dim pTmpFeature As IFeature = pFeatureCursor.NextFeature
                        While Not pTmpFeature Is Nothing
                            count += 1
                            pTmpFeature = pFeatureCursor.NextFeature
                        End While
                        pTmpFeature = Nothing
                        If count > 1 Then _
                            Throw New ApplicationException("Er zijn momenteel meerdere hydranten met LeverancierNr = '" & strLevNr & "' in de databank geregistreerd.")


                        'RW:07-08/2008 - Compare if the input coordinates are the same with the existing coordinates
                        'Compare input coordinates rowCoordX, rowCoordY and feature existing coordinates featureCoordX, featureCoordY
                        Dim featureCoordX As Double = Math.Round(CType(GetAttributeValue(pFeature, "Hydrant", "CoordX"), Double), 3)
                        Dim featureCoordY As Double = Math.Round(CType(GetAttributeValue(pFeature, "Hydrant", "CoordY"), Double), 3)
                        Dim equalLocation As Boolean = (rowCoordX = featureCoordX And rowCoordY = featureCoordY)
                        Debug.WriteLineIf(c_DebugStatus, "equalLocation:" & equalLocation)

                        ' Compare CoordX value.
                        'Dim equalCoordX As Boolean = False
                        'Dim rowCoordX As Double = CType(row.Item("CoordX"), Double)
                        'If rowCoordX = CType(GetAttributeValue(pFeature, "Hydrant", "CoordX"), Double) Then equalCoordX = True
                        'Debug.WriteLineIf(c_DebugStatus, "equalCoordX:" & equalCoordX)

                        ' Compare CoordY value.
                        'Dim equalCoordY As Boolean = False
                        'Dim rowCoordY As Double = CType(row.Item("CoordY"), Double)
                        'If rowCoordY = CType(GetAttributeValue(pFeature, "Hydrant", "CoordY"), Double) Then equalCoordY = True
                        'Debug.WriteLineIf(c_DebugStatus, "equalCoordY:" & equalCoordY)

                        ' Compare Diameter value.
                        Dim equalDiameter As Boolean = False
                        Dim rowLeidingDiameter As Integer = CType(row.Item("LeidingDiameter"), Integer)
                        If rowLeidingDiameter = CType(GetAttributeValue(pFeature, "Hydrant", "Diameter"), Integer) Then equalDiameter = True
                        Debug.WriteLineIf(c_DebugStatus, "equalDiameter:" & equalDiameter)

                        ' Compare HydrantType value.
                        Dim equalHydrantType As Boolean = False
                        Dim rowHydrantType As String = CType(row.Item("HydrantType"), String)
                        If rowHydrantType = CType(GetAttributeValue(pFeature, "Hydrant", "HydrantType"), String) Then equalHydrantType = True
                        Debug.WriteLineIf(c_DebugStatus, "equalHydrantType:" & equalHydrantType)

                        ' Compare LeidingType value.
                        Dim equalLeidingType As Boolean = False
                        Dim rowLeidingType As Integer = CType(row.Item("LeidingType"), Integer)
                        If rowLeidingType = CType(GetAttributeValue(pFeature, "Hydrant", "LeidingType"), Integer) Then equalLeidingType = True
                        Debug.WriteLineIf(c_DebugStatus, "equalLeidingType:" & equalLeidingType)

                        ' Compare LeidingNr value.
                        Dim equalLeidingNr As Boolean = False
                        Dim rowLeidingNr As Integer = CType(row.Item("LeidingNummer"), Integer)
                        If rowLeidingNr = CType(GetAttributeValue(pFeature, "Hydrant", "LeidingNr"), Integer) Then equalLeidingNr = True
                        Debug.WriteLineIf(c_DebugStatus, "equalLeidingNr:" & equalLeidingNr)

                        ' If matching feature has same key attributes ...
                        '   ... insert new feature partly based on the dataset row and partly based on the matched feature.
                        '   ... change the state attribute of the matching feature.
                        '   ... change just a single attribute of the matching feature.

                        'If equalCoordX And equalCoordY Then
                        If equalLocation Then

                            If equalDiameter And equalHydrantType And equalLeidingType Then
                                'No new feature is added
                                If equalLeidingNr Then

                                    '=====================================================================
                                    '== Action: No action required.                                     ==
                                    '=====================================================================

                                    ' Increase "OK" counter.
                                    Call IncrementStatusCounter(1) 'ok (not updated)

                                Else

                                    '=====================================================================
                                    '== Action: Update LeidingNr of matching existing feature.          ==
                                    '=====================================================================

                                    ' Update "LeidingNr" attribute.
                                    Call SetAttributeValue(pFeature, "Hydrant", "LeidingNr", rowLeidingNr)

                                    ' Increase "OK" counter.
                                    Call IncrementStatusCounter(1) 'ok (although updated)

                                End If

                            Else

                                '==================================================================
                                '== Status of new feature NAKIJKEN ATTRIBUTEN                    ==
                                '==================================================================
                                ' statusOfNewFeature = 9
                                statusOfNewFeature = CType(pDomainMgr.CodeValue("nakijken_a"), Integer)
                            End If

                        Else
                            'RW:07-08/2008 - The distance between new location and old location.
                            Dim distanceBetweenPoints As Double = _
                                Math.Sqrt(Math.Pow((featureCoordX - rowCoordX), 2) + Math.Pow((featureCoordY - rowCoordY), 2))

                            If equalDiameter And equalHydrantType And equalLeidingType Then

                                If distanceBetweenPoints <= 0.3 Then
                                    'Set status for the new feature 1 - ACTIEF
                                    'if all attributes are the same and the location difference is less equal than 0.3 m
                                    statusOfNewFeature = CType(pDomainMgr.CodeValue("actief"), Integer)
                                Else
                                    'Set status of new feature 8 - NAKIJKEN COORDINATEN
                                    statusOfNewFeature = CType(pDomainMgr.CodeValue("nakijken_c"), Integer)
                                End If

                            Else

                                'RW:07-08/2008 - Hydrants with some numbers (99994, 99997, 99998) are used differently
                                'If the location for these hydrants is changed with <= 0.3m the hydrants become actief immidiately
                                'and new attribute values are taken over
                                Dim hydrantBrandweerNr As Long = CType(GetAttributeValue(pFeature, "Hydrant", "BrandweerNr"), Long)
                                Dim hydrantHasSpecialBrandweerNr As Boolean = _
                                    (hydrantBrandweerNr = 99994 Or hydrantBrandweerNr = 99997 Or hydrantBrandweerNr = 99998)

                                If hydrantHasSpecialBrandweerNr And distanceBetweenPoints <= 0.3 Then
                                    'Set status for the new feature 1 - ACTIEF
                                    'if hydrant has special number and the location difference is less equal than 0.3 m
                                    statusOfNewFeature = CType(pDomainMgr.CodeValue("actief"), Integer)
                                Else
                                    'Set status for the new feature to 10 - NAKIJKEN ATTRIBUTEN & COORDINATEN
                                    statusOfNewFeature = CType(pDomainMgr.CodeValue("nakijken_ac"), Integer)
                                End If

                            End If
                        End If
                        'If the status of the new feature got a value 
                        'the existing feature gets status = 3 - HISTORIEK
                        'the new feature get the status that it got from the conditions above
                        If statusOfNewFeature > -1 Then
                            ' Update "EindDatum" attribute.
                            Call SetAttributeValue(pFeature, "Hydrant", "EindDatum", Today.AddDays(-1))

                            ' Update "Status" attribute of the existing feature to 3 - HISTORIEK.
                            Dim statusHistoriek As Integer = CType(pDomainMgr.CodeValue("historiek"), Integer)
                            Call SetAttributeValue(pFeature, "Hydrant", "Status", statusHistoriek.ToString())

                            ' Update legend code attribute value.
                            Call UpdateLegendCode(pFeature)


                            ' Create new feature based on upload row and matching feature.
                            Call CreateNewFeature(pFeatureLayer, row, pFeature, providerCode, statusOfNewFeature.ToString())

                            ' Increase status "historiek" counter.
                            Call IncrementStatusCounter(statusHistoriek)

                            ' Increase status  counter of the new feature.
                            Call IncrementStatusCounter(statusOfNewFeature)

                        End If
                    End If

                    ' Confirm edit operation.
                    pEditor.StopOperation("Import hydrant " & strLevNr)

                Catch ex As Exception
                    ' Exception occured during processing of a single data row.

                    '======================================================================================
                    '== Action: Rollback changes made so far, based on this data row.                    ==
                    '== Action: Flag the data row as a problem record.                                   ==
                    '======================================================================================

                    ' Abort edit operation so that changes are not saved.
                    pEditor.AbortOperation()

                    ' Copy problematic data row to the ErrRec data table.
                    dt_ErrRec.ImportRow(row)

                    dt_ErrRec.Rows.Item(dt_ErrRec.Rows.Count - 1).Item(dt_ErrRec.Columns.Count - 1) = ex.Message

                    ' Increase problem rows counter.
                    Call IncrementStatusCounter(0) 'status 0 = problem row

                Finally
                    ' Release COM objects to avoid error "Cannot open any more tables."
                    ' More info on http://forums.esri.com/Thread.asp?c=93&f=993&t=146570
                    If Not (pQueryFilter Is Nothing) Then
                        ReleaseComObject(pQueryFilter)
                        pQueryFilter = Nothing
                    End If
                    If Not (pFeatureCursor Is Nothing) Then
                        ReleaseComObject(pFeatureCursor)
                        pFeatureCursor = Nothing
                    End If
                    If Not (pFeature Is Nothing) Then
                        ReleaseComObject(pFeature)
                        pFeature = Nothing
                    End If
                    GC.Collect() 'garbage collect
                End Try

            Next

            ' Stop the editing session and save changes.
            pEditor.StopEditing(True)

        Catch ex As Exception
            ErrorHandler(ex)
        Finally
            ' Release COM objects.
            If Not (pEditor Is Nothing) Then
                ReleaseComObject(pEditor)
                pEditor = Nothing
            End If
            'If Not (pWorkspace Is Nothing) Then
            '    ReleaseComObject(pWorkspace)
            '    pWorkspace = Nothing
            'End If
            If Not (pFeatureLayer Is Nothing) Then
                ReleaseComObject(pFeatureLayer)
                pFeatureLayer = Nothing
            End If
            GC.Collect() 'garbage collect
        End Try

    End Sub
    ''' <summary>
    ''' It finds out the sector code of the point with coordinates
    ''' given in parameters.
    ''' It is used during the load of the excel data
    ''' </summary>
    ''' <param name="CoordX">X Coordinate</param>
    ''' <param name="CoordY">Y Coordinate</param>
    ''' <returns>The code of the sector. If nothing is found returns an empty string</returns>
    ''' <remarks>
    ''' RW:07-08/2008 Elton Manoku
    ''' </remarks>
    Private Function GetSectorCodeOfPoint(ByVal CoordX As Double, ByVal CoordY As Double) As String

        Dim sectorCode As String = String.Empty
        Dim pQueryFilter As SpatialFilter = Nothing
        Dim pFeatureCursor As IFeatureCursor = Nothing
        Dim pFeature As IFeature = Nothing

        Try
            Dim featureLayerSector As IFeatureLayer = _
                GetFeatureLayer(m_document.FocusMap, GetLayerName("Sector"))

            Dim pointGeom As IPoint = New Point()
            pointGeom.PutCoords(CoordX, CoordY)
            pQueryFilter = New SpatialFilter
            pQueryFilter.SpatialRel = esriSpatialRelEnum.esriSpatialRelWithin
            pQueryFilter.Geometry = pointGeom
            pQueryFilter.GeometryField = featureLayerSector.FeatureClass.ShapeFieldName
            pFeatureCursor = featureLayerSector.Search(pQueryFilter, Nothing)
            pFeature = pFeatureCursor.NextFeature
            If Not (pFeature Is Nothing) Then
                Dim fldIdx As Integer = pFeature.Fields.FindField("AFKORTING")
                sectorCode = pFeature.Value(fldIdx).ToString()
            End If
        Catch ex As Exception
            Throw ex
        Finally
            ' Release COM objects to avoid error "Cannot open any more tables."
            ' More info on http://forums.esri.com/Thread.asp?c=93&f=993&t=146570
            If Not (pQueryFilter Is Nothing) Then
                ReleaseComObject(pQueryFilter)
                pQueryFilter = Nothing
            End If
            If Not (pFeatureCursor Is Nothing) Then
                ReleaseComObject(pFeatureCursor)
                pFeatureCursor = Nothing
            End If
            If Not (pFeature Is Nothing) Then
                ReleaseComObject(pFeature)
                pFeature = Nothing
            End If
            GC.Collect() 'garbage collect
            GetSectorCodeOfPoint = sectorCode
        End Try
    End Function

    Private Function GetSectorEnvelope(ByVal sectorCode As String) As Double()

        Dim pQueryFilter As QueryFilter = Nothing
        Dim pFeatureCursor As IFeatureCursor = Nothing
        Dim pFeature As IFeature = Nothing
        Dim SectorEnvelopeData(3) As Double
        Try
            Dim featureLayerSector As IFeatureLayer = _
                GetFeatureLayer(m_document.FocusMap, GetLayerName("Sector"))

            pQueryFilter = New QueryFilter()
            pQueryFilter.WhereClause = "AFKORTING = " & CStrSql(sectorCode)
            pFeatureCursor = featureLayerSector.Search(pQueryFilter, Nothing)
            pFeature = pFeatureCursor.NextFeature
            If Not (pFeature Is Nothing) Then
                SectorEnvelopeData(0) = pFeature.Shape.Envelope.XMin
                SectorEnvelopeData(1) = pFeature.Shape.Envelope.YMin
                SectorEnvelopeData(2) = pFeature.Shape.Envelope.XMax
                SectorEnvelopeData(3) = pFeature.Shape.Envelope.YMax
            End If
        Catch ex As Exception
            Throw ex
        Finally
            ' Release COM objects to avoid error "Cannot open any more tables."
            ' More info on http://forums.esri.com/Thread.asp?c=93&f=993&t=146570
            If Not (pQueryFilter Is Nothing) Then
                ReleaseComObject(pQueryFilter)
                pQueryFilter = Nothing
            End If
            If Not (pFeatureCursor Is Nothing) Then
                ReleaseComObject(pFeatureCursor)
                pFeatureCursor = Nothing
            End If
            If Not (pFeature Is Nothing) Then
                ReleaseComObject(pFeature)
                pFeature = Nothing
            End If
            GC.Collect() 'garbage collect
            GetSectorEnvelope = SectorEnvelopeData
        End Try
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Visualise the problem rows data table.
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	22/09/2006	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub ShowProblemRows()

        If ds Is Nothing Then
            ' There is no dataset to export from.
            MsgBox(c_Message_NoExportData, MsgBoxStyle.OkOnly, c_Title_OpladenHydranten)
        ElseIf Not ds.Tables.Contains("PROBLEMS") Then
            ' There is no table to export.
            MsgBox(c_Message_NoExportData, MsgBoxStyle.OkOnly, c_Title_OpladenHydranten)
        ElseIf ds.Tables.Item("PROBLEMS").Rows.Count > 0 Then
            ' Show data table in a new Excel.
            DataTableToXLS(ds.Tables.Item("PROBLEMS"))
        Else
            ' There are no rows to export.
            MsgBox(c_Message_NoExportData, MsgBoxStyle.OkOnly, c_Title_OpladenHydranten)
        End If

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Change the Status attribute of hydrant features from current sector,
    '''     if their identifier is not mentioned in the upload data file.
    ''' </summary>
    ''' <param name="sectorCode">The code of the current sector.</param>
    ''' <param name="providerCode">The code of the provider of the current upload data.</param>
    ''' <remarks>
    '''     Only hydrant features without EindDatum, provided by the same source,
    '''     are evaluated and possibly changed.
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	22/09/2006	Created
    '''     [Kristof Vydt]  28/09/2006  Add try...catch...finally and release COM objects.
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub FlagRemovedHydrants( _
            ByVal sectorCode As String, _
            ByVal providerCode As String)

        ' Declare COM objects.
        Dim pFeature As IFeature = Nothing
        Dim pFeatureCursor As IFeatureCursor = Nothing
        Dim pFeatureLayer As IFeatureLayer = Nothing
        Dim pQueryFilter As QueryFilter = Nothing

        Try
            ' Update progress monitor.
            OnShowProgress1(Me.ProgressBar1.Maximum, "Toekenning status 'Verwijderd' ...")

            ' Get all hydrant features that should be evaluated.
            pFeatureLayer = GetFeatureLayer(m_document.FocusMap, GetLayerName("Hydrant"))
            pQueryFilter = New QueryFilter
            pQueryFilter.WhereClause = _
                "(" & GetAttributeName("Hydrant", "EindDatum") & " IS NULL) and " & _
                "(" & GetAttributeName("Hydrant", "Bron") & "=" & CStrSql(providerCode) & ")"
            pFeatureCursor = pFeatureLayer.Search(pQueryFilter, Nothing)

            ' Loop through this feature cursor.
            pFeature = pFeatureCursor.NextFeature
            If Not pFeature Is Nothing Then
                Dim fldIdxLevNr As Integer = pFeature.Fields.FindField(GetAttributeName("Hydrant", "LeverancierNr"))
                Dim fldIdxStatus As Integer = pFeature.Fields.FindField(GetAttributeName("Hydrant", "Status"))
                While Not pFeature Is Nothing

                    ' Look for a data row with the same LeverancierNr as the feature.
                    Dim levNr As String = CType(pFeature.Value(fldIdxLevNr), String)
                    Dim row() As DataRow = ds.Tables("UPLOAD").Select("LeverancierNummer=" & CStrSql(levNr))

                    ' If no matching data row was found, then change feature state.
                    If row.Length = 0 Then

                        '==================================================================================
                        '== Action: Update existing feature status to VERWIJDERD.                        ==
                        '==================================================================================

                        ' Update "EindDatum" attribute.
                        Call SetAttributeValue(pFeature, "Hydrant", "EindDatum", Today.AddDays(-1))

                        ' Update "Status" attribute.
                        Call SetAttributeValue(pFeature, "Hydrant", "Status", "7")

                        ' Update legend code attribute value.
                        Call UpdateLegendCode(pFeature)

                        ' Increase status "verwijderd" counter.
                        Call IncrementStatusCounter(7)

                    End If

                    ' Next feature.
                    pFeature = pFeatureCursor.NextFeature
                End While
            End If

        Catch ex As Exception
            ErrorHandler(ex)

        Finally
            ' Release COM objects.
            If Not (pQueryFilter Is Nothing) Then
                ReleaseComObject(pQueryFilter)
                pQueryFilter = Nothing
            End If
            If Not (pFeatureCursor Is Nothing) Then
                ReleaseComObject(pFeatureCursor)
                pFeatureCursor = Nothing
            End If
            If Not (pFeature Is Nothing) Then
                ReleaseComObject(pFeature)
                pFeature = Nothing
            End If
            GC.Collect() 'garbage collect

        End Try

    End Sub

    '''' -----------------------------------------------------------------------------
    '''' <summary>
    ''''     The complete, complex, actual upload process (reading from Excel,
    ''''     creating new features, modifying existing features).
    '''' </summary>
    '''' <remarks>
    '''' </remarks>
    '''' <history>
    '''' 	[Kristof Vydt]	16/09/2005	Created
    '''' 	[Kristof Vydt]	23/09/2005	Evaluation codes modified.
    ''''                                 Use CodeMapping when determining HasIdentical_*Type.
    ''''                                 Multiple modifications to ignore <null> values from recordset.
    '''' 	[Kristof Vydt]	10/10/2005	Edit session management reviewed.
    '''' 	[Kristof Vydt]	24/10/2005	Add counter for evaluation "OK" & total number of processed records.
    '''' 	[Kristof Vydt]	25/10/2005	Use c_Message_LeverancierNrNotUnique for ApplicationException.
    '''' 	[Kristof Vydt]	02/08/2006	Replace some global constant by a function that reads the value from config file.
    ''''     [Kristof Vydt]  15/09/2006  Deprecated.
    '''' </history>
    '''' -----------------------------------------------------------------------------
    '    Private Sub StartUpload()

    '        Dim oConn As ADODB.Connection 'connection to import data
    '        Dim rsUpload As ADODB.Recordset 'recordset of import data
    '        Dim TmpStr As String 'string for temporary use
    '        Dim TmpStrArr As String() 'string array for temporary use
    '        Dim FieldIndex As Integer 'attribute index
    '        Dim i As Integer 'loop index
    '        Dim j As Integer 'loop index
    '        Dim Success As Boolean 'indication of success or failure
    '        Dim Attributes As Hashtable = New Hashtable 'hashtable with feature attributes
    '        Dim ListXlsIDs As IList = New ArrayList  'list of LinkIDs in import XLS file
    '        Dim NumberOfProcessedRecords As Long 'number of processed XLS records

    '        'ArcGIS object pointers
    '        Dim pEditor As IEditor2 = GetEditorReference(m_application)
    '        Dim pFLayer As IFeatureLayer 'hydrant layer
    '        Dim pQueryFilter As IQueryFilter 'hydrant query filter
    '        Dim pFCursor As IFeatureCursor 'hydrant cursor
    '        Dim pFeature As IFeature 'hydrant feature

    '        Try
    '            'Validations before starting ...
    '            'An error message is displayed instead of the form, 
    '            'when temporary status hydrants for current sector are detected.

    '            'Read SheetName from import schema file.
    '            Dim SheetName As String '= INIRead(OpenFileDialogSchemaFile.FileName, "Globals", "SheetName")
    '            If Len(SheetName) = 0 Then Throw New MissingImportSchemaValue("Globals", "SheetName")

    '            'Connect to the import data file and read it as a recordset.
    '            oConn = ConnectXLS(OpenFileDialogDataFile.FileName)
    '            rsUpload = ReadXLS(oConn, SheetName)

    '            'Determine the total number of records and set progress monitor (1) max.
    '            Dim RecCount As Integer = rsUpload.RecordCount
    '            If RecCount = -1 Then
    '                rsUpload.MoveLast()
    '                RecCount = rsUpload.RecordCount
    '                If RecCount = -1 Then
    '                    rsUpload.MoveFirst()
    '                    RecCount = 0
    '                    While Not rsUpload.EOF
    '                        RecCount += 1
    '                        rsUpload.MoveNext()
    '                    End While
    '                End If
    '            End If
    '            OnSetMaxProgress1(RecCount)

    '            'Read, parse and store import schema file settings.
    '            'ReadSchemaDefinition(OpenFileDialogSchemaFile.FileName, rsUpload.Fields.Count)

    '            'Determine if there exist already hydrant for this sector in the geodatabase.
    '            'If not, no individual record comparison is required later on,
    '            '        just adding all valid import records.
    '            pFLayer = GetFeatureLayer(m_document.FocusMap, GetLayerName("Hydrant"))
    '            If pFLayer Is Nothing Then Throw New LayerNotFoundException(GetLayerName("Hydrant"))
    '            pFCursor = pFLayer.Search(Nothing, Nothing)
    '            pFeature = pFCursor.NextFeature
    '            Dim SkipComparison As Boolean = (pFeature Is Nothing)
    '            If Not pFCursor Is Nothing Then Marshal.ReleaseComObject(pFCursor)
    '            If Not pFeature Is Nothing Then Marshal.ReleaseComObject(pFeature)

    '            'Loop through the Excel worksheet, adding and updating existing hydrant features.
    '            rsUpload.MoveFirst()
    '            While Not rsUpload.EOF

    '                Try 'start a codeblock to rollback in case of unexpected exceptions

    '                    'Update progress monitor.
    '                    TmpStr = CStr(rsUpload(Columns.Item("LeverancierNr")).Value)
    '                    OnShowProgress1(Me.ProgressBar1.Value + 1, TmpStr)

    '                    'Build a list of LinkIDs that are not supposed to get status "verwijderd".
    '                    ListXlsIDs.Add(TmpStr)

    '                    '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
    '                    ' Evaluate if current record should be processed or ignored.
    '                    '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 

    '                    'Validation: Sector check
    '                    If CStr(rsUpload(Columns.Item("Sector")).Value) <> CStr(Sectors.Item(TextBoxSectorCode.Text)) Then
    '                        GoTo NextImportRecord
    '                    End If

    '                    'Validation: Ignore field values
    '                    For Each ColIndex As Integer In IgnoreFieldValues.Keys
    '                        TmpStrArr = (CStr(IgnoreFieldValues(ColIndex))).Split(c_ListSeparator)
    '                        For i = 0 To TmpStrArr.Length - 1
    '                            If Trim(CStr(rsUpload(ColIndex).Value)) = CStr(TmpStrArr.GetValue(i)) Then
    '                                rsUpload(Columns.Item("Evaluatie")).Value = "genegeerd"
    '                                GoTo NextImportRecord
    '                            End If
    '                        Next
    '                    Next

    '                    'Increment the total number of processed records.
    '                    NumberOfProcessedRecords = NumberOfProcessedRecords + 1

    '                    'Validation: Error field values
    '                    For Each ColIndex As Integer In ErrorFieldValues.Keys
    '                        TmpStrArr = (CStr(ErrorFieldValues(ColIndex))).Split(c_ListSeparator)
    '                        For i = 0 To TmpStrArr.Length - 1
    '                            Dim rsFieldValue As String = ""
    '                            If Not TypeOf rsUpload(ColIndex).Value Is System.DBNull Then
    '                                rsFieldValue = Trim(CStr(rsUpload(ColIndex).Value))
    '                            End If
    '                            If rsFieldValue = CStr(TmpStrArr.GetValue(i)) Then
    '                                rsUpload(Columns.Item("Evaluatie")).Value = "datafout"
    '                                IncrementStatusCounter(0)
    '                                GoTo NextImportRecord
    '                            End If
    '                        Next
    '                    Next

    '                    'Determine record ID
    '                    Dim RecID As String = Trim(CStr(rsUpload.Fields(Columns.Item("LeverancierNr")).Value))
    '                    If RecID = "" Then 'Invalid RecID results in error.
    '                        'Write evaluation to data import file.
    '                        rsUpload(Columns.Item("Evaluatie")).Value = "datafout"

    '                        'Increase status counter.
    '                        IncrementStatusCounter(0)

    '                        Exit Try
    '                        GoTo NextImportRecord
    '                    End If

    '                    'Start edit session on hydrants, to be able to rollback in case of unexpected exceptions.
    '                    EditSessionStart(pEditor, GetFeatureLayer(m_document.FocusMap, GetLayerName("Hydrant")))

    '                    If SkipComparison Then

    '                        '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
    '                        ' Add each import record as new feature in case of empty feature class.
    '                        '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 

    '                        Attributes = New Hashtable
    '                        '- BeginDatum
    '                        Attributes.Add("BeginDatum", CDate(Now()))
    '                        '- Bron
    '                        Attributes.Add("Bron", CStr(Globals.Item("ProviderCode")))
    '                        '- CoordX
    '                        Attributes.Add("CoordX", ReadCoordinates(CStr(rsUpload(Columns.Item("CoordX")).Value)))
    '                        '- CoordY
    '                        Attributes.Add("CoordY", ReadCoordinates(CStr(rsUpload(Columns.Item("CoordY")).Value)))
    '                        '- Diameter
    '                        If Trim(CStr(rsUpload(Columns.Item("Diameter")).Value)) <> "" Then _
    '                            Attributes.Add("Diameter", Trim(CStr(rsUpload(Columns.Item("Diameter")).Value)))
    '                        '- Status
    '                        Attributes.Add("Status", CStr(6))     'status=nieuw
    '                        '- LeverancierNr
    '                        If Trim(CStr(rsUpload(Columns.Item("LeverancierNr")).Value)) <> "" Then _
    '                            Attributes.Add("LeverancierNr", Trim(CStr(rsUpload(Columns.Item("LeverancierNr")).Value)))
    '                        '- LeidingType
    '                        If Trim(CStr(rsUpload(Columns.Item("LeidingType")).Value)) <> "" Then _
    '                            Attributes.Add("LeidingType", CodeMapping(LeidingTypeMapping, Trim(CStr(rsUpload(Columns.Item("LeidingType")).Value))))
    '                        '- LeidingNr
    '                        If Trim(CStr(rsUpload(Columns.Item("LeidingNr")).Value)) <> "" Then _
    '                            Attributes.Add("LeidingNr", Trim(CStr(rsUpload(Columns.Item("LeidingNr")).Value)))
    '                        '- HydrantType
    '                        If Trim(CStr(rsUpload(Columns.Item("HydrantType")).Value)) <> "" Then _
    '                            Attributes.Add("HydrantType", CodeMapping(HydrantTypeMapping, Trim(CStr(rsUpload(Columns.Item("HydrantType")).Value))))
    '                        Success = AddHydrant(pFLayer, Attributes)
    '                        '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 

    '                        'No action for related annotations.

    '                        'Write evaluation to data import file.
    '                        rsUpload(Columns.Item("Evaluatie")).Value = "nieuw"

    '                        'Increase status counter.
    '                        IncrementStatusCounter(6)

    '                    Else

    '                        'Check if active hydrants with this ID already exist.
    '                        pQueryFilter = New QueryFilter
    '                        pQueryFilter.WhereClause = _
    '                            "(" & GetAttributeName("Hydrant", "LeverancierNr") & "='" & RecID & "') and " & _
    '                            "(" & GetAttributeName("Hydrant", "Status") & "='1')"
    '                        pFCursor = pFLayer.Search(pQueryFilter, Nothing)
    '                        pFeature = pFCursor.NextFeature
    '                        If pFeature Is Nothing Then

    '                            '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
    '                            ' Add import record with new <LeverancierNr> as new hydrant feature.
    '                            '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 

    '                            Attributes = New Hashtable
    '                            '- BeginDatum
    '                            Attributes.Add("BeginDatum", CDate(Now()))
    '                            '- Bron
    '                            Attributes.Add("Bron", CStr(Globals.Item("ProviderCode")))
    '                            '- CoordX
    '                            Attributes.Add("CoordX", ReadCoordinates(CStr(rsUpload(Columns.Item("CoordX")).Value)))
    '                            '- CoordY
    '                            Attributes.Add("CoordY", ReadCoordinates(CStr(rsUpload(Columns.Item("CoordY")).Value)))
    '                            '- Diameter
    '                            If Not TypeOf rsUpload(Columns.Item("Diameter")).Value Is System.DBNull Then _
    '                                If Trim(CStr(rsUpload(Columns.Item("Diameter")).Value)) <> "" Then _
    '                                    Attributes.Add("Diameter", Trim(CStr(rsUpload(Columns.Item("Diameter")).Value)))
    '                            '- Status
    '                            Attributes.Add("Status", CStr(6))     'status=nieuw
    '                            '- LeverancierNr
    '                            If Not TypeOf rsUpload(Columns.Item("LeverancierNr")).Value Is System.DBNull Then _
    '                                If Trim(CStr(rsUpload(Columns.Item("LeverancierNr")).Value)) <> "" Then _
    '                                    Attributes.Add("LeverancierNr", Trim(CStr(rsUpload(Columns.Item("LeverancierNr")).Value)))
    '                            '- LeidingType
    '                            If Not TypeOf rsUpload(Columns.Item("LeidingType")).Value Is System.DBNull Then _
    '                                If Trim(CStr(rsUpload(Columns.Item("LeidingType")).Value)) <> "" Then _
    '                                    Attributes.Add("LeidingType", CodeMapping(LeidingTypeMapping, CStr(rsUpload(Columns.Item("LeidingType")).Value)))
    '                            '- LeidingNr
    '                            If Not TypeOf rsUpload(Columns.Item("LeidingNr")).Value Is System.DBNull Then _
    '                                If Trim(CStr(rsUpload(Columns.Item("LeidingNr")).Value)) <> "" Then _
    '                                    Attributes.Add("LeidingNr", Trim(CStr(rsUpload(Columns.Item("LeidingNr")).Value)))
    '                            '- HydrantType
    '                            If Not TypeOf rsUpload(Columns.Item("HydrantType")).Value Is System.DBNull Then _
    '                                If Trim(CStr(rsUpload(Columns.Item("HydrantType")).Value)) <> "" Then _
    '                                    Attributes.Add("HydrantType", CodeMapping(HydrantTypeMapping, Trim(CStr(rsUpload(Columns.Item("HydrantType")).Value))))
    '                            Success = AddHydrant(pFLayer, Attributes)
    '                            '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 

    '                            'No action for related annotations.

    '                            'Write evaluation to data import file.
    '                            rsUpload(Columns.Item("Evaluatie")).Value = "nieuw"

    '                            'Increase status counter.
    '                            IncrementStatusCounter(6)

    '                        Else

    '                            'Compare each attribute of import record to hydrant with same ID.

    '                            Dim hasIdentical_CoordX As Boolean = False 'indication if value in Excel equals feature attribute value
    '                            Dim hasIdentical_CoordY As Boolean = False
    '                            Dim hasIdentical_Diameter As Boolean = False
    '                            Dim hasIdentical_HydrantType As Boolean = False
    '                            Dim hasIdentical_LeidingType As Boolean = False
    '                            Dim hasIdentical_LeidingNr As Boolean = False

    '                            hasIdentical_CoordX = AttributeIsConformRecord(pFeature, GetAttributeName("Hydrant", "CoordX"), rsUpload, "CoordX", Nothing)
    '                            hasIdentical_CoordY = AttributeIsConformRecord(pFeature, GetAttributeName("Hydrant", "CoordY"), rsUpload, "CoordY", Nothing)
    '                            hasIdentical_Diameter = AttributeIsConformRecord(pFeature, GetAttributeName("Hydrant", "Diameter"), rsUpload, "Diameter", Nothing)
    '                            hasIdentical_HydrantType = AttributeIsConformRecord(pFeature, GetAttributeName("Hydrant", "HydrantType"), rsUpload, "HydrantType", HydrantTypeMapping)
    '                            hasIdentical_LeidingType = AttributeIsConformRecord(pFeature, GetAttributeName("Hydrant", "LeidingType"), rsUpload, "LeidingType", LeidingTypeMapping)
    '                            hasIdentical_LeidingNr = AttributeIsConformRecord(pFeature, GetAttributeName("Hydrant", "LeidingNr"), rsUpload, "LeidingNr", Nothing)

    '                            If hasIdentical_Diameter And hasIdentical_HydrantType And hasIdentical_LeidingType And _
    '                                hasIdentical_CoordX And hasIdentical_CoordY Then
    '                                'Distinctive attributes all have the same values as the import record...

    '                                '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
    '                                ' The import record has a known <LeverancierNr> and critical attributes
    '                                ' match the known feature. Make sure it has the same <LeidingNr>.
    '                                '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 

    '                                'No need to create new hydrant features.

    '                                'Copy LeidingNr without further notice, if it is different.
    '                                If Not hasIdentical_LeidingNr Then
    '                                    '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
    '                                    Attributes = New Hashtable
    '                                    Attributes.Add("LeidingNr", Trim(CStr(rsUpload(Columns.Item("LeidingNr")).Value)))
    '                                    Success = ModifyHydrantAttributes(pFeature, Attributes)
    '                                    '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
    '                                End If

    '                                'Write evaluation to data import file.
    '                                rsUpload(Columns.Item("Evaluatie")).Value = "ok"

    '                                'Increase status counter.
    '                                IncrementStatusCounter(1)

    '                            Else
    '                                'Not all distinctive attributes have the same values as the import record...

    '                                '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
    '                                ' The import record has a known <LeverancierNr> but critical attributes
    '                                ' differ from the known feature. Add the record as new hydrant feature.
    '                                '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 

    '                                Attributes = New Hashtable
    '                                '- Aanduiding
    '                                FieldIndex = pFeature.Fields.FindField(GetAttributeName("Hydrant", "Aanduiding"))
    '                                If Not TypeOf pFeature.Value(FieldIndex) Is System.DBNull Then _
    '                                    If CStr(pFeature.Value(FieldIndex)) <> "" Then _
    '                                        Attributes.Add("Aanduiding", CStr(pFeature.Value(FieldIndex)))
    '                                '- BeginDatum
    '                                Attributes.Add("BeginDatum", CDate(Now()))
    '                                '- BrandweerNr
    '                                FieldIndex = pFeature.Fields.FindField(GetAttributeName("Hydrant", "BrandweerNr"))
    '                                If Not TypeOf pFeature.Value(FieldIndex) Is System.DBNull Then _
    '                                    If CStr(pFeature.Value(FieldIndex)) <> "" Then _
    '                                        Attributes.Add("BrandweerNr", CStr(pFeature.Value(FieldIndex)))
    '                                '- Bron
    '                                Attributes.Add("Bron", CStr(Globals.Item("ProviderCode")))
    '                                '- CoordX
    '                                Attributes.Add("CoordX", ReadCoordinates(CStr(rsUpload(Columns.Item("CoordX")).Value)))
    '                                '- CoordY
    '                                Attributes.Add("CoordY", ReadCoordinates(CStr(rsUpload(Columns.Item("CoordY")).Value)))
    '                                '- Diameter
    '                                If Not TypeOf rsUpload(Columns.Item("Diameter")).Value Is System.DBNull Then _
    '                                    If Trim(CStr(rsUpload(Columns.Item("Diameter")).Value)) <> "" Then _
    '                                        Attributes.Add("Diameter", Trim(CStr(rsUpload(Columns.Item("Diameter")).Value)))
    '                                '- HydrantType
    '                                If Not TypeOf rsUpload(Columns.Item("HydrantType")).Value Is System.DBNull Then _
    '                                    If Trim(CStr(rsUpload(Columns.Item("HydrantType")).Value)) <> "" Then _
    '                                        Attributes.Add("HydrantType", CodeMapping(HydrantTypeMapping, Trim(CStr(rsUpload(Columns.Item("HydrantType")).Value))))
    '                                '- LeidingType
    '                                If Not TypeOf rsUpload(Columns.Item("LeidingType")).Value Is System.DBNull Then _
    '                                    If Trim(CStr(rsUpload(Columns.Item("LeidingType")).Value)) <> "" Then _
    '                                        Attributes.Add("LeidingType", CodeMapping(LeidingTypeMapping, Trim(CStr(rsUpload(Columns.Item("LeidingType")).Value))))
    '                                '- Leiding Nr
    '                                If Not TypeOf rsUpload(Columns.Item("LeidingNr")).Value Is System.DBNull Then _
    '                                    If Trim(CStr(rsUpload(Columns.Item("LeidingNr")).Value)) <> "" Then _
    '                                        Attributes.Add("LeidingNr", Trim(CStr(rsUpload(Columns.Item("LeidingNr")).Value)))
    '                                '- LeverancierNr
    '                                If Not TypeOf rsUpload(Columns.Item("LeverancierNr")).Value Is System.DBNull Then _
    '                                    If Trim(CStr(rsUpload(Columns.Item("LeverancierNr")).Value)) <> "" Then _
    '                                        Attributes.Add("LeverancierNr", Trim(CStr(rsUpload(Columns.Item("LeverancierNr")).Value)))
    '                                '- Ligging
    '                                FieldIndex = pFeature.Fields.FindField(GetAttributeName("Hydrant", "Ligging"))
    '                                If Not TypeOf pFeature.Value(FieldIndex) Is System.DBNull Then _
    '                                    If CStr(pFeature.Value(FieldIndex)) <> "" Then _
    '                                        Attributes.Add("Ligging", CStr(pFeature.Value(FieldIndex)))
    '                                '- Postcode
    '                                FieldIndex = pFeature.Fields.FindField(GetAttributeName("Hydrant", "Postcode"))
    '                                If Not TypeOf pFeature.Value(FieldIndex) Is System.DBNull Then _
    '                                    If CStr(pFeature.Value(FieldIndex)) <> "" Then _
    '                                        Attributes.Add("Postcode", CStr(pFeature.Value(FieldIndex)))
    '                                '- Status
    '                                If hasIdentical_Diameter And hasIdentical_LeidingType And hasIdentical_HydrantType Then
    '                                    Attributes.Add("Status", CStr(8)) 'status=nakijken coördinaten
    '                                ElseIf hasIdentical_CoordX And hasIdentical_CoordY Then
    '                                    Attributes.Add("Status", CStr(9)) 'status=nakijken attributen
    '                                Else
    '                                    Attributes.Add("Status", CStr(10))  'status=nakijken attributen en coördinaten
    '                                End If
    '                                '- StraatCode
    '                                FieldIndex = pFeature.Fields.FindField(GetAttributeName("Hydrant", "Straatcode"))
    '                                If Not TypeOf pFeature.Value(FieldIndex) Is System.DBNull Then _
    '                                    If Trim(CStr(pFeature.Value(FieldIndex))) <> "" Then _
    '                                        Attributes.Add("StraatCode", Trim(CStr(pFeature.Value(FieldIndex))))
    '                                '- Straatnaam
    '                                FieldIndex = pFeature.Fields.FindField(GetAttributeName("Hydrant", "Straatnaam"))
    '                                If Not TypeOf pFeature.Value(FieldIndex) Is System.DBNull Then _
    '                                    If Trim(CStr(pFeature.Value(FieldIndex))) <> "" Then _
    '                                        Attributes.Add("Straatnaam", Trim(CStr(pFeature.Value(FieldIndex))))
    '                                Success = AddHydrant(pFLayer, Attributes)
    '                                '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 

    '                                '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
    '                                ' Modify existing hydrant to <historiek>.
    '                                '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 

    '                                Attributes = New Hashtable
    '                                '- Einddatum
    '                                Attributes.Add("EindDatum", CDate(Now().AddDays(-1)))
    '                                '- Status
    '                                Attributes.Add("Status", CStr(3))   'status=historiek
    '                                '- Label(s) remain linked to the same LinkID.
    '                                '- LegendCode is automatically recalculated.
    '                                Success = ModifyHydrantAttributes(pFeature, Attributes)
    '                                '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 

    '                                'TBD: Transfer the annotations to a new LinkID.
    '                                'Dim pAnnoFLayer As IFeatureLayer = GetFeatureLayer(m_document.FocusMap, GetLayerName("HydrantAnno"))
    '                                'Dim LinkFieldIndex As Integer = pFeature.Fields.FindField(GetAttributeName("Hydrant","LinkID"))
    '                                'If LinkFieldIndex = -1 Then Throw New AttributeNotFoundException(GetLayerName("Hydrant"), GetAttributeName("Hydrant","LinkID"))
    '                                'Dim oldLinkID As String = CStr(pFeature.Value(LinkFieldIndex))
    '                                'Dim newLinkID As String = CStr(rsUpload(Columns.Item("LeverancierNr")).Value)
    '                                'RelinkAnnotations(CType(pAnnoFLayer, IAnnotationLayer), oldLinkID, newLinkID)
    '                                'If Not pAnnoFLayer Is Nothing Then Marshal.ReleaseComObject(pAnnoFLayer)

    '                                'Write evaluation to data import file,
    '                                'and update status counter.
    '                                If hasIdentical_Diameter And hasIdentical_LeidingType And hasIdentical_HydrantType Then
    '                                    rsUpload(Columns.Item("Evaluatie")).Value = "nakijken_c" 'status=nakijken coördinaten
    '                                    IncrementStatusCounter(8)
    '                                ElseIf hasIdentical_CoordX And hasIdentical_CoordY Then
    '                                    rsUpload(Columns.Item("Evaluatie")).Value = "nakijken_a" 'status=nakijken attributen
    '                                    IncrementStatusCounter(9)
    '                                Else
    '                                    rsUpload(Columns.Item("Evaluatie")).Value = "nakijken_ac"  'status=nakijken attributen en coördinaten
    '                                    IncrementStatusCounter(10)
    '                                End If

    '                            End If 'End of attributes comparison.

    '                            pFeature = pFCursor.NextFeature
    '                            If Not pFeature Is Nothing Then
    '                                'More than 1 active hydrant found with the same ID as the import record.
    '                                'In that case, the data is corrupted: LeverancierNr ought to be unique for active hydrants.
    '                                Throw New ApplicationException(Replace(c_Message_LeverancierNrNotUnique, "^0", RecID))
    '                            End If

    '                        End If 'End of another active hydrant found.

    '                        If Not pQueryFilter Is Nothing Then Marshal.ReleaseComObject(pQueryFilter)
    '                        If Not pFCursor Is Nothing Then Marshal.ReleaseComObject(pFCursor)
    '                        If Not pFeature Is Nothing Then Marshal.ReleaseComObject(pFeature)

    '                    End If

    '                    'Confirm edit and save changes.
    '                    EditSessionSave(pEditor, "Import hydrant " & RecID)

    '                Catch ex As Exception

    '                    '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
    '                    ' Rollback geodatabase edits in case of unexpected exceptions.
    '                    '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 

    '                    'Abort edit session without saving changes.
    '                    EditSessionAbort(pEditor)

    '                    'Write evaluation to data import file.
    '                    rsUpload(Columns.Item("Evaluatie")).Value = "programmafout"

    '                    'Increase status counter.
    '                    IncrementStatusCounter(0)

    '                End Try 'end of rollback codeblock

    'NextImportRecord:
    '                rsUpload.MoveNext()
    '            End While

    '            'Progress monitor update.
    '            OnShowProgress1(-2, "Eerste doorloop beëindigd")

    '            If Not SkipComparison Then

    '                '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
    '                ' For known hydrants that are not listed in the import file,
    '                ' change status to <verwijderd>.
    '                '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 

    '                'Loop through all active hydrants of current provider.
    '                pQueryFilter = New QueryFilter
    '                pQueryFilter.WhereClause = _
    '                    "(" & GetAttributeName("Hydrant", "Status") & "='1') and " & _
    '                    "(" & GetAttributeName("Hydrant", "Bron") & "='" & CStr(Globals.Item("ProviderCode")) & "')"
    '                pFCursor = pFLayer.Search(pQueryFilter, Nothing)
    '                pFeature = pFCursor.NextFeature
    '                If Not pFeature Is Nothing Then

    '                    'Set progress monitor (2) max to the total number of records.
    '                    RecCount = 0
    '                    While Not pFeature Is Nothing
    '                        RecCount += 1
    '                        pFeature = pFCursor.NextFeature
    '                    End While
    '                    'OnSetMaxProgress2(RecCount)

    '                    'Look for hydrant features are not registered in the XLS.
    '                    pFCursor = pFLayer.Search(pQueryFilter, Nothing)
    '                    pFeature = pFCursor.NextFeature
    '                    If Not pFeature Is Nothing Then

    '                        'Determine field index.
    '                        Dim FieldIndex_linkID As Integer = pFeature.Fields.FindField(GetAttributeName("Hydrant", "LinkID"))
    '                        Dim FieldIndex_status As Integer = pFeature.Fields.FindField(GetAttributeName("Hydrant", "Status"))

    '                        'Start edit session on hydrants.
    '                        EditSessionStart(pEditor, GetFeatureLayer(m_document.FocusMap, GetLayerName("Hydrant")))

    '                        'Loop through the recordset.
    '                        Try
    '                            While Not pFeature Is Nothing

    '                                'Get the feature ID.
    '                                Dim LinkIDValue As String = CStr(pFeature.Value(FieldIndex_linkID))

    '                                'Progress monitor update.
    '                                'OnShowProgress2(Me.ProgressBar2.Value + 1, LinkIDValue)

    '                                'Check if the ID was already processed during integration.
    '                                If ListXlsIDs.IndexOf(LinkIDValue) = -1 Then
    '                                    '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
    '                                    Attributes = New Hashtable
    '                                    '- Status
    '                                    Attributes.Add("Status", CStr(7))   'status=verwijderd
    '                                    Success = ModifyHydrantAttributes(pFeature, Attributes)
    '                                    '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
    '                                    'Increase status counter.
    '                                    IncrementStatusCounter(7)
    '                                End If

    '                                'Loop to next feature in the cursor.
    '                                pFeature = pFCursor.NextFeature

    '                            End While

    '                            'Stop edit session while saving the changes.
    '                            EditSessionSave(pEditor)

    '                        Catch ex As Exception

    '                            'Stop edit session without storing the changes.
    '                            EditSessionAbort(pEditor)

    '                        End Try
    '                    End If 'There were features in the cursor.

    '                    'Progress monitor update.
    '                    'OnShowProgress2(-2, "Tweede doorloop beëindigd")

    '                End If 'There were features.
    '            End If 'SkipComparison

    '            'Release all remaining COMObjects.
    '            If Not pFLayer Is Nothing Then Marshal.ReleaseComObject(pFLayer)
    '            If Not pQueryFilter Is Nothing Then Marshal.ReleaseComObject(pQueryFilter)
    '            If Not pFCursor Is Nothing Then Marshal.ReleaseComObject(pFCursor)
    '            If Not pFeature Is Nothing Then Marshal.ReleaseComObject(pFeature)

    '            'Progress monitor update.
    '            'OnShowProgress2(-2, "Opladen en integreren beëindigd: " & NumberOfProcessedRecords & " records.")

    '        Catch ex As Exception
    '            Throw ex
    '        Finally
    '            'Clean up.
    '            If rsUpload.State = ADODB.ObjectStateEnum.adStateOpen Then rsUpload.Close()
    '            rsUpload = Nothing
    '            If oConn.State = ADODB.ObjectStateEnum.adStateOpen Then oConn.Close()
    '            oConn = Nothing
    '        End Try

    '    End Sub

    '''' -----------------------------------------------------------------------------
    '''' <summary>
    ''''     Read import schema definition file and store parameters in private variables.
    '''' </summary>
    '''' <param name="SchemaFile">
    ''''     The full local path of the import schema difinition file.
    '''' </param>
    '''' <param name="ColumnCount">
    ''''     The number of columns in the recordset based on the Excel-file.
    '''' </param>
    '''' <remarks>
    '''' </remarks>
    '''' <history>
    '''' 	[Kristof Vydt]	16/09/2005	Created
    ''''     [Kristof Vydt]  15/09/2006  Deprecated.
    '''' </history>
    '''' -----------------------------------------------------------------------------
    'Private Sub ReadSchemaDefinition( _
    '        ByVal SchemaFile As String, _
    '        ByVal ColumnCount As Integer)

    '    Dim TmpStr As String 'string for temporary use
    '    Dim TmpInt As Integer 'integer for temporary use
    '    Dim TmpStrArr As String() 'string array for temporary use
    '    Dim i As Integer 'loop index

    '    Try
    '        'Take care: The recordset contains only non-empty columns of the Excel file.
    '        '           Therefore it is necessary to check if column indexes in schema file
    '        '           fall within the range of recordset fields.

    '        'Read [Globals] from import schema file.
    '        Globals = New Hashtable
    '        '- SheetName
    '        'TmpStr = INIRead(OpenFileDialogSchemaFile.FileName, "Globals", "SheetName")
    '        If TmpStr = "" Then Throw New MissingImportSchemaValue("Globals", "SheetName")
    '        Globals.Add("SheetName", TmpStr)
    '        '- ProviderCode
    '        'TmpStr = INIRead(OpenFileDialogSchemaFile.FileName, "Globals", "ProviderCode")
    '        If TmpStr = "" Then Throw New MissingImportSchemaValue("Globals", "ProviderCode")
    '        Globals.Add("ProviderCode", TmpStr)
    '        'TmpStr = INIRead(OpenFileDialogSchemaFile.FileName, "Globals", "DecimalSeparator")
    '        '- DecimalSeparator
    '        If TmpStr = "" Then Throw New MissingImportSchemaValue("Globals", "DecimalSeparator")
    '        If Len(TmpStr) <> 1 Then Throw New InvalidImportSchemaValue("Globals", "DecimalSeparator")
    '        Globals.Add("DecimalSeparator", TmpStr)
    '        '- ThousandsSeparator (optional)
    '        'TmpStr = INIRead(OpenFileDialogSchemaFile.FileName, "Globals", "ThousandsSeparator")
    '        If Len(TmpStr) > 1 Then Throw New InvalidImportSchemaValue("Globals", "ThousandsSeparator")
    '        Globals.Add("ThousandsSeparator", TmpStr)

    '        'Read required [ColumnMapping] from import schema file.
    '        Columns = New Hashtable
    '        TmpInt = GetColumnIndex("BrandweerNr", ColumnCount)
    '        If TmpInt > -1 Then Columns.Add("BrandweerNr", TmpInt)
    '        TmpInt = GetColumnIndex("LeverancierNr", ColumnCount)
    '        If TmpInt > -1 Then Columns.Add("LeverancierNr", TmpInt)
    '        TmpInt = GetColumnIndex("CoordX", ColumnCount)
    '        If TmpInt > -1 Then Columns.Add("CoordX", TmpInt)
    '        TmpInt = GetColumnIndex("CoordY", ColumnCount)
    '        If TmpInt > -1 Then Columns.Add("CoordY", TmpInt)
    '        TmpInt = GetColumnIndex("LeidingNr", ColumnCount)
    '        If TmpInt > -1 Then Columns.Add("LeidingNr", TmpInt)
    '        TmpInt = GetColumnIndex("HydrantType", ColumnCount)
    '        If TmpInt > -1 Then Columns.Add("HydrantType", TmpInt)
    '        TmpInt = GetColumnIndex("Diameter", ColumnCount)
    '        If TmpInt > -1 Then Columns.Add("Diameter", TmpInt)
    '        TmpInt = GetColumnIndex("LeidingType", ColumnCount)
    '        If TmpInt > -1 Then Columns.Add("LeidingType", TmpInt)
    '        TmpInt = GetColumnIndex("Sector", ColumnCount)
    '        If TmpInt > -1 Then Columns.Add("Sector", TmpInt)
    '        TmpInt = GetColumnIndex("Evaluatie", ColumnCount)
    '        If TmpInt > -1 Then Columns.Add("Evaluatie", TmpInt)

    '        'Read optional [IgnoreFieldValues] from import schema file into hashtable.
    '        IgnoreFieldValues = New Hashtable
    '        TmpStr = INIRead(SchemaFile, "IgnoreFieldValues")
    '        If Len(TmpStr) > 0 Then
    '            TmpStr = TmpStr.Replace(ControlChars.NullChar, c_ListSeparator)
    '            TmpStrArr = TmpStr.Split(c_ListSeparator)
    '            For i = 0 To TmpStrArr.Length - 1
    '                If CStr(TmpStrArr.GetValue(i)) <> "" Then
    '                    Dim ColIndex As Integer = CInt(TmpStrArr.GetValue(i))
    '                    If ColIndex < 0 Then Throw New InvalidImportSchemaColumnIndex("IgnoreFieldValues", CStr(TmpStrArr.GetValue(i)))
    '                    If ColIndex > ColumnCount - 1 Then Throw New InvalidImportSchemaColumnIndex("IgnoreFieldValues", CStr(TmpStrArr.GetValue(i)))
    '                    Dim ColValues As String = INIRead(SchemaFile, "IgnoreFieldValues", CStr(ColIndex))
    '                    IgnoreFieldValues.Add(ColIndex, ColValues)
    '                End If
    '            Next
    '        End If

    '        'Read optional [ErrorFieldValues] from import schema file into hashtable.
    '        ErrorFieldValues = New Hashtable
    '        TmpStr = INIRead(SchemaFile, "ErrorFieldValues")
    '        If Len(TmpStr) > 0 Then
    '            TmpStr = TmpStr.Replace(ControlChars.NullChar, c_ListSeparator)
    '            TmpStrArr = TmpStr.Split(c_ListSeparator)
    '            For i = 0 To TmpStrArr.Length - 1
    '                If CStr(TmpStrArr.GetValue(i)) <> "" Then
    '                    Dim ColIndex As Integer = CInt(TmpStrArr.GetValue(i))
    '                    If ColIndex < 0 Then Throw New InvalidImportSchemaColumnIndex("ErrorFieldValues", CStr(TmpStrArr.GetValue(i)))
    '                    If ColIndex > ColumnCount - 1 Then Throw New InvalidImportSchemaColumnIndex("ErrorFieldValues", CStr(TmpStrArr.GetValue(i)))
    '                    Dim ColValues As String = INIRead(SchemaFile, "ErrorFieldValues", CStr(ColIndex))
    '                    ErrorFieldValues.Add(ColIndex, ColValues)
    '                End If
    '            Next
    '        End If

    '        'Read required [SectorMapping] from import schema file.
    '        Sectors = New Hashtable
    '        TmpStr = INIRead(SchemaFile, "SectorMapping")
    '        If Len(TmpStr) = 0 Then Throw New MissingImportSchemaValue("SectorMapping")
    '        TmpStr = TmpStr.Replace(ControlChars.NullChar, c_ListSeparator)
    '        TmpStrArr = TmpStr.Split(c_ListSeparator)
    '        For i = 0 To TmpStrArr.Length - 1
    '            Dim SectorName As String = CStr(TmpStrArr.GetValue(i))
    '            Dim SectorCode As String = INIRead(SchemaFile, "SectorMapping", SectorName)
    '            Sectors.Add(SectorCode, SectorName)
    '        Next

    '        'Read optional [HydrantTypeMapping] from import file.
    '        HydrantTypeMapping = New Hashtable
    '        TmpStr = INIRead(SchemaFile, "HydrantTypeMapping")
    '        If Len(TmpStr) > 0 Then
    '            TmpStr = TmpStr.Replace(ControlChars.NullChar, c_ListSeparator)
    '            TmpStrArr = TmpStr.Split(c_ListSeparator)
    '            For i = 0 To TmpStrArr.Length - 1
    '                If CStr(TmpStrArr.GetValue(i)) <> "" Then
    '                    Dim Label As String = CStr(TmpStrArr.GetValue(i))
    '                    Dim Code As String = INIRead(SchemaFile, "HydrantTypeMapping", Label)
    '                    HydrantTypeMapping.Add(Label, Code)
    '                End If
    '            Next
    '        End If

    '        'Read optional [LeidingTypeMapping] from import schema file.
    '        LeidingTypeMapping = New Hashtable
    '        TmpStr = INIRead(SchemaFile, "LeidingTypeMapping")
    '        If Len(TmpStr) > 0 Then
    '            TmpStr = TmpStr.Replace(ControlChars.NullChar, c_ListSeparator)
    '            TmpStrArr = TmpStr.Split(c_ListSeparator)
    '            For i = 0 To TmpStrArr.Length - 1
    '                If CStr(TmpStrArr.GetValue(i)) <> "" Then
    '                    Dim Label As String = CStr(TmpStrArr.GetValue(i))
    '                    Dim Code As String = INIRead(SchemaFile, "LeidingTypeMapping", Label)
    '                    LeidingTypeMapping.Add(Label, Code)
    '                End If
    '            Next
    '        End If

    '    Catch ex As Exception
    '        Throw ex
    '    End Try
    'End Sub

    '''' -----------------------------------------------------------------------------
    '''' <summary>
    ''''     Return the column index for the specified column name,
    ''''     based on the [ColumnMapping] section of the importschema ini-file.
    '''' </summary>
    '''' <param name="ColumnName">
    ''''     The column name as used in the import schema ini-file.
    '''' </param>
    '''' <param name="ColumnCount">
    ''''     The number of columns, available in the import recordset.
    '''' </param>
    '''' <param name="Required">
    ''''     Indication if the column is mandatory or optional.
    '''' </param>
    '''' <returns>
    ''''     The column index, counting from 0.
    '''' </returns>
    '''' <remarks>
    ''''     The exception InvalidImportSchemaColumnIndex can be thrown.
    ''''     If required, the exception MissingImportSchemaValue can be thrown.
    '''' </remarks>
    '''' <history>
    '''' 	[Kristof Vydt]	16/09/2005	Created
    ''''     [Kristof Vydt]  21/09/2006  Deprecated.
    '''' </history>
    '''' -----------------------------------------------------------------------------
    'Private Function GetColumnIndex( _
    '        ByVal ColumnName As String, _
    '        ByVal ColumnCount As Integer, _
    '        Optional ByVal Required As Boolean = True _
    '        ) As Integer
    '    Try
    '        Dim TmpStr As String '= INIRead(OpenFileDialogSchemaFile.FileName, "ColumnMapping", ColumnName)
    '        If (TmpStr <> "") Then
    '            Dim TmpInt As Integer = CInt(TmpStr)
    '            If TmpInt < 0 Then Throw New InvalidImportSchemaColumnIndex("ColumnMapping", ColumnName, TmpStr)
    '            If TmpInt > ColumnCount - 1 Then Throw New InvalidImportSchemaColumnIndex("ColumnMapping", ColumnName, TmpStr)
    '            Return TmpInt
    '        ElseIf Required Then
    '            Throw New MissingImportSchemaValue("ColumnMapping", ColumnName)
    '        Else
    '            Return -1
    '        End If
    '    Catch ex As Exception
    '        Throw ex
    '    End Try
    'End Function

    '''' -----------------------------------------------------------------------------
    '''' <summary>
    ''''     Convert coordinate value from import file, to double type.
    '''' </summary>
    '''' <param name="Input">
    ''''     The coordinate value to convert.
    '''' </param>
    '''' <returns>
    ''''     The double value.
    '''' </returns>
    '''' <remarks>
    ''''     The DecimalSeparator and ThousandsSeparator schema settings
    ''''     are used for interpretation. Every non-numeric character is
    ''''     just ignored.
    '''' </remarks>
    '''' <history>
    '''' 	[Kristof Vydt]	16/09/2005	Created
    '''' 	[Kristof Vydt]	21/09/2005	Deprecated
    '''' </history>
    '''' -----------------------------------------------------------------------------
    'Private Function ReadCoordinates( _
    '        ByVal Input As String _
    '        ) As Double
    '    Try
    '        Dim DS As Char = CChar(Globals.Item("DecimalSeparator"))
    '        Dim TS As Char = CChar(Globals.Item("ThousandsSeparator"))
    '        Dim IsFoundDS As Boolean = False
    '        Dim SystemDS As Char = CChar((CStr(3 / 2)).Substring(1, 1))
    '        Dim FormattedInput As String = ""
    '        For i As Integer = Len(Input) - 1 To 0 Step -1
    '            If (Not IsFoundDS) And (Input.Substring(i, 1) = DS) Then
    '                IsFoundDS = True
    '                FormattedInput = SystemDS & FormattedInput
    '            Else
    '                Select Case Input.Substring(i, 1)
    '                    Case TS
    '                        'thousands separators are filtered
    '                    Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
    '                        FormattedInput = Input.Substring(i, 1) & FormattedInput
    '                    Case Else
    '                        'invalid characters are filtered
    '                End Select
    '            End If
    '        Next
    '        Return CDbl(FormattedInput)
    '    Catch ex As Exception
    '        Throw ex
    '    End Try
    'End Function

    '''' -----------------------------------------------------------------------------
    '''' <summary>
    ''''     Lookup some value in a code-value-list, and return the code.
    '''' </summary>
    '''' <param name="Mapping">
    ''''     The code-value-list.
    '''' </param>
    '''' <param name="Value">
    ''''     The value to lookup.
    '''' </param>
    '''' <returns>
    ''''     The corresponding code.
    '''' </returns>
    '''' <remarks>
    ''''     If the exect value is not found in the list,
    ''''     a default code registered with value '?' is looked for.
    ''''     If even the default code could not be found, 
    ''''     an empty string is returned.
    '''' </remarks>
    '''' <history>
    '''' 	[Kristof Vydt]	16/09/2005	Created
    ''''     [Kristof Vydt]  15/09/2006  Deprecated.
    '''' </history>
    '''' -----------------------------------------------------------------------------
    'Private Function CodeMapping( _
    '            ByVal Mapping As Hashtable, _
    '            ByVal Value As String _
    '            ) As String
    '    Try
    '        If Mapping.ContainsKey(Value) Then
    '            Return CStr(Mapping(Value))
    '        ElseIf Mapping.ContainsKey("?") Then
    '            Return CStr(Mapping("?"))
    '        Else
    '            Return ""
    '        End If
    '    Catch ex As Exception
    '        Throw ex
    '    End Try
    'End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     If the status code has a counter on the form, increase it.
    ''' </summary>
    ''' <param name="statusCode">
    '''     The status code that must be incremented.
    ''' </param>
    ''' <param name="increment">
    '''     Optional increment size. Default is 1.
    ''' </param>
    ''' <remarks>
    '''     Status code 1 is used for evaluation OK (= import with no or minor changes). 
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    ''' 	[Kristof Vydt]	24/10/2005	Add counter for "OK".
    ''' 	[Kristof Vydt]	28/09/2006	Move refresh to the finally.
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub IncrementStatusCounter( _
                 ByVal statusCode As Integer, _
        Optional ByVal increment As Integer = 1)

        Try
            Select Case statusCode
                Case 0 'error record
                    TextBoxCountStatus0.Text = CStr(CInt(TextBoxCountStatus0.Text) + increment)
                Case 1 'status "actief" - used for "OK"
                    TextBoxCountStatus1.Text = CStr(CInt(TextBoxCountStatus1.Text) + increment)
                Case 2 'status "niet bruikbaar"
                    'no counter for this
                Case 3 'status "historiek"
                    'no counter for this
                Case 4 'status "defect"
                    'no counter for this
                Case 5 'status "in ontwerp"
                    'no counter for this
                Case 6 'status "nieuw"
                    TextBoxCountStatus6.Text = CStr(CInt(TextBoxCountStatus6.Text) + increment)
                Case 7 'status "verwijderd"
                    TextBoxCountStatus7.Text = CStr(CInt(TextBoxCountStatus7.Text) + increment)
                Case 8 'status "nakijken_c"
                    TextBoxCountStatus8.Text = CStr(CInt(TextBoxCountStatus8.Text) + increment)
                Case 9 'status "nakijken_a"
                    TextBoxCountStatus9.Text = CStr(CInt(TextBoxCountStatus9.Text) + increment)
                Case 10 'status "nakijken_ac"
                    TextBoxCountStatus10.Text = CStr(CInt(TextBoxCountStatus10.Text) + increment)
                Case Else
            End Select
        Catch ex As Exception
            Throw ex
        Finally
            Me.Refresh()
        End Try
    End Sub

    '''' -----------------------------------------------------------------------------
    '''' <summary>
    ''''     Compare a feature attribute with a recordset field.
    ''''     Return true if both are equal or equivalent.
    '''' </summary>
    '''' <param name="pFeature">
    ''''     The hydrant feature.
    '''' </param>
    '''' <param name="featureAttributeName">
    ''''     The attribute name of the feature.
    '''' </param>
    '''' <param name="pRecordset">
    ''''     The ADODB hydrants recordset.
    '''' </param>
    '''' <param name="columnIdentifier">
    ''''     Some kind of identifier that can be linked to a column index, 
    ''''     through the Columns hashtable, based on the import schema file.
    '''' </param>
    '''' <param name="mappingTable">
    ''''     [Optional] Code mapping hashtable, to convert recordset values
    ''''     to attribute coded values.
    '''' </param>
    '''' <returns>
    ''''     Boolean
    '''' </returns>
    '''' <remarks>
    ''''     Null values are interpreted as empty strings.
    '''' </remarks>
    '''' <history>
    '''' 	[Kristof Vydt]	22/09/2005	Created
    ''''     [Kristof Vydt]  15/09/2006  Deprecated.
    '''' </history>
    '''' -----------------------------------------------------------------------------
    'Private Function AttributeIsConformRecord( _
    '    ByVal pFeature As IFeature, _
    '    ByVal featureAttributeName As String, _
    '    ByVal pRecordset As ADODB.Recordset, _
    '    ByVal columnIdentifier As String, _
    '    Optional ByVal mappingTable As Hashtable = Nothing) As Boolean

    '    Dim attrIndex As Integer 'feature attribute field index
    '    Dim attrObject As Object 'feature attribute value object
    '    Dim attrValue As String = "" 'feature attribute value text
    '    Dim colIndex As Integer 'recordset column index
    '    Dim fldObject As Object 'recordset field value object
    '    Dim fldValue As String = "" 'recordset field value text

    '    Try

    '        'Get the feature attribute value.
    '        attrIndex = pFeature.Fields.FindField(featureAttributeName)
    '        If attrIndex < 0 Then Throw New AttributeNotFoundException(GetLayerName("Hydrant"), featureAttributeName)
    '        attrObject = pFeature.Value(attrIndex)
    '        If Not TypeOf attrObject Is System.DBNull Then attrValue = CStr(attrObject)

    '        'Get the recordset field value.
    '        colIndex = CInt(Columns.Item(columnIdentifier))
    '        fldObject = pRecordset(colIndex).Value
    '        If Not TypeOf fldObject Is System.DBNull Then fldValue = Trim(CStr(fldObject))

    '        'Map the recordset field value to the corresponding feature attribute code.
    '        If Not mappingTable Is Nothing Then fldValue = CodeMapping(mappingTable, fldValue)

    '        'Compare both values.
    '        Return (attrValue = fldValue)

    '    Catch ex As Exception
    '        Throw ex
    '    End Try

    'End Function

#End Region

End Class
