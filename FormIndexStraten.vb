Option Explicit On 
Option Strict On

#Region " Imports namespaces "
Imports ADODB
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Microsoft.Office.Interop
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.Geodatabase
#End Region

''' -----------------------------------------------------------------------------
''' Project	 : Digipolis.Hydranten.BeheerHydranten
''' Class	 : Hydranten.BeheerHydranten.FormIndexStraten
''' 
''' -----------------------------------------------------------------------------
''' <summary>
'''     GUI for generating street index pages.
''' </summary>
''' <remarks>
''' </remarks>
''' <history>
''' 	[Kristof Vydt]	16/09/2005	Created
''' 	[Kristof Vydt]	26/09/2005	Open *.dot as template, and not the document itself.
''' 	[Kristof Vydt]	10/10/2005	Enable cancel button on error.
'''                                 Try-catch around calling Word template and macro.
''' 	[Kristof Vydt]	17/10/2005	Change order in catch block to enable cancel button on error.
''' 	                            Interpret aanduiding &lt;null&gt; as "".
''' 	                            Interpret diameter &lt;null&gt; as 0.
''' 	[Kristof Vydt]	28/10/2005	Correction of attribute reference in query for Sleepboot during export.
''' 	[Kristof Vydt]	13/07/2006	Read max field lengths from the existing tables, instead of using constants.
'''                                 Private Function GetMap() moved to public in ModuleToolsLib.
''' 	[Kristof Vydt]	02/08/2006	Replace some global constant by a function that reads the value from config file.
'''     [Kristof Vydt]  10/08/2006  Add validity check during export before using danger layers.
'''     [Kristof Vydt]  11/08/2006  Export last letters (p.ex. X, Y, Z) even when there are no streets for them.
'''     [Kristof Vydt]  20/02/2007  Filter NewLines out of the hydrant labels before writing to txt-file, because Word macro cannot handle them.
'''     [Kristof Vydt]  14/03/2007  Add sort column in streetlist lookup recordset &amp; table.
'''                                 Isolate the different export blocks in separate procedures.
'''     [Kristof Vydt]  28/03/2007  Add dummy export line at the beginning of the output file.
'''     [Koen Vermeer]  23/06/2008  Conversion to ArcGIS 9.2
'''     [Koen Vermeer]  26/06/2007  Filter NewLines out of the hydrant labels before writing to txt-file, because Word macro cannot handle them. (Nu ook opgelost voor bovengrondse hydranten)
''' </history>
''' -----------------------------------------------------------------------------
Public NotInheritable Class FormIndexStraten
    Inherits System.Windows.Forms.Form

#Region " Private variables & constants "

    'Constants.
    Private Const c_lookupTableName As String = "straatlijst" 'name of table in personal geodatabase for street lookup info
    Private Const c_refListFieldName As String = "REFERENTIE" 'name of field with list of kwadrant references
    Private Const c_streetFieldName As String = "STRAATNAAM" 'name of field with streetname
    Private Const c_sortFieldName As String = "SORTERING" 'name of field with sort text

    'Locals.
    Private m_application As IMxApplication 'set by constructor
    Private m_document As IMxDocument 'set by constructor
    Private m_refListMaxLength As Integer 'maximum number of characters available for the reference list (a ";"-separated list of kwadrants, each consisting of 2 char raster page + 1 char kwadrant letter)
    Private m_streetMaxLength As Integer 'maximum number of characters available for a streetname
    Private m_sortMaxLength As Integer 'maximum number of characters available for a sortstring
    Private m_hydrantMaxLength As Integer 'maximum number of characters available a hydrant label

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
    Friend WithEvents ButtonOK As System.Windows.Forms.Button
    Friend WithEvents ButtonCancel As System.Windows.Forms.Button
    Friend WithEvents GroupBoxAlphabet As System.Windows.Forms.GroupBox
    Friend WithEvents CheckedListBoxAlphabet As System.Windows.Forms.CheckedListBox
    Friend WithEvents ButtonAll As System.Windows.Forms.Button
    Friend WithEvents ButtonNone As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents LabelUpdate As System.Windows.Forms.Label
    Friend WithEvents CheckBoxUpdateKwadrantInfo As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents LabelProgressMessage As System.Windows.Forms.Label
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents ProgressBar2 As System.Windows.Forms.ProgressBar
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.ButtonOK = New System.Windows.Forms.Button
        Me.ButtonCancel = New System.Windows.Forms.Button
        Me.GroupBoxAlphabet = New System.Windows.Forms.GroupBox
        Me.ButtonNone = New System.Windows.Forms.Button
        Me.ButtonAll = New System.Windows.Forms.Button
        Me.CheckedListBoxAlphabet = New System.Windows.Forms.CheckedListBox
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.LabelUpdate = New System.Windows.Forms.Label
        Me.CheckBoxUpdateKwadrantInfo = New System.Windows.Forms.CheckBox
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.ProgressBar2 = New System.Windows.Forms.ProgressBar
        Me.LabelProgressMessage = New System.Windows.Forms.Label
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar
        Me.GroupBoxAlphabet.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'ButtonOK
        '
        Me.ButtonOK.Location = New System.Drawing.Point(120, 296)
        Me.ButtonOK.Name = "ButtonOK"
        Me.ButtonOK.Size = New System.Drawing.Size(75, 23)
        Me.ButtonOK.TabIndex = 1
        Me.ButtonOK.Text = "OK"
        '
        'ButtonCancel
        '
        Me.ButtonCancel.Location = New System.Drawing.Point(208, 296)
        Me.ButtonCancel.Name = "ButtonCancel"
        Me.ButtonCancel.Size = New System.Drawing.Size(75, 23)
        Me.ButtonCancel.TabIndex = 2
        Me.ButtonCancel.Text = "Annuleren"
        '
        'GroupBoxAlphabet
        '
        Me.GroupBoxAlphabet.Controls.Add(Me.ButtonNone)
        Me.GroupBoxAlphabet.Controls.Add(Me.ButtonAll)
        Me.GroupBoxAlphabet.Controls.Add(Me.CheckedListBoxAlphabet)
        Me.GroupBoxAlphabet.Location = New System.Drawing.Point(8, 0)
        Me.GroupBoxAlphabet.Name = "GroupBoxAlphabet"
        Me.GroupBoxAlphabet.Size = New System.Drawing.Size(280, 128)
        Me.GroupBoxAlphabet.TabIndex = 3
        Me.GroupBoxAlphabet.TabStop = False
        Me.GroupBoxAlphabet.Text = "Alfabet"
        '
        'ButtonNone
        '
        Me.ButtonNone.Location = New System.Drawing.Point(200, 48)
        Me.ButtonNone.Name = "ButtonNone"
        Me.ButtonNone.Size = New System.Drawing.Size(75, 23)
        Me.ButtonNone.TabIndex = 2
        Me.ButtonNone.Text = "Niets"
        '
        'ButtonAll
        '
        Me.ButtonAll.Location = New System.Drawing.Point(200, 16)
        Me.ButtonAll.Name = "ButtonAll"
        Me.ButtonAll.Size = New System.Drawing.Size(75, 23)
        Me.ButtonAll.TabIndex = 1
        Me.ButtonAll.Text = "Alles"
        '
        'CheckedListBoxAlphabet
        '
        Me.CheckedListBoxAlphabet.CheckOnClick = True
        Me.CheckedListBoxAlphabet.ColumnWidth = 40
        Me.CheckedListBoxAlphabet.Dock = System.Windows.Forms.DockStyle.Left
        Me.CheckedListBoxAlphabet.Items.AddRange(New Object() {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"})
        Me.CheckedListBoxAlphabet.Location = New System.Drawing.Point(3, 16)
        Me.CheckedListBoxAlphabet.MultiColumn = True
        Me.CheckedListBoxAlphabet.Name = "CheckedListBoxAlphabet"
        Me.CheckedListBoxAlphabet.Size = New System.Drawing.Size(184, 109)
        Me.CheckedListBoxAlphabet.Sorted = True
        Me.CheckedListBoxAlphabet.TabIndex = 0
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.LabelUpdate)
        Me.GroupBox1.Controls.Add(Me.CheckBoxUpdateKwadrantInfo)
        Me.GroupBox1.Location = New System.Drawing.Point(8, 128)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(280, 88)
        Me.GroupBox1.TabIndex = 7
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Kwadranten"
        '
        'LabelUpdate
        '
        Me.LabelUpdate.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelUpdate.Location = New System.Drawing.Point(8, 40)
        Me.LabelUpdate.Name = "LabelUpdate"
        Me.LabelUpdate.Size = New System.Drawing.Size(264, 40)
        Me.LabelUpdate.TabIndex = 6
        Me.LabelUpdate.Text = "Label Update"
        '
        'CheckBoxUpdateKwadrantInfo
        '
        Me.CheckBoxUpdateKwadrantInfo.Location = New System.Drawing.Point(8, 16)
        Me.CheckBoxUpdateKwadrantInfo.Name = "CheckBoxUpdateKwadrantInfo"
        Me.CheckBoxUpdateKwadrantInfo.Size = New System.Drawing.Size(264, 24)
        Me.CheckBoxUpdateKwadrantInfo.TabIndex = 5
        Me.CheckBoxUpdateKwadrantInfo.Text = "Vernieuw kwadrant referenties"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.ProgressBar2)
        Me.GroupBox2.Controls.Add(Me.LabelProgressMessage)
        Me.GroupBox2.Controls.Add(Me.ProgressBar1)
        Me.GroupBox2.Location = New System.Drawing.Point(8, 216)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(280, 72)
        Me.GroupBox2.TabIndex = 8
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
        'FormIndexStraten
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(292, 328)
        Me.ControlBox = False
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.GroupBoxAlphabet)
        Me.Controls.Add(Me.ButtonCancel)
        Me.Controls.Add(Me.ButtonOK)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.Name = "FormIndexStraten"
        Me.Text = "Straten Index"
        Me.GroupBoxAlphabet.ResumeLayout(False)
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
        InitializeForm()

    End Sub
#End Region

#Region " Initialization procedures "

    'Initial state of the form when loading.
    Private Sub InitializeForm()

        'Activate every letter of the alphabet.
        InitializeCheckedListBox(Me.CheckedListBoxAlphabet, True)

        'Set the labeltext for the checkbox.
        Me.LabelUpdate.Text = _
            "Activeer deze optie enkel indien de stratenlaag of rasterlaag is gewijzigd. " & _
            "Het aanmaken van de index zal meer tijd vergen indien deze optie actief staat."

        'Computing lookup info is mandatory, if lookup table is empty.
        Dim pLookupTable As ITable = GetTable(c_lookupTableName, m_document.FocusMap)
        If pLookupTable Is Nothing Then
            'RW:2008 TODO Create the table

            Throw New TableNotFoundException(c_lookupTableName)
        Else
            If pLookupTable.RowCount(Nothing) = 0 Then
                With Me.CheckBoxUpdateKwadrantInfo
                    .Checked = True
                    .Enabled = False
                End With
            End If
        End If

        'Set focus to the OK button.
        Me.ButtonOK.Focus()

    End Sub

    'Modify the checked state of all items in a CheckedListBox control.
    Private Sub InitializeCheckedListBox( _
        ByVal control As CheckedListBox, _
        ByVal checked As Boolean)

        Try
            Dim i As Integer 'loop index
            For i = 0 To control.Items.Count - 1
                control.SetItemChecked(i, checked)
            Next
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

#End Region

#Region " Form controls events "

    'Activate all letters of the alphabet.
    Private Sub ButtonAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonAll.Click
        InitializeCheckedListBox(Me.CheckedListBoxAlphabet, True)
    End Sub

    'Deactivate all letters of the alphabet.
    Private Sub ButtonNone_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonNone.Click
        InitializeCheckedListBox(Me.CheckedListBoxAlphabet, False)
    End Sub

    'Close the form.
    Private Sub ButtonCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonCancel.Click
        Me.Close()
    End Sub

    'Create the index and close the form after successful result.
    Private Sub ButtonOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonOK.Click

        Dim pStandaloneTable As IStandaloneTable = Nothing
        Dim pTableFields As ITableFields = Nothing
        Dim pLookupRS As ADODB.Recordset = Nothing

        Try

            'Disable button controls.
            Me.ButtonOK.Enabled = False
            Me.ButtonNone.Enabled = False
            Me.ButtonCancel.Enabled = False
            Me.ButtonAll.Enabled = False

            'Read available field lengths from the street lookup table.
            pStandaloneTable = New StandaloneTable
            pStandaloneTable.Table = GetTable(c_lookupTableName, m_document.FocusMap)
            pTableFields = CType(pStandaloneTable, ITableFields)
            With pTableFields
                If .FindField(c_refListFieldName) < 0 Then Throw New AttributeNotFoundException(c_lookupTableName, c_refListFieldName)
                m_refListMaxLength = .Field(.FindField(c_refListFieldName)).Length
                If .FindField(c_streetFieldName) < 0 Then Throw New AttributeNotFoundException(c_lookupTableName, c_streetFieldName)
                m_streetMaxLength = .Field(.FindField(c_streetFieldName)).Length
                If .FindField(c_sortFieldName) < 0 Then Throw New AttributeNotFoundException(c_lookupTableName, c_sortFieldName)
                m_sortMaxLength = .Field(.FindField(c_sortFieldName)).Length
            End With

            'Read hydrant <<aanduiding>> length from the hydrant feature class.
            Dim pFeatureLayer As IFeatureLayer
            pFeatureLayer = GetFeatureLayer(m_document.FocusMap, GetLayerName("Hydrant"))
            If pFeatureLayer Is Nothing Then Throw New LayerNotFoundException(GetLayerName("Hydrant"))
            With pFeatureLayer.FeatureClass.Fields()
                If .FindField(GetAttributeName("Hydrant", "Aanduiding")) < 0 Then Throw New AttributeNotFoundException(GetLayerName("Hydrant"), GetAttributeName("Hydrant", "Aanduiding"))
                m_hydrantMaxLength = .Field(.FindField(GetAttributeName("Hydrant", "Aanduiding"))).Length
            End With

            'Maak een lege ADODB recordset met 2 velden van gepaste grootte.
            pLookupRS = CreateLookupRecordset()

            'Update and store kwadrant references of the streets.
            If Me.CheckBoxUpdateKwadrantInfo.Checked Then
                'Vul de recordset met straatnamen en kwadranten.
                UpdateLookupInfo(pLookupRS, AddressOf OnShowProgress1, AddressOf OnSetMaxProgress1)
                'Bewaar de recordset voor een volgende keer.
                StoreRecordset(pLookupRS, GetTable(c_lookupTableName, m_document.FocusMap))
            End If

            'Read the kwadrant references of the streets.
            If Not Me.CheckBoxUpdateKwadrantInfo.Checked Then
                'Laad de recordset van een vorige keer.
                LoadRecordset(pLookupRS, AddressOf OnShowProgress1, AddressOf OnSetMaxProgress1)
            End If

            'If c_DebugStatus Then MsgBox(pLookupRS.RecordCount & " records in lookup recordset", , "Straten Index")

            'Compose index data and export to txt file.
            ExportRecordset(pLookupRS, AddressOf OnShowProgress2, AddressOf OnSetMaxProgress2)

            'Progress monitor info
            OnShowProgress2(-2, "Export succesvol beëindigd")

            'Enable/disable button controls.
            Me.ButtonOK.Enabled = False
            Me.ButtonNone.Enabled = False
            Me.ButtonCancel.Enabled = True
            Me.ButtonAll.Enabled = False

        Catch ex As Exception

            'Progress monitor info
            OnShowProgress2(-2, "Export niet succesvol beëindigd")

            'Enable/disable button controls.
            Me.ButtonOK.Enabled = False
            Me.ButtonNone.Enabled = False
            Me.ButtonCancel.Enabled = True
            Me.ButtonAll.Enabled = False

            'Pass exception to a higher level.
            'rw 23/07/2008 It was changed to directly throw the existing message in order to make the
            'Error Handeler in the upper event call to treat the initiated error message
            'New ApplicationException("Er is een fout opgetreden.", ex)
            Throw ex
        Finally

            'Try cleaning up the objects.
            If Not pStandaloneTable Is Nothing Then Marshal.ReleaseComObject(pStandaloneTable)
            If Not pTableFields Is Nothing Then Marshal.ReleaseComObject(pTableFields)
            If Not pLookupRS Is Nothing Then
                pLookupRS.Close()
                Marshal.ReleaseComObject(pLookupRS)
            End If

        End Try

        Dim oWord As Word.ApplicationClass = Nothing
        Dim dotPath As String
        Dim iconPath As String

        'Open Word template and run macro.
        Try
            oWord = New Word.ApplicationClass
            dotPath = oWord.Application.Options.DefaultFilePath(Word.WdDefaultFilePath.wdWorkgroupTemplatesPath)
            iconPath = dotPath.Substring(0, dotPath.Length - c_FileDir_Output.Length) & c_FileDir_Icons & "\"
            dotPath &= "\" & c_FileName_WordTemplateIndexStraten
            'oWord.Documents.Open(CType(dotPath, System.Object)) 'to open as a regular document
            oWord.Documents.Add(CType(dotPath, System.Object)) 'to make a new document based on template
            oWord.Run("Statenindex.setIconFolder", CType(iconPath, Object))
            oWord.Run(c_MacroName_IndexStraten)
            oWord.Visible = True

        Catch ex As Exception
            Throw ex

        Finally
            If Not oWord Is Nothing Then oWord = Nothing

        End Try

        'Close form if export finished successfully.
        Me.Close()

    End Sub

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
    '''     Return a recordset to hold links between streets to kwadrants.
    ''' </summary>
    ''' <returns>Empty recordset with 2 columns</returns>
    ''' <remarks>
    '''     Uses some public variables and constants.
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    '''  	[Kristof Vydt]	13/07/2006	Updated the names of some variables/constants.
    ''' 	[Kristof Vydt]	14/03/2007	Add sort column.
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Function CreateLookupRecordset() As ADODB.Recordset

        Try

            'Define a new ADODB recordset.
            Dim rs As New ADODB.Recordset
            rs.Fields.Append(c_refListFieldName, DataTypeEnum.adChar, m_refListMaxLength)
            rs.Fields.Append(c_streetFieldName, DataTypeEnum.adChar, m_streetMaxLength)
            rs.Fields.Append(c_sortFieldName, DataTypeEnum.adChar, m_sortMaxLength)

            'Open the recordset and return it.
            rs.Open()
            CreateLookupRecordset = rs

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Store the records of a recordset in a persistent way.
    '''     For this purpose we use a 'flat' table in a personal geodatabase.
    ''' </summary>
    ''' <param name="Recordset">
    '''     ADODB Recordset to store.
    ''' </param>
    ''' <param name="Table">
    '''     An 'flat' table from an ArcGIS personal geodatabase.
    ''' </param>
    ''' <remarks>
    '''     The table is cleared before saving the recordset in it.
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    '''  	[Kristof Vydt]	13/07/2006	Updated the names of some variables/constants.
    ''' 	[Kristof Vydt]	14/03/2007	Export new sorting column.
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub StoreRecordset( _
            ByVal Recordset As ADODB.Recordset, _
            ByRef Table As ITable)

        Try
            Dim InsertCursor As ICursor
            Dim RowBuffer As IRowBuffer
            Dim RowOID As System.Object

            'Clear the table.
            Table.DeleteSearchedRows(Nothing)

            'Add each record to the table.
            InsertCursor = Table.Insert(True)
            Recordset.MoveFirst()
            While Not Recordset.EOF
                RowBuffer = Table.CreateRowBuffer
                RowBuffer.Value(1) = Trim(CStr(Recordset(c_streetFieldName).Value))
                RowBuffer.Value(2) = Trim(CStr(Recordset(c_refListFieldName).Value))
                RowBuffer.Value(3) = Trim(CStr(Recordset(c_sortFieldName).Value))
                RowOID = InsertCursor.InsertRow(RowBuffer)
                Recordset.MoveNext()
            End While

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Fill an existing recordset with new records.
    '''     For this purpose we use a 'flat' table in a personal geodatabase.
    '''     The original content of the recordset is cleared.
    ''' </summary>
    ''' <param name="pRecordset">
    '''     [out] The recordset to fill.
    ''' </param>
    ''' <param name="ProgressDelegate">
    '''     Delegate for progress monitoring.
    ''' </param>
    ''' <param name="MaxProgressDelegate">
    '''     Delegate for setting maximum progress value.
    ''' </param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    '''  	[Kristof Vydt]	13/07/2006	Updated the names of some variables/constants.
    ''' 	[Kristof Vydt]	14/03/2007	Load new sort column.
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub LoadRecordset( _
            ByRef pRecordset As ADODB.Recordset, _
            ByVal ProgressDelegate As ShowProgress, _
            ByVal MaxProgressDelegate As SetMaxProgress)

        Try
            Dim pTable As ITable
            Dim pCursor As ICursor
            Dim pRow As IRow
            Dim FieldIndex_streetname As Integer
            Dim FieldIndex_kwadrants As Integer
            Dim FieldIndex_sortname As Integer
            Dim RowIndex As Integer

            'Clear passed recordset before filling it up with new records.
            ' pRecordset.Delete() 
            ' --> seems not to work as easily, and since it's not really necessary, just skip this.

            'Get a pointer to the lookup table in the geodatabase.
            pTable = GetTable(c_lookupTableName, m_document.FocusMap)
            MaxProgressDelegate(pTable.RowCount(Nothing))

            'Loop through the table records.
            pCursor = pTable.Search(Nothing, Nothing)
            pRow = pCursor.NextRow
            If Not pRow Is Nothing Then

                'Get the field indices to read from.
                FieldIndex_kwadrants = pCursor.FindField(c_refListFieldName)
                FieldIndex_streetname = pCursor.FindField(c_streetFieldName)
                FieldIndex_sortname = pCursor.FindField(c_sortFieldName)
                If FieldIndex_kwadrants < 0 Then Throw New AttributeNotFoundException(c_lookupTableName, c_refListFieldName)
                If FieldIndex_streetname < 0 Then Throw New AttributeNotFoundException(c_lookupTableName, c_streetFieldName)
                If FieldIndex_sortname < 0 Then Throw New AttributeNotFoundException(c_lookupTableName, c_sortFieldName)

                While Not pRow Is Nothing

                    'Progress monitoring.
                    RowIndex += 1
                    ProgressDelegate(RowIndex, "Kwadranten info inlezen ....")

                    'Add record by record to the recordset.
                    pRecordset.AddNew()
                    pRecordset(c_refListFieldName).Value = pRow.Value(FieldIndex_kwadrants)
                    pRecordset(c_streetFieldName).Value = pRow.Value(FieldIndex_streetname)
                    pRecordset(c_sortFieldName).Value = pRow.Value(FieldIndex_sortname)

                    'Loop with next record.
                    pRow = pCursor.NextRow
                End While

                'Progress monitoring.
                ProgressDelegate(RowIndex, CStr(RowIndex) & " referenties ingelezen.")

            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Deze functie update een ADODB recordset bestaande uit 2 velden:
    '''     straatnaam en een lijst van enkelvoudige kwadrantreferentie.
    ''' </summary>
    ''' <param name="rs">
    '''     The recordset that should be filled.
    ''' </param>
    ''' <param name="ProgressDel">
    '''     Progress monitoring delegate.
    ''' </param>
    ''' <param name="MaxProgressDel">
    '''     Progress monitoring deligate.
    ''' </param>
    ''' <remarks>
    '''     Met "enkelvoudige kwadrantreferentie" wordt de combinatie van één bladzijdegetal en één kwadrantletter bedoeld.
    '''     Deze procedure behandelt steeds alle letters van het alfabet, los van de keuze van de eindgebruiker in FormStratenIndex.
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    '''  	[Kristof Vydt]	13/07/2006	Updated the names of some variables/constants.
    ''' 	[Kristof Vydt]	02/08/2006	Replace some global constant by a function that reads the value from config file.
    ''' 	[Kristof Vydt]	19/04/2007	Release SpatialFilter COM object to avoid unexpected, unspecified errors.
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub UpdateLookupInfo( _
            ByRef rs As ADODB.Recordset, _
            ByVal ProgressDel As ShowProgress, _
            ByVal MaxProgressDel As SetMaxProgress)

        Dim pFCursor_park As IFeatureCursor = Nothing
        Dim pFCursor_raster As IFeatureCursor = Nothing
        Dim pFCursor_straten As IFeatureCursor = Nothing
        Dim pFCursor_water As IFeatureCursor = Nothing

        Try
            Dim arrayKwadranten As IGeometry() = Nothing
            Dim fieldIndex_page As Integer
            Dim fieldIndex_name As Integer
            Dim fieldIndex_name1 As Integer
            Dim fieldIndex_name2 As Integer
            Dim i As Integer
            Dim kwadrantReference As String = ""
            Dim progress As Integer
            'Dim sectorCode As String
            Dim whereClause_straten As String = ""

            Dim pFClass As IFeatureClass
            Dim pFCursor As IFeatureCursor
            Dim pFeature As IFeature
            Dim pFeature_raster As IFeature
            Dim pFLayer_park As IFeatureLayer
            Dim pFLayer_raster As IFeatureLayer
            Dim pFLayer_sector As IFeatureLayer
            Dim pFLayer_straten As IFeatureLayer
            Dim pFLayer_water As IFeatureLayer
            Dim pGeometry_sector As IGeometry
            Dim pQueryFilter As IQueryFilter
            Dim pSpatialFilter As ISpatialFilter

            'Progress monitor
            progress = -1 'empty progress bar
            ProgressDel(progress, "Reference update begint ...")

            'Where-clause o.b.v. geconfigureerde postcodes.
            Dim postcodes As Collection = Config.Postcodes
            For Each postcode As String In postcodes
                If Len(whereClause_straten) > 0 Then whereClause_straten &= " OR "
                whereClause_straten &= "(" & GetAttributeName("Straatassen", "Postcode") & " = '" & postcode & "')"
            Next
            'whereClause_straten = _
            '    "(" & c_AttributeName_straatassen_straatnaam & "<>'ONBEKEND') AND " & _
            '    "(" & c_AttributeName_straatassen_straatnaam & "<>'PAD'     ) AND " & _
            '    "(" & c_AttributeName_straatassen_straatnaam & "<>'SNELWEG' ) AND " & _
            '    "(" & whereClause_straten & ")"
            'MsgBox(whereClause_straten, , "WhereClause voor Straten")

            'De sectoren laag opzoeken.
            pFLayer_sector = GetFeatureLayer(m_document.FocusMap, GetLayerName("Sector"))
            If pFLayer_sector Is Nothing Then Throw New LayerNotFoundException(GetLayerName("Sector"))

            'De raster laag opzoeken.
            pFLayer_raster = GetFeatureLayer(m_document.FocusMap, GetLayerName("Raster"))
            If pFLayer_raster Is Nothing Then Throw New LayerNotFoundException(GetLayerName("Raster"))
            'Bepaal veldindexen van bruikbare attributen.
            fieldIndex_page = pFLayer_raster.FeatureClass.FindField("BLZ_" & Config.SectorCode)
            If fieldIndex_page = 0 Then Throw New AttributeNotFoundException(GetLayerName("Raster"), "BLZ_" & Config.SectorCode)

            'De straten laag opzoeken.
            pFLayer_straten = GetFeatureLayer(m_document.FocusMap, GetLayerName("Straatassen"))
            If pFLayer_straten Is Nothing Then Throw New LayerNotFoundException(GetLayerName("Straatassen"))
            'Veldindex van het naam-attributen.
            fieldIndex_name = pFLayer_straten.FeatureClass.FindField(GetAttributeName("Straatassen", "Straatnaam"))
            If fieldIndex_name = -1 Then Throw New AttributeNotFoundException(GetLayerName("Straatassen"), GetAttributeName("Straatassen", "Straatnaam"))

            'De park laag opzoeken.
            pFLayer_park = GetFeatureLayer(m_document.FocusMap, GetLayerName("Park"))
            If pFLayer_park Is Nothing Then Throw New LayerNotFoundException(GetLayerName("Park"))
            'Veldindex van het naam-attributen.
            fieldIndex_name1 = pFLayer_park.FeatureClass.FindField(GetAttributeName("Park", "Naam"))
            If fieldIndex_name1 = -1 Then Throw New AttributeNotFoundException(GetLayerName("Park"), GetAttributeName("Park", "Naam"))

            'De water laag opzoeken.
            pFLayer_water = GetFeatureLayer(m_document.FocusMap, GetLayerName("Water"))
            If pFLayer_water Is Nothing Then Throw New LayerNotFoundException(GetLayerName("Water"))
            'Veldindex van het naam-attributen.
            fieldIndex_name2 = pFLayer_water.FeatureClass.FindField(GetAttributeName("Water", "Naam"))
            If fieldIndex_name2 = -1 Then Throw New AttributeNotFoundException(GetLayerName("Water"), GetAttributeName("Water", "Naam"))

            'Op basis van de sectorcode kan de sectorpolygoon geselecteerd worden uit de laag 'Sector'.
            pQueryFilter = New QueryFilter
            pQueryFilter.WhereClause = GetAttributeName("Sector", "Afkorting") & " = " & CStrSql(Config.SectorCode)
            pFClass = CType(pFLayer_sector, IFeatureLayer2).FeatureClass()
            pFCursor = pFClass.Search(pQueryFilter, True)
            pFeature = pFCursor.NextFeature()
            'If pFeature Is Nothing Then Exit Sub
            pGeometry_sector = pFeature.ShapeCopy
            Dim pGeometry_sectorTop As ITopologicalOperator2
            pGeometry_sectorTop = CType(pGeometry_sector, ITopologicalOperator2)
            pGeometry_sectorTop.IsKnownSimple_2 = False
            pGeometry_sectorTop.Simplify()

            'Maak een cursor van alle rasterfeatures waarvoor het bladzijdeattribuut is ingevuld.
            pQueryFilter = New QueryFilter
            pQueryFilter.WhereClause = "BLZ_" & Config.SectorCode & " <> ''"
            pFCursor_raster = pFLayer_raster.Search(pQueryFilter, Nothing)

            'Determine featurecount in featurecursor and use this to set maximum progress value
            'TODO: not a nicer way to get featurecount of cursor ?
            'http://forums.esri.com/Thread.asp?c=93&f=982&t=57065&mc=2#143079
            'http://forums.esri.com/Thread.asp?c=93&f=993&t=77173&mc=1#msgid205805
            pFeature_raster = pFCursor_raster.NextFeature
            i = 0
            While Not pFeature_raster Is Nothing
                i += 5
                pFeature_raster = pFCursor_raster.NextFeature
            End While
            MaxProgressDel(i)

            'Loop door de cursor en voor elke rasterfeature ...
            pFCursor_raster = pFLayer_raster.Search(pQueryFilter, Nothing)
            pFeature_raster = pFCursor_raster.NextFeature
            While Not pFeature_raster Is Nothing

                'Progress monitor
                progress += 1
                ProgressDel(progress, "Reference update: " & CStr(pFeature_raster.Value(fieldIndex_page)))
                Debug.Write("Reference update:")

                'Dim fieldIndex As Integer = pFeature_raster.Fields.FindField("BLZ_" & sectorCode)
                'MsgBox(pFeature_raster.Shape.GeometryType() & vbNewLine & CStr(pFeature_raster.Value(fieldIndex)))

                'Splits de envelope van de rasterpolygoon op in vier gelijke kwadrantpolygonen
                'en bepaal de doorsnede met de sectorpolygoon. Dit resulteert in 4 polygonen.
                SplitFeatureEnvelopeIntoKwadrants(pFeature_raster, pGeometry_sector, arrayKwadranten)

                'Voor elke kwadrantpolygoon ...
                For i = 0 To arrayKwadranten.Length - 1

                    'Bepaal de referentie van dit kwadrant.
                    Select Case i
                        Case 0 'Kwadrant A
                            kwadrantReference = CStr(pFeature_raster.Value(fieldIndex_page)) & "A" 'keep 0 at the beginning
                            'kwadrantReference = CStr(CInt(pFeature_raster.Value(fieldIndex_page))) & "A" 'remove 0 at the beginning
                        Case 1 'Kwadrant B
                            kwadrantReference = CStr(pFeature_raster.Value(fieldIndex_page)) & "B" 'keep 0 at the beginning
                            'kwadrantReference = CStr(CInt(pFeature_raster.Value(fieldIndex_page))) & "B" 'remove 0 at the beginning
                        Case 2 'Kwadrant C
                            kwadrantReference = CStr(pFeature_raster.Value(fieldIndex_page)) & "C" 'keep 0 at the beginning
                            'kwadrantReference = CStr(CInt(pFeature_raster.Value(fieldIndex_page))) & "C" 'remove 0 at the beginning
                        Case 3 'Kwadrant D
                            kwadrantReference = CStr(pFeature_raster.Value(fieldIndex_page)) & "D" 'keep 0 at the beginning
                            'kwadrantReference = CStr(CInt(pFeature_raster.Value(fieldIndex_page))) & "D" 'remove 0 at the beginning
                    End Select

                    'Progress monitor
                    progress += 1
                    ProgressDel(progress, "Reference update: " & kwadrantReference)
                    Debug.Write(" " & kwadrantReference & " ")

                    'Spatial query op stratenlaag o.b.v. doorsnedepolygoon. 
                    'Bijkomend wordt een attribuutfilter gebruikt om straten 
                    'met een afwijkende postcode uit te sluiten. 
                    'Dit resulteert in een cursor van overlappende straten.
                    pSpatialFilter = New SpatialFilter
                    pSpatialFilter.Geometry = arrayKwadranten(i)
                    pSpatialFilter.SpatialRel = esriSpatialRelEnum.esriSpatialRelIntersects
                    pSpatialFilter.WhereClause = whereClause_straten
                    Debug.Write("(")
                    'pFCursor_straten = pFLayer_straten.Search(pSpatialFilter, Nothing)
                    pFCursor_straten = pFLayer_straten.Search(pSpatialFilter, True)

                    'Voeg de volledige cursor toe aan de Lookup Recordset.
                    If Not pFCursor_straten Is Nothing Then
                        Debug.Write("S")
                        AddCursorToLookupRecordset(pFCursor_straten, rs, kwadrantReference, fieldIndex_name)
                        Marshal.ReleaseComObject(pFCursor_straten)
                        pFCursor_straten = Nothing
                    End If
                    Debug.Write(")")

                    'Analoog voor de park laag.
                    pSpatialFilter.WhereClause = GetAttributeName("Park", "Naam") & "<>''"
                    Debug.Write("(")
                    'pFCursor_park = pFLayer_park.Search(pSpatialFilter, Nothing)
                    pFCursor_park = pFLayer_park.Search(pSpatialFilter, True)
                    If Not pFCursor_park Is Nothing Then
                        Debug.Write("P")
                        AddCursorToLookupRecordset(pFCursor_park, rs, kwadrantReference, fieldIndex_name1)
                        Marshal.ReleaseComObject(pFCursor_park)
                        pFCursor_park = Nothing
                    End If
                    Debug.Write(")")

                    'Analoog voor de water laag.
                    pSpatialFilter.WhereClause = GetAttributeName("Water", "Naam") & "<>''"
                    Debug.Write("(")
                    'pFCursor_water = pFLayer_water.Search(pSpatialFilter, Nothing)
                    pFCursor_water = pFLayer_water.Search(pSpatialFilter, True)
                    If Not pFCursor_water Is Nothing Then
                        Debug.Write("W")
                        AddCursorToLookupRecordset(pFCursor_water, rs, kwadrantReference, fieldIndex_name2)
                        Marshal.ReleaseComObject(pFCursor_water)
                        pFCursor_water = Nothing
                    End If
                    Debug.Write(")")

                    ' Release SpatialFilter COM object.
                    If Not pSpatialFilter Is Nothing Then
                        Marshal.ReleaseComObject(pSpatialFilter)
                        pSpatialFilter = Nothing
                    End If

                    ' Force garbage collection
                    GC.Collect()

                Next '... volgende kwadrant.

                Debug.WriteLine(" ... finished.")

                '... volgende rasterfeature uit cursor.
                pFeature_raster = pFCursor_raster.NextFeature
            End While

            'Progress monitor
            progress = -2 'full progress bar
            ProgressDel(progress, "Reference update beëindigd")

        Catch ex As Exception
            Throw ex
            'Throw New ApplicationException("Er is een onverwachte fout opgetreden.", ex)

        Finally
            ' Force garbage collection
            If Not (pFCursor_raster Is Nothing) Then Marshal.ReleaseComObject(pFCursor_raster)
            If Not (pFCursor_straten Is Nothing) Then Marshal.ReleaseComObject(pFCursor_straten)
            If Not (pFCursor_park Is Nothing) Then Marshal.ReleaseComObject(pFCursor_park)
            If Not (pFCursor_water Is Nothing) Then Marshal.ReleaseComObject(pFCursor_water)
            GC.Collect()
        End Try

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Add each feature of the the cursor to the recordset.
    ''' </summary>
    ''' <param name="pCursor">
    '''     [in] The feature cursor that must be added to the recordset.
    ''' </param>
    ''' <param name="pRecordset">
    '''     [out] The recordset that has to be expanded.
    ''' </param>
    ''' <param name="kwadrantReference">
    '''     [in] The kwadrant reference that is used for the whole cursor.
    ''' </param>
    ''' <param name="fieldIndex1">
    '''     [in] The field index of the feature attribute of which the value 
    '''          must be copied to the lookup recordset.
    ''' </param>
    ''' <remarks>
    '''     The kwadrant references are sorted in this procedure.
    '''     Therefore, for example, use "01A" and not "1A".
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    ''' 	[Kristof Vydt]	13/07/2006	Updated the names of some variables/constants.
    ''' 	[Kristof Vydt]	14/03/2007	Fill new sorting column when adding new record.
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub AddCursorToLookupRecordset( _
            ByVal pCursor As IFeatureCursor, _
            ByRef pRecordset As ADODB.Recordset, _
            ByVal kwadrantReference As String, _
            ByVal fieldIndex1 As Integer)

        Try
            Dim pFeature As IFeature
            Dim fieldValue1 As String
            Dim searchCondition As String

            'Loop through the complete cursor.
            pFeature = pCursor.NextFeature
            'RW:2008 to find how many are found
            Dim nrOfFoundFeatures As Integer = 0

            While Not pFeature Is Nothing


                'Bewaar straatnaam en kwadrantreferentie in de recordset, 
                'door een nieuwe straatnaam toe te voegen, 
                'of door een nieuwe enkelvoudige kwadrantreferentie toe te voegen 
                'aan een reeds geregistreerde straatnaam. 
                fieldValue1 = CStr(pFeature.Value(fieldIndex1))
                If Trim(fieldValue1) = "" Then Exit While
                nrOfFoundFeatures = nrOfFoundFeatures + 1
                If Not pRecordset.EOF Then pRecordset.MoveFirst()
                searchCondition = c_streetFieldName & "=" & CStrSql(fieldValue1)
                pRecordset.Find(searchCondition)

                If pRecordset.EOF Then
                    'add new record to rs
                    pRecordset.AddNew()
                    pRecordset(c_refListFieldName).Value = kwadrantReference
                    pRecordset(c_streetFieldName).Value = fieldValue1
                    pRecordset(c_sortFieldName).Value = CharFilter(fieldValue1.ToUpper, c_AllowedSortChars)
                Else
                    'get all kwadrants already in rs
                    Dim listKwadrant As String
                    Dim arrayKwadrant As String()
                    listKwadrant = CStr(pRecordset(c_refListFieldName).Value)
                    arrayKwadrant = Split2(listKwadrant, c_ListSeparator, True)
                    If Array.IndexOf(arrayKwadrant, kwadrantReference) = -1 Then
                        'add ref only if kwadrant is not yet added in rs
                        ReDim Preserve arrayKwadrant(arrayKwadrant.Length)
                        arrayKwadrant.SetValue(kwadrantReference, arrayKwadrant.Length - 1)
                        Array.Sort(arrayKwadrant) 'sort references
                        listKwadrant = Concat(arrayKwadrant)
                        If Len(listKwadrant) > m_refListMaxLength Then
                            Throw New RecordsetFieldSizeNotSufficientException(c_refListFieldName)
                        Else
                            pRecordset(c_refListFieldName).Value = listKwadrant
                        End If
                    End If
                End If

                'Next feature from cursor. 
                pFeature = pCursor.NextFeature
            End While
            Debug.Write(nrOfFoundFeatures.ToString())
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Use the lookup recordset (streetnames linked to kwadrants),
    '''     combine it with other themes, and export it to a txt file.
    ''' </summary>
    ''' <param name="rs">
    '''     The lookup recordset which links streetnames to kwadrants.
    ''' </param>
    ''' <param name="ProgressDelegate">
    '''     Progress monitoring delegate.
    ''' </param>
    ''' <param name="MaxProgressDelegate">
    '''     Maximum progress delegate.
    ''' </param>
    ''' <remarks>
    '''     Use of Marshal.ReleaseComObject is required to tackle
    '''     System.Runtime.InteropServices.COMException "Exception from HRESULT: 0x80040213".
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    ''' 	[Kristof Vydt]	17/10/2005	Interpret aanduiding &lt;null&gt; as "??".
    '''                                 Interpret diameter &lt;null&gt; as 0.
    ''' 	[Kristof Vydt]	28/10/2005	Correction of attribute reference in query for Sleepboot.
    '''     [Kristof Vydt]  12/07/2006  Use CStrSql for string values in WhereClauses.
    '''   	[Kristof Vydt]	13/07/2006	Updated the names of some variables/constants.
    ''' 	[Kristof Vydt]	02/08/2006	Replace some global constant by a function that reads the value from config file.
    '''     [Kristof Vydt]  10/08/2006  Add validity check before using danger layers.
    '''     [Kristof Vydt]  11/08/2006  Export last letters if no more streets in the recordset.
    '''     [Kristof Vydt]  14/03/2007  Always export the "g"-code as letter header for the Word macro.
    '''                                 Use recordset filter on the new sort column to obtain subset for 1 letter.
    '''                                 Isolate the different export blocks in separate procedures.
    '''     [Kristof Vydt]  15/03/2007  Continue
    '''     [Kristof Vydt]  28/03/2007  Add dummy export line at the beginning of the output file.
    '''     [Elton Manoku]  24/07/2008  Add prive ondergrond and prive bovengrond in the outputfile.
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub ExportRecordset( _
            ByVal rs As ADODB.Recordset, _
            ByRef ProgressDelegate As ShowProgress, _
            ByRef MaxProgressDelegate As SetMaxProgress)

        Dim StreamWriter As StreamWriter 'for writing to the export file
        'Dim LettersEnum As System.Collections.IEnumerator 'enumeration of all selected letters of the alphabet
        Dim LetterCounter As Integer 'number of processed letters - required for updating the progress bar
        'Dim ThisLetter As String 'one letter from the enumeration
        'Dim ThisStreet As String 'one streetname from the recordset
        'Dim ThisStreetFirstLetter As String 'the first regular letter of the street name
        'Dim ThisKwadrant As String 'one singel raster-kwadrant reference (pagenumber & 1 kwadrant char)
        'Dim ArrayKwadrant As String() 'array of raster-kwadrant references
        'Dim FieldIndex As Integer 'index of feature attribute

        'Dim pFCursor As IFeatureCursor
        'Dim pFeature As IFeature
        Dim pFLayer_hoogspanning As IFeatureLayer = Nothing
        Dim pFLayer_hydrant As IFeatureLayer = Nothing
        Dim pFLayer_instorting As IFeatureLayer = Nothing
        Dim pFLayer_sleepboot As IFeatureLayer = Nothing
        Dim pFLayer_sleutelgebouw As IFeatureLayer = Nothing
        Dim pFLayer_stralingsbron As IFeatureLayer = Nothing
        'Dim pQueryFilter As IQueryFilter

        Try 'first try-block of this procedure

            'Progress monitor information.
            MaxProgressDelegate(Me.CheckedListBoxAlphabet.CheckedItems.Count + 1)
            ProgressDelegate(-1, "Exportbestand wordt voorbereid...")

            'Maak een nieuw txt exportbestand klaar om naar weg te schrijven.
            If File.Exists(c_FilePath_IndexStraten) Then File.Delete(c_FilePath_IndexStraten)
            StreamWriter = New StreamWriter(c_FilePath_IndexStraten, True, System.Text.Encoding.Unicode)

            ' Add an empty line at the begin of file, to avoid problem
            ' with hidden character(s) as a result of unicode encoding.
            StreamWriter.WriteLine("")

        Catch ex As Exception
            Throw New RecreateExportFileException(c_FilePath_IndexStraten)
        End Try

        Try

            'Zoek feature layers op.
            pFLayer_hoogspanning = GetFeatureLayer(m_document.FocusMap, GetLayerName("Hoogspanning"))
            pFLayer_hydrant = GetFeatureLayer(m_document.FocusMap, GetLayerName("Hydrant"))
            pFLayer_instorting = GetFeatureLayer(m_document.FocusMap, GetLayerName("Instorting"))
            pFLayer_sleepboot = GetFeatureLayer(m_document.FocusMap, GetLayerName("Sleepboot"))
            pFLayer_sleutelgebouw = GetFeatureLayer(m_document.FocusMap, GetLayerName("Sleutelgebouw"))
            pFLayer_stralingsbron = GetFeatureLayer(m_document.FocusMap, GetLayerName("Stralingsbron"))

            '********* new begin

            ' Loop door alle geselecteerde letters van het alfabet.
            For Each checkedItem As Object In Me.CheckedListBoxAlphabet.CheckedItems

                '-- Export letterheader -- Code G --
                Dim currentLetter As Char = Convert.ToChar(checkedItem)
                ExportLetterHeader(StreamWriter, currentLetter)

                'Update progress monitor.
                LetterCounter += 1
                ProgressDelegate(LetterCounter, "Export letter " & currentLetter)

                ' Filter op beginletter en sorteer de stratenlijst.
                rs.Filter = c_sortFieldName & " LIKE '" & currentLetter & "*'"
                rs.Sort = c_streetFieldName

                ' Loop door de gefilterde stratenlijst.
                If Not (rs.BOF And rs.EOF) Then
                    rs.MoveFirst()
                    While Not rs.EOF

                        '-- Export straatnaam -------------- Code S --
                        Dim currentStreet As String = Trim(Convert.ToString(rs(c_streetFieldName).Value))
                        ExportStreetName(StreamWriter, currentStreet)

                        '-- Export kwadrant referenties ---- Code K --
                        Dim kwadrantsReferences As String = Trim(Convert.ToString(rs(c_refListFieldName).Value))
                        ExportKwadrants(StreamWriter, kwadrantsReferences)

                        '-- Export ondergrondse hydranten -- Code H --
                        If Not pFLayer_hydrant Is Nothing Then _
                            ExportHydrantenOndergronds(StreamWriter, pFLayer_hydrant, currentStreet)

                        '-- Export hoogspanning ------------ Code P --
                        If Not pFLayer_hoogspanning Is Nothing Then _
                            ExportGevaren(StreamWriter, pFLayer_hoogspanning, currentStreet)

                        '-- Export stralingsbronnen -------- Code B --
                        If Not pFLayer_stralingsbron Is Nothing Then _
                            ExportGevaren(StreamWriter, pFLayer_stralingsbron, currentStreet)

                        '-- Export instortingsgevaar ------- Code I --
                        If Not pFLayer_instorting Is Nothing Then _
                            ExportGevaren(StreamWriter, pFLayer_instorting, currentStreet)

                        '-- Export sleutelgebouwen --------- Code C --
                        If Not pFLayer_sleutelgebouw Is Nothing Then _
                            ExportGevaren(StreamWriter, pFLayer_sleutelgebouw, currentStreet)

                        '-- Export sleepboten -------------- Code W --
                        If Not pFLayer_sleepboot Is Nothing Then _
                            ExportGevaren(StreamWriter, pFLayer_sleepboot, currentStreet)

                        If Not pFLayer_hydrant Is Nothing Then
                            '-- Export bovengrondse hydranten -- Code D --
                            ExportHydrantenBovengronds(StreamWriter, pFLayer_hydrant, currentStreet)
                            'RW:07-08/2008
                            '-- Export prive ondergrondse hydranten -- Code F --
                            ExportHydrantenPrive(StreamWriter, pFLayer_hydrant, currentStreet, "3")
                            'RW:07-08/2008
                            '-- Export prive bovengrondse hydranten -- Code E --
                            ExportHydrantenPrive(StreamWriter, pFLayer_hydrant, currentStreet, "4")

                        End If

                        ' Volgende straat uit de lijst.
                        rs.MoveNext()
                    End While
                End If

                ' Volgende letter die werd aangevinkt.
            Next

            Exit Try
            '********* new end

            ''Sorteer lookup recordset: alfabetisch op straatnaam.
            'rs.Sort = c_streetFieldName
            'rs.MoveFirst()

            ''Voor elke geselecteerde letter ...
            'LettersEnum = Me.CheckedListBoxAlphabet.CheckedItems.GetEnumerator()
            'While LettersEnum.MoveNext()
            '    LetterCounter += 1

            '    'Een door de gebruiker geselecteerde letter.
            '    ThisLetter = Trim(CStr(LettersEnum.Current))

            '    'Eerste/volgende straatnaam uit de lookup recordset.
            '    If Not rs.EOF Then
            '        ThisStreet = Trim((CStr(rs(c_streetFieldName).Value)).ToUpper)
            '        ThisStreetFirstLetter = GetFirstLetter(ThisStreet)
            '    End If

            '    'Update progress monitor.
            '    ProgressDelegate(LetterCounter, "Export letter " & ThisLetter)

            '    'Voor de letterheading in het txt exportbestand ...
            '    'Loop door de stratenrecordset op zoek naar een straatnaam beginnend met die letter.
            '    ' (Gebruik hiervoor geen gelijkheids-operator, maar een groter/kleiner-dan-operator, 
            '    '  zodat de loop stopt als de letter niet wordt gevonden.)
            '    If Not rs.EOF Then
            '        While 0 < String.Compare(ThisLetter, ThisStreetFirstLetter, True)
            '            rs.MoveNext()
            '            If Not rs.EOF Then
            '                ThisStreet = Trim((CStr(rs(c_streetFieldName).Value)).ToUpper)
            '                ThisStreetFirstLetter = GetFirstLetter(ThisStreet)
            '            End If
            '        End While
            '    End If

            '    '-- G --
            '    'Is de letter niet gevonden, schrijf code "g" (met ontbrekende letter in uppercase) naar het export txt bestand. 
            '    'Herneem de loop met volgende letter uit de alfabet-subset.
            '    If rs.EOF Then
            '        StreamWriter.WriteLine("g," & ThisLetter)
            '    ElseIf 0 > String.Compare(ThisLetter, ThisStreetFirstLetter, True) Then
            '        StreamWriter.WriteLine("g," & ThisLetter)
            '    End If

            '    '-- S --
            '    'Is de letter wel gevonden, schrijf code "s" (met straatnaam) naar het export txt bestand.
            '    While 0 = String.Compare(ThisLetter, ThisStreetFirstLetter, True)
            '        StreamWriter.WriteLine("s," & ThisStreet)
            '        StreamWriter.Flush()

            '        '-- K --
            '        'Schrijf code "k" (met één raster-kwadrant-referentie = één rasterpaginanummer en één kwadrantletter)
            '        ArrayKwadrant = Split2(CStr(rs(c_refListFieldName).Value), c_ListSeparator, True)
            '        Array.Sort(ArrayKwadrant)
            '        For Each ThisKwadrant In ArrayKwadrant
            '            While ThisKwadrant.Substring(0, 1) = "0"
            '                ThisKwadrant = ThisKwadrant.Substring(1) 'Verwijder nullen aan het begin.
            '            End While
            '            StreamWriter.WriteLine("k," & ThisKwadrant)
            '        Next
            '        StreamWriter.Flush()

            '        '-- H --
            '        'Voor de lijst van actieve, ondergrondse hydranten per diameter in het txt exportbestand ...
            '        'Zoek alle actieve, ondergrondse hydranten op die aan deze straat zijn geconnecteerd.
            '        pQueryFilter = New QueryFilter
            '        pQueryFilter.WhereClause = "(" & GetAttributeName("Hydrant", "Straatnaam") & " = " & CStrSql(ThisStreet) & ")" _
            '                            & " AND (" & GetAttributeName("Hydrant", "Status") & " = '1')" _
            '                            & " AND (" & GetAttributeName("Hydrant", "HydrantType") & " = '1')"
            '        pFCursor = pFLayer_hydrant.Search(pQueryFilter, Nothing)
            '        pFeature = pFCursor.NextFeature
            '        If Not pFeature Is Nothing Then

            '            'Er zijn hydranten gevonden.
            '            Dim rsHydrant As ADODB.Recordset 'recordset for collecting diameters and labels of hydrants
            '            Dim FieldIndex2 As Integer 'field index for diameter attribute of hydrants

            '            'Open a new recordset for hydrants.
            '            rsHydrant = New ADODB.Recordset
            '            rsHydrant.Fields.Append("diameter", DataTypeEnum.adSmallInt)
            '            rsHydrant.Fields.Append("aanduiding", DataTypeEnum.adChar, m_hydrantMaxLength)
            '            rsHydrant.Open()

            '            'Determine the field indexes.
            '            FieldIndex = pFeature.Fields.FindField(GetAttributeName("Hydrant", "Aanduiding"))
            '            If FieldIndex = -1 Then Throw New AttributeNotFoundException(GetLayerName("Hydrant"), GetAttributeName("Hydrant", "Aanduiding"))
            '            FieldIndex2 = pFeature.Fields.FindField(GetAttributeName("Hydrant", "Diameter"))
            '            If FieldIndex2 = -1 Then Throw New AttributeNotFoundException(GetLayerName("Hydrant"), GetAttributeName("Hydrant", "Diameter"))

            '            'Fill recordset with hydrants.
            '            While Not pFeature Is Nothing
            '                rsHydrant.AddNew()
            '                If TypeOf pFeature.Value(FieldIndex) Is System.DBNull Then
            '                    rsHydrant("aanduiding").Value = CStr("??") 'minimum length = 2, otherwise error in Word macro !
            '                Else
            '                    'Eliminate NewLine characters because Word macro cannot handle them.
            '                    Dim label As String = CStr(pFeature.Value(FieldIndex))
            '                    rsHydrant("aanduiding").Value = Replace(label, vbNewLine, "")
            '                End If
            '                If TypeOf pFeature.Value(FieldIndex2) Is System.DBNull Then
            '                    rsHydrant("diameter").Value = CInt(0)
            '                Else
            '                    rsHydrant("diameter").Value = CInt(pFeature.Value(FieldIndex2))
            '                End If
            '                pFeature = pFCursor.NextFeature
            '            End While

            '            'Sort the hydrants on diameter.
            '            rsHydrant.Sort = "diameter DESC, aanduiding ASC"

            '            'Write the recordset to the txt export file.
            '            rsHydrant.MoveFirst()
            '            While Not rsHydrant.EOF
            '                StreamWriter.WriteLine("h," & CStr(rsHydrant("diameter").Value) & " " & Trim(CStr(rsHydrant("aanduiding").Value)))
            '                rsHydrant.MoveNext()
            '            End While

            '            'Release the recordset object.
            '            rsHydrant.Close()
            '            'rsHydrant = Nothing
            '            If Not rsHydrant Is Nothing Then Marshal.ReleaseComObject(rsHydrant)

            '        Else
            '            'Geen gevonden, schrijf code "h" (met "!!! geen hydranten") naar het txt exportbestand. Ga verder.
            '            StreamWriter.WriteLine(CStr(c_Message_NoLinkedHydrants))
            '        End If
            '        StreamWriter.Flush()

            '        'Release COM objects.
            '        If Not pQueryFilter Is Nothing Then Marshal.ReleaseComObject(pQueryFilter)
            '        If Not pFCursor Is Nothing Then Marshal.ReleaseComObject(pFCursor)
            '        If Not pFeature Is Nothing Then Marshal.ReleaseComObject(pFeature)

            '        '-- P --
            '        'Voor de lijst van gevarenthema "Hoogspanning" in het txt exportbestand ...
            '        'Is de laag van het gevarenthema niet aanwezig in huidige mxd, ga dan verder met volgende gevarenthema.
            '        'Zoek alle hoogspanning features op die aan deze straat zijn geconnecteerd.
            '        While Not (pFLayer_hoogspanning Is Nothing)
            '            Try
            '                If Not pFLayer_hoogspanning.Valid Then _
            '                    Throw New LayerNotValidException(pFLayer_hoogspanning.Name)
            '                pQueryFilter = New QueryFilter
            '                pQueryFilter.WhereClause = "(" & GetAttributeName("Hoogspanning", "Straatnaam") & " = " & CStrSql(ThisStreet) & ")"
            '                pFCursor = pFLayer_hoogspanning.Search(pQueryFilter, Nothing)
            '                'Loop door featurecursor ...
            '                pFeature = pFCursor.NextFeature
            '                If Not pFeature Is Nothing Then
            '                    FieldIndex = pFeature.Fields.FindField(GetAttributeName("Hoogspanning", "Aanduiding"))
            '                    If FieldIndex = -1 Then Throw New AttributeNotFoundException(GetLayerName("Hoogspanning"), GetAttributeName("Hoogspanning", "Aanduiding"))
            '                    While Not pFeature Is Nothing
            '                        'Schrijf code "p" (met 1 aanduiding) naar het txt exportbestand.
            '                        StreamWriter.WriteLine("p," & CStr(pFeature.Value(FieldIndex)))
            '                        pFeature = pFCursor.NextFeature
            '                    End While
            '                    StreamWriter.Flush()
            '                End If
            '                ' Continue with next chapter.
            '                Exit While
            '            Catch ex As LayerNotValidException
            '                Select Case MsgBox(ex.Message, MsgBoxStyle.Exclamation Or MsgBoxStyle.AbortRetryIgnore Or MsgBoxStyle.DefaultButton2, c_Title_OpladenHydranten)
            '                    Case MsgBoxResult.Abort
            '                        ' Abort the procedure.
            '                        Throw New AbortedByUserException
            '                    Case MsgBoxResult.Ignore
            '                        ' Avoid trying again with next feature.
            '                        pFLayer_hoogspanning = Nothing
            '                        ' Exit this part of the procedure.
            '                        Exit While
            '                    Case MsgBoxResult.Retry
            '                        ' The while loop will give it another try.
            '                End Select
            '            Finally
            '                'Release COM objects.
            '                If Not pQueryFilter Is Nothing Then Marshal.ReleaseComObject(pQueryFilter)
            '                If Not pFCursor Is Nothing Then Marshal.ReleaseComObject(pFCursor)
            '                If Not pFeature Is Nothing Then Marshal.ReleaseComObject(pFeature)
            '            End Try
            '        End While

            '        '-- B --
            '        'Voor de lijst van gevarenthema "Stralingsbron" in het txt exportbestand ...
            '        'Is de laag van het gevarenthema niet aanwezig in huidige mxd, ga dan verder met volgende gevarenthema.
            '        'Zoek alle hoogspanning features op die aan deze straat zijn geconnecteerd.
            '        While Not (pFLayer_stralingsbron Is Nothing)
            '            Try
            '                If Not pFLayer_stralingsbron.Valid Then _
            '                    Throw New LayerNotValidException(pFLayer_stralingsbron.Name)
            '                pQueryFilter = New QueryFilter
            '                pQueryFilter.WhereClause = "(" & GetAttributeName("Stralingsbron", "Straatnaam") & " = " & CStrSql(ThisStreet) & ")"
            '                pFCursor = pFLayer_stralingsbron.Search(pQueryFilter, Nothing)
            '                'Loop door featurecursor ...
            '                pFeature = pFCursor.NextFeature
            '                If Not pFeature Is Nothing Then
            '                    FieldIndex = pFeature.Fields.FindField(GetAttributeName("Stralingsbron", "Aanduiding"))
            '                    If FieldIndex = -1 Then Throw New AttributeNotFoundException(GetLayerName("Stralingsbron"), GetAttributeName("Stralingsbron", "Aanduiding"))
            '                    While Not pFeature Is Nothing
            '                        'Schrijf code "b" (met 1 aanduiding) naar het txt exportbestand.
            '                        StreamWriter.WriteLine("b," & CStr(pFeature.Value(FieldIndex)))
            '                        pFeature = pFCursor.NextFeature
            '                    End While
            '                    StreamWriter.Flush()
            '                End If
            '                ' Continue with next chapter.
            '                Exit While
            '            Catch ex As LayerNotValidException
            '                Select Case MsgBox(ex.Message, MsgBoxStyle.Exclamation Or MsgBoxStyle.AbortRetryIgnore Or MsgBoxStyle.DefaultButton2, c_Title_OpladenHydranten)
            '                    Case MsgBoxResult.Abort
            '                        ' Abort the procedure.
            '                        Throw New AbortedByUserException
            '                    Case MsgBoxResult.Ignore
            '                        ' Avoid trying again with next feature.
            '                        pFLayer_stralingsbron = Nothing
            '                        ' Exit this part of the procedure.
            '                        Exit While
            '                    Case MsgBoxResult.Retry
            '                        ' The while loop will give it another try.
            '                End Select
            '            Finally
            '                'Release COM objects.
            '                If Not pQueryFilter Is Nothing Then Marshal.ReleaseComObject(pQueryFilter)
            '                If Not pFCursor Is Nothing Then Marshal.ReleaseComObject(pFCursor)
            '                If Not pFeature Is Nothing Then Marshal.ReleaseComObject(pFeature)
            '            End Try
            '        End While

            '        '-- I --
            '        'Voor de lijst van gevarenthema "Instorting" in het txt exportbestand ...
            '        'Is de laag van het gevarenthema niet aanwezig in huidige mxd, ga dan verder met volgende gevarenthema.
            '        'Zoek alle hoogspanning features op die aan deze straat zijn geconnecteerd.
            '        While Not (pFLayer_instorting Is Nothing)
            '            Try
            '                If Not pFLayer_instorting.Valid Then _
            '                    Throw New LayerNotValidException(pFLayer_instorting.Name)
            '                pQueryFilter = New QueryFilter
            '                pQueryFilter.WhereClause = "(" & GetAttributeName("Instorting", "Straatnaam") & " = " & CStrSql(ThisStreet) & ")"
            '                pFCursor = pFLayer_instorting.Search(pQueryFilter, Nothing)
            '                'Loop door featurecursor ...
            '                pFeature = pFCursor.NextFeature
            '                If Not pFeature Is Nothing Then
            '                    FieldIndex = pFeature.Fields.FindField(GetAttributeName("Instorting", "Aanduiding"))
            '                    If FieldIndex = -1 Then Throw New AttributeNotFoundException(GetLayerName("Instorting"), GetAttributeName("Instorting", "Aanduiding"))
            '                    While Not pFeature Is Nothing
            '                        'Schrijf code "i" (met 1 aanduiding) naar het txt exportbestand.
            '                        StreamWriter.WriteLine("i," & CStr(pFeature.Value(FieldIndex)))
            '                        pFeature = pFCursor.NextFeature
            '                    End While
            '                    StreamWriter.Flush()
            '                End If
            '                ' Continue with next chapter.
            '                Exit While
            '            Catch ex As LayerNotValidException
            '                Select Case MsgBox(ex.Message, MsgBoxStyle.Exclamation Or MsgBoxStyle.AbortRetryIgnore Or MsgBoxStyle.DefaultButton2, c_Title_OpladenHydranten)
            '                    Case MsgBoxResult.Abort
            '                        ' Abort the procedure.
            '                        Throw New AbortedByUserException
            '                    Case MsgBoxResult.Ignore
            '                        ' Avoid trying again with next feature.
            '                        pFLayer_instorting = Nothing
            '                        ' Exit this part of the procedure.
            '                        Exit While
            '                    Case MsgBoxResult.Retry
            '                        ' The while loop will give it another try.
            '                End Select
            '            Finally
            '                'Release COM objects.
            '                If Not pQueryFilter Is Nothing Then Marshal.ReleaseComObject(pQueryFilter)
            '                If Not pFCursor Is Nothing Then Marshal.ReleaseComObject(pFCursor)
            '                If Not pFeature Is Nothing Then Marshal.ReleaseComObject(pFeature)
            '            End Try
            '        End While

            '        '-- C --
            '        'Voor de lijst van gevarenthema "Sleutelgebouw" in het txt exportbestand ...
            '        'Is de laag van het gevarenthema niet aanwezig in huidige mxd, ga dan verder met volgende gevarenthema.
            '        'Zoek alle hoogspanning features op die aan deze straat zijn geconnecteerd.
            '        While Not (pFLayer_sleutelgebouw Is Nothing)
            '            Try
            '                If Not pFLayer_sleutelgebouw.Valid Then _
            '                    Throw New LayerNotValidException(pFLayer_sleutelgebouw.Name)
            '                pQueryFilter = New QueryFilter
            '                pQueryFilter.WhereClause = "(" & GetAttributeName("Sleutelgebouw", "Straatnaam") & " = " & CStrSql(ThisStreet) & ")"
            '                pFCursor = pFLayer_sleutelgebouw.Search(pQueryFilter, Nothing)
            '                'Loop door featurecursor ...
            '                pFeature = pFCursor.NextFeature
            '                If Not pFeature Is Nothing Then
            '                    FieldIndex = pFeature.Fields.FindField(GetAttributeName("Sleutelgebouw", "Aanduiding"))
            '                    If FieldIndex = -1 Then Throw New AttributeNotFoundException(GetLayerName("Sleutelgebouw"), GetAttributeName("Sleutelgebouw", "Aanduiding"))
            '                    While Not pFeature Is Nothing
            '                        'Schrijf code "c" (met 1 aanduiding) naar het txt exportbestand.
            '                        StreamWriter.WriteLine("c," & CStr(pFeature.Value(FieldIndex)))
            '                        pFeature = pFCursor.NextFeature
            '                    End While
            '                    StreamWriter.Flush()
            '                End If
            '                ' Continue with next chapter.
            '                Exit While
            '            Catch ex As LayerNotValidException
            '                Select Case MsgBox(ex.Message, MsgBoxStyle.Exclamation Or MsgBoxStyle.AbortRetryIgnore Or MsgBoxStyle.DefaultButton2, c_Title_OpladenHydranten)
            '                    Case MsgBoxResult.Abort
            '                        ' Abort the procedure.
            '                        Throw New AbortedByUserException
            '                    Case MsgBoxResult.Ignore
            '                        ' Avoid trying again with next feature.
            '                        pFLayer_sleutelgebouw = Nothing
            '                        ' Exit this part of the procedure.
            '                        Exit While
            '                    Case MsgBoxResult.Retry
            '                        ' The while loop will give it another try.
            '                End Select
            '            Finally
            '                'Release COM objects.
            '                If Not pQueryFilter Is Nothing Then Marshal.ReleaseComObject(pQueryFilter)
            '                If Not pFCursor Is Nothing Then Marshal.ReleaseComObject(pFCursor)
            '                If Not pFeature Is Nothing Then Marshal.ReleaseComObject(pFeature)
            '            End Try
            '        End While

            '        '-- W --
            '        'Voor de lijst van gevarenthema "Sleepboot" in het txt exportbestand ...
            '        'Is de laag van het gevarenthema niet aanwezig in huidige mxd, ga dan verder met volgende gevarenthema.
            '        'Zoek alle hoogspanning features op die aan deze straat zijn geconnecteerd.
            '        While Not (pFLayer_sleepboot Is Nothing)
            '            Try
            '                If Not pFLayer_sleepboot.Valid Then _
            '                    Throw New LayerNotValidException(pFLayer_sleepboot.Name)
            '                pQueryFilter = New QueryFilter
            '                pQueryFilter.WhereClause = "(" & GetAttributeName("Sleepboot", "Straatnaam") & " = " & CStrSql(ThisStreet) & ")"
            '                pFCursor = pFLayer_sleepboot.Search(pQueryFilter, Nothing)
            '                ' Loop through featurecursor ...
            '                pFeature = pFCursor.NextFeature
            '                If Not pFeature Is Nothing Then
            '                    FieldIndex = pFeature.Fields.FindField(GetAttributeName("Sleepboot", "Aanduiding"))
            '                    If FieldIndex = -1 Then Throw New AttributeNotFoundException(GetLayerName("Sleepboot"), GetAttributeName("Sleepboot", "Aanduiding"))
            '                    While Not pFeature Is Nothing
            '                        'Schrijf code "w" (met 1 aanduiding) naar het txt exportbestand.
            '                        StreamWriter.WriteLine("w," & CStr(pFeature.Value(FieldIndex)))
            '                        pFeature = pFCursor.NextFeature
            '                    End While
            '                    StreamWriter.Flush()
            '                End If
            '                ' Continue with next chapter.
            '                Exit While
            '            Catch ex As LayerNotValidException
            '                Select Case MsgBox(ex.Message, MsgBoxStyle.Exclamation Or MsgBoxStyle.AbortRetryIgnore Or MsgBoxStyle.DefaultButton2, c_Title_OpladenHydranten)
            '                    Case MsgBoxResult.Abort
            '                        ' Abort the procedure.
            '                        Throw New AbortedByUserException
            '                    Case MsgBoxResult.Ignore
            '                        ' Avoid trying again with next feature.
            '                        pFLayer_sleepboot = Nothing
            '                        ' Exit this part of the procedure.
            '                        Exit While
            '                    Case MsgBoxResult.Retry
            '                        ' The while loop will give it another try.
            '                End Select
            '            Finally
            '                'Release COM objects.
            '                If Not pQueryFilter Is Nothing Then Marshal.ReleaseComObject(pQueryFilter)
            '                If Not pFCursor Is Nothing Then Marshal.ReleaseComObject(pFCursor)
            '                If Not pFeature Is Nothing Then Marshal.ReleaseComObject(pFeature)
            '            End Try
            '        End While

            '        '-- D --
            '        'Zoek alle actieve, bovengrondse hydranten, die aan de huidige straat zijn geconnecteerd.
            '        'Schrijf code "d" (met 1 aanduiding) voor elke bovengrondse hydrant.
            '        'Indien er geen gevonden worden, moet er niets worden weggeschreven.
            '        pQueryFilter = New QueryFilter
            '        pQueryFilter.WhereClause = "(" & GetAttributeName("Hydrant", "Straatnaam") & " = " & CStrSql(ThisStreet) & ")" _
            '                            & " AND (" & GetAttributeName("Hydrant", "Status") & " = '1')" _
            '                            & " AND (" & GetAttributeName("Hydrant", "HydrantType") & " = '2')"
            '        pFCursor = pFLayer_hydrant.Search(pQueryFilter, Nothing)
            '        pFeature = pFCursor.NextFeature
            '        If Not pFeature Is Nothing Then
            '            FieldIndex = pFeature.Fields.FindField(GetAttributeName("Hydrant", "Aanduiding"))
            '            If FieldIndex = -1 Then Throw New AttributeNotFoundException(GetLayerName("Hydrant"), GetAttributeName("Hydrant", "Aanduiding"))
            '            While Not pFeature Is Nothing
            '                If TypeOf pFeature.Value(FieldIndex) Is System.DBNull Then
            '                    StreamWriter.WriteLine("d,??")
            '                Else
            '                    StreamWriter.WriteLine("d," & Trim(CStr(pFeature.Value(FieldIndex))))
            '                End If
            '                pFeature = pFCursor.NextFeature
            '            End While
            '            StreamWriter.Flush()
            '        End If

            '        'Release COM objects.
            '        If Not pQueryFilter Is Nothing Then Marshal.ReleaseComObject(pQueryFilter)
            '        If Not pFCursor Is Nothing Then Marshal.ReleaseComObject(pFCursor)
            '        If Not pFeature Is Nothing Then Marshal.ReleaseComObject(pFeature)

            '        'Loop verder door de stratenrecordset en controleer of volgende straatnaam 
            '        'met dezelfde letter begint. Zo ja, moet die onder zelfde letter heading komen.
            '        rs.MoveNext()
            '        If rs.EOF Then Exit While
            '        ThisStreet = Trim((CStr(rs(c_streetFieldName).Value)).ToUpper)
            '        ThisStreetFirstLetter = GetFirstLetter(ThisStreet)

            '    End While

            '    '... volgende letter die werd aangevinkt.
            'End While




        Catch ex As Exception
            Throw ex

        Finally
            'Sluit het txt export bestand.
            StreamWriter.Flush()
            StreamWriter.Close()
            ProgressDelegate(-2, "Exportbestand gesloten")

            'Release COM & other objects.
            StreamWriter = Nothing
            'LettersEnum = Nothing
            'If Not pFCursor Is Nothing Then Marshal.ReleaseComObject(pFCursor)
            'If Not pFeature Is Nothing Then Marshal.ReleaseComObject(pFeature)
            If Not pFLayer_hoogspanning Is Nothing Then Marshal.ReleaseComObject(pFLayer_hoogspanning)
            If Not pFLayer_hydrant Is Nothing Then Marshal.ReleaseComObject(pFLayer_hydrant)
            If Not pFLayer_instorting Is Nothing Then Marshal.ReleaseComObject(pFLayer_instorting)
            If Not pFLayer_sleepboot Is Nothing Then Marshal.ReleaseComObject(pFLayer_sleepboot)
            If Not pFLayer_sleutelgebouw Is Nothing Then Marshal.ReleaseComObject(pFLayer_sleutelgebouw)
            If Not pFLayer_stralingsbron Is Nothing Then Marshal.ReleaseComObject(pFLayer_stralingsbron)
            'If Not pQueryFilter Is Nothing Then Marshal.ReleaseComObject(pQueryFilter)
        End Try

    End Sub
#End Region

#Region " Export block procedures "

    ' Header voor de geselecteerde letter van het alfabet.
    Private Sub ExportLetterHeader( _
        ByVal StreamWriter As StreamWriter, _
        ByVal currentLetter As Char)

        Try
            StreamWriter.WriteLine("g," & currentLetter)
            StreamWriter.Flush()
        Catch ex As Exception
            ErrorHandler(ex)
        End Try
    End Sub

    'Is de letter wel gevonden, schrijf code "s" (met straatnaam) naar het export txt bestand.
    Private Sub ExportStreetName( _
        ByVal StreamWriter As StreamWriter, _
        ByVal currentStreet As String)

        Try
            StreamWriter.WriteLine("s," & currentStreet)
            StreamWriter.Flush()
        Catch ex As Exception
            ErrorHandler(ex)
        End Try
    End Sub

    'Schrijf code "k" (met één raster-kwadrant-referentie = één rasterpaginanummer en één kwadrantletter)
    Private Sub ExportKwadrants( _
        ByVal StreamWriter As StreamWriter, _
        ByVal allKwadrants As String)

        Try
            Dim kwadrantsArray As String() = Split2(allKwadrants, c_ListSeparator, True)
            Array.Sort(kwadrantsArray)
            For Each currentKwadrant As String In kwadrantsArray
                While currentKwadrant.Substring(0, 1) = "0"
                    currentKwadrant = currentKwadrant.Substring(1) 'Verwijder nullen aan het begin.
                End While
                StreamWriter.WriteLine("k," & currentKwadrant)
            Next
            StreamWriter.Flush()

        Catch ex As Exception
            ErrorHandler(ex)
        End Try
    End Sub


    ' Export actieve, ondergrondse hydranten per diameter, gekoppeld aan huidige straat.
    Private Sub ExportHydrantenOndergronds( _
        ByVal StreamWriter As StreamWriter, _
        ByVal hydrantLayer As IFeatureLayer, _
        ByVal currentStreet As String)

        Dim pQueryFilter As QueryFilter = Nothing
        Dim pCursor As IFeatureCursor = Nothing
        Dim pHydrant As IFeature = Nothing

        Try

            'Zoek alle actieve, ondergrondse hydranten op die aan deze straat zijn geconnecteerd.
            pQueryFilter = New QueryFilter
            pQueryFilter.WhereClause = "(" & GetAttributeName("Hydrant", "Straatnaam") & " = " & CStrSql(currentStreet) & ")" _
                                & " AND (" & GetAttributeName("Hydrant", "Status") & " = '1')" _
                                & " AND (" & GetAttributeName("Hydrant", "HydrantType") & " = '1')"
            pCursor = hydrantLayer.Search(pQueryFilter, Nothing)
            pHydrant = pCursor.NextFeature

            'Er zijn hydranten gevonden.
            If Not pHydrant Is Nothing Then

                'Open a new recordset for collecting diameters and labels of hydrants.
                Dim rsHydrant As New ADODB.Recordset
                rsHydrant.Fields.Append("diameter", DataTypeEnum.adSmallInt)
                rsHydrant.Fields.Append("aanduiding", DataTypeEnum.adChar, m_hydrantMaxLength)
                rsHydrant.Open()

                'Determine the field indexes.
                Dim fldName1 As String = GetAttributeName("Hydrant", "Aanduiding")
                Dim fldName2 As String = GetAttributeName("Hydrant", "Diameter")
                If fldName1 Is Nothing Then Throw New ApplicationException("Attribuut 'Aanduiding' van laag 'Hydrant' ontbreekt in configuratie.")
                If fldName2 Is Nothing Then Throw New ApplicationException("Attribuut 'Diameter' van laag 'Hydrant' ontbreekt in configuratie.")
                Dim fldIndex1 As Integer = pHydrant.Fields.FindField(fldName1)
                Dim fldIndex2 As Integer = pHydrant.Fields.FindField(fldName2)
                If fldIndex1 = -1 Then Throw New AttributeNotFoundException(GetLayerName("Hydrant"), GetAttributeName("Hydrant", "Aanduiding"))
                If fldIndex2 = -1 Then Throw New AttributeNotFoundException(GetLayerName("Hydrant"), GetAttributeName("Hydrant", "Diameter"))

                'Fill recordset with hydrants.
                While Not pHydrant Is Nothing
                    rsHydrant.AddNew()

                    'Hydrant aanduiding.
                    If TypeOf pHydrant.Value(fldIndex1) Is System.DBNull Then
                        rsHydrant("aanduiding").Value = CStr("??") 'minimum length = 2, otherwise error in Word macro !
                    Else
                        'Eliminate NewLine characters because Word macro cannot handle them.
                        Dim aanduiding As String = CStr(pHydrant.Value(fldIndex1))
                        rsHydrant("aanduiding").Value = Replace(aanduiding, vbNewLine, "")
                    End If

                    'Hydrant diameter.
                    If TypeOf pHydrant.Value(fldIndex2) Is System.DBNull Then
                        rsHydrant("diameter").Value = CInt(0)
                    Else
                        rsHydrant("diameter").Value = CInt(pHydrant.Value(fldIndex2))
                    End If

                    pHydrant = pCursor.NextFeature
                End While

                'Sort the hydrants on diameter.
                rsHydrant.Sort = "diameter DESC, aanduiding ASC"

                'Write the recordset to the txt export file.
                rsHydrant.MoveFirst()
                While Not rsHydrant.EOF
                    StreamWriter.WriteLine("h," & CStr(rsHydrant("diameter").Value) & " " & Trim(CStr(rsHydrant("aanduiding").Value)))
                    rsHydrant.MoveNext()
                End While

                'Release the recordset object.
                rsHydrant.Close()

            Else
                ' Indien geen hydranten gevonden:
                ' Schrijf code "h" (met "!!! geen hydranten") naar het txt exportbestand.
                StreamWriter.WriteLine("h," & Convert.ToString(c_Message_NoLinkedHydrants))
            End If

            StreamWriter.Flush()

        Catch ex As Exception
            ErrorHandler(ex)

        Finally
            'Release COM objects.
            If Not pQueryFilter Is Nothing Then Marshal.ReleaseComObject(pQueryFilter)
            If Not pCursor Is Nothing Then Marshal.ReleaseComObject(pCursor)
            If Not pHydrant Is Nothing Then Marshal.ReleaseComObject(pHydrant)
        End Try
    End Sub


    ' Exporteer alle features uit een gevarenlaag die aan deze straat zijn gelinkt.
    Private Sub ExportGevaren( _
        ByRef StreamWriter As StreamWriter, _
        ByRef pLayer As IFeatureLayer, _
        ByVal currentStreet As String)

        Dim pQueryFilter As QueryFilter = Nothing
        Dim pCursor As IFeatureCursor = Nothing
        Dim pFeature As IFeature = Nothing

        Try

            ' Gewoon negeren indien gevarenlaag niet gekend.
            If (pLayer Is Nothing) Then Exit Sub

            ' Exception indien gevarenlaag niet valid.
            If Not pLayer.Valid Then Throw New LayerNotValidException(pLayer.Name)

            ' Bepaal de naam van het straatnaam attribuut van de gevarenlaag.
            Dim streetFldName As String = String.Empty
            Select Case pLayer.Name
                Case GetLayerName("Hoogspanning")
                    streetFldName = GetAttributeName("Hoogspanning", "Straatnaam")
                Case GetLayerName("Stralingsbron")
                    streetFldName = GetAttributeName("Stralingsbron", "Straatnaam")
                Case GetLayerName("Instorting")
                    streetFldName = GetAttributeName("Instorting", "Straatnaam")
                Case GetLayerName("Sleutelgebouw")
                    streetFldName = GetAttributeName("Sleutelgebouw", "Straatnaam")
                Case GetLayerName("Sleepboot")
                    streetFldName = GetAttributeName("Sleepboot", "Straatnaam")
                Case Else
                    Throw New ApplicationException("De kaartlaag met de naam '" & pLayer.Name & "' wordt niet ondersteund door de procedure ExportGevaren().")
            End Select

            ' Filter gevarenfeatures op basis van straatnaam.
            pQueryFilter = New QueryFilter
            pQueryFilter.WhereClause = "(" & streetFldName & " = " & CStrSql(currentStreet) & ")"
            pCursor = pLayer.Search(pQueryFilter, Nothing)
            pFeature = pCursor.NextFeature

            ' Minstens 1 gevaren feature gevonden.
            If Not pFeature Is Nothing Then

                ' Bepaal veldindex van het te exporteren attribuut.
                Dim fldName As String = String.Empty
                Select Case pLayer.Name
                    Case GetLayerName("Hoogspanning")
                        fldName = GetAttributeName("Hoogspanning", "Aanduiding")
                        If fldName Is Nothing Then Throw New ApplicationException( _
                            "Attribuut 'Aanduiding' van laag 'Hoogspanning' ontbreekt in configuratie.")
                    Case GetLayerName("Stralingsbron")
                        fldName = GetAttributeName("Stralingsbron", "Aanduiding")
                        If fldName Is Nothing Then Throw New ApplicationException( _
                            "Attribuut 'Aanduiding' van laag 'Stralingsbron' ontbreekt in configuratie.")
                    Case GetLayerName("Instorting")
                        fldName = GetAttributeName("Instorting", "Aanduiding")
                        If fldName Is Nothing Then Throw New ApplicationException( _
                            "Attribuut 'Aanduiding' van laag 'Instorting' ontbreekt in configuratie.")
                    Case GetLayerName("Sleutelgebouw")
                        fldName = GetAttributeName("Sleutelgebouw", "Aanduiding")
                        If fldName Is Nothing Then Throw New ApplicationException( _
                            "Attribuut 'Aanduiding' van laag 'Sleutelgebouw' ontbreekt in configuratie.")
                    Case GetLayerName("Sleepboot")
                        fldName = GetAttributeName("Sleepboot", "Aanduiding")
                        If fldName Is Nothing Then Throw New ApplicationException( _
                            "Attribuut 'Aanduiding' van laag 'Sleepboot' ontbreekt in configuratie.")
                    Case Else
                        Throw New ApplicationException("De kaartlaag met de naam '" & pLayer.Name & "' wordt niet ondersteund door de procedure ExportGevaren().")
                End Select
                Dim fldIndex As Integer = pFeature.Fields.FindField(fldName)
                If fldIndex = -1 Then Throw New AttributeNotFoundException(pLayer.Name, fldName)

                ' Loop door hoogspanning featurecursor ...
                While Not pFeature Is Nothing

                    ' Schrijf 1 lettercode met attribuut inhoud naar het txt exportbestand.
                    Select Case pLayer.Name
                        Case GetLayerName("Hoogspanning")
                            StreamWriter.WriteLine("p," & Convert.ToString(pFeature.Value(fldIndex)))
                        Case GetLayerName("Stralingsbron")
                            StreamWriter.WriteLine("b," & Convert.ToString(pFeature.Value(fldIndex)))
                        Case GetLayerName("Instorting")
                            StreamWriter.WriteLine("i," & Convert.ToString(pFeature.Value(fldIndex)))
                        Case GetLayerName("Sleutelgebouw")
                            StreamWriter.WriteLine("c," & Convert.ToString(pFeature.Value(fldIndex)))
                        Case GetLayerName("Sleepboot")
                            StreamWriter.WriteLine("w," & Convert.ToString(pFeature.Value(fldIndex)))
                        Case Else
                            Throw New ApplicationException("De kaartlaag met de naam '" & pLayer.Name & "' wordt niet ondersteund door de procedure ExportGevaren().")
                    End Select

                    ' Lees volgende feature uit.
                    pFeature = pCursor.NextFeature
                End While

                StreamWriter.Flush()

            End If

        Catch ex As LayerNotValidException

            Select Case MsgBox(ex.Message, MsgBoxStyle.Exclamation Or MsgBoxStyle.AbortRetryIgnore Or MsgBoxStyle.DefaultButton2, c_Title_ExportStratenindex)
                Case MsgBoxResult.Abort
                    ' Aborted by the user.
                    Throw New AbortedByUserException
                Case MsgBoxResult.Ignore
                    ' Avoid trying again with next feature by setting layer pointer to nothing.
                    pLayer = Nothing
                Case MsgBoxResult.Retry
                    ' Give it another try.
                    ExportGevaren(StreamWriter, pLayer, currentStreet)
            End Select

            ' Other exceptions are just propagated to calling procedure.

        Finally
            'Release COM objects.
            If Not pQueryFilter Is Nothing Then Marshal.ReleaseComObject(pQueryFilter)
            If Not pCursor Is Nothing Then Marshal.ReleaseComObject(pCursor)
            If Not pFeature Is Nothing Then Marshal.ReleaseComObject(pFeature)

        End Try
    End Sub

    ' Export actieve, bovengrondse hydranten, gekoppeld aan huidige straat.
    Private Sub ExportHydrantenBovengronds( _
        ByVal StreamWriter As StreamWriter, _
        ByVal hydrantLayer As IFeatureLayer, _
        ByVal currentStreet As String)

        Dim pQueryFilter As QueryFilter = Nothing
        Dim pCursor As IFeatureCursor = Nothing
        Dim pHydrant As IFeature = Nothing

        Try

            'Zoek alle actieve, bovengrondse hydranten op die aan deze straat zijn geconnecteerd.
            pQueryFilter = New QueryFilter
            pQueryFilter.WhereClause = "(" & GetAttributeName("Hydrant", "Straatnaam") & " = " & CStrSql(currentStreet) & ")" _
                                & " AND (" & GetAttributeName("Hydrant", "Status") & " = '1')" _
                                & " AND (" & GetAttributeName("Hydrant", "HydrantType") & " = '2')"
            pCursor = hydrantLayer.Search(pQueryFilter, Nothing)
            pHydrant = pCursor.NextFeature

            ' Er zijn hydranten gevonden.
            If Not pHydrant Is Nothing Then

                ' Bepaal het te exporteren attribuut.
                Dim fldName As String = GetAttributeName("Hydrant", "Aanduiding")
                If fldName Is Nothing Then Throw New ApplicationException("Attribuut 'Aanduiding' van laag 'Hydrant' ontbreekt in configuratie.")
                Dim fldIndex As Integer = pHydrant.Fields.FindField(fldName)
                If fldIndex = -1 Then Throw New AttributeNotFoundException(hydrantLayer.Name, fldName)

                ' Loop door gevonden hydranten.
                While Not pHydrant Is Nothing

                    ' Exporteer attribuut waarde.
                    If TypeOf pHydrant.Value(fldIndex) Is System.DBNull Then
                        StreamWriter.WriteLine("d,??") 'minimum length = 2, otherwise error in Word macro !
                    Else
                        'Eliminate NewLine characters because Word macro cannot handle them.
                        Dim aanduiding As String = Replace(CStr(pHydrant.Value(fldIndex)), vbNewLine, " ")
                        StreamWriter.WriteLine("d," & Trim(aanduiding))
                    End If

                    ' Lees volgende hydrant in.
                    pHydrant = pCursor.NextFeature
                End While

            Else
                ' Indien geen hydranten gevonden:
                ' Schrijf code "d" (met "!!! geen hydranten") naar het txt exportbestand.
                'StreamWriter.WriteLine("d," & Convert.ToString(c_Message_NoLinkedHydrants))
            End If

            StreamWriter.Flush()

        Catch ex As Exception
            ErrorHandler(ex)

        Finally
            'Release COM objects.
            If Not pQueryFilter Is Nothing Then Marshal.ReleaseComObject(pQueryFilter)
            If Not pCursor Is Nothing Then Marshal.ReleaseComObject(pCursor)
            If Not pHydrant Is Nothing Then Marshal.ReleaseComObject(pHydrant)

        End Try
    End Sub

    'RW: 07-08/2008
    ' Export actieve, prive hydranten, gekoppeld aan huidige straat.
    ' 3 - ondergrond -> code f
    ' 4 - bovengrond -> code e
    Private Sub ExportHydrantenPrive( _
        ByVal StreamWriter As StreamWriter, _
        ByVal hydrantLayer As IFeatureLayer, _
        ByVal currentStreet As String, _
        ByVal hydrantType As String)

        Dim pQueryFilter As QueryFilter = Nothing
        Dim pCursor As IFeatureCursor = Nothing
        Dim pHydrant As IFeature = Nothing

        Try

            'Zoek alle actieve, hydranten (3 - prive_ondergrondse of 4 - prive_bovengrondse afhankelijk van hydrantType parameter) 
            'op die aan deze straat zijn geconnecteerd.
            pQueryFilter = New QueryFilter
            pQueryFilter.WhereClause = "(" & GetAttributeName("Hydrant", "Straatnaam") & " = " & CStrSql(currentStreet) & ")" _
                                & " AND (" & GetAttributeName("Hydrant", "Status") & " = '1')" _
                                & " AND (" & GetAttributeName("Hydrant", "HydrantType") & " = '" & hydrantType & "')"
            pCursor = hydrantLayer.Search(pQueryFilter, Nothing)
            pHydrant = pCursor.NextFeature

            'Code to use in the macro
            Dim codeForMacro As String = IIf(hydrantType.Equals("3"), "f", "e").ToString()

            ' Er zijn hydranten gevonden.
            If Not pHydrant Is Nothing Then

                ' Bepaal het te exporteren attribuut.
                Dim fldName As String = GetAttributeName("Hydrant", "Aanduiding")
                If fldName Is Nothing Then Throw New ApplicationException("Attribuut 'Aanduiding' van laag 'Hydrant' ontbreekt in configuratie.")
                Dim fldIndex As Integer = pHydrant.Fields.FindField(fldName)
                If fldIndex = -1 Then Throw New AttributeNotFoundException(hydrantLayer.Name, fldName)

                ' Loop door gevonden hydranten.
                While Not pHydrant Is Nothing

                    ' Exporteer attribuut waarde.
                    If TypeOf pHydrant.Value(fldIndex) Is System.DBNull Then
                        StreamWriter.WriteLine(codeForMacro & ",??") 'minimum length = 2, otherwise error in Word macro !
                    Else
                        'Eliminate NewLine characters because Word macro cannot handle them.
                        Dim aanduiding As String = Replace(CStr(pHydrant.Value(fldIndex)), vbNewLine, " ")
                        StreamWriter.WriteLine(codeForMacro & "," & Trim(aanduiding))
                    End If

                    ' Lees volgende hydrant in.
                    pHydrant = pCursor.NextFeature
                End While

            Else
                ' Indien geen hydranten gevonden:
                ' Schrijf code "d" (met "!!! geen hydranten") naar het txt exportbestand.
                'StreamWriter.WriteLine("d," & Convert.ToString(c_Message_NoLinkedHydrants))
            End If

            StreamWriter.Flush()

        Catch ex As Exception
            ErrorHandler(ex)

        Finally
            'Release COM objects.
            If Not pQueryFilter Is Nothing Then Marshal.ReleaseComObject(pQueryFilter)
            If Not pCursor Is Nothing Then Marshal.ReleaseComObject(pCursor)
            If Not pHydrant Is Nothing Then Marshal.ReleaseComObject(pHydrant)

        End Try
    End Sub

#End Region

End Class
