Option Explicit On 
Option Strict On

Imports ADODB
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.Geodatabase

Public NotInheritable Class FormStratenIndex
    Inherits System.Windows.Forms.Form

    'Constants.
    Private Const c_LookupRS_first_label As String = "REFERENTIE" 'Hold separated list of individual kwadrant references.
    Private Const c_LookupRS_first_maxsize As Integer = 150 'Maximum number of characters for a ";"-separated list of kwadrant references (2 char raster page + 1 char kwadrant letter).
    Private Const c_LookupRS_second_label As String = "STRAATNAAM" 'Hold streetname label.
    Private Const c_LookupRS_second_maxsize As Integer = 70 'Maximum number of characters for a streetname.
    Private Const c_LookupRS_tablename As String = "straatlijst" 'Personal geodatabase tablename to store lookup information.
    Private Const c_MaxSize_hydranten_aanduiding As Integer = 60 'Maximum number of characters of a hydrant label.

    'Locals.
    Private m_application As IMxApplication 'set by constructor
    Private m_document As IMxDocument 'set by constructor

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
        Me.LabelProgressMessage = New System.Windows.Forms.Label
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar
        Me.ProgressBar2 = New System.Windows.Forms.ProgressBar
        Me.GroupBoxAlphabet.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'ButtonOK
        '
        Me.ButtonOK.Location = New System.Drawing.Point(120, 312)
        Me.ButtonOK.Name = "ButtonOK"
        Me.ButtonOK.TabIndex = 1
        Me.ButtonOK.Text = "OK"
        '
        'ButtonCancel
        '
        Me.ButtonCancel.Location = New System.Drawing.Point(208, 312)
        Me.ButtonCancel.Name = "ButtonCancel"
        Me.ButtonCancel.TabIndex = 2
        Me.ButtonCancel.Text = "Annuleren"
        '
        'GroupBoxAlphabet
        '
        Me.GroupBoxAlphabet.Controls.Add(Me.ButtonNone)
        Me.GroupBoxAlphabet.Controls.Add(Me.ButtonAll)
        Me.GroupBoxAlphabet.Controls.Add(Me.CheckedListBoxAlphabet)
        Me.GroupBoxAlphabet.Location = New System.Drawing.Point(8, 8)
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
        Me.ButtonNone.TabIndex = 2
        Me.ButtonNone.Text = "Niets"
        '
        'ButtonAll
        '
        Me.ButtonAll.Location = New System.Drawing.Point(200, 16)
        Me.ButtonAll.Name = "ButtonAll"
        Me.ButtonAll.TabIndex = 1
        Me.ButtonAll.Text = "Alles"
        '
        'CheckedListBoxAlphabet
        '
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
        Me.GroupBox1.Location = New System.Drawing.Point(8, 144)
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
        Me.GroupBox2.Location = New System.Drawing.Point(8, 232)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(280, 72)
        Me.GroupBox2.TabIndex = 8
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Vooruitgang"
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
        'ProgressBar2
        '
        Me.ProgressBar2.Enabled = False
        Me.ProgressBar2.Location = New System.Drawing.Point(4, 51)
        Me.ProgressBar2.Name = "ProgressBar2"
        Me.ProgressBar2.Size = New System.Drawing.Size(272, 16)
        Me.ProgressBar2.TabIndex = 9
        '
        'FormStratenIndex
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(292, 342)
        Me.ControlBox = False
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.GroupBoxAlphabet)
        Me.Controls.Add(Me.ButtonCancel)
        Me.Controls.Add(Me.ButtonOK)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.Name = "FormStratenIndex"
        Me.Text = "Straten Index"
        Me.GroupBoxAlphabet.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "Overloaded constructor"
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

#Region "Initialization procedures"

    'Initial state of the form when loading.
    Private Sub InitializeForm()

        'Activate every letter of the alphabet.
        InitializeCheckedListBox(Me.CheckedListBoxAlphabet, True)

        'Set the labeltext for the checkbox.
        Me.LabelUpdate.Text = _
            "Activeer deze optie enkel indien de stratenlaag of rasterlaag is gewijzigd. " & _
            "Het aanmaken van de index zal meer tijd vergen indien deze optie actief staat."

        'Computing lookup info is mandatory, if lookup table is empty.
        Dim pLookupTable As ITable = GetTable(c_LookupRS_tablename)
        If pLookupTable Is Nothing Then
            Throw New TableNotFoundException(c_LookupRS_tablename)
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

#Region "Form controls events"

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
        Try
            Dim pLookupRS As ADODB.Recordset

            'Disable button controls.
            Me.ButtonOK.Enabled = False
            Me.ButtonNone.Enabled = False
            Me.ButtonCancel.Enabled = False
            Me.ButtonAll.Enabled = False

            'Maak een lege ADODB recordset met 2 velden van gepaste grootte.
            pLookupRS = CreateLookupRecordset()

            'Update and store kwadrant references of the streets.
            If Me.CheckBoxUpdateKwadrantInfo.Checked Then
                'Vul de recordset met straatnamen en kwadranten.
                UpdateLookupInfo(pLookupRS, AddressOf OnShowProgress1, AddressOf OnSetMaxProgress1)
                'Bewaar de recordset voor een volgende keer.
                StoreRecordset(pLookupRS, GetTable(c_LookupRS_tablename))
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

        Catch ex As Exception

            'Progress monitor info
            OnShowProgress2(-2, "Export niet succesvol beëindigd")

            'Pass exception to a higher level.
            Throw ex

        Finally

            'Enable button controls.
            Me.ButtonOK.Enabled = True
            Me.ButtonNone.Enabled = True
            Me.ButtonCancel.Enabled = True
            Me.ButtonAll.Enabled = True

        End Try

        'Open Word template and run macro.
        Dim oWord As Word.ApplicationClass
        Dim dotPath As String
        oWord = New Word.ApplicationClass
        dotPath = oWord.Application.Options.DefaultFilePath(Word.WdDefaultFilePath.wdWorkgroupTemplatesPath)
        dotPath &= "\" & c_FileName_WordTemplateIndexStraten
        oWord.Documents.Open(CType(dotPath, System.Object))
        oWord.Run(c_MacroName_IndexStraten)
        oWord.Visible = True

        'Close form if export finished successfully.
        Me.Close()

    End Sub

#End Region

#Region "Progress information management"

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
    ''' 	[ex00764]	12/08/2005	Created
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
    ''' 	[ex00764]	12/08/2005	Created
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

    'Return a recordset to hold links between streets to kwadrants.
    Private Function CreateLookupRecordset() As ADODB.Recordset
        Try
            Dim rs As New ADODB.Recordset
            rs.Fields.Append(c_LookupRS_first_label, DataTypeEnum.adChar, c_LookupRS_first_maxsize)
            rs.Fields.Append(c_LookupRS_second_label, DataTypeEnum.adChar, c_LookupRS_second_maxsize)
            rs.Open()
            CreateLookupRecordset = rs
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Get a table from a personal geodatabase.
    ''' </summary>
    ''' <param name="tableName">
    '''     The name of the table to return.
    ''' </param>
    ''' <returns>
    '''     The table with the specified name.
    ''' </returns>
    ''' <remarks>
    '''     Nothing is returned if no table with specified name is found.
    ''' </remarks>
    ''' <history>
    ''' 	[ex00764]	12/08/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Function GetTable( _
            ByVal tableName As String _
            ) As ITable

        Try
            Dim pTable As ITable
            Dim pFLayer As IFeatureLayer
            Dim pWorkspace As IWorkspace
            Dim pEnumDataset As IEnumDataset
            Dim pDataset As IDataset

            'Determine the workspace of current sector.
            pFLayer = GetFeatureLayer(m_document.FocusMap, c_LayerName_hydrant)
            If pFLayer Is Nothing Then Throw New LayerNotFoundException(c_LayerName_hydrant)
            Dim pFClass As IFeatureClass = pFLayer.FeatureClass
            pDataset = CType(pFLayer.FeatureClass, IDataset)
            pWorkspace = pDataset.Workspace

            'Check all datasets of the workspace.
            pEnumDataset = pWorkspace.Datasets(esriDatasetType.esriDTTable)
            pDataset = pEnumDataset.Next
            While Not pDataset Is Nothing
                If pDataset.Name = tableName Then
                    Return CType(pDataset, ITable)
                End If
                pDataset = pEnumDataset.Next
            End While

            'No dataset with the specified name.
            Return Nothing

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
    ''' 	[ex00764]	12/08/2005	Created
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
                RowBuffer.Value(1) = Trim(CStr(Recordset(c_LookupRS_second_label).Value))
                RowBuffer.Value(2) = Trim(CStr(Recordset(c_LookupRS_first_label).Value))
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
    ''' <param name="rs">
    '''     [out] The recordset to fill.
    ''' </param>
    ''' <param name="ProgressDel">
    '''     Delegate for progress monitoring.
    ''' </param>
    ''' <param name="MaxProgressDel">
    '''     Delegate for setting maximum progress value.
    ''' </param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[ex00764]	11/08/2005	Created
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
            Dim RowIndex As Integer

            'Clear passed recordset before filling it up with new records.
            ' pRecordset.Delete() 
            ' --> seems not to work as easily, and since it's not really necessary, just skip this.

            'Get a pointer to the lookup table in the geodatabase.
            pTable = GetTable(c_LookupRS_tablename)
            MaxProgressDelegate(pTable.RowCount(Nothing))

            'Loop through the table records.
            pCursor = pTable.Search(Nothing, Nothing)
            pRow = pCursor.NextRow
            If Not pRow Is Nothing Then

                'Get the field indices to read from.
                FieldIndex_kwadrants = pRow.Fields.FindField(c_LookupRS_first_label)
                FieldIndex_streetname = pRow.Fields.FindField(c_LookupRS_second_label)
                If FieldIndex_kwadrants < 0 Then Throw New AttributeNotFoundException(c_LookupRS_tablename, c_LookupRS_first_label)
                If FieldIndex_streetname < 0 Then Throw New AttributeNotFoundException(c_LookupRS_tablename, c_LookupRS_second_label)
                While Not pRow Is Nothing

                    'Progress monitoring.
                    RowIndex += 1
                    ProgressDelegate(RowIndex, "Kwadranten info inlezen ....")

                    'Add record by record to the recordset.
                    pRecordset.AddNew()
                    pRecordset(c_LookupRS_first_label).Value = pRow.Value(FieldIndex_kwadrants)
                    pRecordset(c_LookupRS_second_label).Value = pRow.Value(FieldIndex_streetname)

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
    ''' 	[ex00764]	10/08/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub UpdateLookupInfo( _
            ByRef rs As ADODB.Recordset, _
            ByVal ProgressDel As ShowProgress, _
            ByVal MaxProgressDel As SetMaxProgress)

        Try
            Dim arrayKwadranten As IGeometry()
            Dim arrayPostcode As String()
            Dim fieldIndex_page As Integer
            Dim fieldIndex_name As Integer
            Dim fieldIndex_name1 As Integer
            Dim fieldIndex_name2 As Integer
            Dim i As Integer
            Dim kwadrantReference As String
            Dim progress As Integer
            Dim sectorCode As String
            Dim whereClause_straten As String

            Dim pFClass As IFeatureClass
            Dim pFCursor As IFeatureCursor
            Dim pFCursor_park As IFeatureCursor
            Dim pFCursor_raster As IFeatureCursor
            Dim pFCursor_straten As IFeatureCursor
            Dim pFCursor_water As IFeatureCursor
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

            'Op basis van de bestandsnaam kan de code van de sector uit het configuratiebestand worden uitgelezen. 
            sectorCode = GetSectorCode(m_document)

            'Op basis van de bestandsnaam kan een lijst van bijhorende postcodes uit het configuratiebestand worden uitgelezen. 
            arrayPostcode = GetSectorPostcodes(m_document)
            For i = 0 To arrayPostcode.Length - 1
                If Len(whereClause_straten) > 0 Then whereClause_straten &= " OR "
                whereClause_straten &= "(" & c_AttributeName_straatassen_postcode & " = '" & arrayPostcode(i) & "')"
            Next
            'whereClause_straten = _
            '    "(" & c_AttributeName_straatassen_straatnaam & "<>'ONBEKEND') AND " & _
            '    "(" & c_AttributeName_straatassen_straatnaam & "<>'PAD'     ) AND " & _
            '    "(" & c_AttributeName_straatassen_straatnaam & "<>'SNELWEG' ) AND " & _
            '    "(" & whereClause_straten & ")"
            'MsgBox(whereClause_straten, , "WhereClause voor Straten")

            'De sectoren laag opzoeken.
            pFLayer_sector = GetFeatureLayer(m_document.FocusMap, c_LayerName_sector)
            If pFLayer_sector Is Nothing Then Throw New LayerNotFoundException(c_LayerName_sector)

            'De raster laag opzoeken.
            pFLayer_raster = GetFeatureLayer(m_document.FocusMap, c_LayerName_raster)
            If pFLayer_raster Is Nothing Then Throw New LayerNotFoundException(c_LayerName_raster)
            'Bepaal veldindexen van bruikbare attributen.
            fieldIndex_page = pFLayer_raster.FeatureClass.FindField("BLZ_" & sectorCode)
            If fieldIndex_page = 0 Then Throw New AttributeNotFoundException(c_LayerName_raster, "BLZ_" & sectorCode)

            'De straten laag opzoeken.
            pFLayer_straten = GetFeatureLayer(m_document.FocusMap, c_LayerName_straatassen)
            If pFLayer_straten Is Nothing Then Throw New LayerNotFoundException(c_LayerName_straatassen)
            'Veldindex van het naam-attributen.
            fieldIndex_name = pFLayer_straten.FeatureClass.FindField(c_AttributeName_straatassen_straatnaam)
            If fieldIndex_name = -1 Then Throw New AttributeNotFoundException(c_LayerName_straatassen, c_AttributeName_straatassen_straatnaam)

            'De park laag opzoeken.
            pFLayer_park = GetFeatureLayer(m_document.FocusMap, c_LayerName_park)
            If pFLayer_park Is Nothing Then Throw New LayerNotFoundException(c_LayerName_park)
            'Veldindex van het naam-attributen.
            fieldIndex_name1 = pFLayer_park.FeatureClass.FindField(c_AttributeName_park_naam)
            If fieldIndex_name1 = -1 Then Throw New AttributeNotFoundException(c_LayerName_park, c_AttributeName_park_naam)

            'De water laag opzoeken.
            pFLayer_water = GetFeatureLayer(m_document.FocusMap, c_LayerName_water)
            If pFLayer_water Is Nothing Then Throw New LayerNotFoundException(c_LayerName_water)
            'Veldindex van het naam-attributen.
            fieldIndex_name2 = pFLayer_water.FeatureClass.FindField(c_AttributeName_water_naam)
            If fieldIndex_name2 = -1 Then Throw New AttributeNotFoundException(c_LayerName_water, c_AttributeName_water_naam)

            'Op basis van de sectorcode kan de sectorpolygoon geselecteerd worden uit de laag 'Sector'.
            pQueryFilter = New QueryFilter
            pQueryFilter.WhereClause = c_AttributeName_sector_afkorting & " = '" & sectorCode & "'"
            pFClass = CType(pFLayer_sector, IFeatureLayer2).FeatureClass()
            pFCursor = pFClass.Search(pQueryFilter, True)
            pFeature = pFCursor.NextFeature()
            'If pFeature Is Nothing Then Exit Sub
            pGeometry_sector = pFeature.ShapeCopy

            'Maak een cursor van alle rasterfeatures waarvoor het bladzijdeattribuut is ingevuld.
            pQueryFilter = New QueryFilter
            pQueryFilter.WhereClause = "BLZ_" & sectorCode & " <> ''"
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

                    'Spatial query op stratenlaag o.b.v. doorsnedepolygoon. 
                    'Bijkomend wordt een attribuutfilter gebruikt om straten 
                    'met een afwijkende postcode uit te sluiten. 
                    'Dit resulteert in een cursor van overlappende straten.
                    pSpatialFilter = New SpatialFilter
                    pSpatialFilter.Geometry = arrayKwadranten(i)
                    pSpatialFilter.SpatialRel = esriSpatialRelEnum.esriSpatialRelIntersects
                    pSpatialFilter.WhereClause = whereClause_straten
                    pFCursor_straten = pFLayer_straten.Search(pSpatialFilter, Nothing)

                    'Voeg de volledige cursor toe aan de Lookup Recordset.
                    If Not pFCursor_straten Is Nothing Then _
                        AddCursorToLookupRecordset(pFCursor_straten, rs, kwadrantReference, fieldIndex_name)

                    'Analoog voor de park laag.
                    pSpatialFilter.WhereClause = c_AttributeName_park_naam & "<>''"
                    pFCursor_park = pFLayer_park.Search(pSpatialFilter, Nothing)
                    If Not pFCursor_park Is Nothing Then _
                        AddCursorToLookupRecordset(pFCursor_park, rs, kwadrantReference, fieldIndex_name1)

                    'Analoog voor de water laag.
                    pSpatialFilter.WhereClause = c_AttributeName_water_naam & "<>''"
                    pFCursor_water = pFLayer_water.Search(pSpatialFilter, Nothing)
                    If Not pFCursor_water Is Nothing Then _
                        AddCursorToLookupRecordset(pFCursor_water, rs, kwadrantReference, fieldIndex_name2)

                Next '... volgende kwadrant.

                '... volgende rasterfeature uit cursor.
                pFeature_raster = pFCursor_raster.NextFeature
            End While

            'Progress monitor
            progress = -2 'full progress bar
            ProgressDel(progress, "Reference update beëindigd")

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Add each feature of the the cursor to the recordset.
    ''' </summary>
    ''' <param name="cursor">
    '''     [in] The feature cursor that must be added to the recordset.
    ''' </param>
    ''' <param name="recordset">
    '''     [out] The recordset that has to be expanded.
    ''' </param>
    ''' <param name="kwadrant">
    '''     [in] The kwadrant reference that is used for the whole cursor.
    ''' </param>
    ''' <remarks>
    '''     The kwadrant references are sorted in this procedure.
    '''     Therefore, for example, use "01A" and not "1A".
    ''' </remarks>
    ''' <history>
    ''' 	[ex00764]	11/08/2005	Created
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
            While Not pFeature Is Nothing


                'Bewaar straatnaam en kwadrantreferentie in de recordset, 
                'door een nieuwe straatnaam toe te voegen, 
                'of door een nieuwe enkelvoudige kwadrantreferentie toe te voegen 
                'aan een reeds geregistreerde straatnaam. 
                fieldValue1 = CStr(pFeature.Value(fieldIndex1))
                If Trim(fieldValue1) = "" Then Exit While
                If Not pRecordset.EOF Then pRecordset.MoveFirst()
                If InStr(fieldValue1, "'") > 0 Then
                    searchCondition = c_LookupRS_second_label & "='" & Replace(fieldValue1, "'", "''") & "'"
                Else
                    searchCondition = c_LookupRS_second_label & "='" & fieldValue1 & "'"
                End If
                pRecordset.Find(searchCondition)

                If pRecordset.EOF Then
                    'add new record to rs
                    pRecordset.AddNew()
                    pRecordset(c_LookupRS_first_label).Value = kwadrantReference
                    pRecordset(c_LookupRS_second_label).Value = fieldValue1
                Else
                    'get all kwadrants already in rs
                    Dim listKwadrant As String
                    Dim arrayKwadrant As String()
                    listKwadrant = CStr(pRecordset(c_LookupRS_first_label).Value)
                    arrayKwadrant = Split2(listKwadrant, c_ListSeparator, True)
                    If Array.IndexOf(arrayKwadrant, kwadrantReference) = -1 Then
                        'add ref only if kwadrant is not yet added in rs
                        ReDim Preserve arrayKwadrant(arrayKwadrant.Length)
                        arrayKwadrant.SetValue(kwadrantReference, arrayKwadrant.Length - 1)
                        Array.Sort(arrayKwadrant) 'sort references
                        listKwadrant = Concat(arrayKwadrant)
                        If Len(listKwadrant) > c_LookupRS_first_maxsize Then
                            Throw New RecordsetFieldSizeNotSufficientException(c_LookupRS_first_label)
                        Else
                            pRecordset(c_LookupRS_first_label).Value = listKwadrant
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
    ''' 	[ex00764]	12/08/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub ExportRecordset( _
            ByVal rs As ADODB.Recordset, _
            ByRef ProgressDelegate As ShowProgress, _
            ByRef MaxProgressDelegate As SetMaxProgress)

        Dim StreamWriter As StreamWriter 'for writing to the export file
        Dim LettersEnum As System.Collections.IEnumerator 'enumeration of all selected letters of the alphabet
        Dim LetterCounter As Integer 'number of processed letters - required for updating the progress bar
        Dim ThisLetter As String 'one letter from the enumeration
        Dim ThisStreet As String 'one streetname from the recordset
        Dim ThisKwadrant As String 'one singel raster-kwadrant reference (pagenumber & 1 kwadrant char)
        Dim ArrayKwadrant As String() 'array of raster-kwadrant references
        Dim FieldIndex As Integer 'index of feature attribute

        Dim pFCursor As IFeatureCursor
        Dim pFeature As IFeature
        Dim pFLayer_hoogspanning As IFeatureLayer
        Dim pFLayer_hydrant As IFeatureLayer
        Dim pFLayer_instorting As IFeatureLayer
        Dim pFLayer_sleepboot As IFeatureLayer
        Dim pFLayer_sleutelgebouw As IFeatureLayer
        Dim pFLayer_stralingsbron As IFeatureLayer
        Dim pQueryFilter As IQueryFilter

        Try 'first try-block of this procedure

            'Progress monitor information.
            MaxProgressDelegate(Me.CheckedListBoxAlphabet.CheckedItems.Count)
            ProgressDelegate(-1, "Exportbestand wordt voorbereid...")

            'Maak een nieuw txt exportbestand klaar om naar weg te schrijven.
            If File.Exists(c_FilePath_IndexStraten) Then _
                File.Delete(c_FilePath_IndexStraten)
            StreamWriter = New StreamWriter(c_FilePath_IndexStraten)

        Catch ex As Exception
            Throw New RecreateExportFileException(c_FilePath_IndexStraten)
        End Try

        Try

            'Zoek feature layers op.
            pFLayer_hoogspanning = GetFeatureLayer(m_document.FocusMap, c_LayerName_hoogspanning)
            pFLayer_hydrant = GetFeatureLayer(m_document.FocusMap, c_LayerName_hydrant)
            pFLayer_instorting = GetFeatureLayer(m_document.FocusMap, c_LayerName_instorting)
            pFLayer_sleepboot = GetFeatureLayer(m_document.FocusMap, c_LayerName_sleepboot)
            pFLayer_sleutelgebouw = GetFeatureLayer(m_document.FocusMap, c_LayerName_sleutelgebouw)
            pFLayer_stralingsbron = GetFeatureLayer(m_document.FocusMap, c_LayerName_stralingsbron)

            'Sorteer lookup recordset: alfabetisch op straatnaam.
            rs.Sort = c_LookupRS_second_label
            rs.MoveFirst()

            'Loop door alle gevraagde letters van het alfabet. Voor elke letter ...
            LettersEnum = Me.CheckedListBoxAlphabet.CheckedItems.GetEnumerator()
            While LettersEnum.MoveNext()
                LetterCounter += 1
                ThisLetter = Trim(CStr(LettersEnum.Current))
                ThisStreet = Trim((CStr(rs(c_LookupRS_second_label).Value)).ToUpper)
                ProgressDelegate(LetterCounter, "Export letter " & ThisLetter)

                'Voor de letterheading in het txt exportbestand ...
                'Loop door de stratenrecordset op zoek naar een straatnaam beginnend met die letter.
                ' (Gebruik hiervoor geen gelijkheids-operator, maar een groter/kleiner-dan-operator, 
                '  zodat de loop stopt als de letter niet wordt gevonden.)
                While 0 < String.Compare(ThisLetter, ThisStreet.Substring(0, 1), True)
                    rs.MoveNext()
                    ThisStreet = Trim((CStr(rs(c_LookupRS_second_label).Value)).ToUpper)
                End While

                '-- G --
                'Is de letter niet gevonden, schrijf code "g" (met ontbrekende letter in uppercase) naar het export txt bestand. 
                'Herneem de loop met volgende letter uit de alfabet-subset.
                If 0 > String.Compare(ThisLetter, ThisStreet.Substring(0, 1), True) Then _
                    StreamWriter.WriteLine("g," & ThisLetter)

                '-- S --
                'Is de letter wel gevonden, schrijf code "s" (met straatnaam) naar het export txt bestand.
                While 0 = String.Compare(ThisLetter, ThisStreet.Substring(0, 1), True)
                    StreamWriter.WriteLine("s," & ThisStreet)
                    StreamWriter.Flush()

                    '-- K --
                    'Schrijf code "k" (met één raster-kwadrant-referentie = één rasterpaginanummer en één kwadrantletter)
                    ArrayKwadrant = Split2(CStr(rs(c_LookupRS_first_label).Value), c_ListSeparator, True)
                    Array.Sort(ArrayKwadrant)
                    For Each ThisKwadrant In ArrayKwadrant
                        While ThisKwadrant.Substring(0, 1) = "0"
                            ThisKwadrant = ThisKwadrant.Substring(1) 'Verwijder nullen aan het begin.
                        End While
                        StreamWriter.WriteLine("k," & ThisKwadrant)
                    Next
                    StreamWriter.Flush()

                    '-- H --
                    'Voor de lijst van actieve, ondergrondse hydranten per diameter in het txt exportbestand ...
                    'Zoek alle actieve, ondergrondse hydranten op die aan deze straat zijn geconnecteerd.
                    pQueryFilter = New QueryFilter
                    pQueryFilter.WhereClause = "(" & c_AttributeName_hydrant_straatnaam & " = '" & ThisStreet & "')" _
                                        & " AND (" & c_AttributeName_hydrant_status & " = '1')" _
                                        & " AND (" & c_AttributeName_hydrant_hydranttype & " = '1')"
                    pFCursor = pFLayer_hydrant.Search(pQueryFilter, Nothing)
                    pFeature = pFCursor.NextFeature
                    If Not pFeature Is Nothing Then

                        'Er zijn hydranten gevonden.
                        Dim rsHydrant As ADODB.Recordset 'recordset for collecting diameters and labels of hydrants
                        Dim FieldIndex2 As Integer 'field index for diameter attribute of hydrants

                        'Open a new recordset for hydrants.
                        rsHydrant = New ADODB.Recordset
                        rsHydrant.Fields.Append("diameter", DataTypeEnum.adSmallInt)
                        rsHydrant.Fields.Append("aanduiding", DataTypeEnum.adChar, c_MaxSize_hydranten_aanduiding)
                        rsHydrant.Open()

                        'Determine the field indexes.
                        FieldIndex = pFeature.Fields.FindField(c_AttributeName_hydrant_aanduiding)
                        If FieldIndex = -1 Then Throw New AttributeNotFoundException(c_LayerName_hydrant, c_AttributeName_hydrant_aanduiding)
                        FieldIndex2 = pFeature.Fields.FindField(c_AttributeName_hydrant_diameter)
                        If FieldIndex2 = -1 Then Throw New AttributeNotFoundException(c_LayerName_hydrant, c_AttributeName_hydrant_diameter)

                        'Fill recordset with hydrants.
                        While Not pFeature Is Nothing
                            rsHydrant.AddNew()
                            rsHydrant("aanduiding").Value = CStr(pFeature.Value(FieldIndex))
                            rsHydrant("diameter").Value = CInt(pFeature.Value(FieldIndex2))
                            pFeature = pFCursor.NextFeature
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
                        'rsHydrant = Nothing
                        If Not rsHydrant Is Nothing Then Marshal.ReleaseComObject(rsHydrant)

                    Else
                        'Geen gevonden, schrijf code "h" (met "!!! geen hydranten") naar het txt exportbestand. Ga verder.
                        StreamWriter.WriteLine("h,!!! geen hydranten")
                    End If
                    StreamWriter.Flush()

                    'Release COM objects.
                    If Not pQueryFilter Is Nothing Then Marshal.ReleaseComObject(pQueryFilter)
                    If Not pFCursor Is Nothing Then Marshal.ReleaseComObject(pFCursor)
                    If Not pFeature Is Nothing Then Marshal.ReleaseComObject(pFeature)

                    '-- P --
                    'Voor de lijst van gevarenthema "Hoogspanning" in het txt exportbestand ...
                    'Is de laag van het gevarenthema niet aanwezig in huidige mxd, ga dan verder met volgende gevarenthema.
                    'Zoek alle hoogspanning features op die aan deze straat zijn geconnecteerd.
                    If Not pFLayer_hoogspanning Is Nothing Then
                        pQueryFilter = New QueryFilter
                        pQueryFilter.WhereClause = "(" & c_AttributeName_hoogspanning_straatnaam & " = '" & ThisStreet & "')"
                        pFCursor = pFLayer_hoogspanning.Search(pQueryFilter, Nothing)
                        'Loop door featurecursor ...
                        pFeature = pFCursor.NextFeature
                        If Not pFeature Is Nothing Then
                            FieldIndex = pFeature.Fields.FindField(c_AttributeName_hoogspanning_aanduiding)
                            If FieldIndex = -1 Then Throw New AttributeNotFoundException(c_LayerName_hoogspanning, c_AttributeName_hoogspanning_aanduiding)
                            While Not pFeature Is Nothing
                                'Schrijf code "p" (met 1 aanduiding) naar het txt exportbestand.
                                StreamWriter.WriteLine("p," & CStr(pFeature.Value(FieldIndex)))
                                pFeature = pFCursor.NextFeature
                            End While
                            StreamWriter.Flush()
                        End If

                        'Release COM objects.
                        If Not pQueryFilter Is Nothing Then Marshal.ReleaseComObject(pQueryFilter)
                        If Not pFCursor Is Nothing Then Marshal.ReleaseComObject(pFCursor)
                        If Not pFeature Is Nothing Then Marshal.ReleaseComObject(pFeature)
                    End If

                    '-- B --
                    'Voor de lijst van gevarenthema "Stralingsbron" in het txt exportbestand ...
                    'Is de laag van het gevarenthema niet aanwezig in huidige mxd, ga dan verder met volgende gevarenthema.
                    'Zoek alle hoogspanning features op die aan deze straat zijn geconnecteerd.
                    If Not pFLayer_stralingsbron Is Nothing Then
                        pQueryFilter = New QueryFilter
                        pQueryFilter.WhereClause = "(" & c_AttributeName_stralingsbron_straatnaam & " = '" & ThisStreet & "')"
                        pFCursor = pFLayer_stralingsbron.Search(pQueryFilter, Nothing)
                        'Loop door featurecursor ...
                        pFeature = pFCursor.NextFeature
                        If Not pFeature Is Nothing Then
                            FieldIndex = pFeature.Fields.FindField(c_AttributeName_stralingsbron_aanduiding)
                            If FieldIndex = -1 Then Throw New AttributeNotFoundException(c_LayerName_stralingsbron, c_AttributeName_stralingsbron_aanduiding)
                            While Not pFeature Is Nothing
                                'Schrijf code "b" (met 1 aanduiding) naar het txt exportbestand.
                                StreamWriter.WriteLine("b," & CStr(pFeature.Value(FieldIndex)))
                                pFeature = pFCursor.NextFeature
                            End While
                            StreamWriter.Flush()
                        End If

                        'Release COM objects.
                        If Not pQueryFilter Is Nothing Then Marshal.ReleaseComObject(pQueryFilter)
                        If Not pFCursor Is Nothing Then Marshal.ReleaseComObject(pFCursor)
                        If Not pFeature Is Nothing Then Marshal.ReleaseComObject(pFeature)
                    End If

                    '-- I --
                    'Voor de lijst van gevarenthema "Instorting" in het txt exportbestand ...
                    'Is de laag van het gevarenthema niet aanwezig in huidige mxd, ga dan verder met volgende gevarenthema.
                    'Zoek alle hoogspanning features op die aan deze straat zijn geconnecteerd.
                    If Not pFLayer_instorting Is Nothing Then
                        pQueryFilter = New QueryFilter
                        pQueryFilter.WhereClause = "(" & c_AttributeName_instorting_straatnaam & " = '" & ThisStreet & "')"
                        pFCursor = pFLayer_instorting.Search(pQueryFilter, Nothing)
                        'Loop door featurecursor ...
                        pFeature = pFCursor.NextFeature
                        If Not pFeature Is Nothing Then
                            FieldIndex = pFeature.Fields.FindField(c_AttributeName_instorting_aanduiding)
                            If FieldIndex = -1 Then Throw New AttributeNotFoundException(c_LayerName_instorting, c_AttributeName_instorting_aanduiding)
                            While Not pFeature Is Nothing
                                'Schrijf code "i" (met 1 aanduiding) naar het txt exportbestand.
                                StreamWriter.WriteLine("i," & CStr(pFeature.Value(FieldIndex)))
                                pFeature = pFCursor.NextFeature
                            End While
                            StreamWriter.Flush()
                        End If

                        'Release COM objects.
                        If Not pQueryFilter Is Nothing Then Marshal.ReleaseComObject(pQueryFilter)
                        If Not pFCursor Is Nothing Then Marshal.ReleaseComObject(pFCursor)
                        If Not pFeature Is Nothing Then Marshal.ReleaseComObject(pFeature)
                    End If

                    '-- C --
                    'Voor de lijst van gevarenthema "Sleutelgebouw" in het txt exportbestand ...
                    'Is de laag van het gevarenthema niet aanwezig in huidige mxd, ga dan verder met volgende gevarenthema.
                    'Zoek alle hoogspanning features op die aan deze straat zijn geconnecteerd.
                    If Not pFLayer_sleutelgebouw Is Nothing Then
                        pQueryFilter = New QueryFilter
                        pQueryFilter.WhereClause = "(" & c_AttributeName_sleutelgebouw_straatnaam & " = '" & ThisStreet & "')"
                        pFCursor = pFLayer_sleutelgebouw.Search(pQueryFilter, Nothing)
                        'Loop door featurecursor ...
                        pFeature = pFCursor.NextFeature
                        If Not pFeature Is Nothing Then
                            FieldIndex = pFeature.Fields.FindField(c_AttributeName_sleutelgebouw_aanduiding)
                            If FieldIndex = -1 Then Throw New AttributeNotFoundException(c_LayerName_sleutelgebouw, c_AttributeName_sleutelgebouw_aanduiding)
                            While Not pFeature Is Nothing
                                'Schrijf code "c" (met 1 aanduiding) naar het txt exportbestand.
                                StreamWriter.WriteLine("c," & CStr(pFeature.Value(FieldIndex)))
                                pFeature = pFCursor.NextFeature
                            End While
                            StreamWriter.Flush()
                        End If

                        'Release COM objects.
                        If Not pQueryFilter Is Nothing Then Marshal.ReleaseComObject(pQueryFilter)
                        If Not pFCursor Is Nothing Then Marshal.ReleaseComObject(pFCursor)
                        If Not pFeature Is Nothing Then Marshal.ReleaseComObject(pFeature)
                    End If

                    '-- W --
                    'Voor de lijst van gevarenthema "Sleepboot" in het txt exportbestand ...
                    'Is de laag van het gevarenthema niet aanwezig in huidige mxd, ga dan verder met volgende gevarenthema.
                    'Zoek alle hoogspanning features op die aan deze straat zijn geconnecteerd.
                    If Not pFLayer_sleepboot Is Nothing Then
                        pQueryFilter = New QueryFilter
                        pQueryFilter.WhereClause = "(" & c_AttributeName_sleutelgebouw_straatnaam & " = '" & ThisStreet & "')"
                        pFCursor = pFLayer_sleepboot.Search(pQueryFilter, Nothing)
                        'Loop door featurecursor ...
                        pFeature = pFCursor.NextFeature
                        If Not pFeature Is Nothing Then
                            FieldIndex = pFeature.Fields.FindField(c_AttributeName_sleepboot_aanduiding)
                            If FieldIndex = -1 Then Throw New AttributeNotFoundException(c_LayerName_sleepboot, c_AttributeName_sleepboot_aanduiding)
                            While Not pFeature Is Nothing
                                'Schrijf code "w" (met 1 aanduiding) naar het txt exportbestand.
                                StreamWriter.WriteLine("w," & CStr(pFeature.Value(FieldIndex)))
                                pFeature = pFCursor.NextFeature
                            End While
                            StreamWriter.Flush()
                        End If

                        'Release COM objects.
                        If Not pQueryFilter Is Nothing Then Marshal.ReleaseComObject(pQueryFilter)
                        If Not pFCursor Is Nothing Then Marshal.ReleaseComObject(pFCursor)
                        If Not pFeature Is Nothing Then Marshal.ReleaseComObject(pFeature)
                    End If

                    '-- D --
                    'Zoek alle actieve, bovengrondse hydranten, die aan de huidige straat zijn geconnecteerd.
                    'Schrijf code "d" (met 1 aanduiding) voor elke bovengrondse hydrant.
                    'Indien er geen gevonden worden, moet er niets worden weggeschreven.
                    pQueryFilter = New QueryFilter
                    pQueryFilter.WhereClause = "(" & c_AttributeName_hydrant_straatnaam & " = '" & ThisStreet & "')" _
                                        & " AND (" & c_AttributeName_hydrant_status & " = '1')" _
                                        & " AND (" & c_AttributeName_hydrant_hydranttype & " = '2')"
                    pFCursor = pFLayer_hydrant.Search(pQueryFilter, Nothing)
                    pFeature = pFCursor.NextFeature
                    If Not pFeature Is Nothing Then
                        FieldIndex = pFeature.Fields.FindField(c_AttributeName_hydrant_aanduiding)
                        If FieldIndex = -1 Then Throw New AttributeNotFoundException(c_LayerName_hydrant, c_AttributeName_hydrant_aanduiding)
                        While Not pFeature Is Nothing
                            StreamWriter.WriteLine("d," & CStr(pFeature.Value(FieldIndex)))
                            pFeature = pFCursor.NextFeature
                        End While
                        StreamWriter.Flush()
                    End If

                    'Release COM objects.
                    If Not pQueryFilter Is Nothing Then Marshal.ReleaseComObject(pQueryFilter)
                    If Not pFCursor Is Nothing Then Marshal.ReleaseComObject(pFCursor)
                    If Not pFeature Is Nothing Then Marshal.ReleaseComObject(pFeature)

                    'Loop verder door de stratenrecordset en controleer of volgende straatnaam 
                    'met dezelfde letter begint. Zo ja, moet die onder zelfde letter heading komen.
                    rs.MoveNext()
                    If rs.EOF Then Exit While
                    ThisStreet = Trim((CStr(rs(c_LookupRS_second_label).Value)).ToUpper)

                End While

                '... volgende letter die werd aangevinkt.
            End While

        Catch ex As Exception
            Throw ex
        Finally
            'Sluit het txt export bestand.
            StreamWriter.Flush()
            StreamWriter.Close()
            ProgressDelegate(-2, "Exportbestand gesloten")

            'Release COM & other objects.
            StreamWriter = Nothing
            LettersEnum = Nothing 
            If Not pFCursor Is Nothing Then Marshal.ReleaseComObject(pFCursor)
            If Not pFeature Is Nothing Then Marshal.ReleaseComObject(pFeature)
            If Not pFLayer_hoogspanning Is Nothing Then Marshal.ReleaseComObject(pFLayer_hoogspanning)
            If Not pFLayer_hydrant Is Nothing Then Marshal.ReleaseComObject(pFLayer_hydrant)
            If Not pFLayer_instorting Is Nothing Then Marshal.ReleaseComObject(pFLayer_instorting)
            If Not pFLayer_sleepboot Is Nothing Then Marshal.ReleaseComObject(pFLayer_sleepboot)
            If Not pFLayer_sleutelgebouw Is Nothing Then Marshal.ReleaseComObject(pFLayer_sleutelgebouw)
            If Not pFLayer_stralingsbron Is Nothing Then Marshal.ReleaseComObject(pFLayer_stralingsbron)
            If Not pQueryFilter Is Nothing Then Marshal.ReleaseComObject(pQueryFilter)
        End Try

    End Sub

End Class
