Option Explicit On 
Option Strict On

#Region " Imports namespaces "
Imports System.Drawing.Color
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Microsoft.VisualBasic.MsgBoxStyle
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.CartoUI
Imports ESRI.ArcGIS.Display
Imports ESRI.ArcGIS.Editor
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Geometry
#End Region

'All kinds of procedures that can be used all over the project.
Module ModuleToolsLib

    Private m_LayerNames As Hashtable 'layer names hashtable (list of key-value pairs)
    Private m_AttributeNames As Hashtable 'attribute names hashtable (list of key-value pairs)
    Private m_DomainNames As Hashtable 'domain names hashtable (list of key-value pairs)

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Handle exceptions to the user.
    ''' </summary>
    ''' <param name="err">
    '''     The exception object.
    ''' </param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    ''' 	[Kristof Vydt]	11/08/2006	Msgbox style modified.
    '''     [Koen Vermeer]  23/06/2008  Conversion to ArcGIS 9.2
    ''' 	[Elton Manoku]	28/11/2008	To see changes search for RW:
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub ErrorHandler( _
        ByVal err As Exception)

        Try
            Dim message As String     ' Messagebox prompt
            Dim title As String       ' Messagebox title
            Dim style As MsgBoxStyle  ' Messagebox style

            If TypeOf (err) Is LayerNotFoundException Then
                message = "Laag '" & DirectCast(err, LayerNotFoundException).LayerName() & "' werd niet gevonden."

            Else 'Not specified error handler.
                message = err.Message
            End If

            'Display exception info.
            title = "Hydrantenbeheer: Fout opgetreden."
            style = MsgBoxSetForeground Or Exclamation Or YesNo Or DefaultButton2
            message = message & vbNewLine & vbNewLine & "Klik 'Yes' voor gedetailleerde informatie over deze fout."
            If vbYes = MsgBox(message, style, title) Then ShowStackTrace(err)

        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Show the exception stack trace to the user.
    ''' </summary>
    ''' <param name="ex">
    '''     The exception object.
    ''' </param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    ''' 	[Kristof Vydt]	19/04/2007	Restructured to avoid exceptions.
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub ShowStackTrace(ByVal ex As Exception)

        ' Start with empty message.
        Dim MsgBoxText As String = String.Empty

        ' Add exception source to the message.
        If Not (ex.Source.Equals(String.Empty)) Then
            If Len(MsgBoxText) > 0 Then MsgBoxText = MsgBoxText & vbNewLine & vbNewLine
            MsgBoxText = MsgBoxText & "Source: " & ex.Source
        End If

        ' Add stack trace to the message.
        If Not (ex.StackTrace.Equals(String.Empty)) Then
            If Len(MsgBoxText) > 0 Then MsgBoxText = MsgBoxText & vbNewLine & vbNewLine
            MsgBoxText = MsgBoxText & "StackTrace: " & vbNewLine & ex.StackTrace
        End If

        ' Add base exception message to the message.
        If Not (ex.GetBaseException Is Nothing) Then
            If Len(MsgBoxText) > 0 Then MsgBoxText = MsgBoxText & vbNewLine & vbNewLine
            MsgBoxText = MsgBoxText & "BaseException: " & ex.GetBaseException.ToString
        End If

        ' Add inner exception to the message.
        If Not (ex.InnerException Is Nothing) Then
            If Len(MsgBoxText) > 0 Then MsgBoxText = MsgBoxText & vbNewLine & vbNewLine
            MsgBoxText = MsgBoxText & "InnerException: " & ex.InnerException.ToString
        End If

        ' Show all info messages.
        MsgBox( _
            MsgBoxText, _
            CType(MsgBoxSetForeground + Information + OKOnly, MsgBoxStyle), _
            ex.GetType.ToString)

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Get some value (preferably a string) into some control (textbox, combobox, datetimepicker).
    '''     An optional checkbox can be passed to check if the result is probably correct.
    ''' </summary>
    ''' <param name="SomeControl">
    '''     [in] TextBox, DateTimePicker or ComboBox control
    ''' </param>
    ''' <param name="SomeValue">
    '''     [in] The value to display in the control. It can be text, a numeric value or a date.
    ''' </param>
    ''' <param name="UseCheckBox">
    '''     [in][optional] CheckBox control.
    ''' </param>
    ''' <remarks>
    '''     If the value is valid according to the control type,
    '''     the available checkbox will be checked.
    '''     Unvalid values are displayed as empty strings.
    '''     The ComboBox uses a visual syntax: {code}:{label}.
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub SetEditBoxValue( _
            ByVal SomeControl As Control, _
            ByVal SomeValue As Object, _
            Optional ByVal UseCheckBox As CheckBox = Nothing)

        Try
            If TypeOf SomeControl Is TextBox Then

                'Copy a string value into a TextBox control.
                'If value is nothing, set empty string as text.
                Dim SomeTextBox As TextBox = CType(SomeControl, TextBox)
                If TypeOf SomeValue Is System.DBNull Then
                    SomeTextBox.Text = CStr("")
                    If Not UseCheckBox Is Nothing Then _
                        UseCheckBox.Checked = False
                Else
                    SomeTextBox.Text = CStr(SomeValue)
                    If Not UseCheckBox Is Nothing Then _
                        UseCheckBox.Checked = True
                End If

            ElseIf TypeOf SomeControl Is DateTimePicker Then

                'Copy a date string value into a DateTimePicker control.
                'If value is nothing, set empty string as text.
                Dim SomeDateBox As DateTimePicker = CType(SomeControl, DateTimePicker)
                If TypeOf SomeValue Is System.DBNull Then
                    SomeDateBox.Text = CStr("")
                    If Not UseCheckBox Is Nothing Then
                        UseCheckBox.Checked = False
                    End If
                Else
                    SomeDateBox.Text = CStr(SomeValue)
                    If Not UseCheckBox Is Nothing Then
                        UseCheckBox.Checked = True
                    End If
                End If

            ElseIf TypeOf SomeControl Is ComboBox Then

                'Select the item that starts with the specified value & ":".
                'If no matching item is found, set the specified value as text.
                Dim SomeComboBox As ComboBox = CType(SomeControl, ComboBox)
                Dim SomeTextValue As String
                Dim ItemIndex As Integer
                If TypeOf SomeValue Is System.DBNull Then
                    SomeTextValue = CStr("")
                Else
                    SomeTextValue = CStr(SomeValue)
                End If
                ItemIndex = SomeComboBox.FindString(SomeTextValue & ":")
                If SomeComboBox.FindString(SomeTextValue & ":") > -1 Then
                    SomeComboBox.SelectedIndex = ItemIndex
                    If Not UseCheckBox Is Nothing Then
                        UseCheckBox.Checked = True
                    End If
                Else
                    SomeComboBox.Text = SomeTextValue
                    If Not UseCheckBox Is Nothing Then
                        UseCheckBox.Checked = False
                    End If
                End If

            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Return the text before the double point, of the selected combobox item.
    ''' </summary>
    ''' <param name="SomeComboBox">
    '''     A ComboBox control.
    ''' </param>
    ''' <returns></returns>
    ''' <remarks>
    '''     The ComboBox should use a visual syntax {code}:{label},
    '''     otherwise an empty string is returned.
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function GetComboBoxCodeValue( _
            ByVal SomeComboBox As ComboBox _
            ) As String

        Try
            Dim ComboBoxText As String
            Dim SelectedCode As String
            Dim SeparatorPosition As Integer

            ComboBoxText = SomeComboBox.Text
            SeparatorPosition = InStr(ComboBoxText, ":")
            If SeparatorPosition > 0 Then
                SelectedCode = Left(ComboBoxText, SeparatorPosition - 1)
            Else
                SelectedCode = ""
            End If
            GetComboBoxCodeValue = SelectedCode

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Zoom to the specified feature with a fixed buffer extent.
    '''     Mark the spot on the map by setting a reusable marker graphic.
    ''' </summary>
    ''' <param name="pFeature">
    '''     The feature you want to zoom to on the map.
    ''' </param>
    ''' <param name="pMxDocument">
    '''     The ArcMap document you are working in.
    ''' </param>
    ''' <remarks>
    '''     If marker could not be found in focus map, a new one is created.
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    ''' 	[Kristof Vydt]	24/10/2005	Adjust zoomextent for polygons to 3x objectextent.
    ''' 	[Kristof Vydt]	18/08/2006	Eliminate MarkerElement as parameter. Recover existing one based on its name.
    '''     [Kristof Vydt]  20/02/2007  Introduce the use of c_ZoomPolygonBuffer.
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub MarkAndZoomTo( _
        ByVal pFeature As IFeature, _
        ByVal pMxDocument As IMxDocument, _
        Optional ByVal zoomToFeature As Boolean = False)

        Dim pActiveView As IActiveView
        Dim pElement As IElement
        Dim pElementProp As IElementProperties
        Dim pEnvelope As IEnvelope
        Dim pFont As System.Drawing.Font
        Dim pFontDisp As stdole.IFontDisp
        Dim pGraphics As IGraphicsContainer
        Dim pMarkerElement As IMarkerElement
        Dim pMarkerSymbol As ICharacterMarkerSymbol
        Dim pPoint As IPoint
        Dim rgbColor As IRgbColor

        Try
            pEnvelope = pFeature.Extent
            pActiveView = pMxDocument.ActivatedView()
            If pFeature.Shape.GeometryType = esriGeometryType.esriGeometryPoint Then
                'buffer around the point.
                pEnvelope.Expand(c_ZoomPointBuffer, c_ZoomPointBuffer, False)
            Else 'expand the feature extent
                pEnvelope.Expand(c_ZoomPolygonBuffer, c_ZoomPolygonBuffer, True)
            End If

            If (zoomToFeature) Then
                ' Zoom to this specific feature.
                pActiveView.Extent = pEnvelope
            End If

            ' Determine new point location for the marker.
            pPoint = New PointClass
            pPoint.X = (pEnvelope.XMax + pEnvelope.XMin) / 2
            pPoint.Y = (pEnvelope.YMax + pEnvelope.YMin) / 2
            Debug.WriteLine("Markerpoint: { " & pPoint.X & " ; " & pPoint.Y & " }")

            ' Retrieve existing marker element.
            pMarkerElement = GetMarkerElement(c_MarkerTag, pMxDocument)

            ' Create new marker if none was found.
            If pMarkerElement Is Nothing Then

                ' Select font with marker symbols.
                pFont = New System.Drawing.Font("ESRI Cartography", 18)
                pFontDisp = FontConverter.FontToOLEFont(pFont)

                ' Set the marker color.
                rgbColor = New RgbColorClass
                rgbColor.Red = 100
                rgbColor.Green = 150
                rgbColor.Blue = 100

                ' Create a marker symbol.
                pMarkerSymbol = New CharacterMarkerSymbol
                pMarkerSymbol.Color = rgbColor
                pMarkerSymbol.CharacterIndex = 72
                pMarkerSymbol.Font = pFontDisp
                pMarkerSymbol.Size = 32

                ' Create a new marker element.
                pMarkerElement = New MarkerElementClass
                pMarkerElement.Symbol = pMarkerSymbol
                pElement = CType(pMarkerElement, IElement)
                pElement.Geometry = pPoint

                ' Tag the marker to recognize it later.
                pElementProp = CType(pElement, IElementProperties)
                pElementProp.Name = c_MarkerTag

                ' Add new marker to graphics container.
                pGraphics = CType(pMxDocument.FocusMap, IGraphicsContainer)
                pGraphics.AddElement(pElement, 0)
            End If

            ' Move the marker to its new location.
            If Not pMarkerElement Is Nothing Then
                pElement = CType(pMarkerElement, IElement)
                pElement.Geometry = pPoint
            End If

            ' Refresh active view.
            pActiveView.Refresh()

        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Returns the concatinated array elements.
    ''' </summary>
    ''' <param name="StringArray">
    '''     Array of strings.
    ''' </param>
    ''' <param name="Separator">
    '''     [optional] String to put between 2 array string elements.
    ''' </param>
    ''' <returns>
    '''     Concatinated string.
    ''' </returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function Concat( _
            ByVal StringArray As String(), _
            Optional ByVal Separator As String = c_ListSeparator _
            ) As String
        Concat = ""
        Dim i As Integer 'loop index
        For i = 0 To StringArray.Length - 1
            If i > 0 Then Concat = Concat & Separator
            Concat = Concat & StringArray(i)
        Next
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Return feature selection set of one feature layer.
    ''' </summary>
    ''' <param name="pMxDocument">
    '''     The ArcMap document you are working with.
    ''' </param>
    ''' <param name="pFeatureLayer">
    '''     The feature layer you want to know the selection set of.
    ''' </param>
    ''' <returns>
    '''     The selection set of the specified feature.
    ''' </returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function LayerSelectionSet( _
            ByVal pMxDocument As IMxDocument, _
            ByVal pFeatureLayer As IFeatureLayer _
            ) As ISelectionSet

        Try
            Dim pMap As IMap
            Dim pFeatureSelection As IFeatureSelection
            Dim pSelectionSet As ISelectionSet

            pMap = pMxDocument.FocusMap
            pFeatureSelection = CType(pFeatureLayer, IFeatureSelection)
            pSelectionSet = pFeatureSelection.SelectionSet
            LayerSelectionSet = pSelectionSet

        Catch ex As Exception
            Throw ex
        End Try

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Determine the Brandweer Sector Name.
    ''' </summary>
    ''' <param name="pMxDocument">
    '''     ArcMap document you are working in.
    ''' </param>
    ''' <returns>
    '''     A string representing a sector by its name.
    ''' </returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    '''     [Kristof Vydt]  22/02/2007  Use XML-based configuration.
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function GetSectorName( _
                ByVal pMxDocument As IMxDocument _
                ) As String

        'TODO: Eliminate argument.

        ' Check availability of configuration.
        If Config Is Nothing Then Throw New ApplicationException("No configuration loaded.")

        ' Return configured layer name.
        Return Config.SectorName()

        'Try
        '    Dim DocumentInfo As IDocumentInfo
        '    Dim DocumentTitle As String
        '    Dim SectorName As String

        '    'Get the document title.
        '    If Not TypeOf pMxDocument Is IDocumentInfo Then Exit Function
        '    DocumentInfo = CType(pMxDocument, IDocumentInfo)
        '    DocumentTitle = DocumentInfo.DocumentTitle()

        '    'Remove the file extension.
        '    If Right(DocumentTitle, 4) = ".mxd" Then _
        '        SectorName = Left(DocumentTitle, Len(DocumentTitle) - 4)

        '    'Validate output.
        '    If SectorName = "" Then Throw New SectorNameException
        '    GetSectorName = SectorName

        'Catch ex As Exception
        '    Throw ex
        'End Try

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Determine the Brandweer Sector Code
    ''' </summary>
    ''' <param name="pMxDocument">
    '''     ArcMap document you are working in.
    ''' </param>
    ''' <returns>
    '''     A string representing a sector by its code, 
    '''     consisting of 1/2/3 uppercase characters.
    ''' </returns>
    ''' <remarks>
    '''     Uses the configuration ini-file.
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    '''     [Kristof Vydt]  22/02/2007  Use XML-based configuration.
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function GetSectorCode( _
                ByVal pMxDocument As IMxDocument _
                ) As String

        'TODO: Eliminate argument.

        ' Check availability of configuration.
        If Config Is Nothing Then Throw New ApplicationException("No configuration loaded.")

        ' Return configured layer name.
        Return Config.SectorCode()

    End Function

    'Public Function GetSectorCode( _
    '    ByVal pMxDocument As IMxDocument _
    '    ) As String

    '    Try
    '        Dim SectorName As String
    '        Dim SectorCode As String

    '        'Determine the sector name.
    '        SectorName = GetSectorName(pMxDocument)

    '        'Look for the corresponding code in the configuration file.
    '        SectorCode = (INIRead(g_FilePath_Config, "SectorCodes", SectorName)).ToUpper

    '        'Validate the output.
    '        If SectorCode = "" Then Throw New SectorCodeException(SectorName)
    '        GetSectorCode = SectorCode

    '    Catch ex As Exception
    '        Throw ex
    '    End Try

    'End Function

    '''' -----------------------------------------------------------------------------
    '''' <summary>
    '''' Return all postcodes for the current sector.
    '''' </summary>
    '''' <param name="pMxDocument"></param>
    '''' <returns></returns>
    '''' <remarks>
    '''' </remarks>
    '''' <history>
    '''' 	[Kristof Vydt]	16/09/2005	Created
    '''' 	[Kristof Vydt]	22/02/2007	Deprecated
    '''' </history>
    '''' -----------------------------------------------------------------------------
    'Public Function GetSectorPostcodes( _
    '            ByVal pMxDocument As IMxDocument _
    '            ) As Collection

    '    Try
    '        Dim sectorName As String
    '        Dim postcodeList As String

    '        'Determine the sector name.
    '        sectorName = GetSectorName(pMxDocument)

    '        'Look for the corresponding postcodelist in the configuration file.
    '        postcodeList = INIRead(g_FilePath_Config, "Postcodes", sectorName)

    '        'Validate the output.
    '        If postcodeList = "" Then Throw New SectorPostcodeException(sectorName)

    '        'Split string to array.
    '        GetSectorPostcodes = postcodeList.Split(c_ListSeparator)

    '    Catch ex As Exception
    '        Throw ex
    '    End Try

    'End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Generate a new LeverancierNummer.
    ''' </summary>
    ''' <param name="pMxDocument">
    '''     The ArcMap document you are working with.
    ''' </param>
    ''' <returns>
    '''     The new LeverancierNummer as string.
    ''' </returns>
    ''' <remarks>
    '''     LeverancierNummer = sectorcode + incrementing autonumber
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    ''' 	[Kristof Vydt]	02/08/2006	Replace some global constant by a function that reads the value from config file.
    ''' 	[Elton Manoku]	27/11/2008	Add the finally statement
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function NewLerancierNr( _
            ByVal pMxDocument As IMxDocument _
            ) As String
        Dim pFeatureCursor As IFeatureCursor = Nothing
        Try
            'Determine the code of the current BrandweerSector.
            Dim SectorCode As String = GetSectorCode(pMxDocument)

            'Determine the highest number currently in use for this sector.
            Dim pLayer As IFeatureLayer = GetFeatureLayer(pMxDocument.FocusMap, GetLayerName("Hydrant"))
            Dim pQueryFilter As IQueryFilter = New QueryFilter
            pQueryFilter.WhereClause = GetAttributeName("Hydrant", "LeverancierNr") & " LIKE '" & SectorCode & "*'"
            pFeatureCursor = pLayer.FeatureClass.Search(pQueryFilter, True)

            'TODO: Moet toch efficienter te doen zijn:
            'opvragen van enkel de grootste in plaats van elke record te doorlopen.
            'From IQueryFilter reference:
            'ORDER BY cannot be used with ArcObjects. If ordered results are required 
            'you need to use ITableSort. A method that allows the use of ORDER BY is 
            'planned for a future release.

            Dim CurrentNr As Integer = 0
            Dim HighestNr As Integer = 0
            Dim CurrentCode As String = ""
            Dim pFeature As IFeature = pFeatureCursor.NextFeature
            If Not pFeature Is Nothing Then
                Dim FieldIndex As Integer = pFeature.Fields.FindField(GetAttributeName("Hydrant", "LeverancierNr"))

                If FieldIndex > -1 Then
                    While Not pFeature Is Nothing
                        CurrentCode = CStr(pFeature.Value(FieldIndex))
                        CurrentNr = CInt(Mid(CurrentCode, Len(SectorCode) + 1))
                        HighestNr = MaxInt(HighestNr, CurrentNr)
                        pFeature = pFeatureCursor.NextFeature
                    End While
                End If
            End If
            'Compose the new code.
            Return SectorCode & CStr(HighestNr + 1)

        Catch ex As Exception
            Throw ex
        Finally
            'RW:2008
            If Not pFeatureCursor Is Nothing Then Marshal.ReleaseComObject(pFeatureCursor)
        End Try

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Return a new value for a specified numeric feature attribute.
    '''     This can be the highest current value + 1 or a missing value in the range.
    ''' </summary>
    ''' <param name="featureLayer">
    '''     The feature layer to analyse.
    ''' </param>
    ''' <param name="attributeName">
    '''     The name of the attribute to analyse.
    ''' </param>
    ''' <param name="fillGaps">
    '''     Optional indicator to fill gaps in existing numbering,
    '''     returning a missing value below the maximum.
    ''' </param>
    ''' <returns>
    '''     Nothing in case no (valid) values were found.
    '''     The highest number (integer, double, ...) that was found.
    ''' </returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function NextUniqueAttributeValue( _
            ByVal featureLayer As IFeatureLayer, _
            ByVal attributeName As String, _
            Optional ByVal fillGaps As Boolean = False _
            ) As Object
        Try
            'Input validation.
            If featureLayer Is Nothing Then Throw New ArgumentNullException("featureLayer")
            Dim featureCursor As IFeatureCursor = featureLayer.FeatureClass.Search(Nothing, Nothing)
            Dim feature As IFeature = featureCursor.NextFeature
            Dim fieldIndex As Integer = featureLayer.FeatureClass.Fields.FindField(attributeName)
            If fieldIndex < 0 Then Throw New ArgumentException("Attribuut niet gevonden.", "attributeName")
            Dim rangeValue As Boolean() 'store which values are detected

            'Initial values.
            Dim currentValue As Integer = 0
            Dim highestValue As Integer = 0
            ReDim rangeValue(0)
            rangeValue(0) = True 'register value 0 as true so that it is never returned

            'Loop all attribute values.
            While Not feature Is Nothing
                'Determine attribute value of current feature in the loop.
                If TypeOf feature.Value(fieldIndex) Is System.DBNull Then
                    currentValue = 0
                ElseIf Trim(CStr(feature.Value(fieldIndex))) = "" Then
                    currentValue = 0
                Else
                    currentValue = CInt(feature.Value(fieldIndex))
                End If
                'Determine current highest attribute value.
                highestValue = MaxInt(highestValue, currentValue)
                If fillGaps Then
                    'Expand range of values.
                    While currentValue > rangeValue.Length - 1
                        ReDim Preserve rangeValue(rangeValue.Length)
                        rangeValue(rangeValue.Length - 1) = False
                    End While
                    'Store current atribute value in the range of values.
                    rangeValue(currentValue) = True
                End If
                'Loop to next feature.
                feature = featureCursor.NextFeature
            End While

            'Return first missing value from the range.
            If fillGaps Then
                For i As Integer = 0 To rangeValue.Length - 1
                    If rangeValue(i) = False Then
                        Return i
                        Exit Function
                    End If
                Next
            End If

            'Return highest + 1.
            Return highestValue + 1
            Exit Function

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Return the biggest of 2 integer.
    ''' </summary>
    ''' <param name="FirstInt">integer</param>
    ''' <param name="SecondInt">integer</param>
    ''' <returns>integer</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Function MaxInt( _
            ByVal FirstInt As Integer, _
            ByVal SecondInt As Integer _
            ) As Integer

        If SecondInt > FirstInt Then
            MaxInt = SecondInt
        Else
            MaxInt = FirstInt
        End If
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Return the feature layer with the specified name.
    ''' </summary>
    ''' <param name="pMap">
    '''     The map you are working with.
    ''' </param>
    ''' <param name="LayerName">
    '''     The name of the feature layer you are looking for.
    ''' </param>
    ''' <returns>
    '''     The feature layer with the spacified name.
    '''     Nothing if the specified name could not be found.
    ''' </returns>
    ''' <remarks>
    '''     This function calls another function with the same name, but different parameters.
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function GetFeatureLayer( _
            ByVal pMap As IMap, _
            ByVal LayerName As String _
            ) As IFeatureLayer

        GetFeatureLayer = Nothing
        Try
            Dim i As Integer
            Dim pLayer As ILayer
            Dim pMatchingLayer As IFeatureLayer

            'Loop through the map legend.
            For i = 0 To pMap.LayerCount - 1
                pLayer = pMap.Layer(i)
                pMatchingLayer = GetFeatureLayer(pLayer, LayerName)
                If Not pMatchingLayer Is Nothing Then
                    GetFeatureLayer = pMatchingLayer
                    Exit Function
                End If
            Next

        Catch ex As Exception
            Throw ex
        End Try

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Return the feature layer with the specified name.
    ''' </summary>
    ''' <param name="pLayer"></param>
    ''' <param name="LayerName">
    '''     The name of the feature layer you are looking for.
    ''' </param>
    ''' <returns>
    '''     The feature layer with the specified name.
    '''     Nothing if the specified name could not be found.
    ''' </returns>
    ''' <remarks>
    '''     This function is called another function with the same name, but different parameters.
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    ''' 	[Kristof Vydt]	23/09/2005	Use case-insensitive layername comparison.
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Function GetFeatureLayer( _
            ByVal pLayer As ILayer, _
            ByVal LayerName As String _
            ) As IFeatureLayer

        GetFeatureLayer = Nothing
        Try
            Dim i As Integer 'loop index
            Dim pCompositeLayer As ICompositeLayer
            Dim pSubLayer As ILayer
            Dim pMatchingLayer As IFeatureLayer

            'Loop within layer groups.
            If TypeOf pLayer Is IGroupLayer Then
                pCompositeLayer = CType(pLayer, ICompositeLayer)
                For i = 0 To pCompositeLayer.Count - 1
                    pSubLayer = pCompositeLayer.Layer(i)
                    pMatchingLayer = GetFeatureLayer(pSubLayer, LayerName)
                    If Not pMatchingLayer Is Nothing Then
                        GetFeatureLayer = pMatchingLayer
                        Exit Function
                    End If
                Next
            End If

            'Return the input layer if it's a feature layer with matching name.
            If TypeOf pLayer Is IFeatureLayer Then
                If LCase(pLayer.Name) = LCase(LayerName) Then
                    GetFeatureLayer = CType(pLayer, IFeatureLayer)
                End If
            End If

        Catch ex As Exception
            Throw ex
        End Try

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Activate an ArcMap tool from a toolbar.
    ''' </summary>
    ''' <param name="pDocument">
    '''     The ArcGIS document you're working with.
    ''' </param>
    ''' <param name="ToolProgID">
    '''     The ProgID of the tool you want to activate.
    ''' </param>
    ''' <remarks>
    '''     To find the ProgID of a built-in command, menu, or toolbar in ArcMap,
    '''     refer to the following technical document: 
    '''     "Captions, names, and GUIDs of built-in commands, menus, and toolbars in ArcMap"
    '''     http://edndoc.esri.com/arcobjects/9.0/ArcGISDevHelp/TechnicalDocuments/Guids/ArcMapIds.htm
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub ActivateTool( _
            ByVal pDocument As IDocument, _
            ByVal ToolProgID As String)

        Try
            Dim pUID As UID
            Dim pCommandItem As ICommandItem

            pUID = New UID
            pUID.Value = ToolProgID
            pCommandItem = pDocument.CommandBars.Find(pUID)
            If Not (pCommandItem Is Nothing) Then _
            pCommandItem.Execute()

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    '''' -----------------------------------------------------------------------------
    '''' <summary>
    ''''     Determine the path of the configuration file 
    ''''     and store it in a global variable.
    '''' </summary>
    '''' <param name="pApplication">
    ''''     The ArcGIS application object you're working with.
    '''' </param>
    '''' <remarks>
    ''''     Throws a ConfigFileNotFoundException if the configuration file could not be found.
    '''' </remarks>
    '''' <history>
    '''' 	[Kristof Vydt]	16/09/2005	Created
    '''' 	[Kristof Vydt]	08/03/2007	Deprecated
    '''' </history>
    '''' -----------------------------------------------------------------------------
    'Public Function GetConfigFilePath( _
    '    ByVal pApplication As ESRI.ArcGIS.Framework.IApplication _
    '    ) As String

    '    Try

    '        Dim pTemplates As ESRI.ArcGIS.Framework.ITemplates
    '        Dim mxdIndex As Integer
    '        Dim filePath As String
    '        Dim folderPath As String

    '        'Get the location of the current mxd file.
    '        'This is the last one in the templates collection.
    '        pTemplates = pApplication.Templates
    '        mxdIndex = pTemplates.Count - 1
    '        filePath = pTemplates.Item(mxdIndex).ToString

    '        'Extract the mxd folder path.
    '        folderPath = System.IO.Path.GetDirectoryName(filePath)

    '        'The configuration file is located in the same folder as the mxd.
    '        filePath = folderPath & "\" & c_FileName_Config

    '        'Throw an exception if the configuration file does not exist.
    '        If System.IO.File.Exists(filePath) Then
    '            GetConfigFilePath = filePath
    '        Else
    '            Throw New ConfigFileNotFoundException
    '        End If

    '    Catch ex As Exception
    '        Throw ex

    '    End Try

    'End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Returns a reference to the editor extension.
    ''' </summary>
    ''' <param name="mxAppl">
    '''     ArcMap application to get the editor extension from.
    ''' </param>
    ''' <returns>
    '''     Editor
    ''' </returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    ''' 	[Kristof Vydt]	23/09/2005	Changed parameter from IApplication to IMxApplication.
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function GetEditorReference( _
            ByVal mxAppl As IMxApplication _
            ) As IEditor2

        Try
            Dim pApplication As IApplication 'ArcGIS application object
            Dim pEditor As IEditor2 'Editor object
            Dim pUID As UID

            pUID = New UID
            pUID.Value = "esriEditor.Editor"
            pApplication = CType(mxAppl, IApplication)
            pEditor = CType(pApplication.FindExtensionByCLSID(pUID), IEditor2)

            Return pEditor

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Starts an edit session of the specified feature layer.
    ''' </summary>
    ''' <param name="pEditor">
    '''     The editor pointer to use.
    ''' </param>
    ''' <param name="pFeatureLayer">
    '''     The feature layer to edit.
    ''' </param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    ''' 	[Kristof Vydt]	23/09/2005	Edit session management reviewed.
    ''' 	[Kristof Vydt]	17/07/2006	Use global parameter for message box prompt text.
    ''' 	[Elton Manoku]	23/07/2008	Added parameter startAlsoEditOperation because
    '''                                 in the case of importing excel data, it is not required to have a start operation
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub EditSessionStart( _
        ByVal pEditor As IEditor2, _
        ByVal pFeatureLayer As IFeatureLayer, _
        ByVal startAlsoEditOperation As Boolean)

        Dim pDataset As IDataset 'The dataset of the requested layer.
        Dim pReqWorkspace As IWorkspace 'The workspace of the requiested layer.
        '    Dim pCurWorkspace As IWorkspace 'The worspace that is currently being edited.

        Try
            'Stop an edit session in progress.
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

            'Determine the edit workspace.
            pDataset = CType(pFeatureLayer.FeatureClass, IDataset)
            pReqWorkspace = pDataset.Workspace

            'Start editing the requiested workspace.
            pEditor.StartEditing(pReqWorkspace)

            If startAlsoEditOperation Then
                'Start operation.
                pEditor.StartOperation()
            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Terminate an edit session while storing the changes.
    ''' </summary>
    ''' <param name="pEditor">
    '''     The editor object.
    ''' </param>
    ''' <param name="menuText">
    '''     Optional text to add to the undo list.
    ''' </param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	23/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub EditSessionSave( _
        ByVal pEditor As IEditor2, _
        Optional ByVal menuText As String = "Edit session")
        Try
            pEditor.StopOperation(menuText) 'Stop operation.
            pEditor.StopEditing(True) 'Save changes.
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Terminate an edit session without storing the changes.
    ''' </summary>
    ''' <param name="pEditor">
    '''     The editor object.
    ''' </param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	23/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub EditSessionAbort( _
            ByVal pEditor As IEditor2)
        Try
            pEditor.AbortOperation() 'Abort operation.
            pEditor.StopEditing(False) 'Save changes.
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Query a feature layer, returning a feature cursor.
    ''' </summary>
    ''' <param name="pFeatureLayer">
    '''     The feature layer that is to be queried.
    ''' </param>
    ''' <param name="searchGeometry">
    '''     The geometry of the spatial filter.
    ''' </param>
    ''' <param name="spatialRelation">
    '''     The spatial relation of the geometry.
    ''' </param>
    ''' <param name="whereClause">
    '''     [optional] Attribute WhereClause.
    ''' </param>
    ''' <returns>
    '''     Cursor of feature.
    ''' </returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function SearchFeatureLayer( _
            ByVal pFeatureLayer As IFeatureLayer, _
            ByVal searchGeometry As IGeometry, _
            ByVal spatialRelation As esriSpatialRelEnum, _
            Optional ByVal whereClause As String = "" _
            ) As IFeatureCursor

        Try
            Dim pSpatialFilter As ISpatialFilter
            Dim pFeatureClass As IFeatureClass
            Dim pFeatureCursor As IFeatureCursor

            'Create a spatial query filter.
            pSpatialFilter = New SpatialFilter

            'Set spatial filter properties.
            pFeatureClass = pFeatureLayer.FeatureClass
            pSpatialFilter.GeometryField = pFeatureClass.ShapeFieldName
            pSpatialFilter.Geometry = searchGeometry
            pSpatialFilter.SpatialRel = esriSpatialRelEnum.esriSpatialRelRelation
            pSpatialFilter.WhereClause = whereClause

            'Perform the query and get the resulting cursor.
            pFeatureCursor = pFeatureLayer.Search(pSpatialFilter, True)

            'Return cursor.
            SearchFeatureLayer = pFeatureCursor

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Split the envelope of a feature, into 4 equal kwadrants.
    ''' </summary>
    ''' <param name="pFeature">
    '''     [in] The feature of which the envelope is derived.
    ''' </param>
    ''' <param name="pClip">
    '''     [in] Each kwadrant is clipped to this geometry.
    ''' </param>
    ''' <param name="arrayKwadrant">
    '''     [out] The array of 4 kwadrants.
    ''' </param>
    ''' <remarks>
    '''     The order of the kwadrants in the array is A-B-C-D.
    '''     Set pClip = Nothing if no clipping is required.
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    ''' 	[Elton Manoku]	28/11/2008	RW:2008 New shapes that are generated are simplified 
    '''                                 and are set with the spatial reference of the raster features.
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub SplitFeatureEnvelopeIntoKwadrants( _
            ByVal pFeature As IFeature, _
            ByVal pClip As IGeometry, _
            ByRef arrayKwadrant As IGeometry())

        Try

            'Envelope of geometry.
            Dim pEnvelope As IEnvelope
            pEnvelope = New EnvelopeClass
            pFeature.Shape.QueryEnvelope(pEnvelope)
            Dim spatialReference As ISpatialReference = pFeature.Shape.SpatialReference
            'Calculate border and center coordinates.
            Dim minX As Double
            Dim minY As Double
            Dim midX As Double
            Dim midY As Double
            Dim maxX As Double
            Dim maxY As Double
            pEnvelope.QueryCoords(minX, minY, maxX, maxY)
            midX = (minX + maxX) / 2
            midY = (minY + maxY) / 2

            'Build a set of corner points.
            'pts(0) is not used ! As a result, 
            'the numeric keyboard layout equals the points configuration.
            Dim pts As IPoint()
            ReDim pts(9)
            Dim i As Integer
            For i = 1 To 9
                pts(i) = New Point
            Next
            pts(1).PutCoords(minX, minY)
            pts(2).PutCoords(midX, minY)
            pts(3).PutCoords(maxX, minY)
            pts(4).PutCoords(minX, midY)
            pts(5).PutCoords(midX, midY)
            pts(6).PutCoords(maxX, midY)
            pts(7).PutCoords(minX, maxY)
            pts(8).PutCoords(midX, maxY)
            pts(9).PutCoords(maxX, maxY)

            'Reallocate kwadrants.
            Dim pKwadrant As IPointCollection
            'Dim pKwadrant As IPolygon
            Dim pOperator As ITopologicalOperator2
            Dim pIntersection As IGeometry
            ReDim arrayKwadrant(3)

            'First kwadrant [A| ]
            '               [ | ]
            pKwadrant = New Polygon
            pKwadrant.AddPoint(pts(4))
            pKwadrant.AddPoint(pts(7))
            pKwadrant.AddPoint(pts(8))
            pKwadrant.AddPoint(pts(5))
            pKwadrant.AddPoint(pts(4))
            'RW:2008 
            pOperator = CType(GetSimplifiedGeometry(pKwadrant, spatialReference), ITopologicalOperator2)
            If Not pClip Is Nothing Then

                'RW:2008 
                'pOperator = CType(pKwadrant, ITopologicalOperator2)
                pIntersection = pOperator.Intersect(pClip, esriGeometryDimension.esriGeometry2Dimension)
                arrayKwadrant(0) = pIntersection
            Else
                arrayKwadrant(0) = CType(pOperator, IPolygon)
            End If

            'Second kwadrant [ |B]
            '                [ | ]
            pKwadrant = New Polygon
            pKwadrant.AddPoint(pts(5))
            pKwadrant.AddPoint(pts(8))
            pKwadrant.AddPoint(pts(9))
            pKwadrant.AddPoint(pts(6))
            pKwadrant.AddPoint(pts(5))
            'RW:2008 
            pOperator = CType(GetSimplifiedGeometry(pKwadrant, spatialReference), ITopologicalOperator2)

            If Not pClip Is Nothing Then
                'RW:2008 
                'pOperator = CType(pKwadrant, ITopologicalOperator2)
                pIntersection = pOperator.Intersect(pClip, esriGeometryDimension.esriGeometry2Dimension)
                arrayKwadrant(1) = pIntersection
            Else
                arrayKwadrant(1) = CType(pOperator, IPolygon)
            End If

            'Third kwadrant [ | ]
            '               [ |C]
            pKwadrant = New Polygon
            pKwadrant.AddPoint(pts(2))
            pKwadrant.AddPoint(pts(5))
            pKwadrant.AddPoint(pts(6))
            pKwadrant.AddPoint(pts(3))
            pKwadrant.AddPoint(pts(2))
            'RW:2008 
            pOperator = CType(GetSimplifiedGeometry(pKwadrant, spatialReference), ITopologicalOperator2)
            If Not pClip Is Nothing Then
                'RW:2008 
                'pOperator = CType(pKwadrant, ITopologicalOperator2)
                pIntersection = pOperator.Intersect(pClip, esriGeometryDimension.esriGeometry2Dimension)
                arrayKwadrant(2) = pIntersection
            Else
                arrayKwadrant(2) = CType(pOperator, IPolygon)
            End If

            'Fourth kwadrant [ | ]
            '                [D| ]
            pKwadrant = New Polygon
            pKwadrant.AddPoint(pts(1))
            pKwadrant.AddPoint(pts(4))
            pKwadrant.AddPoint(pts(5))
            pKwadrant.AddPoint(pts(2))
            pKwadrant.AddPoint(pts(1))
            'RW:2008 
            pOperator = CType(GetSimplifiedGeometry(pKwadrant, spatialReference), ITopologicalOperator2)

            If Not pClip Is Nothing Then
                'RW:2008 
                'pOperator = CType(pKwadrant, ITopologicalOperator2)
                pIntersection = pOperator.Intersect(pClip, esriGeometryDimension.esriGeometry2Dimension)
                arrayKwadrant(3) = pIntersection
            Else
                arrayKwadrant(3) = CType(pOperator, IPolygon)
            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    ''' <summary>
    ''' 
    ''' This function gets a geometry and makes sure that it gets the spatial reference 
    ''' and it is simplyfied. Must be done if the geometry will be used in topological operations
    ''' </summary>
    ''' <param name="geom"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' <history>
    '''      [Elton Manoku] 28/11/2008 RW:2008 Created
    ''' </history>
    Public Function GetSimplifiedGeometry(ByVal pointCollection As IPointCollection, ByVal spatialRef As ISpatialReference) As IGeometry
        Dim pOperator As ITopologicalOperator2 = CType(pointCollection, ITopologicalOperator2)
        pOperator.IsKnownSimple_2 = False
        pOperator.Simplify()
        Dim geom As IGeometry = CType(pOperator, IGeometry)
        geom.SpatialReference = spatialRef
        geom.SnapToSpatialReference()
        GetSimplifiedGeometry = geom
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Split some text into an array of strings.
    ''' </summary>
    ''' <param name="input">
    '''     The text to split.
    ''' </param>
    ''' <param name="separator">
    '''     The text that indicates the separation between the parts of text.
    ''' </param>
    ''' <param name="trimming">
    '''     By setting this argument to true, 
    '''     spaces at the beginning and at the end of each part of text, are removed.
    ''' </param>
    ''' <returns>
    '''     An array of strings.
    ''' </returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function Split2( _
            ByVal input As String, _
            ByVal separator As String, _
            ByVal trimming As Boolean _
            ) As String()

        Try
            Dim i As Integer 'loop index
            Dim output As String() 'output arrayof strings

            output = Split(input, separator)
            If trimming Then
                For i = output.GetLowerBound(0) To output.GetUpperBound(0)
                    output(i) = Trim(output(i).ToString)
                Next
            End If
            Return output

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Return the selected feature from selectionset.
    ''' </summary>
    ''' <param name="SelectionSet">
    '''     The selectionset pointer.
    ''' </param>
    ''' <param name="SelectedIndex">
    '''     The index of the feature in the selectionset.
    ''' </param>
    ''' <returns>
    '''     The indexed feature.
    ''' </returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    '''     [Elton Manoku]  27/11/2008 Added finally statement to clear memory
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function GetSelectedFeature( _
            ByVal SelectionSet As ISelectionSet, _
            ByVal SelectedIndex As Integer _
            ) As IFeature

        Dim LoopIndex As Integer
        Dim pCursor As ICursor = Nothing
        Dim pFCursor As IFeatureCursor = Nothing
        Dim pFeature As IFeature = Nothing
        Try

            If SelectionSet Is Nothing Then Return Nothing 'no selectionset
            If Not SelectedIndex > -1 Then Return Nothing 'not a valid index

            'Loop through selectionset.
            SelectionSet.Search(Nothing, False, pCursor)
            pFCursor = CType(pCursor, IFeatureCursor)
            pFeature = pFCursor.NextFeature
            If pFeature Is Nothing Then Return Nothing 'empty selectionset
            LoopIndex = 0
            While Not pFeature Is Nothing
                'Exit loop if requested index is located.
                If LoopIndex = SelectedIndex Then Exit While
                pFeature = pFCursor.NextFeature
                LoopIndex = LoopIndex + 1
            End While
            'Result
            If LoopIndex <> SelectedIndex Then Return Nothing 'requested index not found
            Return pFeature

        Catch ex As Exception
            Throw ex
        Finally
            'RW:2008
            If Not pFCursor Is Nothing Then Marshal.ReleaseComObject(pFCursor)
            If Not pCursor Is Nothing Then Marshal.ReleaseComObject(pCursor)
        End Try

    End Function

    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Let a feature flash on the map.
    ''' </summary>
    ''' <param name="pFeature">
    '''     The feature you want to see flashing on the map.
    ''' </param>
    ''' <param name="pMxDoc">
    '''     The document you are working in.
    ''' </param>
    ''' <param name="FlashTimes">
    '''     [optional] Number of times the feature should flash.
    ''' </param>
    ''' <param name="Interval">
    '''     [optional] The interval in milliseconds between two flashes.
    ''' </param>
    ''' <remarks>
    '''     Sub Sleep Lib "kernel32" required.
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub FlashFeature( _
            ByVal pFeature As IFeature, _
            ByVal pMxDoc As IMxDocument, _
            Optional ByVal FlashTimes As Integer = 3, _
            Optional ByVal Interval As Integer = 300)

        Try

            ' Create identify object using the feature.
            Dim pRowIdentifyObj As FeatureIdentifyObject = New FeatureIdentifyObject
            pRowIdentifyObj.Feature = pFeature
            Dim pIdentifyObj As IIdentifyObj = CType(pRowIdentifyObj, IIdentifyObj)

            ' Flash the feature several times.
            For i As Integer = 0 To FlashTimes - 1
                pIdentifyObj.Flash(pMxDoc.ActiveView.ScreenDisplay)
                Application.DoEvents()
                Sleep(Interval)
            Next i

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Change the casing of a string:
    '''     only first character of each word in uppercase,
    '''     allt the rest in lowercase.
    ''' </summary>
    ''' <param name="inputText">
    '''     The string we want to modify.
    ''' </param>
    ''' <returns>
    '''     The modified string.
    ''' </returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function MixedCasing(ByVal inputText As String) As String
        Try
            Dim position As Integer 'character position
            Dim outputText As String = Nothing 'building the return value
            Dim character As Char 'one character
            Dim isNewWord As Boolean = True 'indicates begin of a new word
            If inputText Is Nothing Then Return ""
            If Len(inputText) = 0 Then Return ""
            For position = 0 To Len(inputText) - 1
                character = CChar(inputText.Substring(position, 1))
                If isNewWord Then
                    outputText &= UCase(character)
                Else
                    outputText &= LCase(character)
                End If
                isNewWord = (character = " "c)
            Next
            Return outputText
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Return the current MapBook object.
    ''' </summary>
    ''' <param name="pApp">
    '''     The current ArcGIS application.
    ''' </param>
    ''' <returns>
    '''     The current MapBook object.
    ''' </returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	28/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function GetMapBook(ByRef pApp As ESRI.ArcGIS.Framework.IApplication) As DSMapBookPrj.IDSMapBook
        Dim pMapBookExt As IExtension 'DSMapBookUIPrj.DSMapBookExt
        Dim pMapBook As DSMapBookPrj.IDSMapBook
        Try
            pMapBookExt = pApp.FindExtensionByName("DevSample_MapBook")
            If pMapBookExt Is Nothing Then
                MsgBox("Map Book code not installed properly!!", , "Map Book Extension Not Found!!!")
                Return Nothing
                Exit Function
            End If
            pMapBook = CType(CType(pMapBookExt, DSMapBookUIPrj.DSMapBookExt).MapBook, DSMapBookPrj.IDSMapBook)
            Return pMapBook
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Modify the visibility of feature layers in the ArcMap document
    ''' </summary>
    ''' <param name="pMxDoc">
    '''     The current ArcMap document.
    ''' </param>
    ''' <param name="layerNames">
    '''     An string array of feature layer names.
    ''' </param>
    ''' <param name="setVisible">
    '''     Boolean to set layer visibility.
    ''' </param>
    ''' <param name="forceOthers">
    '''     Boolean to force feature layers not present in layerNames
    '''     to visibility opposite to setVisible.
    ''' </param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	28/09/2005	Created
    ''' 	[Kristof Vydt]	29/09/2005	Refresh map and TOC at the end.
    '''                                 Convert layernames in array to lowercase.
    ''' 	[Kristof Vydt]	26/09/2005	Extract the core of the method to another seperate method SetLayerVisibility.
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub SetLayerVisibility( _
        ByVal pMxDoc As IMxDocument, _
        ByVal layerNames As String(), _
        ByVal setVisible As Boolean, _
        ByVal forceOthers As Boolean)

        Dim enumLayer As IEnumLayer
        Dim pLayer As ILayer
        '    Dim layerName As String

        Try
            enumLayer = pMxDoc.FocusMap.Layers
            pLayer = enumLayer.Next

            'Convert layernames array to all lowercase.
            For i As Integer = 0 To layerNames.Length - 1
                layerNames(i) = LCase(layerNames(i))
            Next

            'Loop through all map layers.
            While Not pLayer Is Nothing
                SetLayerVisibility(pLayer, layerNames, setVisible, forceOthers)
                pLayer = enumLayer.Next
            End While

            'Refresh active map and table of contents.
            pMxDoc.UpdateContents()
            pMxDoc.ActiveView.Refresh()

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Modify the visibility of feature layers in the specified layer
    ''' </summary>
    ''' <param name="pLayer">
    '''     Some layer pointer.
    ''' </param>
    ''' <param name="layerNames">
    '''     An string array of LOWERCASE (!) feature layer names.
    ''' </param>
    ''' <param name="setVisible">
    '''     Boolean to set layer visibility.
    ''' </param>
    ''' <param name="forceOthers">
    '''     Boolean to force feature layers not present in layerNames
    '''     to visibility opposite to setVisible.
    ''' </param>
    ''' <remarks>
    '''     Only feature layers or group layers are recognised.
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	26/10/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub SetLayerVisibility( _
            ByVal pLayer As ILayer, _
            ByVal layerNames As String(), _
            ByVal setVisible As Boolean, _
            ByVal forceOthers As Boolean)

        Try

            'Group layer...
            If TypeOf pLayer Is IGroupLayer Then
                'Force group layer visibility.
                pLayer.Visible = setVisible
                'Access the composite layers.
                'Dim pCompositeLayer As ICompositeLayer = CType(pLayer, ICompositeLayer)
                'For i As Integer = 0 To pCompositeLayer.Count - 1
                '    'Seperate method call for each sublayer.
                '    SetLayerVisibility(pCompositeLayer.Layer(i), layerNames, setVisible, forceOthers)
                'Next
                '--> There is no need to process the composite layers
                '    because this method is called for each layer individually.
            End If

            'Feature layer...
            If TypeOf pLayer Is IFeatureLayer Then
                'Compare lowercased layername with list layerNames.
                If System.Array.IndexOf(layerNames, LCase(pLayer.Name)) > -1 Then
                    'Adjust visibility for feature layers mentioned in the array.
                    pLayer.Visible = setVisible
                ElseIf forceOthers Then
                    'Force visibility for feature layers not mentioned in the array.
                    pLayer.Visible = Not setVisible
                End If
            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Convert the input string for use in an SQL WhereClause statement.
    ''' </summary>
    ''' <param name="sSource"></param>
    ''' <returns></returns>
    ''' <remarks>
    '''     The characters that must be escaped, and the escape character,
    '''     depending on the type of database (i.e. SQL dialect).
    '''     For Access, single quotes should be replaced by 2 single quotes.
    '''     Example: 
    '''         pQueryFilter.WhereClause = [FieldName] &amp; " = " &amp; CStrSql("BSG Oil' LTD")  
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	12/07/2006	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function CStrSql(ByVal sSource As String) As String

        CStrSql = "'" + Replace(sSource, "'", "''") + "'"

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Return the first letter of the input string.
    ''' </summary>
    ''' <param name="sSource">Input string</param>
    ''' <param name="bAllowLowerCase">Allow returning lowercase</param>
    ''' <param name="bAllowUpperCase">Allow returning uppercase</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	12/07/2006	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function GetFirstLetter( _
            ByVal sSource As String, _
            Optional ByVal bAllowLowerCase As Boolean = True, _
            Optional ByVal bAllowUpperCase As Boolean = True) As Char

        Dim pos As Integer 'character position in the text string
        Dim letter As Char 'character at pos
        Const upper As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        Const lower As String = "abcdefghijklmnopqrstuvwxyz"

        Try
            For pos = 0 To sSource.Length - 1
                letter = sSource.Chars(pos)
                If bAllowLowerCase And (lower.IndexOf(letter) > -1) Then
                    Return letter 'return the first lowercase character of the input text
                    Exit Function
                ElseIf bAllowUpperCase And (upper.IndexOf(letter) > -1) Then
                    Return letter 'return the first uppercase character of the input text
                    Exit Function
                End If
            Next
        Catch ex As Exception
            Throw ex
        End Try

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Get a (standalone) table from a personal geodatabase.
    ''' </summary>
    ''' <param name="tableName">
    '''     The name of the table to return.
    ''' </param>
    ''' <param name="pMap">
    '''     The current map.
    ''' </param>
    ''' <returns>
    '''     The table with the specified name.
    ''' </returns>
    ''' <remarks>
    '''     Nothing is returned if no table with specified name is found.
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    '''     [Kristof Vydt]  13/07/2006  Moved from FormIndexStraten to here, making it public.
    ''' 	[Kristof Vydt]	02/08/2006	Replace some global constant by a function that reads the value from config file.
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function GetTable( _
            ByVal tableName As String, _
            ByVal pMap As IMap _
            ) As ITable

        Try
            '    Dim pTable As ITable
            Dim pFLayer As IFeatureLayer
            Dim pWorkspace As IWorkspace
            Dim pEnumDataset As IEnumDataset
            Dim pDataset As IDataset

            'Determine the workspace of current sector.
            pFLayer = GetFeatureLayer(pMap, GetLayerName("Hydrant"))
            If pFLayer Is Nothing Then Throw New LayerNotFoundException(GetLayerName("Hydrant"))
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
    '''     Return the exact layer name from configuration file.
    ''' </summary>
    ''' <param name="layer">The layer codename.</param>
    ''' <returns>The exact layer name from the configuration file.</returns>
    ''' <remarks>
    '''     Values are cached in private hashtable m_LayerNames.
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	02/08/2006	Created
    '''     [Kristof Vydt]  22/02/2007  Use XML-based configuration.
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function GetLayerName(ByVal layer As String) As String

        ' Check availability of configuration.
        If Config Is Nothing Then Throw New ApplicationException("No configuration loaded.")

        ' Return configured layer name.
        Return Config.LayerName(layer)

        'Dim hashKey As String 'hashtable key
        'Dim hashValue As String 'hashtable value

        'Try

        '    hashKey = layer
        '    If m_LayerNames Is Nothing Then m_LayerNames = New Hashtable
        '    If m_LayerNames.ContainsKey(hashKey) Then
        '        'Read from hashtable.
        '        hashValue = CStr(m_LayerNames.Item(hashKey))
        '    Else
        '        'Add to hashtable.
        '        hashValue = INIRead(g_FilePath_Config, "LayerNames", hashKey)
        '        m_LayerNames.Add(hashKey, hashValue)
        '    End If
        '    GetLayerName = hashValue

        'Catch ex As Exception
        '    Throw ex
        'End Try

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Return the exact attribute name from configuration file.
    ''' </summary>
    ''' <param name="layer">The layer key.</param>
    ''' <param name="attribute">The attribute key.</param>
    ''' <returns>The exact attribute name from the configuration file.</returns>
    ''' <remarks>
    '''     Values are cached in private hashtable m_AttributeNames.
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	2/08/2006	Created
    '''     [Kristof Vydt]  22/02/2007  Use XML-based configuration.
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function GetAttributeName(ByVal layer As String, ByVal attribute As String) As String

        ' Check availability of configuration.
        If Config Is Nothing Then Throw New ApplicationException("No configuration loaded.")

        ' Return configured layer name.
        Return Config.AttributeName(layer, attribute)

        'Dim hashKey As String 'hashtable key
        'Dim hashValue As String 'hashtable value

        'Try

        '    hashKey = layer & "_" & attribute
        '    If m_AttributeNames Is Nothing Then m_AttributeNames = New Hashtable
        '    If m_AttributeNames.ContainsKey(hashKey) Then
        '        'Read from hashtable.
        '        hashValue = CStr(m_AttributeNames.Item(hashKey))
        '    Else
        '        'Add to hashtable.
        '        hashValue = INIRead(g_FilePath_Config, "AttributeNames", hashKey)
        '        m_AttributeNames.Add(hashKey, hashValue)
        '    End If
        '    GetAttributeName = hashValue

        'Catch ex As Exception
        '    Throw ex
        'End Try

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Return the exact domain name from configuration file.
    ''' </summary>
    ''' <param name="domain">The domain codename.</param>
    ''' <returns>The exact domain name from the configuration file.</returns>
    ''' <remarks>
    '''     Values are cached in private hashtable m_DomainNames.
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	2/08/2006	Created
    '''     [Kristof Vydt]  22/02/2007  Use XML-based configuration.
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function GetDomainName(ByVal domain As String) As String

        ' Check availability of configuration.
        If Config Is Nothing Then Throw New ApplicationException("No configuration loaded.")

        ' Return configured domain name.
        Return Config.DomainName(domain)

        'Dim hashKey As String 'hashtable key
        'Dim hashValue As String 'hashtable value

        'Try

        '    hashKey = domain
        '    If m_DomainNames Is Nothing Then m_DomainNames = New Hashtable
        '    If m_DomainNames.ContainsKey(hashKey) Then
        '        'Read from hashtable.
        '        hashValue = CStr(m_DomainNames.Item(hashKey))
        '    Else
        '        'Add to hashtable.
        '        hashValue = INIRead(g_FilePath_Config, "DomainNames", hashKey)
        '        m_DomainNames.Add(hashKey, hashValue)
        '    End If
        '    GetDomainName = hashValue

        'Catch ex As Exception
        '    Throw ex
        'End Try

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Return first marker element with specified name on focus map.
    ''' </summary>
    ''' <param name="markerName">Name of the marker.</param>
    ''' <param name="mxDoc">Current ArcMap document.</param>
    ''' <returns>Marker element or Nothing.</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	18/08/2006	Created
    ''' 	[Kristof Vydt]	29/09/2006	Get next element in container if first one doesn't fit.
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function GetMarkerElement( _
        ByVal markerName As String, _
        ByRef mxDoc As IMxDocument) _
        As IMarkerElement

        Dim pElement As IElement
        Dim pGraphics As IGraphicsContainer
        Dim pMarkerElement As IMarkerElement = Nothing

        Try
            pGraphics = CType(mxDoc.FocusMap, IGraphicsContainer)
            pGraphics.Reset()
            pElement = pGraphics.Next
            While Not pElement Is Nothing
                If CType(pElement, IElementProperties).Name = "BRANDWEER" Then
                    pMarkerElement = CType(pElement, IMarkerElement)
                    Exit While
                End If
                pElement = pGraphics.Next
            End While
            Return pMarkerElement

        Catch ex As Exception
            Throw ex
        End Try

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Get the workspace of a feature layer.
    ''' </summary>
    ''' <param name="mxAppl">ArcMap application</param>
    ''' <param name="featureLayer">Feature layer that is used in the ArcMap application</param>
    ''' <returns>Workspace of the feature layer</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	8/09/2006	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Function GetLayerWorkspace( _
            ByRef mxAppl As IMxApplication, _
            ByRef featureLayer As IFeatureLayer _
            ) As IWorkspace

        Try

            Dim pArcGisApplication As IApplication = CType(mxAppl, IApplication)
            Dim pArcMapDocument As IMxDocument = CType(pArcGisApplication.Document, IMxDocument)
            Dim pFeatureDataset As IDataset = CType(featureLayer.FeatureClass, IDataset)
            GetLayerWorkspace = pFeatureDataset.Workspace

        Catch ex As Exception
            Throw ex
        End Try

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Read a single feature attribute value.
    ''' </summary>
    ''' <param name="feature">The feature to read from.</param>
    ''' <param name="codedLayerName">The feature layer name as used in the configuration file.</param>
    ''' <param name="codedAttributeName">The feature attribute name as used in the configuration file.</param>
    ''' <returns>The attribute value object.</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	22/09/2006	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function GetAttributeValue( _
            ByVal feature As IFeature, _
            ByVal codedLayerName As String, _
            ByVal codedAttributeName As String _
            ) As Object

        GetAttributeValue = Nothing
        ' Abort if no feature.
        If feature Is Nothing Then Exit Function

        ' Determine the attribute field index.
        Dim realAttributeName As String = GetAttributeName(codedLayerName, codedAttributeName)
        If realAttributeName = "" Then Exit Function
        Dim attributeIndex As Integer = feature.Fields.FindField(realAttributeName)

        ' Abort if attribute not found.
        If attributeIndex < 0 Then Exit Function

        ' Return attribute value object.
        Return feature.Value(attributeIndex)

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Overwrite a single feature attribute value.
    ''' </summary>
    ''' <param name="feature">The feature to be updated.</param>
    ''' <param name="codedLayerName">The feature layer name as used in the configuration file.</param>
    ''' <param name="codedAttributeName">The feature attribute name as used in the configuration file.</param>
    ''' <param name="newValue">The new feature attribute value object.</param>
    ''' <returns>The actual new feature attribute value object.</returns>
    ''' <remarks>
    '''     The required edit session management is not included in this method.
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	22/09/2006	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function SetAttributeValue( _
            ByVal feature As IFeature, _
            ByVal codedLayerName As String, _
            ByVal codedAttributeName As String, _
            ByVal newValue As Object _
            ) As Object

        SetAttributeValue = Nothing
        ' Abort if no feature.
        If feature Is Nothing Then Exit Function

        ' Determine the attribute field index.
        Dim realAttributeName As String = GetAttributeName(codedLayerName, codedAttributeName)
        If realAttributeName = "" Then Exit Function
        Dim attributeIndex As Integer = feature.Fields.FindField(realAttributeName)

        ' Abort if attribute not found.
        If attributeIndex < 0 Then Exit Function

        ' Update attribute value.
        feature.Value(attributeIndex) = newValue

        ' Write changes to database.
        feature.Store()

        ' Return attribute value object.
        Return feature.Value(attributeIndex)

    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Force layers visibility if current focus map.
    ''' </summary>
    ''' <param name="mxDoc">ArcMap document</param>
    ''' <param name="layerName">Name of the layer</param>
    ''' <param name="visible">Boolean</param>
    ''' <remarks>
    ''' Current TOC is updated when layer visibility is changed.
    ''' If a layer must be visible, the parent layers are also set visible.
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	9/03/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub EnforceLayerVisibility( _
            ByVal mxDoc As IMxDocument, _
            ByVal layerName As String, _
            ByVal visible As Boolean)

        Try

            ' Find the feature layer on the map based on the layername.
            Dim pLayer As IFeatureLayer = GetFeatureLayer(mxDoc.FocusMap, layerName)
            If pLayer Is Nothing Then Throw New LayerNotFoundException(layerName)
            If Not pLayer.Valid Then Throw New LayerNotValidException(layerName)

            ' Set feature layer visibility if different from current situation.
            ' Update TOC if layer visibility is changed.
            If pLayer.Visible <> visible Then
                pLayer.Visible = visible
                mxDoc.CurrentContentsView.Refresh(pLayer)
                mxDoc.ActiveView.PartialRefresh(esriViewDrawPhase.esriViewGeography, _
                    pLayer, mxDoc.ActiveView.Extent)
            End If

            ' Set parents visibility if layer must be visible.
            If visible Then
                For Each pAncestor As ILayer In FindAncestors(pLayer, mxDoc)
                    If Not pAncestor.Visible Then
                        pAncestor.Visible = True
                        mxDoc.CurrentContentsView.Refresh(pAncestor)
                        mxDoc.ActiveView.PartialRefresh(esriViewDrawPhase.esriViewGeography, _
                            pAncestor, mxDoc.ActiveView.Extent)
                    End If
                Next
            End If

        Catch ex As LayerNotFoundException
            ' Ignore.
        Catch ex As LayerNotValidException
            ' Ignore.
        Catch ex As Exception
            ErrorHandler(ex)
        End Try

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Return a subset of the original input, including only the specified characters.
    ''' </summary>
    ''' <param name="input">The text to be filtered.</param>
    ''' <param name="charlist">A string with all allowed characters.</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	14/03/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function CharFilter(ByVal input As String, ByVal charlist As String) As String

        CharFilter = ""
        Try
            Dim result As String = String.Empty
            For i As Integer = 0 To input.Length - 1
                If charlist.IndexOf(input.Substring(i, 1)) > -1 Then
                    result &= input.Substring(i, 1)
                End If
            Next
            Return result
        Catch ex As Exception
            ErrorHandler(ex)
        End Try

    End Function

End Module

