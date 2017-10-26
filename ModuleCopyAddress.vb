Option Strict On
Option Explicit On 

#Region " Imports namespaces "
Imports System.Windows.Forms
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Carto.esriViewDrawPhase
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.Geodatabase
#End Region

''' -----------------------------------------------------------------------------
''' <summary>
'''     This module is part of the "Copy Address" functionality,
'''     used by the form "Beheer van speciale gebouwen".
''' </summary>
''' <remarks>
''' </remarks>
''' <history>
''' 	[Kristof Vydt]	16/09/2005	Created
''' 	[Kristof Vydt]	11/10/2005	Introduce attemptsCounter to skip one out of 2 eventhandler calls.
''' 	[Kristof Vydt]	24/10/2005	Deactivate after filling the calling form.
''' 	[Kristof Vydt]	27/10/2005	Bugs in coordination with calling form solved.
''' 	[Kristof Vydt]	02/08/2006	Replace some global constant by a function that reads the value from config file.
'''     [Kristof Vydt]  20/02/2007  Rewrite storing &amp; resetting layer visibility/selectability, now including parent group layers.
'''     [Koen Vermeer]  23/06/2008  Conversion to ArcGIS 9.2
''' </history>
''' -----------------------------------------------------------------------------
Module ModuleCopyAddress

#Region " Private variables "

    'Store the status of this functionality.
    Private m_Active As Boolean = False

    'Keep track of the number of calls to this functionality.
    Private m_AttemptsCounter As Long = 0

    'Declare the delegate (event listener)
    Private MapSelectionChanged As IActiveViewEvents_SelectionChangedEventHandler

    'Pointer to ArcMap Document objects.
    Private m_MxDocument As IMxDocument

    'Pointer to the form that called the connect feature functionality.
    Private m_CallingForm As FormBeheerGebouwen

    ' Collection of (group)layers to restore visibility/selectability
    ' on deactivation. The collections are filled on activation.
    Private m_NonVisibleLayers As Collection
    Private m_NonSelectableLayers As Collection

#End Region

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Activate the CopyAddress functionality.
    '''     Modify ArcGIS status so that the user can select a feature to copy from.
    ''' </summary>
    ''' <param name="pMxDocument">
    '''     The ArcMap document you are working with.
    ''' </param>
    ''' <param name="pForm">
    '''     The form where the result must be written to.
    ''' </param>
    ''' <remarks>
    '''     Since this functionality is used in only one form,
    '''     there is no interface specifically designed for this purpose.
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    ''' 	[Kristof Vydt]	26/09/2005	Only trigger map refresh if required.
    ''' 	[Kristof Vydt]	27/10/2005	Refresh TOC at the end.
    ''' 	[Kristof Vydt]	02/08/2006	Replace some global constant by a function that reads the value from config file.
    ''' 	[Kristof Vydt]	31/08/2006	Refresh map only if layer visibility is changed.
    '''                                 Always refresh contents pane.
    '''     [Kristof Vydt]  20/02/2007  Rewrite the part to store layer visibility/selectability, now including parent group layers.
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub CopyAddressFunctionality_Activate( _
        ByRef pMxDocument As IMxDocument, _
        ByRef pForm As FormBeheerGebouwen)

        Try

            Dim pEventMap As Map
            Dim pMap As IMap
            Dim pLayer As IFeatureLayer

            'Store input pointers in global variables.
            m_MxDocument = pMxDocument
            m_CallingForm = pForm

            'Create an instance of the delegate, add it to SelectionChanged event
            pEventMap = CType(pMxDocument.FocusMap, Map)
            MapSelectionChanged = New IActiveViewEvents_SelectionChangedEventHandler(AddressOf CopyAddressFunctionality_OnMapSelectionChanged)
            AddHandler pEventMap.SelectionChanged, MapSelectionChanged
            pEventMap = Nothing

            'Remember that this functionality is now active.
            m_Active = True

            ' Get the current map.
            pMap = pMxDocument.FocusMap

            ' Clear the collections of manipulated layers.
            m_NonVisibleLayers = New Collection
            m_NonSelectableLayers = New Collection

            ' Make "Hoofdgebouw" layer visible and selectable.
            pLayer = GetFeatureLayer(pMap, GetLayerName("Hoofdgebouw"))
            If pLayer Is Nothing Then
                ' Feature layer not found.
                Throw New LayerNotFoundException(GetLayerName("Hoofdgebouw"))
            Else
                ' Set feature layer visibility.
                If Not pLayer.Visible Then
                    m_NonVisibleLayers.Add(pLayer)
                    pLayer.Visible = True
                    pMxDocument.CurrentContentsView.Refresh(pLayer)
                End If
                ' Set feature layer selectability.
                If Not pLayer.Selectable Then
                    m_NonSelectableLayers.Add(pLayer)
                    pLayer.Selectable = True
                    pMxDocument.CurrentContentsView.Refresh(pLayer)
                End If
                ' Set parents visibility.
                For Each pAncestor As ILayer In FindAncestors(pLayer, m_MxDocument)
                    If Not pAncestor.Visible Then
                        m_NonVisibleLayers.Add(pAncestor)
                        pAncestor.Visible = True
                        pMxDocument.CurrentContentsView.Refresh(pAncestor)
                    End If
                Next
            End If

            ' Make "Straatassen" layer visible and selectable.
            pLayer = GetFeatureLayer(pMap, GetLayerName("Straatassen"))
            If pLayer Is Nothing Then
                ' Feature layer not found.
                Throw New LayerNotFoundException(GetLayerName("Straatassen"))
            Else
                ' Set feature layer visibility.
                If Not pLayer.Visible Then
                    m_NonVisibleLayers.Add(pLayer)
                    pLayer.Visible = True
                    m_MxDocument.CurrentContentsView.Refresh(pLayer)
                End If
                ' Set feature layer selectability.
                If Not pLayer.Selectable Then
                    m_NonSelectableLayers.Add(pLayer)
                    pLayer.Selectable = True
                    m_MxDocument.CurrentContentsView.Refresh(pLayer)
                End If
                ' Set parents visibility.
                For Each pAncestor As ILayer In FindAncestors(pLayer, m_MxDocument)
                    If Not pAncestor.Visible Then
                        m_NonVisibleLayers.Add(pAncestor)
                        pAncestor.Visible = True
                        m_MxDocument.CurrentContentsView.Refresh(pAncestor)
                    End If
                Next
            End If

            ' Refresh the active view if layer visibility has changed.
            If m_NonVisibleLayers.Count > 0 Then _
                pMxDocument.ActiveView.PartialRefresh(esriViewGeography, Nothing, Nothing)

            ' Activate the SelectFeature tool.
            ActivateTool(CType(pMxDocument, IDocument), "esriArcMapUI.SelectFeaturesTool")

            'The rest is done by event handler method CopyAddressFunctionality_OnMapSelectionChanged().

        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Copy address attributes of selected feature.
    ''' </summary>
    ''' <remarks>
    '''     This procedure is triggered by the delegate, when the user has made
    '''     a selection after activating the "Copy Address" functionality.
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    ''' 	[Kristof Vydt]	23/09/2005	Handling of invalid selections is reviewed.
    ''' 	[Kristof Vydt]	11/10/2005	Skip 1 out of 2 attempts.
    ''' 	[Kristof Vydt]	27/10/2005	No deactivation in this method.
    ''' 	[Kristof Vydt]	02/08/2006	Replace some global constant by a function that reads the value from config file.
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub CopyAddressFunctionality_OnMapSelectionChanged()

        Try

            Dim pMap As IMap
            Dim pFLayer As IFeatureLayer
            Dim pSelectionSet As ISelectionSet
            Dim pCursor As ICursor = Nothing
            Dim pFCursor As IFeatureCursor
            Dim pFeature As IFeature
        '    Dim statusFieldIndex As Integer
            '    Dim statusCodeValue As String

            'Deactivate the functionality.
            'If m_Active Then CopyAddressFunctionality_Deactivate()

            'Exit if there is no document.
            If m_MxDocument Is Nothing Then Exit Sub

            'Skip first (and every odd) attempt because every change to 
            'the map selection, triggers 2 times the event.
            m_AttemptsCounter = m_AttemptsCounter + 1
            If m_AttemptsCounter Mod 2 = 1 Then Exit Sub

            'Retrieve feature selection - Hoofdgebouw.
            pMap = m_MxDocument.FocusMap
            pFLayer = GetFeatureLayer(pMap, GetLayerName("Hoofdgebouw"))
            pSelectionSet = LayerSelectionSet(m_MxDocument, pFLayer)

            Select Case pSelectionSet.Count
                Case 0 'No features selected.
                    'Just continue and look for selection in the streets layer.

                Case 1 'One feature selected

                    'Get the feature.
                    pSelectionSet.Search(Nothing, False, pCursor)
                    pFCursor = CType(pCursor, IFeatureCursor)
                    pFeature = pFCursor.NextFeature

                    'Pass this one to the calling form.
                    ReturnAttributes(pFeature, _
                        huisnummerIndex:=pFeature.Fields.FindField(GetAttributeName("Hoofdgebouw", "Huisnr")), _
                        postcodeIndex:=pFeature.Fields.FindField(GetAttributeName("Hoofdgebouw", "Postcode")), _
                        straatcodeIndex:=pFeature.Fields.FindField(GetAttributeName("Hoofdgebouw", "Straatcode")), _
                        straatnaamIndex:=pFeature.Fields.FindField(GetAttributeName("Hoofdgebouw", "Straatnaam")))
                    Exit Sub

                Case Else 'More than one feature selected.
                    Throw New ApplicationException(c_Message_CopyAddressFromMultipleObjects)
            End Select

            'Retrieve feature selection - Straatassen.
            pMap = m_MxDocument.FocusMap
            pFLayer = GetFeatureLayer(pMap, GetLayerName("Straatassen"))
            pSelectionSet = LayerSelectionSet(m_MxDocument, pFLayer)

            Select Case pSelectionSet.Count
                Case 0 'No features selected.
                    Throw New ApplicationException(c_Message_NoFeaturesToCopyAddress)

                Case 1 'One feature selected

                    'Get the feature.
                    pSelectionSet.Search(Nothing, False, pCursor)
                    pFCursor = CType(pCursor, IFeatureCursor)
                    pFeature = pFCursor.NextFeature

                    'Pass this one to the calling form.
                    ReturnAttributes(pFeature, _
                        postcodeIndex:=pFeature.Fields.FindField(GetAttributeName("Straatassen", "Postcode")), _
                        straatcodeIndex:=pFeature.Fields.FindField(GetAttributeName("Straatassen", "Straatcode")), _
                        straatnaamIndex:=pFeature.Fields.FindField(GetAttributeName("Straatassen", "Straatnaam")))
                    Exit Sub

                Case Else 'More than one feature selected.
                    Throw New ApplicationException(c_Message_CopyAddressFromMultipleObjects)
            End Select

        Catch ex As Exception
            ErrorHandler(ex)
        End Try
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Deactivate the "Copy Address" functionality.
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    ''' 	[Kristof Vydt]	23/09/2005	Made public. Exit if no MxDocument or not active.
    ''' 	[Kristof Vydt]	26/09/2005	Only trigger map refresh if required.
    ''' 	[Kristof Vydt]	27/10/2005	Refresh TOC at the end.
    '''                                 Uncheck toolbutton on calling form.
    ''' 	[Kristof Vydt]	02/08/2006	Replace some global constant by a function that reads the value from config file.
    ''' 	[Kristof Vydt]	31/08/2006	Clear selection and refresh map and contents pane at the end.
    '''     [Krisotf Vydt]  20/02/2007  Rewrite the part to reset layer visibility/selectability, now including parent group layers.
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub CopyAddressFunctionality_Deactivate()
        Try
            Dim pEventMap As Map
            '    Dim pMap As IMap
            Dim pFeatureLayer As IFeatureLayer
            Dim pLayer As ILayer

            ' Exit if the functionality is not active.
            If Not m_Active Then Exit Sub

            ' Exit if there is no document.
            If m_MxDocument Is Nothing Then Exit Sub

            ' Remove handler after feature selection is retrieved.
            pEventMap = CType(m_MxDocument.FocusMap, Map)
            RemoveHandler pEventMap.SelectionChanged, MapSelectionChanged
            pEventMap = Nothing

            ' Remember that this functionality is not longer active.
            m_Active = False

            ' Uncheck toolbutton on calling form.
            m_CallingForm.CheckBoxCopy.Checked = False

            ' Clear selection on focus map.
            m_MxDocument.FocusMap.ClearSelection()

            ' Reset layers that were set visible during activation.
            For Each pLayer In m_NonVisibleLayers
                pLayer.Visible = False
                m_MxDocument.CurrentContentsView.Refresh(pLayer)
            Next

            ' Reset layers that were set selectable during activation.
            For Each pFeatureLayer In m_NonSelectableLayers
                pFeatureLayer.Selectable = False
                m_MxDocument.CurrentContentsView.Refresh(pFeatureLayer)
            Next

            ' Refresh the active view.
            m_MxDocument.ActiveView.PartialRefresh(esriViewGeoSelection Or esriViewGeography, Nothing, Nothing)

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Retrieve address attribute values of the feature to copy from,
    '''     and display it in the calling form.
    ''' </summary>
    ''' <param name="pFeature">
    '''     The feature that you are copying from.
    ''' </param>
    ''' <remarks>
    '''     Because the "Copy Address" functionality is only used
    '''     in one form, no specific interface has been introduced.
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    ''' 	[Kristof Vydt]	23/09/2005	Exit if there is no calling form.
    ''' 	[Kristof Vydt]	24/10/2005	Deactivate at the end.
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub ReturnAttributes( _
        ByVal pFeature As IFeature, _
        Optional ByVal straatnaamIndex As Integer = -1, _
        Optional ByVal straatcodeIndex As Integer = -1, _
        Optional ByVal huisnummerIndex As Integer = -1, _
        Optional ByVal postcodeIndex As Integer = -1)

        Try
            Dim fieldIndex As Integer
            Dim attributeValue As String
            Dim attributeControl As Control
            Dim compositeAddress As String

            'Exit if form or feature is missing.
            If m_CallingForm Is Nothing Then Exit Sub
            If pFeature Is Nothing Then Exit Sub

            '- Huisnummer
            fieldIndex = huisnummerIndex
            If fieldIndex < 0 Then
                attributeValue = ""
            ElseIf TypeOf pFeature.Value(fieldIndex) Is System.DBNull Then
                attributeValue = ""
            Else
                attributeValue = CStr(pFeature.Value(fieldIndex))
            End If
            attributeControl = m_CallingForm.TextBoxHuisnr
            If attributeValue <> attributeControl.Text Then _
                SetEditBoxValue(attributeControl, attributeValue)
            compositeAddress = attributeValue

            '- Postcode
            fieldIndex = postcodeIndex
            If fieldIndex < 0 Then
                attributeValue = ""
            ElseIf TypeOf pFeature.Value(fieldIndex) Is System.DBNull Then
                attributeValue = ""
            Else
                attributeValue = CStr(pFeature.Value(fieldIndex))
            End If
            attributeControl = m_CallingForm.TextBoxPostcode
            If attributeValue <> attributeControl.Text Then _
                SetEditBoxValue(attributeControl, attributeValue)

            '- Straatcode
            fieldIndex = straatcodeIndex
            If fieldIndex < 0 Then
                attributeValue = ""
            ElseIf TypeOf pFeature.Value(fieldIndex) Is System.DBNull Then
                attributeValue = ""
            Else
                attributeValue = CStr(pFeature.Value(fieldIndex))
            End If
            attributeControl = m_CallingForm.TextBoxStraatcode
            If attributeValue <> attributeControl.Text Then _
                SetEditBoxValue(attributeControl, attributeValue)

            '- Straatnaam
            fieldIndex = straatnaamIndex
            If fieldIndex < 0 Then
                attributeValue = ""
            ElseIf TypeOf pFeature.Value(fieldIndex) Is System.DBNull Then
                attributeValue = ""
            Else
                attributeValue = CStr(pFeature.Value(fieldIndex))
            End If
            attributeControl = m_CallingForm.TextBoxStraatnaam
            If attributeValue <> attributeControl.Text Then _
                SetEditBoxValue(attributeControl, attributeValue)
            compositeAddress = Trim(attributeValue & " " & compositeAddress)

            '- Samengesteld adres
            attributeValue = MixedCasing(compositeAddress)
            attributeControl = m_CallingForm.TextBoxAanduiding
            SetEditBoxValue(attributeControl, attributeValue)

            'Deactivate.
            CopyAddressFunctionality_Deactivate()

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

End Module
