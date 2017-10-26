Option Explicit On 
Option Strict On

#Region " Import namespaces "
Imports System.Windows.Forms
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Carto.esriViewDrawPhase
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.Geodatabase
#End Region

''' -----------------------------------------------------------------------------
''' <summary>
'''     This module is part of the "Connect Feature" functionality,
'''     used by the forms "Beheer hydranten" and "Beheer gevaren".
''' </summary>
''' <remarks>
''' </remarks>
''' <history>
''' 	[Kristof Vydt]	16/09/2005	Created
''' 	[Kristof Vydt]	11/10/2005	Introduce attemptsCounter to skip one out of 2 eventhandler calls.
''' 	[Kristof Vydt]	24/10/2005	Deactivate after filling the calling form.
''' 	[Kristof Vydt]	01/08/2006	Finetune refresh of map and contents pane on activate/deactivate.
''' 	[Kristof Vydt]	02/08/2006	Replace some global constant by a function that reads the value from config file.
'''     [Kristof Vydt]  20/02/2007  Rewrite storing &amp; resetting layer visibility/selectability, now including parent group layers.
'''     [Koen Vermeer]  23/06/2008  Conversion to ArcGIS 9.2
''' </history>
''' -----------------------------------------------------------------------------
Module ModuleConnectFeature

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
    Private m_CallingForm As IConnectFeature

    'A keyword referring to the layer that the user wants to connect to.
    ' Equals "straat" in case at least one street has been selected.
    ' Equals "dok" in case no streets but at least one dock has been selected.
    ' Equals "park" in case no streets and no docks but at least one park has been selected.
    'This keyword determines what attributes are displayed to the user in case of multiple selection.
    'This keyword determines what attributes are copied to the management form that called this connect feature functionality.
    Private m_ConnectionType As String

    ' Collection of (group)layers to restore visibility/selectability
    ' on deactivation. The collections are filled on activation.
    Private m_NonVisibleLayers As Collection
    Private m_NonSelectableLayers As Collection

#End Region

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Activate the ConnectFeature functionality.
    '''     Modify ArcGIS status so that the user can select a connecting feature.
    ''' </summary>
    ''' <param name="pMxDocument">
    '''     The ArcMap document you are working with.
    ''' </param>
    ''' <param name="pForm">
    '''     The form where the result must be written to.
    ''' </param>
    ''' <remarks>
    '''     The interface IConnectFeature is specifically designed for this purpose.
    '''     Each form that calls this procedure, must implement that interface.
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    ''' 	[Kristof Vydt]	26/09/2005	Only trigger map refresh if required.
    ''' 	[Kristof Vydt]	01/08/2006	Refresh map only if layer visibility is changed.
    '''                                 Always refresh contents pane.
    ''' 	[Kristof Vydt]	02/08/2006	Replace some global constant by a function that reads the value from config file.
    '''     [Kristof Vydt]  20/02/2007  Rewrite the part to store layer visibility/selectability, now including parent group layers.
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub ConnectFeatureFunctionality_Activate( _
        ByRef pMxDocument As IMxDocument, _
        ByRef pForm As IConnectFeature)

        Try

            Dim pEventMap As Map
            Dim pMap As IMap
            Dim pLayer As IFeatureLayer

            'Store input pointers in global variables.
            m_MxDocument = pMxDocument
            m_CallingForm = pForm

            'Create an instance of the delegate, add it to SelectionChanged event
            pEventMap = CType(pMxDocument.FocusMap, Map)
            MapSelectionChanged = New IActiveViewEvents_SelectionChangedEventHandler(AddressOf ConnectFeatureFunctionality_OnMapSelectionChanged)
            AddHandler pEventMap.SelectionChanged, MapSelectionChanged
            pEventMap = Nothing

            'Remember that this functionality is now active.
            m_Active = True

            ' Get the current map.
            pMap = pMxDocument.FocusMap

            ' Clear the collections of manipulated layers.
            m_NonVisibleLayers = New Collection
            m_NonSelectableLayers = New Collection

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

            ' Make "Dokken" layer visible and selectable.
            pLayer = GetFeatureLayer(pMap, GetLayerName("Water"))
            If pLayer Is Nothing Then
                ' Feature layer not found.
                Throw New LayerNotFoundException(GetLayerName("Water"))
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

            ' Make "Parken" layer visible and selectable.
            pLayer = GetFeatureLayer(pMap, GetLayerName("Park"))
            If pLayer Is Nothing Then
                ' Feature layer not found.
                Throw New LayerNotFoundException(GetLayerName("Park"))
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

            'The rest is done by event handler method ConnectFeatureFunctionality_OnMapSelectionChanged().

        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Connect to selected feature.
    ''' </summary>
    ''' <remarks>
    '''     This procedure is triggered by the delegate, when the user has made
    '''     a selection after activating the ConnectFeature functionality.
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    ''' 	[Kristof Vydt]	23/09/2005	Exit if there is no MxDocument.
    ''' 	[Kristof Vydt]	05/10/2005	No longer deactivate.
    ''' 	[Kristof Vydt]	11/10/2005	Skip 1 out of 2 attempts.
    ''' 	[Kristof Vydt]	02/08/2006	Replace some global constant by a function that reads the value from config file.
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub ConnectFeatureFunctionality_OnMapSelectionChanged()

        Try

            Dim pMap As IMap
            Dim pFLayer As IFeatureLayer
            Dim pSelectionSet As ISelectionSet
            Dim pCursor As ICursor = Nothing
            Dim pFCursor As IFeatureCursor
            Dim pFeature As IFeature

            'Deactivate the feature functionality.
            'If m_Active Then ConnectFeatureFunctionality_Deactivate()

            'Exit if there is no document.
            If m_MxDocument Is Nothing Then Exit Sub

            'Skip first (and every odd) attempt because every change to 
            'the map selection, triggers 2 times the event.
            m_AttemptsCounter = m_AttemptsCounter + 1
            If m_AttemptsCounter Mod 2 = 1 Then Exit Sub

            'Retrieve feature selection - Streets.
            Dim LayerName As String = GetLayerName("Straatassen")
            pMap = m_MxDocument.FocusMap
            pFLayer = GetFeatureLayer(pMap, LayerName)
            If pFLayer Is Nothing Then _
                Throw New LayerNotFoundException(LayerName)
            pSelectionSet = LayerSelectionSet(m_MxDocument, pFLayer)

            If pSelectionSet.Count = 1 Then 'One street

                'Debug.WriteLine("Connect feature: one street ...")
                'MsgBox("Connect feature: one street ...")

                m_ConnectionType = "straat"
                pFLayer = GetFeatureLayer(pMap, GetLayerName("Straatassen"))

                'Get the first feature.
                pSelectionSet.Search(Nothing, False, pCursor)
                pFCursor = CType(pCursor, IFeatureCursor)
                pFeature = pFCursor.NextFeature
                'Pass this one to the calling form.
                ReturnFeature(pFeature)

            ElseIf pSelectionSet.Count > 1 Then 'Several streets

                'Debug.WriteLine("Connect feature: several streets ...")
                'MsgBox("Connect feature: several streets ...")

                m_ConnectionType = "straat"

                Dim ListForm As FormConnectFeatureList
                ListForm = New FormConnectFeatureList(pSelectionSet, m_ConnectionType, m_MxDocument)
                ListForm.ShowDialog()
                'From here, the ListForm takes over, showing the complete selectionset.
                'The user selects one feature from the listbox.

            ElseIf pSelectionSet.Count = 0 Then 'No streets.

                'Debug.WriteLine("Connect feature: no streets ...")
                'MsgBox("Connect feature: no streets ...")

                'Retrieve feature selection - Docks.
                pFLayer = GetFeatureLayer(pMap, GetLayerName("Water"))
                pSelectionSet = LayerSelectionSet(m_MxDocument, pFLayer)

                If pSelectionSet.Count = 1 Then 'One dock

                    'Debug.WriteLine("Connect feature: one dock ...")
                    'MsgBox("Connect feature: one dock ...")

                    m_ConnectionType = "dok"

                    'Get the first feature.
                    pSelectionSet.Search(Nothing, False, pCursor)
                    pFCursor = CType(pCursor, IFeatureCursor)
                    pFeature = pFCursor.NextFeature
                    'Pass this one to the calling form.
                    ReturnFeature(pFeature)

                ElseIf pSelectionSet.Count > 1 Then 'Several docks

                    'Debug.WriteLine("Connect feature: several docks ...")
                    'MsgBox("Connect feature: several docks ...")

                    m_ConnectionType = "dok"

                    Dim ListForm As FormConnectFeatureList
                    ListForm = New FormConnectFeatureList(pSelectionSet, m_ConnectionType, m_MxDocument)
                    ListForm.ShowDialog()
                    'From here, the ListForm takes over, showing the complete selectionset.
                    'The user selects one feature from the listbox.

                ElseIf pSelectionSet.Count = 0 Then 'No docks

                    'Debug.WriteLine("Connect feature: no docks ...")
                    'MsgBox("Connect feature: no docks ...")

                    'Retrieve feature selection - Parks.
                    pFLayer = GetFeatureLayer(pMap, GetLayerName("Park"))
                    pSelectionSet = LayerSelectionSet(m_MxDocument, pFLayer)

                    If pSelectionSet.Count = 1 Then 'One park

                        'Debug.WriteLine("Connect feature: one park ...")
                        'MsgBox("Connect feature: one park ...")

                        m_ConnectionType = "park"

                        'Get the first feature.
                        pSelectionSet.Search(Nothing, False, pCursor)
                        pFCursor = CType(pCursor, IFeatureCursor)
                        pFeature = pFCursor.NextFeature
                        'Pass this one to the calling form.
                        ReturnFeature(pFeature)

                    ElseIf pSelectionSet.Count > 1 Then 'Several parks

                        'Debug.WriteLine("Connect feature: several parks ...")
                        'MsgBox("Connect feature: several parks ...")

                        m_ConnectionType = "park"

                        Dim ListForm As FormConnectFeatureList
                        ListForm = New FormConnectFeatureList(pSelectionSet, m_ConnectionType, m_MxDocument)
                        ListForm.ShowDialog()
                        'From here, the ListForm takes over, showing the complete selectionset.
                        'The user selects one feature from the listbox.

                    ElseIf pSelectionSet.Count = 0 Then 'No parks

                        'Debug.WriteLine("Connect feature: no parks ...")
                        'MsgBox("Connect feature: no parks ...")

                        'No selection at all in one of the connecting layers.
                        Throw New ApplicationException( _
                            c_Message_NoConnectFeatureSelection)

                    End If
                End If
            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Deactivate the ConnectFeature functionality.
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    ''' 	[Kristof Vydt]	23/09/2005	Made public. Exit if no MxDocument or not active.
    ''' 	[Kristof Vydt]	26/09/2005	Only trigger map refresh if required.
    ''' 	[Kristof Vydt]	05/10/2005	Uncheck toolbutton on calling form.
    ''' 	[Kristof Vydt]	01/08/2006	Clear selection and refresh map and contents pane at the end.
    ''' 	[Kristof Vydt]	02/08/2006	Replace some global constant by a function that reads the value from config file.
    '''     [Krisotf Vydt]  20/02/2007  Rewrite the part to reset layer visibility/selectability, now including parent group layers.
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub ConnectFeatureFunctionality_Deactivate()

        Try

            Dim pEventMap As Map
            '    Dim pMap As IMap
            Dim pLayer As ILayer
            Dim pFeatureLayer As IFeatureLayer

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
            m_CallingForm.Toolbutton.Checked = False

            ' Clear selection.
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
    '''     Retrieve attribute info of the connected feature,
    '''     and display it in the calling form.
    ''' </summary>
    ''' <param name="pFeature">
    '''     The feature that you are connecting to.
    ''' </param>
    ''' <remarks>
    '''     Specifically for this procedure, IConnectFeature has been introduced.
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    ''' 	[Kristof Vydt]	23/09/2005	Exit if there is no calling form.
    ''' 	[Kristof Vydt]	24/10/2005	Deactivate at the end.
    ''' 	[Kristof Vydt]	02/08/2006	Replace some global constant by a function that reads the value from config file.
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub ReturnFeature(ByVal pFeature As IFeature)

        Try

            Dim pFields As IFields
            '    Dim pField As IField
            Dim FieldIndex As Integer

            'Exit if form or feature is missing.
            If m_CallingForm Is Nothing Then Exit Sub
            If pFeature Is Nothing Then Exit Sub

            Select Case m_ConnectionType
                Case "straat" 'Copy value of attributes: straatnaam, straatcode, postcode.

                    pFields = pFeature.Fields
                    FieldIndex = pFields.FindField(GetAttributeName("Straatassen", "Straatnaam"))
                    m_CallingForm.Straatnaam = CStr(pFeature.Value(FieldIndex))
                    FieldIndex = pFields.FindField(GetAttributeName("Straatassen", "Straatcode"))
                    m_CallingForm.Straatcode = CStr(pFeature.Value(FieldIndex))
                    FieldIndex = pFields.FindField(GetAttributeName("Straatassen", "Postcode"))
                    m_CallingForm.Postcode = CStr(pFeature.Value(FieldIndex))

                Case "dok" 'Copy value of attribute : naam.

                    pFields = pFeature.Fields
                    FieldIndex = pFields.FindField(GetAttributeName("Water", "Naam"))
                    m_CallingForm.Straatnaam = CStr(pFeature.Value(FieldIndex))
                    m_CallingForm.Straatcode = ""
                    m_CallingForm.Postcode = ""

                Case "park" 'Copy value of attribute : naam.

                    pFields = pFeature.Fields
                    FieldIndex = pFields.FindField(GetAttributeName("Park", "Naam"))
                    m_CallingForm.Straatnaam = CStr(pFeature.Value(FieldIndex))
                    m_CallingForm.Straatcode = ""
                    m_CallingForm.Postcode = ""

            End Select

            'Deactivate.
            ConnectFeatureFunctionality_Deactivate()

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

End Module
