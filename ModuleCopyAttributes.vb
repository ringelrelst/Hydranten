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
'''     This module is part of the "Copy Attributes" functionality,
'''     used by the form "Beheer hydranten".
''' </summary>
''' <remarks>
''' </remarks>
''' <history>
''' 	[Kristof Vydt]	16/09/2005	Created
''' 	[Kristof Vydt]	11/10/2005	Introduce attemptsCounter to skip one out of 2 eventhandler calls.
''' 	[Kristof Vydt]	24/10/2005	Deactivate after filling the calling form.
''' 	[Kristof Vydt]	27/10/2005	Refresh TOC at activate &amp; deactivate.
''' 	[Kristof Vydt]	02/08/2006	Replace some global constant by a function that reads the value from config file.
'''     [Kristof Vydt]  20/02/2007  Rewrite storing &amp; resetting layer visibility/selectability, now including parent group layers.
'''     [Koen Vermeer]  23/06/2008  Conversion to ArcGIS 9.2
''' </history>
''' -----------------------------------------------------------------------------
Module ModuleCopyAttributes

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
    Private m_CallingForm As FormBeheerHydranten

    ' Collection of (group)layers to restore visibility/selectability
    ' on deactivation. The collections are filled on activation.
    Private m_NonVisibleLayers As Collection
    Private m_NonSelectableLayers As Collection

#End Region

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Activate the CopyAttributes functionality.
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
    Public Sub CopyAttributesFunctionality_Activate( _
        ByRef pMxDocument As IMxDocument, _
        ByRef pForm As FormBeheerHydranten)

        Try

            Dim pEventMap As Map
            Dim pMap As IMap
            Dim pLayer As IFeatureLayer

            'Store input pointers in global variables.
            m_MxDocument = pMxDocument
            m_CallingForm = pForm

            'Create an instance of the delegate, add it to SelectionChanged event
            pEventMap = CType(pMxDocument.FocusMap, Map)
            MapSelectionChanged = New IActiveViewEvents_SelectionChangedEventHandler(AddressOf CopyAttributesFunctionality_OnMapSelectionChanged)
            AddHandler pEventMap.SelectionChanged, MapSelectionChanged
            pEventMap = Nothing

            'Remember that this functionality is now active.
            m_Active = True

            ' Get the current map.
            pMap = pMxDocument.FocusMap

            ' Clear the collections of manipulated layers.
            m_NonVisibleLayers = New Collection
            m_NonSelectableLayers = New Collection

            ' Make "Hydrant" layer visible and selectable.
            pLayer = GetFeatureLayer(pMap, GetLayerName("Hydrant"))
            If pLayer Is Nothing Then
                ' Feature layer not found.
                Throw New LayerNotFoundException(GetLayerName("Hydrant"))
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

            'The rest is done by event handler method CopyFeatureFunctionality_OnMapSelectionChanged().

        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Copy attributes of selected feature.
    ''' </summary>
    ''' <remarks>
    '''     This procedure is triggered by the delegate, when the user has made
    '''     a selection after activating the "Copy Attributes" functionality.
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    ''' 	[Kristof Vydt]	23/09/2005	Handling of invalid selections is reviewed.
    ''' 	[Kristof Vydt]	05/10/2005	No longer deactivate.
    ''' 	[Kristof Vydt]	11/10/2005	Skip 1 out of 2 attempts.
    ''' 	[Kristof Vydt]	02/08/2006	Replace some global constant by a function that reads the value from config file.
    '''     [Kristof Vydt]  31/08/2006  Filter on status "verwijderd" within selectset, before checking on >1 feature.
    ''' 	[Kristof Vydt]	22/03/2007	Use the new CodedValueDomainManager instead of the deprecated ModuleDomainAccess.
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub CopyAttributesFunctionality_OnMapSelectionChanged()

        Try

            Dim pCursor As ICursor = Nothing
            Dim pFeature As IFeature
            Dim pFeatureCursor As IFeatureCursor
            Dim pFeatureLayer As IFeatureLayer
            Dim pMap As IMap
            Dim pQueryFilter As IQueryFilter
            Dim pSelectionSet As ISelectionSet

            ' Deactivate the functionality.
            'If m_Active Then CopyAttributesFunctionality_Deactivate()

            ' Exit if there is no document.
            If m_MxDocument Is Nothing Then Exit Sub

            ' Skip first (and every odd) attempt because every change to 
            ' the map selection, triggers 2 times the event.
            m_AttemptsCounter = m_AttemptsCounter + 1
            If m_AttemptsCounter Mod 2 = 1 Then Exit Sub

            ' Retrieve feature selection from layer "Hydrant".
            pMap = m_MxDocument.FocusMap
            pFeatureLayer = GetFeatureLayer(pMap, GetLayerName("Hydrant"))
            pSelectionSet = LayerSelectionSet(m_MxDocument, pFeatureLayer)

            ' Abort if no hydrants selected.
            If pSelectionSet.Count = 0 Then _
                Throw New ApplicationException(c_Message_NoHydrantToCopyAttributes)

            ' Get selected features with status "verwijderd".
            Dim domainMgr As New CodedValueDomainManager(pFeatureLayer, "Status")
            Dim attrValue As String = domainMgr.CodeValue("verwijderd")
            Dim attrName As String = GetAttributeName("Hydrant", "Status")
            pQueryFilter = New QueryFilter
            pQueryFilter.WhereClause = attrName & "=" & CStrSql(attrValue)
            pSelectionSet.Search(pQueryFilter, False, pCursor)

            ' Get the first feature from cursor.
            pFeatureCursor = CType(pCursor, IFeatureCursor)
            pFeature = pFeatureCursor.NextFeature

            ' Abort if no valid feature found.
            If pFeature Is Nothing Then _
                Throw New ApplicationException(c_Message_NoHydrantToCopyAttributes)

            ' Abort if more than one valid feature found.
            Dim pTmpFeature As IFeature = pFeatureCursor.NextFeature
            If Not pTmpFeature Is Nothing Then _
                Throw New ApplicationException(c_Message_MultipleHydrantsToCopyAttributes)

            'Pass the feature to the calling form.
            ReturnAttributes(pFeature)

        Catch ex As Exception
            ErrorHandler(ex)
        End Try
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Deactivate the "Copy Attributes" functionality.
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    ''' 	[Kristof Vydt]	23/09/2005	Made public. Exit if no MxDocument or not active.
    ''' 	[Kristof Vydt]	26/09/2005	Only trigger map refresh if required.
    ''' 	[Kristof Vydt]	05/10/2005	Uncheck toolbutton on calling form.
    ''' 	[Kristof Vydt]	27/10/2005	Refresh TOC at the end.
    ''' 	[Kristof Vydt]	02/08/2006	Replace some global constant by a function that reads the value from config file.
    ''' 	[Kristof Vydt]	31/08/2006	Clear selection and refresh map and contents pane at the end.
    '''     [Krisotf Vydt]  20/02/2007  Rewrite the part to reset layer visibility/selectability, now including parent group layers.
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub CopyAttributesFunctionality_Deactivate()
        Try

            Dim pEventMap As Map
            '    Dim pMap As IMap
            Dim pLayer As ILayer
            Dim pFeatureLayer As IFeatureLayer

            'Exit if the functionality is not active.
            If Not m_Active Then Exit Sub

            'Exit if there is no document.
            If m_MxDocument Is Nothing Then Exit Sub

            'Remove handler after feature selection is retrieved.
            pEventMap = CType(m_MxDocument.FocusMap, Map)
            RemoveHandler pEventMap.SelectionChanged, MapSelectionChanged
            pEventMap = Nothing

            'Remember that this functionality is not longer active.
            m_Active = False

            'Uncheck toolbutton on calling form.
            m_CallingForm.CheckBoxCopy.Checked = False

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
    '''     Retrieve attribute values of the feature to copy from,
    '''     and display it in the calling form.
    ''' </summary>
    ''' <param name="pFeature">
    '''     The feature that you are copying from.
    ''' </param>
    ''' <remarks>
    '''     Because the "Copy Attributes" functionality is only used
    '''     in one form, no specific interface has been introduced.
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    ''' 	[Kristof Vydt]	23/09/2005	Exit if there is no calling form.
    ''' 	[Kristof Vydt]	24/10/2005	Deactivate at the end.
    ''' 	[Kristof Vydt]	25/10/2005	Remember the hydrant that was copied from in the calling form.
    ''' 	[Kristof Vydt]	02/08/2006	Replace some global constant by a function that reads the value from config file.
    '''     [Kristof Vydt]  28/09/2006  Replace incorrect reference to TextBoxLeidingID into ComboBoxLeidingType.
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub ReturnAttributes(ByVal pFeature As IFeature)
        Try

            Dim pFields As IFields
            '    Dim pField As IField
            Dim FieldIndex As Integer
            Dim AttributeValue As String
            Dim AttributeControl As Control

            'Exit if form or feature is missing.
            If m_CallingForm Is Nothing Then Exit Sub
            If pFeature Is Nothing Then Exit Sub

            'Remember the feature in the calling form.
            'When saving changes in that form, the user will be asked 
            'to automatically modify this feature to "historiek".
            m_CallingForm.SetCopyFrom(pFeature)

            'Copy value of several attributes.
            pFields = pFeature.Fields

            '- Aanduiding
            FieldIndex = pFields.FindField(GetAttributeName("Hydrant", "Aanduiding"))
            AttributeValue = CStr(pFeature.Value(FieldIndex))
            AttributeControl = m_CallingForm.TextBoxAanduiding
            If AttributeValue <> AttributeControl.Text Then
                SetEditBoxValue(AttributeControl, AttributeValue)
                'm_CallingForm.MarkAsChanged(AttributeControl)
            End If

            '- BrandweerNummer
            FieldIndex = pFields.FindField(GetAttributeName("Hydrant", "BrandweerNr"))
            AttributeValue = CStr(pFeature.Value(FieldIndex))
            AttributeControl = m_CallingForm.TextBoxBrandweerID
            If AttributeValue <> AttributeControl.Text Then
                SetEditBoxValue(AttributeControl, AttributeValue)
                'm_CallingForm.MarkAsChanged(AttributeControl)
            End If

            '- Bron
            FieldIndex = pFields.FindField(GetAttributeName("Hydrant", "Bron"))
            AttributeValue = CStr(pFeature.Value(FieldIndex))
            AttributeControl = m_CallingForm.ComboBoxBron
            If AttributeValue <> AttributeControl.Text Then
                SetEditBoxValue(AttributeControl, AttributeValue)
                'm_CallingForm.MarkAsChanged(AttributeControl)
            End If

            '- Ligging
            FieldIndex = pFields.FindField(GetAttributeName("Hydrant", "Ligging"))
            AttributeValue = CStr(pFeature.Value(FieldIndex))
            AttributeControl = m_CallingForm.ComboBoxLigging
            If AttributeValue <> AttributeControl.Text Then
                SetEditBoxValue(AttributeControl, AttributeValue)
                'm_CallingForm.MarkAsChanged(AttributeControl)
            End If

            '- LeidingType
            FieldIndex = pFields.FindField(GetAttributeName("Hydrant", "LeidingType"))
            AttributeValue = CStr(pFeature.Value(FieldIndex))
            AttributeControl = m_CallingForm.ComboBoxLeidingType
            If AttributeValue <> AttributeControl.Text Then
                SetEditBoxValue(AttributeControl, AttributeValue)
                'm_CallingForm.MarkAsChanged(AttributeControl)
            End If

            '- Straatnaam
            FieldIndex = pFields.FindField(GetAttributeName("Hydrant", "Straatnaam"))
            AttributeValue = CStr(pFeature.Value(FieldIndex))
            AttributeControl = m_CallingForm.TextBoxStraatnaam
            If AttributeValue <> AttributeControl.Text Then
                SetEditBoxValue(AttributeControl, AttributeValue)
                'm_CallingForm.MarkAsChanged(AttributeControl)
            End If

            '- Straatcode
            FieldIndex = pFields.FindField(GetAttributeName("Hydrant", "Straatcode"))
            AttributeValue = CStr(pFeature.Value(FieldIndex))
            AttributeControl = m_CallingForm.TextBoxStraatcode
            If AttributeValue <> AttributeControl.Text Then
                SetEditBoxValue(AttributeControl, AttributeValue)
                'm_CallingForm.MarkAsChanged(AttributeControl)
            End If

            'Modify source feature only on saving changes to destination feature.
            'This is taken care of in the calling form, when saving the changes.

            'Deactivate.
            CopyAttributesFunctionality_Deactivate()

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

End Module
