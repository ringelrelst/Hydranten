#Region "Import namespaces"
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
#End Region

''' -----------------------------------------------------------------------------
''' <summary>
''' Procedures to determine the parent layer of a specified layer,
''' or all ancestors of a specified layer.
''' </summary>
''' <remarks>
''' </remarks>
''' <history>
''' 	[Kristof Vydt]	20/02/2007	Created
'''     [Koen Vermeer]  23/06/2008  Conversion to ArcGIS 9.2
''' </history>
''' -----------------------------------------------------------------------------
Module ModuleFindAncestors

#Region "Private variables"
    Private _MxDoc As IMxDocument
#End Region

#Region "Public procedures"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Find all parent layers of a specific layer.
    ''' </summary>
    ''' <param name="pLayer">A layer object.</param>
    ''' <param name="pMxDocument">ArcMap document object.</param>
    ''' <returns>A collection of group layers.</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	20/02/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function FindAncestors( _
        ByVal pLayer As ILayer, _
        ByVal pMxDocument As IMxDocument) As Collection

        ' Store document resource as private variable.
        _MxDoc = pMxDocument

        ' Prepare an empty collection.
        FindAncestors = New Collection

        ' Find first parent of the layer.
        Dim pParentLayer As ILayer
        pParentLayer = FindParentLayer(pLayer)

        ' Add all parent layers of the parent layer.
        Do Until pParentLayer Is Nothing
            FindAncestors.Add(pParentLayer)
            pParentLayer = FindParentLayer(pParentLayer)
        Loop

    End Function

#End Region

#Region "Private procedures"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Find the parent layer of a specified layer.
    ''' </summary>
    ''' <param name="pChildLayer">The child layer.</param>
    ''' <returns>The parent layer.</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	20/02/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Function FindParentLayer( _
        ByVal pChildLayer As ILayer) As ILayer

        FindParentLayer = Nothing

        Dim pMxDoc As IMxDocument
        pMxDoc = _MxDoc
        If pMxDoc Is Nothing Then Exit Function

        ' IGrouplayer UID
        Dim pUID As New ESRI.ArcGIS.esriSystem.UID
        pUID.Value = "{EDAD6644-1810-11D1-86AE-0000F8751720}"
        On Error Resume Next

        ' Loop through all group layers of the focus map.
        Dim pEnumLayer As IEnumLayer
        pEnumLayer = pMxDoc.FocusMap.Layers(pUID, True)
        If Not pEnumLayer Is Nothing Then
            Dim pLayer As ILayer
            pLayer = pEnumLayer.Next
            Do Until pLayer Is Nothing

                ' Check if group layer contains the child layer.
                Dim pCandidate As ICompositeLayer
                pCandidate = CType(pLayer, ICompositeLayer)
                If IsParentLayer(pChildLayer, pCandidate) Then

                    ' Return the candidate layer as the valid parent layer. 
                    Dim pParent As ILayer
                    pParent = CType(pCandidate, ILayer)
                    FindParentLayer = pParent
                    Exit Function
                End If

                ' Loop to next group layer.
                pLayer = pEnumLayer.Next
            Loop
        End If

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Return True if first layer is child of second layer.
    ''' </summary>
    ''' <param name="pChildLayer">The child layer.</param>
    ''' <param name="pCandidate">The candidate parent layer.</param>
    ''' <returns>True or False</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	20/02/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Function IsParentLayer( _
        ByVal pChildLayer As ILayer, _
        ByVal pCandidate As ICompositeLayer) As Boolean

        ' Loop through composing layers.
        Dim l As Integer
        For l = 0 To pCandidate.Count - 1

            ' Return True if child layer matches.
            If pCandidate.Layer(l) Is pChildLayer Then
                IsParentLayer = True
                Exit Function
            End If
        Next l

        ' Return False if none of the composing layers matches the child layer.
        IsParentLayer = False

    End Function

#End Region

End Module
