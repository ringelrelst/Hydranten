Option Explicit On 
Option Strict On

Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Display
Imports ESRI.ArcGIS.Editor
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Geometry

Module ModuleAnnotations

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Return the first selected feature from the specified annotation layer.
    ''' </summary>
    ''' <param name="pAnnotationLayer">
    '''     The annotation layer from which we want to retrieve one feature.
    ''' </param>
    ''' <returns>
    '''     An annotation feature.
    ''' </returns>
    ''' <remarks>
    '''     This procedure is used during development only, and has no functional
    '''     value for the application in production environment.
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function GetSelectedAnnotation( _
        ByVal pAnnotationLayer As IAnnotationLayer _
        ) As IAnnotationFeature

        Dim pSelectionSet As ISelectionSet
        Dim pCursor As ICursor = Nothing
        Dim pFeatureCursor As IFeatureCursor
        Dim pFeature As IFeature
        Dim pAnnotationFeature As IAnnotationFeature

        Try

            pSelectionSet = CType(pAnnotationLayer, IFeatureSelection).SelectionSet
            If pSelectionSet.Count < 1 Then
                MsgBox("Geen annotatie geselecteerd.")
                Throw New System.ApplicationException
            End If

            pSelectionSet.Search(Nothing, Nothing, pCursor)
            pFeatureCursor = CType(pCursor, IFeatureCursor)
            pFeature = pFeatureCursor.NextFeature

            pAnnotationFeature = CType(pFeature, IAnnotationFeature)

            GetSelectedAnnotation = pAnnotationFeature

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Get all linked annotation feature.
    ''' </summary>
    ''' <param name="annoLayer">
    '''     The annotation layer to search.
    ''' </param>
    ''' <param name="linkField">
    '''     The link attribute name.
    ''' </param>
    ''' <param name="linkValue">
    '''     The link ID string.
    ''' </param>
    ''' <returns>
    '''     List of annotation features.
    ''' </returns>
    ''' <remarks>
    '''     Nothing is returned if linkValue or linkField is empty.
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    ''' 	[Kristof Vydt]	23/09/2005	Return Nothing if annotation layer is Nothing.
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function GetLinkedAnnotations( _
        ByVal annoLayer As IAnnotationLayer, _
        ByVal linkField As String, _
        ByVal linkValue As String _
        ) As IList

        Dim annoList As New ArrayList
        Dim pFLayer As IFeatureLayer
        Dim pQueryFilter As IQueryFilter
        Dim pFCursor As IFeatureCursor
        Dim pFeature As IFeature

        Try
            'Do not start search if one of the parameters is empty.
            If annoLayer Is Nothing Then Return Nothing
            If Len(Trim(linkField)) = 0 Then Return Nothing
            If Len(Trim(linkValue)) = 0 Then Return Nothing

            'Check if there are single quotes in the LinkID.
            linkValue = Replace(linkValue, "'", "''")

            'Feature cursor based on LinkID.
            pFLayer = CType(annoLayer, IFeatureLayer)
            pQueryFilter = New QueryFilter
            pQueryFilter.WhereClause = linkField & "='" & linkValue & "'"
            pFCursor = pFLayer.Search(pQueryFilter, False)

            'List of annotation features.
            pFeature = pFCursor.NextFeature
            While Not pFeature Is Nothing
                annoList.Add(pFeature)
                pFeature = pFCursor.NextFeature
            End While

            Return annoList

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Update an existing annotation feature.
    ''' </summary>
    ''' <param name="annoFeat">
    '''     The annotation feature to be updated.
    ''' </param>
    ''' <param name="annoParams">
    '''     A hashtable of annotation parameters that need to be updated.
    ''' </param>
    ''' <remarks>
    '''     Currently, the following parameters are supported:
    '''     - TextString [string] The display text on the map.
    '''     - Angle      [double] The angle for displaying the text.
    '''     - LinkField  [string] The linking attribute name.
    '''     - LinkValue  [string] The linking value that is stored as attribute.
    '''     There is no validation of the values of these parameters.
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub UpdateAnno( _
        ByRef annoFeat As IAnnotationFeature, _
        ByVal annoParams As Hashtable)

        Dim TextString As String
        Dim Angle As Double
        Dim LinkID As String
        Dim FieldIndex As Integer

        Try
            If TypeOf annoFeat.Annotation Is ITextElement Then

                'Clone the text element.
                Dim pTElement As ITextElement = CType(annoFeat.Annotation, ITextElement)
                pTElement = CType(CType(pTElement, IClone).Clone, ITextElement)

                'Clone the text symbol.
                Dim pTSymbol As ITextSymbol = CType(pTElement.Symbol, ITextSymbol)
                pTSymbol = CType(CType(pTSymbol, IClone).Clone, ITextSymbol)

                'Set the text string.
                If annoParams.ContainsKey("TextString") Then
                    TextString = CStr(annoParams.Item("TextString"))
                    pTElement.Text = TextString
                End If

                'Set the angle.
                If annoParams.ContainsKey("Angle") Then
                    Angle = CDbl(annoParams.Item("Angle"))
                    pTSymbol.Angle = Angle
                End If

                'Assign the updated text symbol to the text element.
                pTElement.Symbol = pTSymbol

                'Assign the updated text element to the annotation feature.
                annoFeat.Annotation = CType(pTElement, IElement)

                'Store the updated feature.
                Dim pFeature As IFeature
                pFeature = CType(annoFeat, IFeature)

                'Assign LinkID.
                If annoParams.ContainsKey("LinkValue") And annoParams.ContainsKey("LinkField") Then
                    FieldIndex = pFeature.Fields.FindField(CStr(annoParams.Item("LinkField")))
                    If FieldIndex > -1 Then
                        LinkID = CStr(annoParams.Item("LinkValue"))
                        pFeature.Value(FieldIndex) = LinkID
                    End If
                End If

                'Save changes to feature.
                pFeature.Store()

            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Update all annotation features with specified LinkID.
    ''' </summary>
    ''' <param name="annoLayer">
    '''     The annotation feature layer to be updated.
    ''' </param>
    ''' <param name="annoParams">
    '''     A hashtable of annotation parameters that need to be updated.
    ''' </param>
    ''' <param name="linkField">
    '''     The linking attribute name.
    ''' </param>
    ''' <param name="linkValue">
    '''     The linking value to identify the annotation features.
    ''' </param>
    ''' <remarks>
    '''     Cfr UpdateAnno(ByRef AnnoFeat As IAnnotationFeature, ByVal Params As Hashtable)
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    ''' 	[Kristof Vydt]	23/09/2005	Added checks on parameters.
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub UpdateAnno( _
        ByRef annoLayer As IAnnotationLayer, _
        ByVal annoParams As Hashtable, _
        ByVal linkField As String, _
        ByVal linkValue As String)

        'Abort if one of the parameters is empty.
        If annoLayer Is Nothing Then Exit Sub
        If annoParams Is Nothing Then Exit Sub
        If annoParams.Count < 1 Then Exit Sub
        If Len(linkField) < 1 Then Exit Sub
        If Len(linkValue) < 1 Then Exit Sub

        'Update each matching annotation feature.
        Dim annoList As IList = GetLinkedAnnotations(annoLayer, linkField, linkValue)
        For Each annoFeature As IAnnotationFeature In annoList
            UpdateAnno(annoFeature, annoParams)
        Next

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Add an annotation feature with specified properties.
    ''' </summary>
    ''' <param name="annoLayer">
    '''     The annotation layer where the new annotation feature is to be added.
    ''' </param>
    ''' <param name="annoClassName">
    '''     The name of the annotation class.
    ''' </param>
    ''' <param name="symbolName">
    '''     The name of the annotation symbol.
    ''' </param>
    ''' <param name="pointGeom">
    '''     The location (as point geometry object) where to add the new annotation feature.
    ''' </param>
    ''' <param name="textString">
    '''     The visual text of the new annotation feature.
    ''' </param>
    ''' <param name="linkField">
    '''     The link attribute name of the annotation layer. If linkField is Nothing,
    '''     the link attribute of the new annotation feature, will have value Null.
    ''' </param>
    ''' <param name="linkValue">
    '''     The ID of a referencing feature. If linkValue is Nothing, 
    '''     the link attribute of the new annotation feature, will have value Null.
    ''' </param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    ''' 	[Kristof Vydt]	21/10/2005	Add annoClassName.
    ''' 	[Kristof Vydt]	24/10/2005	Add symbolName.
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub AddAnno( _
        ByRef annoLayer As IAnnotationLayer, _
        ByVal annoClassName As String, _
        ByVal symbolName As String, _
        ByVal pointGeom As IPoint, _
        ByVal textString As String, _
        ByVal linkField As String, _
        ByVal linkValue As String)

        Try
            'Create a new feature.
            Dim pFClass As IFeatureClass = CType(annoLayer, IFeatureLayer).FeatureClass
            Dim pFeature As IFeature = pFClass.CreateFeature

            'Assign AnnotationClassID.
            Dim pAnnotationFeature2 As IAnnotationFeature2 = CType(pFeature, IAnnotationFeature2)
            Dim annoClassID As Integer = GetAnnoClassID(annoLayer, annoClassName)
            pAnnotationFeature2.AnnotationClassID = annoClassID

            'Assign link ID.
            If (Not linkValue = Nothing) And (Not linkField = Nothing) Then
                Dim FieldIndex As Integer = pFeature.Fields.FindField(linkField)
                If FieldIndex > -1 Then pFeature.Value(FieldIndex) = linkValue
            End If

            'Create a text element.
            Dim pTextElement As ITextElement = New TextElement
            pTextElement.Text = textString

            'Give it the correct location.
            Dim pElement As IElement = CType(pTextElement, IElement)
            pElement.Geometry = CType(pointGeom, IGeometry)

            'Give it the correct symbology.
            Dim pGroupSymbolElement As IGroupSymbolElement = CType(pTextElement, IGroupSymbolElement)
            Dim symbolClassID As Integer = GetSymbolID(annoLayer, symbolName)
            pGroupSymbolElement.SymbolID = symbolClassID

            'Store the new annotation feature.
            Dim pAnnotationFeature As IAnnotationFeature = CType(pFeature, IAnnotationFeature)
            pAnnotationFeature.Annotation = pElement
            pFeature.Store()

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Try to found out some info on the existing annotation feature.
    ''' </summary>
    ''' <param name="pAnnotationFeature">
    '''     The annotation feature that has to be analysed.
    ''' </param>
    ''' <remarks>
    '''     This procedure is used during development only, and has no functional
    '''     value for the application in production environment.
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub AnalyseAnnotation( _
        ByVal pAnnotationFeature As IAnnotationFeature)

        '    Dim pGeometry As IGeometry
        Dim pElement As IElement
        Dim pTextElement As ITextElement
        Dim pSymbolCollectionElement As ISymbolCollectionElement

        Try
            ''Analyse geometry
            'pGeometry = pAnnotationFeature.Annotation.Geometry
            'Select Case pGeometry.GeometryType.ToString
            '    Case "esriGeometryPolygon"
            '        Dim pPolygon As IPolygon
            '        pPolygon = CType(pGeometry, IPolygon)
            '        MsgBox("Geometry type = Polygon" & vbNewLine & "Length = " & pPolygon.Length.ToString, , "Geometry")
            '    Case "esriGeometryPolyline"
            '        Dim pPolyline As IPolyline
            '        pPolyline = CType(pGeometry, IPolyline)
            '        MsgBox("Geometry type = Polyline" & vbNewLine & "Length = " & pPolyline.Length.ToString, , "Geometry")
            '    Case Else
            '        MsgBox("Geometry type = " & pGeometry.GeometryType.ToString, , "Geometry")
            'End Select

            'Analyse annotation text symbol element
            pElement = pAnnotationFeature.Annotation
            pTextElement = CType(pElement, ITextElement)
            pSymbolCollectionElement = CType(pElement, ISymbolCollectionElement)
            MsgBox("SymbolID  = " & pSymbolCollectionElement.SymbolID & vbNewLine & _
                   "Text      = " & pSymbolCollectionElement.Text & vbNewLine & _
                   "FlipAngle = " & pSymbolCollectionElement.FlipAngle _
                   , , "SymbolCollectionElement")

            'Analyse feature attributes.
            Dim pFeature As IFeature = CType(pAnnotationFeature, IFeature)
            Dim angleFieldIndex As Integer = pFeature.Fields.FindField("Angle")
            Dim flipAngleFieldIndex As Integer = pFeature.Fields.FindField("FlipAngle")
            Dim textFieldIndex As Integer = pFeature.Fields.FindField("TextString")
            MsgBox("TextString = " & CStr(pFeature.Value(textFieldIndex)) & vbNewLine & _
                   "Angle      = " & CStr(pFeature.Value(angleFieldIndex)) & vbNewLine & _
                   "FlipAngle  = " & CStr(pFeature.Value(flipAngleFieldIndex)) _
                   , , "Feature attributes")

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Modify all annotations currently linked to one LinkID, 
    '''     to link from now on to another LinkID.
    ''' </summary>
    ''' <param name="pAnnoLayer">
    '''     The annotation feature layer.
    ''' </param>
    ''' <param name="OldLinkID">
    '''     The current LinkID that has to be changed.
    ''' </param>
    ''' <param name="NewLinkID">
    '''     The LinkID to use from now on.
    ''' </param>
    ''' <returns>
    '''     The number of annotations that were altered.
    ''' </returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    ''' 	[Kristof Vydt]	02/08/2006	Replace some global constant by a function that reads the value from config file.
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function RelinkAnnotations( _
        ByRef pAnnoLayer As IAnnotationLayer, _
        ByVal OldLinkID As String, _
        ByVal NewLinkID As String _
        ) As Integer
        Try
            'Validation of input.
            If pAnnoLayer Is Nothing Then Throw New ArgumentNullException("pAnnoLayer")
            If Len(OldLinkID) = 0 Then Throw New ArgumentException("Noodzakelijk argument is leeg.", "OldLinkID")
            If Len(NewLinkID) = 0 Then Throw New ArgumentException("Noodzakelijk argument is leeg.", "NewLinkID")

            'Get cursor of features to be changed.
            Dim counter As Integer = 0 'the number of annotations that are linked to the new ID
            Dim pQueryFilter As IQueryFilter = New QueryFilter
            pQueryFilter.WhereClause = GetAttributeName("HydrantAnno", "LinkID") & "='" & OldLinkID & "'"
            Dim pFCursor As IFeatureCursor = CType(pAnnoLayer, IFeatureLayer).FeatureClass.Update(pQueryFilter, Nothing)
            Dim pFeature As IFeature = pFCursor.NextFeature
            If Not pFeature Is Nothing Then

                'Determine LinkID index.
                Dim FieldIndex As Integer = pFeature.Fields.FindField(GetAttributeName("HydrantAnno", "LinkID"))
                If FieldIndex < 0 Then Throw New AttributeNotFoundException(CType(pAnnoLayer, ILayer).Name, GetAttributeName("HydrantAnno", "LinkID"))

                'Loop through cursor and update LinkID.
                While Not pFeature Is Nothing
                    counter += 1
                    pFeature.Value(FieldIndex) = NewLinkID
                    pFCursor.UpdateFeature(pFeature)
                    pFeature = pFCursor.NextFeature
                End While
            End If

            'Return number of modified features.
            Return counter

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Delete all annotation features from the specified annotation layer,
    '''     which are linked to the specified ID.
    ''' </summary>
    ''' <param name="annoLayer">
    '''     The annotation layer to query.
    ''' </param>
    ''' <param name="linkField">
    '''     The link attribute name.
    ''' </param>
    ''' <param name="linkValue">
    '''     The link ID string.
    ''' </param>
    ''' <returns>
    '''     The number of removed annotation features.
    ''' </returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function RemoveLinkedAnnotations( _
        ByVal annoLayer As IAnnotationLayer, _
        ByVal linkField As String, _
        ByVal linkValue As String _
        ) As Integer

        Dim annoList As IList
        Dim annoFeat As IFeature
        Dim annoCounter As Integer = 0

        Try
            annoList = GetLinkedAnnotations(annoLayer, linkField, linkValue)
            For Each annoFeat In annoList
                annoFeat.Delete()
                annoCounter += 1
            Next
            Return annoCounter

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Return the AnnotationClassID that fits the class name.
    ''' </summary>
    ''' <param name="pAnnoLayer">
    '''     Annotation layer pointer.
    ''' </param>
    ''' <param name="sAnnoClassName">
    '''     The name of the annotation class.
    ''' </param>
    ''' <returns>
    '''     The ID of the annotation class.
    ''' </returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	21/10/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function GetAnnoClassID( _
            ByVal pAnnoLayer As IAnnotationLayer, _
            ByVal sAnnoClassName As String _
            ) As Integer

        Dim pGroupLayer As ICompositeLayer = CType(pAnnoLayer, ICompositeLayer)
        For i As Integer = 0 To pGroupLayer.Count - 1
            Dim pAnnoSubLayer As IAnnotationSublayer = CType(pGroupLayer.Layer(i), IAnnotationSublayer)
            If pGroupLayer.Layer(i).Name = sAnnoClassName Then
                GetAnnoClassID = pAnnoSubLayer.AnnotationClassID
                Exit Function
            End If
        Next
        GetAnnoClassID = 0 'in case annotation class name was not found
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Return the SymbolID that fits the name.
    ''' </summary>
    ''' <param name="pAnnoLayer">
    '''     Annotation layer pointer.
    ''' </param>
    ''' <param name="sSymbolName">
    '''     The name of the symbol.
    ''' </param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	24/10/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function GetSymbolID( _
                ByVal pAnnoLayer As IAnnotationLayer, _
                ByVal sSymbolName As String _
                ) As Integer

        Dim pFeatLayer As IFeatureLayer = CType(pAnnoLayer, IFeatureLayer)
        Dim pFeatClass As IFeatureClass = CType(pFeatLayer.FeatureClass, IFeatureClass)
        Dim pAnnoClass As IAnnoClass = CType(pFeatClass.Extension, IAnnoClass)
        Dim pSymbolCollection As ISymbolCollection = CType(pAnnoClass.SymbolCollection, ISymbolCollection)
        pSymbolCollection.Reset()
        Dim pSymbolIdentifier As ISymbolIdentifier2 = CType(pSymbolCollection.Next, ISymbolIdentifier2)
        While Not pSymbolIdentifier Is Nothing
            If pSymbolIdentifier.Name = sSymbolName Then
                GetSymbolID = pSymbolIdentifier.ID
                Exit Function
            End If
            pSymbolIdentifier = CType(pSymbolCollection.Next, ISymbolIdentifier2)
        End While
        GetSymbolID = 0 'in case symbol name was not found
    End Function

End Module
