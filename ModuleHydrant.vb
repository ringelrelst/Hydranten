Option Explicit On 
Option Strict On

Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Geometry

Module ModuleHydrant

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Create a new hydrant feature to add to the feature layer,
    '''     based on a single upload data row and an existing feature.
    ''' </summary>
    ''' <param name="featLayer">The hydrants feature layer.</param>
    ''' <param name="row">The hydrant upload data row.</param>
    ''' <param name="matchingFeature">The matching existing hydrant feature.</param>
    ''' <param name="providerCode">The code of the data row provider.</param>
    ''' <param name="statusCode">The status code for the new hydrant feature.</param>
    ''' <remarks>
    '''     If there is no matching feature available, matchingFeature is Nothing.
    '''     See also remark of AddHydrant().
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	22/09/2006	Created
    ''' 	[Elton Manoku]	03/12/2008	Coordinates are rounded with 3 digits after comma
    ''' 
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub CreateNewFeature( _
            ByVal featLayer As IFeatureLayer, _
            ByVal row As DataRow, _
            ByVal matchingFeature As IFeature, _
            ByVal providerCode As String, _
            ByVal statusCode As String)

        Dim attrHashTbl As New Hashtable 'attributes hash table
        Dim fldIdx As Integer            'attribute field index
        Dim tmpVal As Object
        Dim success As Boolean           'success or failure

        Try
            ' Combine info from data row and matching feature 
            ' to a hashtable of useful attributes.
            ' Start with info from the data row.

            '- BeginDatum (today)
            attrHashTbl.Add("BeginDatum", CDate(Now()))
            '- Bron
            attrHashTbl.Add("Bron", CStr(providerCode))
            '- CoordX
            attrHashTbl.Add("CoordX", Math.Round(CDbl(row.Item("CoordX")), 3))
            '- CoordY
            attrHashTbl.Add("CoordY", Math.Round(CDbl(row.Item("CoordY")), 3))
            '- Diameter (optional)
            tmpVal = row.Item("LeidingDiameter")
            If IsNumeric(tmpVal) Then attrHashTbl.Add("Diameter", CInt(tmpVal))
            '- EindDatum (null)
            '- HydrantType
            attrHashTbl.Add("HydrantType", CInt(row.Item("HydrantType")))
            '- LegendeCode (calculated as part of AddHydrant)
            '- LeidingNr (optional)
            tmpVal = Trim(CStr(row.Item("LeidingNummer")))
            If Len(tmpVal) > 0 Then attrHashTbl.Add("LeidingNr", tmpVal)
            '- LeidingType
            attrHashTbl.Add("LeidingType", CInt(row.Item("LeidingType")))
            '- LeverancierNr
            tmpVal = Trim(CStr(row.Item("LeverancierNummer")))
            attrHashTbl.Add("LeverancierNr", tmpVal)
            '- Status
            attrHashTbl.Add("Status", CStr(statusCode))

            ' Add additional info from matching feature (if available).
            If Not matchingFeature Is Nothing Then

                '- Aanduiding (optional)
                fldIdx = matchingFeature.Fields.FindField(GetAttributeName("Hydrant", "Aanduiding"))
                tmpVal = matchingFeature.Value(fldIdx)
                If Not TypeOf tmpVal Is System.DBNull Then _
                    If CStr(tmpVal) <> "" Then _
                        attrHashTbl.Add("Aanduiding", CStr(tmpVal))
                '- BrandweerNr (optional)
                fldIdx = matchingFeature.Fields.FindField(GetAttributeName("Hydrant", "BrandweerNr"))
                tmpVal = matchingFeature.Value(fldIdx)
                If Not TypeOf tmpVal Is System.DBNull Then _
                    If CStr(tmpVal) <> "" Then _
                        attrHashTbl.Add("BrandweerNr", CStr(tmpVal))
                '- Ligging (optional)
                fldIdx = matchingFeature.Fields.FindField(GetAttributeName("Hydrant", "Ligging"))
                tmpVal = matchingFeature.Value(fldIdx)
                If Not TypeOf tmpVal Is System.DBNull Then _
                    If CStr(tmpVal) <> "" Then _
                        attrHashTbl.Add("Ligging", CStr(tmpVal))
                '- Postcode (optional)
                fldIdx = matchingFeature.Fields.FindField(GetAttributeName("Hydrant", "Postcode"))
                tmpVal = matchingFeature.Value(fldIdx)
                If Not TypeOf tmpVal Is System.DBNull Then _
                    If CStr(tmpVal) <> "" Then _
                        attrHashTbl.Add("Postcode", CStr(tmpVal))
                '- Straatcode (optional)
                fldIdx = matchingFeature.Fields.FindField(GetAttributeName("Hydrant", "Straatcode"))
                tmpVal = matchingFeature.Value(fldIdx)
                If Not TypeOf tmpVal Is System.DBNull Then _
                    If CStr(tmpVal) <> "" Then _
                        attrHashTbl.Add("StraatCode", CStr(tmpVal))
                '- Straatnaam (optional)
                fldIdx = matchingFeature.Fields.FindField(GetAttributeName("Hydrant", "Straatnaam"))
                tmpVal = matchingFeature.Value(fldIdx)
                If Not TypeOf tmpVal Is System.DBNull Then _
                    If CStr(tmpVal) <> "" Then _
                        attrHashTbl.Add("StraatNaam", CStr(tmpVal))

            End If

            ' Add new feature defined by hashtable.
            success = AddHydrant(featLayer, attrHashTbl)

        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Add a new feature to the feature layer.
    ''' </summary>
    ''' <param name="FeatLayer">
    '''     The feature layer to add to.
    ''' </param>
    ''' <param name="Attributes">
    '''     A hashtable with all required attributes.
    ''' </param>
    ''' <returns>
    '''     Success or failure boolean.
    ''' </returns>
    ''' <remarks>
    '''     In case an exception is thrown in this function, false is returned.
    '''     Available attribute keys:
    '''     - Aanduiding    [string]  [optional]
    '''     - Diameter      [integer] [required]
    '''     - BrandweerNr   [integer] [optional]
    '''     - CoordX        [double]  [required]
    '''     - CoordY        [double]  [required]
    '''     - LeverancierNr [string]  [optional]
    '''     - LeidingType   [string]  [optional]
    '''     - LeidingNr     [string]  [optional]
    '''     - BeginDatum    [date]    [optional]
    '''     - EindDatum     [date]    [optional]
    '''     - StraatNaam    [string]  [optional]
    '''     - StraatCode    [string]  [optional]
    '''     - Postcode      [string]  [optional]
    '''     - Status        [string]  [required]
    '''     - Ligging       [string]  [required]
    '''     - HydrantType   [string]  [required]
    '''     - Bron          [string]  [optional]
    '''     The optional/required state is only based on the implementation of this 
    '''     function. Be aware of additional restrictions defined in the geodatabase.
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    ''' 	[Kristof Vydt]	02/08/2006	Replace some global constant by a function that reads the value from config file.
    ''' 	[Kristof Vydt]	22/02/2007	Private because only addressed by ModuleHydrant.CreateNewFeature().
    '''                                 Use HydrantLegendCode(feature) based on XML configuration.
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Function AddHydrant( _
            ByVal FeatLayer As IFeatureLayer, _
            ByVal Attributes As Hashtable _
            ) As Boolean

        Try

            'Create a new feature.
            Dim pFeature As IFeature = FeatLayer.FeatureClass.CreateFeature

            'Set coordinates.
            Dim pPoint As IPoint = New PointClass
            pPoint.X = CDbl(Attributes("CoordX"))
            pPoint.Y = CDbl(Attributes("CoordY"))
            pFeature.Shape = pPoint

            'Set all available attributes.
            If Attributes.ContainsKey("Aanduiding") Then pFeature.Value(pFeature.Fields.FindField(GetAttributeName("Hydrant", "Aanduiding"))) = CStr(Attributes.Item("Aanduiding"))
            If Attributes.ContainsKey("BeginDatum") Then pFeature.Value(pFeature.Fields.FindField(GetAttributeName("Hydrant", "BeginDatum"))) = FormatDateTime(CDate(Attributes.Item("BeginDatum")), DateFormat.ShortDate)
            If Attributes.ContainsKey("BrandweerNr") Then pFeature.Value(pFeature.Fields.FindField(GetAttributeName("Hydrant", "BrandweerNr"))) = CInt(Attributes.Item("BrandweerNr"))
            If Attributes.ContainsKey("Bron") Then pFeature.Value(pFeature.Fields.FindField(GetAttributeName("Hydrant", "Bron"))) = CStr(Attributes.Item("Bron"))
            If Attributes.ContainsKey("CoordX") Then pFeature.Value(pFeature.Fields.FindField(GetAttributeName("Hydrant", "CoordX"))) = CDbl(Attributes.Item("CoordX"))
            If Attributes.ContainsKey("CoordY") Then pFeature.Value(pFeature.Fields.FindField(GetAttributeName("Hydrant", "CoordY"))) = CDbl(Attributes.Item("CoordY"))
            If Attributes.ContainsKey("Diameter") Then pFeature.Value(pFeature.Fields.FindField(GetAttributeName("Hydrant", "Diameter"))) = CInt(Attributes.Item("Diameter"))
            If Attributes.ContainsKey("EindDatum") Then pFeature.Value(pFeature.Fields.FindField(GetAttributeName("Hydrant", "EindDatum"))) = FormatDateTime(CDate(Attributes.Item("EindDatum")), DateFormat.ShortDate)
            If Attributes.ContainsKey("HydrantType") Then pFeature.Value(pFeature.Fields.FindField(GetAttributeName("Hydrant", "HydrantType"))) = CStr(Attributes.Item("HydrantType"))
            If Attributes.ContainsKey("LeidingNr") Then pFeature.Value(pFeature.Fields.FindField(GetAttributeName("Hydrant", "LeidingNr"))) = CStr(Attributes.Item("LeidingNr"))
            If Attributes.ContainsKey("LeidingType") Then pFeature.Value(pFeature.Fields.FindField(GetAttributeName("Hydrant", "LeidingType"))) = CStr(Attributes.Item("LeidingType"))
            If Attributes.ContainsKey("LeverancierNr") Then pFeature.Value(pFeature.Fields.FindField(GetAttributeName("Hydrant", "LeverancierNr"))) = CStr(Attributes.Item("LeverancierNr"))
            If Attributes.ContainsKey("Ligging") Then pFeature.Value(pFeature.Fields.FindField(GetAttributeName("Hydrant", "Ligging"))) = CStr(Attributes.Item("Ligging"))
            If Attributes.ContainsKey("Postcode") Then pFeature.Value(pFeature.Fields.FindField(GetAttributeName("Hydrant", "Postcode"))) = CStr(Attributes.Item("Postcode"))
            If Attributes.ContainsKey("Status") Then pFeature.Value(pFeature.Fields.FindField(GetAttributeName("Hydrant", "Status"))) = CStr(Attributes.Item("Status"))
            If Attributes.ContainsKey("StraatCode") Then pFeature.Value(pFeature.Fields.FindField(GetAttributeName("Hydrant", "Straatcode"))) = CStr(Attributes.Item("StraatCode"))
            If Attributes.ContainsKey("StraatNaam") Then pFeature.Value(pFeature.Fields.FindField(GetAttributeName("Hydrant", "Straatnaam"))) = CStr(Attributes.Item("StraatNaam"))

            'Set LegendCode attribute.
            Dim FieldIndex As Integer = pFeature.Fields.FindField(GetAttributeName("Hydrant", "Legende"))
            pFeature.Value(FieldIndex) = HydrantLegendCode(pFeature)

            'Save the new feature.
            pFeature.Store()

            'Return success.
            Return True

        Catch ex As Exception
            Throw ex
        End Try

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Change attributes of a given hydrant feature, based on an attributes list.
    ''' </summary>
    ''' <param name="pFeature">
    '''     The feature to be changed.
    ''' </param>
    ''' <param name="Attributes">
    '''     The list of attributes and their values.
    ''' </param>
    ''' <returns>
    '''     Success or failure boolean.
    ''' </returns>
    ''' <remarks>
    '''     In case an exception is thrown in this function, false is returned.
    '''     Available attribute keys:
    '''     - Aanduiding    [string]
    '''     - Diameter      [integer]
    '''     - BrandweerNr   [integer]
    '''     - CoordX        [double]
    '''     - CoordY        [double]
    '''     - LeverancierNr [string]
    '''     - LeidingType   [string]
    '''     - LeidingNr     [string]
    '''     - BeginDatum    [date]  
    '''     - EindDatum     [date]  
    '''     - StraatNaam    [string]
    '''     - StraatCode    [string]
    '''     - Postcode      [string]
    '''     - Status        [string]
    '''     - Ligging       [string]
    '''     - HydrantType   [string]
    '''     - Bron          [string]
    '''     Be aware of restrictions defined in the geodatabase.
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    ''' 	[Kristof Vydt]	02/08/2006	Replace some global constant by a function that reads the value from config file.
    ''' 	[Kristof Vydt]	22/02/2007	Use UpdateLegendCode(feature) instead of HydrantLegendCode().
    ''' 	[Elton Manoku]	28/11/2008	If the einddatum is smaller than the begindatum then set the einddatum equal to begindatum.
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function ModifyHydrantAttributes( _
            ByVal pFeature As IFeature, _
            ByVal Attributes As Hashtable _
            ) As Boolean

        Try

            'Set all available attributes.
            If Attributes.ContainsKey("Aanduiding") Then pFeature.Value(pFeature.Fields.FindField(GetAttributeName("Hydrant", "Aanduiding"))) = CStr(Attributes.Item("Aanduiding"))
            If Attributes.ContainsKey("Diameter") Then pFeature.Value(pFeature.Fields.FindField(GetAttributeName("Hydrant", "Diameter"))) = CInt(Attributes.Item("Diameter"))
            If Attributes.ContainsKey("BrandweerNr") Then pFeature.Value(pFeature.Fields.FindField(GetAttributeName("Hydrant", "BrandweerNr"))) = CInt(Attributes.Item("BrandweerNr"))
            If Attributes.ContainsKey("CoordX") Then pFeature.Value(pFeature.Fields.FindField(GetAttributeName("Hydrant", "CoordX"))) = CDbl(Attributes.Item("CoordX"))
            If Attributes.ContainsKey("CoordY") Then pFeature.Value(pFeature.Fields.FindField(GetAttributeName("Hydrant", "CoordY"))) = CDbl(Attributes.Item("CoordY"))
            If Attributes.ContainsKey("LeverancierNr") Then pFeature.Value(pFeature.Fields.FindField(GetAttributeName("Hydrant", "LeverancierNr"))) = CStr(Attributes.Item("LeverancierNr"))
            If Attributes.ContainsKey("LeidingType") Then pFeature.Value(pFeature.Fields.FindField(GetAttributeName("Hydrant", "LeidingType"))) = CStr(Attributes.Item("LeidingType"))
            If Attributes.ContainsKey("LeidingNr") Then pFeature.Value(pFeature.Fields.FindField(GetAttributeName("Hydrant", "LeidingNr"))) = CStr(Attributes.Item("LeidingNr"))
            If Attributes.ContainsKey("BeginDatum") Then pFeature.Value(pFeature.Fields.FindField(GetAttributeName("Hydrant", "BeginDatum"))) = FormatDateTime(CDate(Attributes.Item("BeginDatum")), DateFormat.ShortDate)
            If Attributes.ContainsKey("EindDatum") Then pFeature.Value(pFeature.Fields.FindField(GetAttributeName("Hydrant", "EindDatum"))) = FormatDateTime(CDate(Attributes.Item("EindDatum")), DateFormat.ShortDate)
            If Attributes.ContainsKey("StraatNaam") Then pFeature.Value(pFeature.Fields.FindField(GetAttributeName("Hydrant", "Straatnaam"))) = CStr(Attributes.Item("StraatNaam"))
            If Attributes.ContainsKey("StraatCode") Then pFeature.Value(pFeature.Fields.FindField(GetAttributeName("Hydrant", "Straatcode"))) = CStr(Attributes.Item("StraatCode"))
            If Attributes.ContainsKey("Postcode") Then pFeature.Value(pFeature.Fields.FindField(GetAttributeName("Hydrant", "Postcode"))) = CStr(Attributes.Item("Postcode"))
            If Attributes.ContainsKey("Status") Then pFeature.Value(pFeature.Fields.FindField(GetAttributeName("Hydrant", "Status"))) = CStr(Attributes.Item("Status"))
            If Attributes.ContainsKey("Ligging") Then pFeature.Value(pFeature.Fields.FindField(GetAttributeName("Hydrant", "Ligging"))) = CStr(Attributes.Item("Ligging"))
            If Attributes.ContainsKey("HydrantType") Then pFeature.Value(pFeature.Fields.FindField(GetAttributeName("Hydrant", "HydrantType"))) = CStr(Attributes.Item("HydrantType"))
            If Attributes.ContainsKey("Bron") Then pFeature.Value(pFeature.Fields.FindField(GetAttributeName("Hydrant", "Bron"))) = CStr(Attributes.Item("Bron"))

            ' Update Legend attribute.
            Call UpdateLegendCode(pFeature)

            ' Save the new feature.
            pFeature.Store()

            ' Return success.
            Return True

        Catch ex As Exception
            Throw ex

            'Return failure.
            Return False
        End Try
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Update the legende code attribute value of a hydrant feature.
    ''' </summary>
    ''' <param name="hydrant">The hydrant feature to evaluate.</param>
    ''' <returns>True if legend attribute recieved a new value.</returns>
    ''' <remarks>
    '''     The required edit session management is not included in this method.
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	22/09/2006	Created
    ''' 	[Kristof Vydt]	13/10/2006	Correct for the fact that diameter attribute value type is Short.
    ''' 	[Kristof Vydt]	22/02/2007	Use HydrantLegendCode(feature) based on XML configuration.
    '''                                 Return True if new value is stored, instead of returning the updated attribute value.
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function UpdateLegendCode( _
            ByRef hydrant As IFeature _
            ) As Boolean

        Try
            '' Read attribute values.
            'attrStatus = GetAttributeValue(hydrantFeature, "Hydrant", "Status")
            'attrHydrType = GetAttributeValue(hydrantFeature, "Hydrant", "Hydranttype")
            'attrLigging = GetAttributeValue(hydrantFeature, "Hydrant", "Ligging")
            'attrDiameter = GetAttributeValue(hydrantFeature, "Hydrant", "Diameter")

            '' Set specific values in case of wrong type (e.g. System.DBNull).
            'If Not TypeOf attrStatus Is String Then attrStatus = ""
            'If Not TypeOf attrHydrType Is String Then attrHydrType = ""
            'If Not TypeOf attrLigging Is String Then attrLigging = ""
            'If Not IsNumeric(attrDiameter) Then attrDiameter = 0
            'If Not ((TypeOf attrDiameter Is Integer) Or (TypeOf attrDiameter Is Short)) Then attrDiameter = 0

            '' Determine the updated legend code attribute.
            'attrLegende = HydrantLegendCodeEx(CStr(attrStatus), _
            '    CStr(attrHydrType), CStr(attrLigging), CInt(attrDiameter))

            '' Update and return the legend code attribute for current feature.
            'Return SetAttributeValue(hydrantFeature, "Hydrant", "Legende", attrLegende)

            ' Current value for the legend attribute.
            Dim current As Integer = CInt(GetAttributeValue(hydrant, "Hydrant", "Legende"))

            ' Determine updated value for the legend attribute.
            Dim updated As Integer = HydrantLegendCode(hydrant)

            ' Compare both values.
            If current = updated Then
                ' No update of the attribute required.
                Return False
            Else
                ' Set the updated attribute value.
                SetAttributeValue(hydrant, "Hydrant", "Legende", updated)
                Return True
            End If

        Catch ex As Exception
            ErrorHandler(ex)
        End Try

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Determine the legend code for 1 hydrant feature.
    ''' </summary>
    ''' <param name="hydrant">a hydrant feature</param>
    ''' <returns>legend code integer</returns>
    ''' <remarks>
    '''     Decision algorithm based on rules defined in configuration file.
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    ''' 	[Kristof Vydt]	24/10/2005	Read default value from ini file.
    '''     [Kristof Vydt]  21/02/2007  Rewrite to use XML config file.
    ''' 	[Kristof Vydt]	09/03/2007	Modify to support result formula.
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function HydrantLegendCode( _
        ByVal hydrant As IFeature) As Integer

        Try

            ' Get the list of decision rules from the configuration.
            Dim rules As Collection = config.LegendRules

            ' Loop through each rule until one is applicable.
            For Each rule As LegendRule In rules

                ' If all conditions for this rule are met,
                ' then return result value of this rule.
                If rule.Comply(hydrant) Then
                    Select Case rule.ResultType
                        Case LegendRule.ResultTypeEnumType.FixedValue

                            ' Fixed value result.
                            Return rule.ResultValue

                        Case LegendRule.ResultTypeEnumType.AttributeBasedCalculation

                            ' Feature attribute value.
                            Dim attrIdx As Integer = hydrant.Fields.FindField(rule.SeedReference)
                            Dim attrVal As Object = hydrant.Value(attrIdx)

                            ' Apply formula calculations.
                            Dim seed As Object = attrVal
                            For Each calc As Calculation In rule.Calculations
                                seed = calc.Execute(seed)
                            Next

                            ' Return attribute based formula result.
                            Return Convert.ToInt16(seed)

                    End Select
                End If

            Next rule

            ' No matching rule. - Return default legend code.
            Dim defaultCode As Integer = config.DefaultLegendCode
            Return defaultCode

        Catch ex As Exception
            Throw ex
        End Try

    End Function

    'Public Function HydrantLegendCode( _
    '    ByVal StatusCode As String, _
    '    ByVal hydrantTypeCode As String, _
    '    ByVal LiggingCode As String, _
    '    ByVal Diameter As Integer _
    '    ) As Integer

    '    Try
    '        Dim Code As String = "" 'the final value to return
    '        Dim TmpStr As String 'temporary string used for holding ini-file values
    '        Dim Rules As String() 'array of rule keys from ini-file
    '        Dim RuleDef As String() 'array of (1 or more) values for one rule
    '        Dim NumberOfRules As Integer
    '        Dim i As Integer 'loop index

    '        'Set a default value for the legend code.
    '        Code = INIRead(g_FilePath_Config, "HydrantenLegendRules", "Default")

    '        'Get the list of decision rules from the Config.ini file.
    '        TmpStr = INIRead(g_FilePath_Config, "HydrantenLegendRules") ' get all keys in section
    '        TmpStr = TmpStr.Replace(ControlChars.NullChar, "|"c) ' change embedded NULLs to pipe chars
    '        Rules = TmpStr.Split("|"c)
    '        NumberOfRules = Rules.Length

    '        'Loop through this list, one by one, until a matching rule is found.
    '        For i = 0 To NumberOfRules - 1

    '            'Filter out the line that defines the default value, instead of a rule.
    '            If Rules(i) <> "Default" Then

    '                'Retrieve the definition of the current rule.
    '                TmpStr = INIRead(g_FilePath_Config, "HydrantenLegendRules", Rules(i))
    '                RuleDef = TmpStr.Split(";"c)

    '                'Check the Status condition of current rule.
    '                If StatusCode = RuleDef(0) Then
    '                    'Check if there is a second condition.
    '                    If RuleDef.Length > 1 Then

    '                        'Check the HydrantType condition of current rule.
    '                        If hydrantTypeCode = RuleDef(1) Then
    '                            'Check if there is a third condition.
    '                            If RuleDef.Length > 2 Then

    '                                'Check the Ligging condition of current rule.
    '                                If LiggingCode = RuleDef(2) Then
    '                                    'Rule does match. Continue with legend code of current rule.
    '                                    Code = Rules(i)
    '                                    Exit For
    '                                End If

    '                            Else 'There is no third condition.
    '                                'Rule does match. Continue with legend code of current rule.
    '                                Code = Rules(i)
    '                                Exit For
    '                            End If
    '                        End If

    '                    Else 'There is no second condition.
    '                        'Rule does match. Continue with legend code of current rule.
    '                        Code = Rules(i)
    '                        Exit For
    '                    End If 'end of If RuleDef.Length > 1
    '                End If 'end of If StatusCode = RuleDef(0)
    '            End If 'end of If Rules(i) <> "Default"
    '        Next

    '        'Does the legend code require any processing ?
    '        'Return the resulting value.
    '        Try
    '            Select Case Code
    '                Case "0k+d"
    '                    Return CInt(Diameter)
    '                Case "1k+d"
    '                    Return 1000 + CInt(Diameter)
    '                Case "2k+d"
    '                    Return 2000 + CInt(Diameter)
    '                Case "3k+d"
    '                    Return 3000 + CInt(Diameter)
    '                Case "4k+d"
    '                    Return 4000 + CInt(Diameter)
    '                Case "5k+d"
    '                    Return 5000 + CInt(Diameter)
    '                Case "6k+d"
    '                    Return 6000 + CInt(Diameter)
    '                Case "7k+d"
    '                    Return 7000 + CInt(Diameter)
    '                Case "8k+d"
    '                    Return 8000 + CInt(Diameter)
    '                Case "9k+d"
    '                    Return 9000 + CInt(Diameter)
    '                Case Else
    '                    Return CInt(Code)
    '            End Select
    '        Catch
    '            'Use default legend code value in case no matching rule could be found, or in case of error.
    '            Return CInt(INIRead(g_FilePath_Config, "HydrantenLegendRules", "Default"))
    '        End Try

    '    Catch ex As Exception
    '        Throw ex
    '    End Try

    'End Function

End Module
