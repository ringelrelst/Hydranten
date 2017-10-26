Option Explicit On 
Option Strict On

#Region "Import namaspaces"

Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI

#End Region

''' -----------------------------------------------------------------------------
''' Project	 : Digipolis.Hydranten.BeheerHydranten
''' Class	 : Hydranten.BeheerHydranten.AppSettings
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Class to access application settings.
''' </summary>
''' <remarks>
''' Settings are read from Config.xml. Also the ArcMap mxd document is used.
''' </remarks>
''' <history>
''' 	[Kristof Vydt]	21/02/2007	Created
''' 	[Kristof Vydt]	08/03/2007	Validate configuration XML against XSD.
'''                                 Rewrite parts after restructuring XSD.
''' 	[Kristof Vydt]	14/03/2007	Correct AttributeName.
''' 	[Kristof Vydt]	21/03/2007	Throw exception when missing configuration setting.
'''     [Koen Vermeer]  23/06/2008  Conversion to ArcGIS 9.2
''' </history>
''' -----------------------------------------------------------------------------
Public Class AppSettings

#Region "Private variables"
    Private _mxdFileName As String          'loacl var to hold the mxd file name without extension
    Private _configFilePath As String       'local var to hold the config file path
    Private _configSchemaFilePath As String 'local var to hold the config schema file path
#End Region

#Region "Public properties"

    Public ReadOnly Property SectorName() As String
        Get
            SectorName = ""

            ' Load the xml file.
            Dim xmlDoc As New Xml.XmlDocument
            xmlDoc.Load(_configFilePath)

            ' Query for a value.
            Dim node As Xml.XmlNode = xmlDoc.DocumentElement.SelectSingleNode( _
                "/configuration/appSettings/sectorList/sector[mxdFile=""" & _mxdFileName & """]/sectorName")

            ' Return the value or nothing if it doesn't exist.
            If Not node Is Nothing Then Return node.InnerText

            ' Missing configuration setting.
            If node Is Nothing Then Throw New IncompleteConfigurationException( _
                "Geen sectornaam gevonden voor huidig document (file '" & _mxdFileName & "') in het configuratiebestand.")

        End Get
    End Property

    Public ReadOnly Property SectorCode() As String
        Get
            SectorCode = Nothing

            ' Load the xml file.
            Dim xmlDoc As New Xml.XmlDocument
            xmlDoc.Load(_configFilePath)

            ' Query for a value.
            Dim node As Xml.XmlNode = xmlDoc.DocumentElement.SelectSingleNode( _
                "/configuration/appSettings/sectorList/sector[mxdFile=""" & _mxdFileName & """]/sectorCode")

            ' Return the value or nothing if it doesn't exist.
            If Not node Is Nothing Then Return node.InnerText

            ' Missing configuration setting.
            If node Is Nothing Then Throw New IncompleteConfigurationException( _
                "Geen sectorcode gevonden voor huidig document (key '" & _mxdFileName & "') in het configuratiebestand.")

        End Get
    End Property

    Public ReadOnly Property MxdFile() As String
        Get
            Return _mxdFileName
        End Get
        'Set(ByVal Value As String)
        '    _mxdFileName = Value
        '    'TODO: Throw an exception if the mxd file is not bound to any sector.
        'End Set
    End Property

    Public ReadOnly Property Postcodes() As Collection
        Get

            ' Load the xml file.
            Dim xmlDoc As New Xml.XmlDocument
            xmlDoc.Load(_configFilePath)

            ' Query for a sector node.
            Dim node As Xml.XmlNode = xmlDoc.DocumentElement.SelectSingleNode( _
                "/configuration/appSettings/sectorList/sector[mxdFile=""" & _mxdFileName & """]")
            If node Is Nothing Then Return Nothing

            ' Query for a list of postcodes.
            Dim nodeList As Xml.XmlNodeList = node.SelectNodes("postcode")

            ' Fill collection with postal codes.
            Dim coll As New Collection
            If nodeList.Count > 0 Then
                Dim ienum As IEnumerator = nodeList.GetEnumerator
                While (ienum.MoveNext)
                    Dim postcode As Xml.XmlNode = CType(ienum.Current, Xml.XmlNode)
                    coll.Add(postcode.InnerText)
                End While
            End If

            ' Return collection.
            Return coll

        End Get
    End Property

    Public ReadOnly Property BuildingLayers() As Collection
        Get

            ' Load the xml file.
            Dim xmlDoc As New Xml.XmlDocument
            xmlDoc.Load(_configFilePath)

            ' Query for a list of layer names.
            Dim nodeList As Xml.XmlNodeList = xmlDoc.SelectNodes( _
                "/configuration/appSettings/buildingList/layer")

            ' Fill collection of layer names.
            Dim coll As New Collection
            If nodeList.Count > 0 Then
                Dim ienum As IEnumerator = nodeList.GetEnumerator
                While (ienum.MoveNext)
                    Dim layer As Xml.XmlNode = CType(ienum.Current, Xml.XmlNode)
                    If Not layer.Attributes("name") Is Nothing Then
                        coll.Add(layer.Attributes("name").Value)
                    ElseIf Not layer.Attributes("key") Is Nothing Then
                        coll.Add(GetLayerName(layer.Attributes("key").Value))
                    End If
                End While
            End If

            ' Return collection.
            Return coll

        End Get
    End Property

    Public ReadOnly Property DangerLayers() As Collection
        Get

            ' Load the xml file.
            Dim xmlDoc As New Xml.XmlDocument
            xmlDoc.Load(_configFilePath)

            ' Query for a list of layer names.
            Dim nodeList As Xml.XmlNodeList = xmlDoc.SelectNodes( _
                "/configuration/appSettings/dangerList/layer")

            ' Fill collection of layer names.
            Dim coll As New Collection
            If nodeList.Count > 0 Then
                Dim ienum As IEnumerator = nodeList.GetEnumerator
                While (ienum.MoveNext)
                    Dim layer As Xml.XmlNode = CType(ienum.Current, Xml.XmlNode)
                    If Not layer.Attributes("name") Is Nothing Then
                        coll.Add(layer.Attributes("name").Value)
                    ElseIf Not layer.Attributes("key") Is Nothing Then
                        coll.Add(GetLayerName(layer.Attributes("key").Value))
                    End If
                End While
            End If

            ' Return collection.
            Return coll

        End Get
    End Property

    Public ReadOnly Property LegendRules() As Collection
        Get

            ' Load the xml file.
            Dim xmlDoc As New Xml.XmlDocument
            xmlDoc.Load(_configFilePath)

            ' Query for a list of <legendRule> nodes.
            Dim nodeList As Xml.XmlNodeList = xmlDoc.SelectNodes( _
                "/configuration/appSettings/legendRuleList/legendRule")

            ' Add each <legendRule> to a collection.
            Dim coll As New Collection
            If nodeList.Count > 0 Then
                Dim ienum As IEnumerator = nodeList.GetEnumerator
                While (ienum.MoveNext)
                    Dim elem As Xml.XmlElement = CType(ienum.Current, Xml.XmlElement)
                    Dim rule As LegendRule = New LegendRule(elem)
                    coll.Add(rule)
                End While
            End If

            ' Return collection.
            Return coll

        End Get
    End Property

    Public ReadOnly Property DefaultLegendCode() As Integer
        Get

            ' Load the xml file.
            Dim xmlDoc As New Xml.XmlDocument
            xmlDoc.Load(_configFilePath)

            ' Query for <defaultLegend>.
            Dim node As Xml.XmlNode = xmlDoc.SelectSingleNode( _
                "/configuration/appSettings/legendRuleList[@defaultValue]")

            ' Return the value or nothing if it doesn't exist.
            If Not node Is Nothing Then Return CInt(node.Attributes("defaultValue").Value)

            ' Missing configuration setting.
            If node Is Nothing Then Throw New IncompleteConfigurationException( _
                "Geen default legende code gevonden in het configuratiebestand.")

        End Get
    End Property

#End Region

#Region "Public methods"

    ' Constructor
    <CLSCompliant(False)> _
    Public Sub New( _
        ByVal mxApp As IMxApplication)

        InitializeConfigFile(mxApp)

    End Sub

    ' Get the name of a map layer.
    Public Function LayerName( _
        ByVal layerKey As String) As String

        LayerName = Nothing

        ' Load the xml file.
        Dim xmlDoc As New Xml.XmlDocument
        xmlDoc.Load(_configFilePath)

        ' Query for a sector specific value.
        Dim node As Xml.XmlNode = xmlDoc.DocumentElement.SelectSingleNode( _
            "/configuration/appSettings/sectorList/sector[mxdFile=""" & _mxdFileName & """]/layerList/layer[@key=""" & layerKey & """]")

        ' Return the value or continu if it doesn't exist.
        If Not node Is Nothing Then
            Dim attrib As Xml.XmlAttribute = node.Attributes("name")
            If Not attrib Is Nothing Then Return attrib.Value
        End If

        ' Query for a global value.
        node = xmlDoc.DocumentElement.SelectSingleNode( _
            "/configuration/appSettings/layerList/layer[@key=""" & layerKey & """]")

        ' Return the value or nothing if it doesn't exist.
        If Not node Is Nothing Then
            Dim attrib As Xml.XmlAttribute = node.Attributes("name")
            If Not attrib Is Nothing Then Return attrib.Value
        End If

        ' Missing configuration setting.
        If node Is Nothing Then Throw New IncompleteConfigurationException( _
            "Geen laag (key '" & layerKey & "') in het configuratiebestand.")

    End Function

    ' Get the name of an attribute field.
    Public Function AttributeName( _
            ByVal layerKey As String, _
            ByVal fieldKey As String) As String

        AttributeName = ""

        ' Load the xml file.
        Dim xmlDoc As New Xml.XmlDocument
        xmlDoc.Load(_configFilePath)

        ' Query for a sector specific value.
        Dim node As Xml.XmlNode = xmlDoc.DocumentElement.SelectSingleNode( _
            "/configuration/appSettings/sectorList/sector[mxdFile=""" & _mxdFileName & """]/layerList/layer[@key=""" & layerKey & """]/attribute[@key=""" & fieldKey & """]")

        ' Return the value or continu if it doesn't exist.
        If Not node Is Nothing Then
            Dim attrib As Xml.XmlAttribute = node.Attributes("name")
            If Not attrib Is Nothing Then Return attrib.Value
        End If

        ' Query for a global value.
        node = xmlDoc.DocumentElement.SelectSingleNode( _
            "/configuration/appSettings/layerList/layer[@key=""" & layerKey & """]/attribute[@key=""" & fieldKey & """]")

        ' Return the value or nothing if it doesn't exist.
        If Not node Is Nothing Then
            Dim attrib As Xml.XmlAttribute = node.Attributes("name")
            If Not attrib Is Nothing Then Return attrib.Value
        End If

        ' Missing configuration setting.
        If node Is Nothing Then Throw New IncompleteConfigurationException( _
            "Geen attribuut (key '" & fieldKey & "') gevonden voor laag (key '" & layerKey & "') in het configuratiebestand.")

    End Function

    ' Get the name of a attribute domain.
    Public Function DomainName( _
        ByVal domainKey As String) As String

        DomainName = ""

        ' Load the xml file.
        Dim xmlDoc As New Xml.XmlDocument
        xmlDoc.Load(_configFilePath)

        ' Query for a global value.
        Dim node As Xml.XmlNode = xmlDoc.DocumentElement.SelectSingleNode( _
            "/configuration/appSettings/domainList/domain[@key=""" & domainKey & """]")

        ' Return the value or nothing if it doesn't exist.
        If Not node Is Nothing Then
            Dim attrib As Xml.XmlAttribute = node.Attributes("name")
            If Not attrib Is Nothing Then Return attrib.Value
        End If

        ' Missing configuration setting.
        If node Is Nothing Then Throw New IncompleteConfigurationException( _
            "Geen domein (key '" & domainKey & "') gevonden in het configuratiebestand.")

    End Function

    ' Get list of visible and hidden map layers.
    Public Function QueryLayerVisibility() As Hashtable

        ' Load the xml file.
        Dim xmlDoc As New Xml.XmlDocument
        xmlDoc.Load(_configFilePath)

        ' Query sector-specific layer settings.
        Dim nodeList As Xml.XmlNodeList = xmlDoc.SelectNodes( _
            "/configuration/appSettings/sectorList/sector[mxdFile=""" & _mxdFileName & """]/layerList/layer")

        ' Built a hashtable of layer names.
        Dim tbl As New Hashtable
        If nodeList.Count > 0 Then
            Dim ienum As IEnumerator = nodeList.GetEnumerator
            While (ienum.MoveNext)

                ' A layer without specification of visibility is of no importance here.
                Dim layer As Xml.XmlElement = CType(ienum.Current, Xml.XmlElement)
                If layer.HasAttribute("visible") Then

                    ' Determine layer visibility.
                    Dim visible As Boolean = Convert.ToBoolean(layer.GetAttribute("visible"))

                    ' Determine layer name.
                    Dim name As String = String.Empty
                    If layer.HasAttribute("name") Then
                        name = layer.GetAttribute("name")
                    ElseIf layer.HasAttribute("key") Then
                        name = Me.LayerName(layer.GetAttribute("key"))
                    End If

                    ' Add layer name to hashtable.
                    If Not name Is Nothing Then tbl.Add(name, visible)

                End If
            End While
        End If

        ' Query global settings.
        nodeList = xmlDoc.SelectNodes( _
            "/configuration/appSettings/layerList/layer")

        ' Expand the hashtable with missing layers from global settings.
        If nodeList.Count > 0 Then
            Dim ienum As IEnumerator = nodeList.GetEnumerator
            While (ienum.MoveNext)

                ' A layer without specification of visibility is of no importance here.
                Dim layer As Xml.XmlElement = CType(ienum.Current, Xml.XmlElement)
                If layer.HasAttribute("visible") Then

                    ' Determine layer visibility.
                    Dim visible As Boolean = Convert.ToBoolean(layer.GetAttribute("visible"))

                    ' Determine layer name.
                    Dim name As String = String.Empty
                    If layer.HasAttribute("name") Then
                        name = layer.GetAttribute("name")
                    ElseIf layer.HasAttribute("key") Then
                        name = Me.LayerName(layer.GetAttribute("key"))
                    End If

                    ' Add new layer name to hashtable.
                    If Not name Is Nothing AndAlso _
                        Not tbl.ContainsKey(name) Then _
                            tbl.Add(name, visible)

                End If

            End While
        End If

        ' Return hashtable.
        Return tbl

    End Function

#End Region

#Region "Private methods"

    ' Initialize the apps config file
    Private Sub InitializeConfigFile( _
        ByVal mxApp As IMxApplication)

        Try

            ' Store config file path in local var.
            Dim sb As New System.Text.StringBuilder
            sb.Append(GetMxdFolder(mxApp))
            sb.Append("\")
            sb.Append(c_FileName_Config)
            _configFilePath = sb.ToString

            ' Throw an exception if the configuration file does not exist.
            If Not System.IO.File.Exists(_configFilePath) Then _
                Throw New FileNotFoundException(_configFilePath)

            ' Store config schema file path in local var.
            sb = New System.Text.StringBuilder
            sb.Append(GetMxdFolder(mxApp))
            sb.Append("\")
            sb.Append(c_FileName_ConfigSchema)
            _configSchemaFilePath = sb.ToString

            ' Throw an exception if the configuration schema file does not exist.
            If Not System.IO.File.Exists(_configSchemaFilePath) Then _
                Throw New FileNotFoundException(_configSchemaFilePath)

        Catch ex As Exception
            ErrorHandler(ex)
        End Try

        Try
            ' Read the config file.
            Dim oXmlTextReader As Xml.XmlTextReader
            oXmlTextReader = New Xml.XmlTextReader(_configFilePath)

            ' Validator.
            Dim oXmlValReader As Xml.XmlValidatingReader
            oXmlValReader = New Xml.XmlValidatingReader(oXmlTextReader)
            oXmlValReader.ValidationType = Xml.ValidationType.Schema
            AddHandler oXmlValReader.ValidationEventHandler, _
                AddressOf Me.ValidationError

            ' Link config file to the schema file.
            Dim oSchemaCollection As New Xml.Schema.XmlSchemaCollection
            oSchemaCollection.Add(Nothing, _configSchemaFilePath)
            oXmlValReader.Schemas.Add(Nothing, _configSchemaFilePath)

            oXmlTextReader = Nothing
            oXmlValReader = Nothing
            oSchemaCollection = Nothing

        Catch ex As Exception
            ErrorHandler(ex)
        End Try

        ' Store mxd file name in local var.
        _mxdFileName = GetMxdFile(mxApp)

    End Sub

    ' Configuration schema validation error
    Private Sub SchemaValidationError( _
        ByVal oSender As Object, ByVal oArgs As Xml.Schema.ValidationEventArgs)
        ErrorHandler(New ApplicationException("Fout bij het inlezen van het configuratie schema:" & _
            vbNewLine & oArgs.Message))
    End Sub

    ' Configuration file validation error
    Private Sub ValidationError( _
        ByVal oSender As Object, ByVal oArgs As Xml.Schema.ValidationEventArgs)
        ErrorHandler(New ApplicationException("Fout bij het inlezen van het configuratie bestand:" & _
            vbNewLine & oArgs.Message))
    End Sub

    ' Get the file name of the MXD document.
    Private Function GetMxdFile( _
        ByVal mxApp As IMxApplication) As String

        GetMxdFile = ""

        Dim templates As ITemplates
        Dim mxdIndex As Integer
        Dim filePath As String

        Try

            'Get the location of the current mxd file.
            'This is the last one in the templates collection.
            templates = CType(mxApp, IApplication).Templates
            mxdIndex = templates.Count - 1
            filePath = templates.Item(mxdIndex).ToString

            'Extract the mxd file name without extension.
            GetMxdFile = System.IO.Path.GetFileNameWithoutExtension(filePath)

        Catch ex As Exception
            ErrorHandler(ex)
        End Try
    End Function

    ' Get the folder path of the MXD document.
    Private Function GetMxdFolder( _
        ByVal mxApp As IMxApplication) As String

        GetMxdFolder = ""

        Dim templates As ITemplates
        Dim mxdIndex As Integer
        Dim filePath As String

        Try

            'Get the location of the current mxd file.
            'This is the last one in the templates collection.
            templates = CType(mxApp, IApplication).Templates
            mxdIndex = templates.Count - 1
            filePath = templates.Item(mxdIndex).ToString

            'Extract the mxd folder path.
            GetMxdFolder = System.IO.Path.GetDirectoryName(filePath)

        Catch ex As Exception
            ErrorHandler(ex)
        End Try
    End Function

#End Region

End Class

''' -----------------------------------------------------------------------------
''' Project	 : Digipolis.Hydranten.BeheerHydranten
''' Class	 : Hydranten.BeheerHydranten.LegendRule
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' A single hydrant legend code rule.
''' </summary>
''' <remarks>
''' The class model is based on the configuration XML schema model.
''' </remarks>
''' <history>
''' 	[Kristof Vydt]	21/02/2007	Created
''' 	[Kristof Vydt]	09/03/2007	Include attribute based result formula.
''' </history>
''' -----------------------------------------------------------------------------
Public Class LegendRule

    Public Enum ResultTypeEnumType
        FixedValue 'the result of the rule is a single, fixed value
        AttributeBasedCalculation 'the result of the rule is a formula, based on a single attribute value, with optional calculations
    End Enum

#Region "Local variables"
    Private _conditions As Collection 'set of multiple AttributeValueCondition objects
    Private _resultType As ResultTypeEnumType 'fixed value or attribute based formula
    Private _resultValue As Integer 'fixed legend code result value if each conditions is fullfilled
    Private _formulaSeedAttributeName As String 'the name of the attribute that is the seed for the attribute based calculation
    Private _formulaCalculations As Collection 'set of multiple Calculation objects that are executed on the attribute value for the attribute based calculation
#End Region

#Region "Public properties"

    Public ReadOnly Property Conditions() As Collection
        Get
            Return _conditions
        End Get
    End Property

    Public ReadOnly Property Condition(ByVal index As Integer) As AttributeValueCondition
        Get
            If index < _conditions.Count Then
                Return CType(_conditions(index), AttributeValueCondition)
            Else
                Throw New IndexOutOfRangeException
            End If
        End Get
    End Property

    Public ReadOnly Property ResultType() As ResultTypeEnumType
        Get
            Return _resultType
        End Get
    End Property

    Public ReadOnly Property ResultValue() As Integer
        Get
            If _resultType = ResultTypeEnumType.FixedValue Then
                Return _resultValue
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public ReadOnly Property SeedReference() As String
        Get
            If _resultType = ResultTypeEnumType.AttributeBasedCalculation Then
                Return _formulaSeedAttributeName
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public ReadOnly Property Calculations() As Collection
        Get
            If _resultType = ResultTypeEnumType.AttributeBasedCalculation Then
                Return _formulaCalculations
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public ReadOnly Property Calculation(ByVal index As Integer) As Calculation
        Get
            If _resultType = ResultTypeEnumType.AttributeBasedCalculation Then
                If index < _formulaCalculations.Count Then
                    Return CType(_formulaCalculations(index), Calculation)
                Else
                    Throw New IndexOutOfRangeException
                End If
            Else
                Return Nothing
            End If
        End Get
    End Property

#End Region

#Region "Public methods"

    ' Constructor
    Public Sub New( _
        ByVal ruleNode As Xml.XmlElement)

        Try

            ' Validate input node.
            If Not ruleNode.Name = "legendRule" Then _
                Throw New ArgumentException("Constructor argument is not <legendRule>")

            ' Query for list of conditions.
            _conditions = New Collection
            Dim nodeList As Xml.XmlNodeList = ruleNode.SelectNodes("attributeValueCondition")
            If nodeList.Count > 0 Then
                Dim ienum As IEnumerator = nodeList.GetEnumerator
                While (ienum.MoveNext)
                    Dim node As Xml.XmlNode = CType(ienum.Current, Xml.XmlNode)
                    Dim condition As AttributeValueCondition = New AttributeValueCondition(node)
                    _conditions.Add(condition)
                End While
            End If

            Select Case ruleNode.LastChild.Name

                Case "resultValue"
                    _resultType = ResultTypeEnumType.FixedValue

                    ' Query for result value.
                    _resultValue = Convert.ToInt16(ruleNode.SelectSingleNode("resultValue").InnerText)

                Case "resultFormula"
                    _resultType = ResultTypeEnumType.AttributeBasedCalculation

                    ' Query for result formula seed value reference.
                    Dim attrRef As Xml.XmlNode = ruleNode.SelectSingleNode("resultFormula/attribute")
                    If attrRef Is Nothing Then Throw New ArgumentException("Missing <attribute> in result formula.")
                    Dim attrName As String = String.Empty
                    If Not attrRef.Attributes("name") Is Nothing Then
                        attrName = attrRef.Attributes("name").Value
                    ElseIf Not attrRef.Attributes("key") Is Nothing Then
                        attrName = GetAttributeName("Hydrant", attrRef.Attributes("key").Value)
                    End If
                    If attrName.Equals(String.Empty) Then Throw New ApplicationException("Attribute name could not be determined.")
                    _formulaSeedAttributeName = attrName

                    ' Query for result formula calculations.
                    _formulaCalculations = New Collection
                    Dim calculationList As Xml.XmlNodeList = ruleNode.SelectNodes("resultFormula/calculation")
                    If calculationList.Count > 0 Then
                        Dim ienum As IEnumerator = calculationList.GetEnumerator
                        While (ienum.MoveNext)
                            Dim node As Xml.XmlNode = CType(ienum.Current, Xml.XmlNode)
                            Dim calc As Calculation = New Calculation(node)
                            _formulaCalculations.Add(calc)
                        End While
                    End If

                Case Else

            End Select

        Catch ex As Exception
            ErrorHandler(ex)
        End Try

    End Sub

    ' Add a new condition to the legend rule object.
    Public Sub AddCondition( _
        ByVal condition As AttributeValueCondition)
        _conditions.Add(condition)
    End Sub

    ' Check if a hydrant feature fullfill all conditions of the legend rule.
    <CLSCompliant(False)> _
    Public Function Comply( _
        ByVal hydrant As ESRI.ArcGIS.Geodatabase.IFeature) As Boolean

        Try

            ' Loop through each condition of the rule.
            For Each condition As AttributeValueCondition In Me.Conditions
                ' If a single condition is not met, then return False.
                If Not condition.Comply(hydrant) Then Return False
            Next condition

            ' If all conditions are met, then return True.
            Return True

        Catch ex As Exception
            ErrorHandler(ex)
        End Try
    End Function

#End Region

End Class

''' -----------------------------------------------------------------------------
''' Project	 : Digipolis.Hydranten.BeheerHydranten
''' Class	 : Hydranten.BeheerHydranten.AttributeValueCondition
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' A single attribute value condition, as part of a hydrant legend code rule.
''' </summary>
''' <remarks>
''' The class model is based on the configuration XML schema model.
''' </remarks>
''' <history>
''' 	[Kristof Vydt]	21/02/2007	Created
''' 	[Kristof Vydt]	08/03/2007	Implement additional operators
''' </history>
''' -----------------------------------------------------------------------------
Public Class AttributeValueCondition

    Public Enum LogicalOperatorEnumType
        IsEqual          '=
        IsLess           '<
        IsLessOrEqual    '</=
        IsGreater        '>
        IsGreaterOrEqual '>/=
    End Enum

#Region "Local variables"
    Private _attributeName As String 'attribute name
    Private _operator As LogicalOperatorEnumType 'how the attribute value and the comparison value should relate
    Private _comparison As String 'value to which the attribute is to be compared
#End Region

#Region "Public properties"

    ' The name of the feature attribute that is used in the condition.
    Public ReadOnly Property AttributeName() As String
        Get
            Return _attributeName
        End Get
    End Property

    ' The logical operator of the condition.
    Public ReadOnly Property [Operator]() As LogicalOperatorEnumType
        Get
            Return _operator
        End Get
    End Property

    ' The comparison value of the condition.
    Public ReadOnly Property ComparisonValue() As String
        Get
            Return _comparison
        End Get
    End Property

#End Region

#Region "Public methods"

    ' Constructor.
    Public Sub New( _
        ByVal conditionNode As Xml.XmlNode)

        Try

            ' Validate input node.
            If Not conditionNode.Name = "attributeValueCondition" Then _
                Throw New ArgumentException("Constructor argument is not <attributeValueCondition>")

            ' Query for attribute reference.
            Dim attrRef As Xml.XmlNode = conditionNode.SelectSingleNode("attribute")
            If attrRef Is Nothing Then Throw New ArgumentException("Missing <attribute> in attribute value condition.")
            Dim attrName As String = String.Empty
            If Not attrRef.Attributes("name") Is Nothing Then
                attrName = attrRef.Attributes("name").Value
            ElseIf Not attrRef.Attributes("key") Is Nothing Then
                attrName = GetAttributeName("Hydrant", attrRef.Attributes("key").Value)
            End If
            If attrName.Equals(String.Empty) Then Throw New ApplicationException("Attribute name could not be determined.")
            _attributeName = attrName


            ' Query for logical operator.
            Select Case conditionNode.SelectSingleNode("operator").InnerText
                Case "isEqual"
                    _operator = LogicalOperatorEnumType.IsEqual
                Case "isGreater"
                    _operator = LogicalOperatorEnumType.IsGreater
                Case "isGreaterOrEqual"
                    _operator = LogicalOperatorEnumType.IsGreaterOrEqual
                Case "isLess"
                    _operator = LogicalOperatorEnumType.IsLess
                Case "isLessOrEqual"
                    _operator = LogicalOperatorEnumType.IsLessOrEqual
                Case Else
                    Throw New ApplicationException("Invalid logical operator for legend rule.")
            End Select

            ' Query for the comparison value.
            _comparison = conditionNode.SelectSingleNode("comparisonValue").InnerText

        Catch ex As Exception
            ErrorHandler(ex)
        End Try

    End Sub

    ' Test if the condition is fullfilled for a specif hydrant feature.
    <CLSCompliant(False)> _
    Public Function Comply( _
        ByVal hydrant As ESRI.ArcGIS.Geodatabase.IFeature) As Boolean

        Try

            ' Check availability of configuration.
            If Config Is Nothing Then Throw New ApplicationException("No configuration loaded.")

            ' Feature attribute referenced by the condition.
            Dim name As String = Me.AttributeName
            Dim index As Integer = hydrant.Fields.FindField(name)
            Dim value As String = CStr(hydrant.Value(index))
            'Dim type As System.Type = value.GetType()

            ' Compare.
            Select Case Me.[Operator]
                Case AttributeValueCondition.LogicalOperatorEnumType.IsEqual
                    If value = Me.ComparisonValue Then Return True
                Case AttributeValueCondition.LogicalOperatorEnumType.IsGreater
                    If value > Me.ComparisonValue Then Return True
                Case AttributeValueCondition.LogicalOperatorEnumType.IsGreaterOrEqual
                    If value >= Me.ComparisonValue Then Return True
                Case AttributeValueCondition.LogicalOperatorEnumType.IsLess
                    If value < Me.ComparisonValue Then Return True
                Case AttributeValueCondition.LogicalOperatorEnumType.IsLessOrEqual
                    If value <= Me.ComparisonValue Then Return True
                Case Else
                    Throw New ApplicationException("Unsupported logical operator in legend rule.")
            End Select

        Catch ex As Exception
            ErrorHandler(ex)
        End Try
    End Function

#End Region

End Class

''' -----------------------------------------------------------------------------
''' Project	 : Digipolis.Hydranten.BeheerHydranten
''' Class	 : Hydranten.BeheerHydranten.Calculation
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' A single calculation as part of a legend rule result formula.
''' </summary>
''' <remarks>
''' The class model is based on the configuration XML schema model.
''' </remarks>
''' <history>
''' 	[Kristof Vydt]	9/03/2007	Created
''' </history>
''' -----------------------------------------------------------------------------
Public Class Calculation

    Public Enum ArithmaticOperatorEnumType
        Add       '+
        Substract '-
        Multiply  '*
        Divide    '/
    End Enum

#Region "Local variables"
    Private _operator As ArithmaticOperatorEnumType
    Private _parameter1 As Decimal
#End Region

#Region "Public properties"

    Public ReadOnly Property [Operator]() As ArithmaticOperatorEnumType
        Get
            Return _operator
        End Get
    End Property

    Public ReadOnly Property Parameter() As Decimal
        Get
            Return _parameter1
        End Get
    End Property

#End Region

#Region "Public methods"

    ' Constructor.
    Public Sub New( _
        ByVal calculationNode As Xml.XmlNode)

        Try

            ' Validate input node.
            If Not calculationNode.Name = "calculation" Then _
                Throw New ArgumentException("Constructor argument is not <calculation>.")

            ' Query for calculation parameters.
            Dim parameterNode As Xml.XmlNode = calculationNode.SelectSingleNode("parameter")
            If parameterNode Is Nothing Then _
                Throw New ArgumentException("Missing <parameter> in constructor argument.")
            _parameter1 = Convert.ToDecimal(parameterNode.InnerText)

            ' Query for logical operator.
            Select Case calculationNode.SelectSingleNode("operation").InnerText
                Case "add"
                    _operator = ArithmaticOperatorEnumType.Add
                Case "substract"
                    _operator = ArithmaticOperatorEnumType.Substract
                Case "multiply"
                    _operator = ArithmaticOperatorEnumType.Multiply
                Case "divide"
                    _operator = ArithmaticOperatorEnumType.Divide
                Case Else
                    Throw New ArgumentException("Invalid arithmatic operator for legend result fomula.")
            End Select

        Catch ex As Exception
            ErrorHandler(ex)
        End Try

    End Sub

    ' Execute the calculation on the specified seed value.
    Public Function Execute( _
        ByVal seed As Object) As Object

        Execute = Nothing
        Try

            Select Case _operator
                Case ArithmaticOperatorEnumType.Add
                    Return Convert.ToDecimal(seed) + _parameter1
                Case ArithmaticOperatorEnumType.Substract
                    Return Convert.ToDecimal(seed) - _parameter1
                Case ArithmaticOperatorEnumType.Multiply
                    Return Convert.ToDecimal(seed) * _parameter1
                Case ArithmaticOperatorEnumType.Divide
                    Return Convert.ToDecimal(seed) / _parameter1
                Case Else
                    Throw New ArgumentException("Unknown arithmatic operator for legend result fomula.")
            End Select

            Return Nothing

        Catch ex As Exception
            ErrorHandler(ex)
        End Try

    End Function

#End Region

End Class