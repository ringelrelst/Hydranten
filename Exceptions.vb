Option Explicit On 
Option Strict On

''' -----------------------------------------------------------------------------
''' Project	 : Digipolis.Hydranten.BeheerHydranten
''' Class	 : Hydranten.BeheerHydranten.IncompleteConfigurationException
''' 
''' -----------------------------------------------------------------------------
''' <summary>
'''     Application setting is missing in configuration.
''' </summary>
''' <remarks>
''' Provided message should clearly describe what is missing.
''' </remarks>
''' <history>
''' 	[Kristof Vydt]	21/03/2007	Created
''' </history>
''' -----------------------------------------------------------------------------
Public Class IncompleteConfigurationException
    Inherits System.ApplicationException

    Public Sub New(ByVal message As String)
        MyBase.New(message)
    End Sub

End Class

''' -----------------------------------------------------------------------------
''' Project	 : Digipolis.Hydranten.BeheerHydranten
''' Class	 : Hydranten.BeheerHydranten.FileNotFoundException
''' 
''' -----------------------------------------------------------------------------
''' <summary>
'''     A file could not be found.
''' </summary>
''' <remarks>
'''     The config file and the config schema file should reside  
'''     in the same folder as the used mxd document.
''' </remarks>
''' <history>
''' 	[Kristof Vydt]	16/09/2005	Created
''' 	[Kristof Vydt]	08/03/2007	Generalise ConfigFileNotFoundException to FileNotFoundException
'''     [Koen Vermeer]  23/06/2008  Conversion to ArcGIS 9.2
''' </history>
''' -----------------------------------------------------------------------------
Public Class FileNotFoundException
    Inherits System.ApplicationException

    Private m_filePath As String = ""

    Public Sub New(ByVal file As String)
        MyBase.New()
        m_filePath = file
    End Sub

    Public Overrides ReadOnly Property Message() As String
        Get
            Return "'" & m_filePath & "' is niet gevonden."
        End Get
    End Property
End Class

''' -----------------------------------------------------------------------------
''' Project	 : Digipolis.Hydranten.BeheerHydranten
''' Class	 : Hydranten.BeheerHydranten.LayerNotFoundException
''' 
''' -----------------------------------------------------------------------------
''' <summary>
'''     The layer could not be found.
''' </summary>
''' <remarks>
'''     Check if the layername (global constants in ModuleGlobals) can be found
'''     in the TOC of the current map.
''' </remarks>
''' <history>
''' 	[Kristof Vydt]	16/09/2005	Created
''' 	[Kristof Vydt]	10/08/2006	Layer name added in exception message.
''' </history>
''' -----------------------------------------------------------------------------
Public Class LayerNotFoundException
    Inherits System.ApplicationException

    Private m_LayerName As String

    Public Sub New(ByVal layerName As String)
        MyBase.New()
        m_LayerName = layerName
    End Sub

    Public Overrides ReadOnly Property Message() As String
        Get
            Return "De laag '" & m_LayerName & "' werd niet gevonden in de kaart."
        End Get
    End Property

    Public ReadOnly Property LayerName() As String
        Get
            Return m_LayerName
        End Get
    End Property

End Class

''' -----------------------------------------------------------------------------
''' Project	 : Digipolis.Hydranten.BeheerHydranten
''' Class	 : Hydranten.BeheerHydranten.LayerNotValidException
''' 
''' -----------------------------------------------------------------------------
''' <summary>
'''     The layer is not valid.
''' </summary>
''' <remarks>
'''     Check if the layer is not greyed out in the TOC.
'''     There might be a problem of finding the source of the layer.
''' </remarks>
''' <history>
''' 	[Kristof Vydt]	10/08/2006	Created
''' </history>
''' -----------------------------------------------------------------------------
Public Class LayerNotValidException
    Inherits System.ApplicationException

    Private m_LayerName As String

    Public Sub New(ByVal layerName As String)
        MyBase.New()
        m_LayerName = layerName
    End Sub

    Public Overrides ReadOnly Property Message() As String
        Get
            Return "De laag '" & m_LayerName & "' is niet geldig."
        End Get
    End Property

    Public ReadOnly Property LayerName() As String
        Get
            Return m_LayerName
        End Get
    End Property

End Class

''' -----------------------------------------------------------------------------
''' Project	 : Digipolis.Hydranten.BeheerHydranten
''' Class	 : Hydranten.BeheerHydranten.SectorNameException
''' 
''' -----------------------------------------------------------------------------
''' <summary>
'''     No Brandweer Sector name.
''' </summary>
''' <remarks>
'''     Check the filename of the mxd document.
'''     Make sure no short notation with tilde is used in mxd filename.
''' </remarks>
''' <history>
''' 	[Kristof Vydt]	16/09/2005	Created
''' </history>
''' -----------------------------------------------------------------------------
Public Class SectorNameException
    Inherits System.ApplicationException

    Public Overrides ReadOnly Property Message() As String
        Get
            Message = "De naam van de brandweer sector kan niet worden afgeleid uit de bestandsnaam."
        End Get
    End Property

End Class

''' -----------------------------------------------------------------------------
''' Project	 : Digipolis.Hydranten.BeheerHydranten
''' Class	 : Hydranten.BeheerHydranten.SectorCodeException
''' 
''' -----------------------------------------------------------------------------
''' <summary>
'''     No Brandweer Sector code.
''' </summary>
''' <remarks>
'''     Check the [SectorCodes] section of the configuration file.
'''     Make sure no short notation with tilde is used in mxd filename.
''' </remarks>
''' <history>
''' 	[Kristof Vydt]	16/09/2005	Created
''' </history>
''' -----------------------------------------------------------------------------
Public Class SectorCodeException
    Inherits System.ApplicationException

    Private m_SectorName As String

    Public Sub New(ByVal sectorName As String)
        m_SectorName = sectorName
    End Sub

    Public Overrides ReadOnly Property Message() As String
        Get
            Message = "Er werd geen code gevonden voor sector '" & m_SectorName & "'."
        End Get
    End Property
End Class

''' -----------------------------------------------------------------------------
''' Project	 : Digipolis.Hydranten.BeheerHydranten
''' Class	 : Hydranten.BeheerHydranten.SectorPostcodeException
''' 
''' -----------------------------------------------------------------------------
''' <summary>
'''     No Brandweer Sector postcodes.
''' </summary>
''' <remarks>
'''     Check the configuration file.
''' </remarks>
''' <history>
''' 	[Kristof Vydt]	16/09/2005	Created
''' </history>
''' -----------------------------------------------------------------------------
Public Class SectorPostcodeException
    Inherits System.ApplicationException

    Private m_SectorName As String

    Public Sub New(ByVal sectorName As String)
        m_SectorName = sectorName
    End Sub

    Public Overrides ReadOnly Property Message() As String
        Get
            Message = "Er werden geen postcodes gevonden voor sector '" & m_SectorName & "'."
        End Get
    End Property
End Class

''' -----------------------------------------------------------------------------
''' Project	 : Digipolis.Hydranten.BeheerHydranten
''' Class	 : Hydranten.BeheerHydranten.TableNotFoundException
''' 
''' -----------------------------------------------------------------------------
''' <summary>
'''     A table with specified name is not found.
''' </summary>
''' <remarks>
''' </remarks>
''' <history>
''' 	[Kristof Vydt]	16/09/2005	Created
''' </history>
''' -----------------------------------------------------------------------------
Public Class TableNotFoundException
    Inherits System.ApplicationException

    Private m_TableName As String

    Public Sub New(ByVal tableName As String)
        MyBase.New()
        m_TableName = tableName
    End Sub


    Public Overrides ReadOnly Property Message() As String
        Get
            Return "Er kan geen tabel gevonden worden met de naam '" & m_TableName & "'."
        End Get
    End Property
End Class

''' -----------------------------------------------------------------------------
''' Project	 : Digipolis.Hydranten.BeheerHydranten
''' Class	 : Hydranten.BeheerHydranten.AttributeNotFoundException
''' 
''' -----------------------------------------------------------------------------
''' <summary>
'''     The attribute with specified name could not be found in the layer.
''' </summary>
''' <remarks>
''' </remarks>
''' <history>
''' 	[Kristof Vydt]	16/09/2005	Created
''' </history>
''' -----------------------------------------------------------------------------
Public Class AttributeNotFoundException
    Inherits System.ApplicationException

    Private m_LayerName As String
    Private m_AttributeName As String

    Public Sub New(ByVal layerName As String, ByVal attributeName As String)
        MyBase.New()
        m_LayerName = layerName
        m_AttributeName = attributeName
    End Sub

    Public Overrides ReadOnly Property Message() As String
        Get
            Return "Geen attribuut met naam '" & m_AttributeName & "' in laag '" & m_LayerName & "'."
        End Get
    End Property
End Class

''' -----------------------------------------------------------------------------
''' Project	 : Digipolis.Hydranten.BeheerHydranten
''' Class	 : Hydranten.BeheerHydranten.RecordsetFieldSizeNotSufficientException
''' 
''' -----------------------------------------------------------------------------
''' <summary>
'''     The size of an ADODB recordset field, is not sufficient
'''     to store some value.
''' </summary>
''' <remarks>
'''     Possible solution: set c_LookupRS_{*}_maxsize to a higher value
''' </remarks>
''' <history>
''' 	[Kristof Vydt]	16/09/2005	Created
''' 	[Kristof Vydt]	23/11/2005	Language-error corrected in message.
''' </history>
''' -----------------------------------------------------------------------------
Public Class RecordsetFieldSizeNotSufficientException
    Inherits System.ApplicationException

    Private m_FieldName As String

    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub New(ByVal fieldName As String)
        MyBase.New()
        m_FieldName = fieldName
    End Sub

    Public Overrides ReadOnly Property Message() As String
        Get
            If Len(m_FieldName) > 0 Then
                Return ("Grootte van het recordveld '" & m_FieldName & "' is niet toereikend.")
            Else
                Return ("Grootte van het recordveld is niet toereikend.")
            End If
        End Get
    End Property
End Class

''' -----------------------------------------------------------------------------
''' Project	 : Digipolis.Hydranten.BeheerHydranten
''' Class	 : Hydranten.BeheerHydranten.AttributeSizeNotSufficientException
''' 
''' -----------------------------------------------------------------------------
''' <summary>
'''     The size of a feature attribute, is not sufficient to store some value.
''' </summary>
''' <remarks>
'''     Possible solution: increase size limitation of attribute to a higher value
''' </remarks>
''' <history>
''' 	[Kristof Vydt]	16/09/2005	Created
''' 	[Kristof Vydt]	23/11/2005	Language-error in message corrected.
''' </history>
''' -----------------------------------------------------------------------------
Public Class AttributeSizeNotSufficientException
    Inherits System.ApplicationException

    Private m_FeatureClassName As String
    Private m_AttributeName As String

    Public Sub New(ByVal featureClassName As String, ByVal attributeName As String)
        MyBase.New()
        m_FeatureClassName = featureClassName
        m_AttributeName = attributeName
    End Sub

    Public Overrides ReadOnly Property Message() As String
        Get
            Return ("Grootte van het attribuut '" & m_AttributeName & "' van feature '" & m_FeatureClassName & "' is niet toereikend om de door u ingegeven waarde te bewaren.")
        End Get
    End Property
End Class

''' -----------------------------------------------------------------------------
''' Project	 : Digipolis.Hydranten.BeheerHydranten
''' Class	 : Hydranten.BeheerHydranten.RecreateExportFileException
''' 
''' -----------------------------------------------------------------------------
''' <summary>
'''     An export txt file could not be recreated.
''' </summary>
''' <remarks>
'''     Possible reasons:
'''     - The folder does not exist.
'''     - The file exists and is read-only.
'''     - The file exists and is locked by another program.
'''     - A new file cannot be created in the folder.
''' </remarks>
''' <history>
''' 	[Kristof Vydt]	16/09/2005	Created
''' </history>
''' -----------------------------------------------------------------------------
Public Class RecreateExportFileException
    Inherits System.ApplicationException

    Private m_FilePath As String

    Public Sub New(ByVal FilePath As String)
        MyBase.New()
        m_FilePath = FilePath
    End Sub

    Public Overrides ReadOnly Property Message() As String
        Get
            Return "Het exportbestand " & m_FilePath & " kan niet worden aangemaakt."
        End Get
    End Property
End Class

''' -----------------------------------------------------------------------------
''' Project	 : Digipolis.Hydranten.BeheerHydranten
''' Class	 : Hydranten.BeheerHydranten.HydrantsWithTemporaryStatus
''' 
''' -----------------------------------------------------------------------------
''' <summary>
'''     Some hydrants have a temporary status ('nieuw', 'verwijderd', 'nakijken_{a|c|ac}').
'''     Uploading for a sector is not allowed when there are temporary status'.
''' </summary>
''' <remarks>
''' </remarks>
''' <history>
''' 	[Kristof Vydt]	16/09/2005	Created
''' 	[Kristof Vydt]	10/10/2005	Message depends on the count of temp qstatus hydrants.
'''     [Kristof Vydt]  15/08/2006  Modify message.
''' </history>
''' -----------------------------------------------------------------------------
Public Class HydrantsWithTemporaryStatus
    Inherits System.ApplicationException

    Private m_SectorName As String
    Private m_TempStatusCount As Integer

    Public Sub New(ByVal count As Integer, ByVal sector As String)
        MyBase.New()
        m_TempStatusCount = count
        m_SectorName = sector
    End Sub

    Public Overrides ReadOnly Property Message() As String
        Get
            If m_TempStatusCount = 1 Then
                Return "Voor de sector " & m_SectorName & " werd " & CStr(m_TempStatusCount) & " hydrant met een tijdelijke status gevonden." & vbNewLine & _
                       "Zolang er tijdelijke statussen bestaan, kunnen er geen hydranten worden opgeladen."
            Else
                Return "Voor de sector " & m_SectorName & " werden " & CStr(m_TempStatusCount) & " hydranten met een tijdelijke status gevonden." & vbNewLine & _
                       "Zolang er tijdelijke statussen bestaan, kunnen er geen hydranten worden opgeladen."
            End If
        End Get
    End Property
End Class

'''' -----------------------------------------------------------------------------
'''' Project	 : Digipolis.Hydranten.BeheerHydranten
'''' Class	 : Hydranten.BeheerHydranten.MissingImportSchemaValue
'''' 
'''' -----------------------------------------------------------------------------
'''' <summary>
''''     It is not possible to determine a value for a required parameter
''''     in the import schema ini-file. The key is missing or has empty value.
'''' </summary>
'''' <remarks>
'''' </remarks>
'''' <history>
'''' 	[Kristof Vydt]	16/09/2005	Created
'''' 	[Kristof Vydt]	29/09/2006	Deprecated
'''' </history>
'''' -----------------------------------------------------------------------------
'Public Class MissingImportSchemaValue
'    Inherits System.ApplicationException

'    Private m_Section As String
'    Private m_Key As String

'    Public Sub New(ByVal SectionName As String)
'        MyBase.New()
'        m_Section = SectionName
'        m_Key = ""
'    End Sub

'    Public Sub New(ByVal SectionName As String, ByVal KeyName As String)
'        MyBase.New()
'        m_Section = SectionName
'        m_Key = KeyName
'    End Sub

'    Public Overrides ReadOnly Property Message() As String
'        Get
'            If m_Key = "" Then
'                Return "De sectie '" & m_Section & "' ontbreekt in het importschema."
'            Else
'                Return "De parameter '" & m_Key & "' in sectie '" & m_Section & "' ontbreekt in het importschema, of heeft een niet-toegelaten lege waarde."
'            End If
'        End Get
'    End Property
'End Class

'''' -----------------------------------------------------------------------------
'''' Project	 : Digipolis.Hydranten.BeheerHydranten
'''' Class	 : Hydranten.BeheerHydranten.InvalidImportSchemaValue
'''' 
'''' -----------------------------------------------------------------------------
'''' <summary>
''''     The value of a parameter in the import schema ini-file
''''     does not have the expected content/format.
'''' </summary>
'''' <remarks>
'''' </remarks>
'''' <history>
'''' 	[Kristof Vydt]	16/09/2005	Created
'''' 	[Kristof Vydt]	29/09/2006	Deprecated
'''' </history>
'''' -----------------------------------------------------------------------------
'Public Class InvalidImportSchemaValue
'    Inherits System.ApplicationException

'    Private m_Section As String
'    Private m_Key As String
'    Private m_Value As String

'    Public Sub New(ByVal SectionName As String, ByVal KeyName As String, Optional ByVal KeyValue As String = "")
'        MyBase.New()
'        m_Section = SectionName
'        m_Key = KeyName
'        m_Value = KeyValue
'    End Sub

'    Public Overrides ReadOnly Property Message() As String
'        Get
'            Return "De parameter '" & m_Key & "' in sectie '" & m_Section & "' heeft een ongeldige waarde."
'        End Get
'    End Property
'End Class

'''' -----------------------------------------------------------------------------
'''' Project	 : Digipolis.Hydranten.BeheerHydranten
'''' Class	 : Hydranten.BeheerHydranten.InvalidImportSchemaColumnIndex
'''' 
'''' -----------------------------------------------------------------------------
'''' <summary>
''''     An invalid column index in the import schema ini-file.
'''' </summary>
'''' <remarks>
''''     Often this refers to an invalid column index used as key name or value,
''''     or to text that is used as key name or value when a column index is expected.
''''     A valid column index is an integer ranging from 0 to (number of columns in Excel -1).
'''' </remarks>
'''' <history>
'''' 	[Kristof Vydt]	16/09/2005	Created
'''' 	[Kristof Vydt]	29/09/2006	Deprecated
'''' </history>
'''' -----------------------------------------------------------------------------
'Public Class InvalidImportSchemaColumnIndex
'    Inherits System.ApplicationException

'    Private m_Section As String
'    Private m_Key As String
'    Private m_Value As String

'    Public Sub New(ByVal SectionName As String, ByVal KeyName As String, ByVal KeyValue As String)
'        MyBase.New()
'        m_Section = SectionName
'        m_Key = KeyName
'        m_Value = KeyValue
'    End Sub

'    Public Sub New(ByVal SectionName As String, ByVal KeyName As String)
'        MyBase.New()
'        m_Section = SectionName
'        m_Key = KeyName
'        m_Value = Nothing
'    End Sub

'    Public Overrides ReadOnly Property Message() As String
'        Get
'            If m_Value Is Nothing Then
'                Return "De parameter '" & m_Key & "' in sectie '" & m_Section & "' is geen geldig kolom index."
'            Else
'                Return "De waarde '" & m_Value & "' voor parameter '" & m_Key & "' in sectie '" & m_Section & "' is geen geldig kolom index."
'            End If
'        End Get
'    End Property
'End Class

''' -----------------------------------------------------------------------------
''' Project	 : Digipolis.Hydranten.BeheerHydranten
''' Class	 : Hydranten.BeheerHydranten.AbortedByUserException
''' 
''' -----------------------------------------------------------------------------
''' <summary>
'''     The procedure was aborted by the user.
''' </summary>
''' <remarks>
''' </remarks>
''' <history>
''' 	[Kristof Vydt]	11/08/2006	Created
''' </history>
''' -----------------------------------------------------------------------------
Public Class AbortedByUserException
    Inherits System.ApplicationException

    Public Sub New()
        MyBase.New()
    End Sub

    Public Overrides ReadOnly Property Message() As String
        Get
            Return "Afgebroken door de gebruiker."
        End Get
    End Property
End Class