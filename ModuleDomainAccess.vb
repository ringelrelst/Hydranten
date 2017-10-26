Option Explicit On 
Option Strict On

Imports System.Windows.Forms
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geodatabase

''' -----------------------------------------------------------------------------
''' Project	 : Digipolis.Hydranten.BeheerHydranten
''' Class	 : Hydranten.BeheerHydranten.CodedValueDomainManager
''' 
''' -----------------------------------------------------------------------------
''' <summary>
'''     Access a coded value domain form ArcGIS.
''' </summary>
''' <remarks>
'''     Introduced to replace ModuleDomainAccess.
''' </remarks>
''' <history>
''' 	[Kristof Vydt]	21/03/2007	Created
'''     [Koen Vermeer]  23/06/2008  Conversion to ArcGIS 9.2
''' </history>
''' -----------------------------------------------------------------------------
Class CodedValueDomainManager

#Region "Local variables"
    Private m_domain As ICodedValueDomain
#End Region

#Region "Constructors"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Constructor based on a workspace.
    ''' </summary>
    ''' <param name="pWorkspace">The ArcGIS workspace the domain belongs to.</param>
    ''' <param name="sDomainKey">The domain keyword as configured.</param>
    ''' <remarks>
    '''     Created to replace ModuleDomainAccess.GetCodedValueDomain.
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	21/03/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub New( _
            ByVal pWorkspace As IWorkspace, _
            ByVal sDomainKey As String)

        ' QI for the IWorkspaceDomains.
        Dim pWSDomains As IWorkspaceDomains
        pWSDomains = CType(pWorkspace, IWorkspaceDomains)

        ' Get the domain name with the key specified.
        Dim sDomainName As String
        sDomainName = GetDomainName(sDomainKey)

        ' Get the domain with the name specified.
        Dim pDomain As IDomain
        pDomain = pWSDomains.DomainByName(sDomainName)

        ' Check if is a coded value domain.
        If pDomain.Type <> esriDomainType.esriDTCodedValue Then
            Throw New ApplicationException("Geen coded value domein gevonden (key '" & sDomainKey & "').")
        End If

        ' QI for the ICodedValueDomain interface.
        Dim pCodedValueDomain As ICodedValueDomain
        pCodedValueDomain = CType(pDomain, ICodedValueDomain)

        ' Store domain in local variable.
        m_domain = pCodedValueDomain

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Constructor based on a feature.
    ''' </summary>
    ''' <param name="pFeature">The ArcGIS feature that uses the coded value domain.</param>
    ''' <param name="sDomainKey">The domain keyword as configured.</param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	21/03/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub New( _
            ByVal pFeature As IFeature, _
            ByVal sDomainKey As String)

        ' Get the workspace of the specified feature.
        Dim pTable As ITable = pFeature.Table
        Dim pDataset As IDataset = CType(pTable, IDataset)
        Dim pWorkspace As IWorkspace = pDataset.Workspace

        '-- The following code is just a copy of the other constructor. --

        ' QI for the IWorkspaceDomains.
        Dim pWSDomains As IWorkspaceDomains
        pWSDomains = CType(pWorkspace, IWorkspaceDomains)

        ' Get the domain name with the key specified.
        Dim sDomainName As String
        sDomainName = GetDomainName(sDomainKey)

        ' Get the domain with the name specified.
        Dim pDomain As IDomain
        pDomain = pWSDomains.DomainByName(sDomainName)

        ' Check if is a coded value domain.
        If pDomain.Type <> esriDomainType.esriDTCodedValue Then
            Throw New ApplicationException("Geen coded value domein gevonden (key '" & sDomainKey & "').")
        End If

        ' QI for the ICodedValueDomain interface.
        Dim pCodedValueDomain As ICodedValueDomain
        pCodedValueDomain = CType(pDomain, ICodedValueDomain)

        ' Store domain in local variable.
        m_domain = pCodedValueDomain

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Constructor based on a feature layer.
    ''' </summary>
    ''' <param name="pLayer">The ArcGIS feature layer that uses the coded value domain.</param>
    ''' <param name="sDomainKey">The domain keyword as configured.</param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	22/03/2007	Created
    ''' 	[Kristof Vydt]	19/04/2007	Throw exception if no domain is found.
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub New( _
            ByVal pLayer As IFeatureLayer, _
            ByVal sDomainKey As String)

        ' Get the workspace of the specified feature layer.
        Dim pFeatureClass As IFeatureClass = pLayer.FeatureClass()
        Dim pDataset As IDataset = CType(pFeatureClass, IDataset)
        Dim pWorkspace As IWorkspace = pDataset.Workspace

        '-- The following code is just a copy of the other constructor. --

        ' QI for the IWorkspaceDomains.
        Dim pWSDomains As IWorkspaceDomains
        pWSDomains = CType(pWorkspace, IWorkspaceDomains)

        ' Get the domain name with the key specified.
        Dim sDomainName As String
        sDomainName = GetDomainName(sDomainKey)

        ' Get the domain with the name specified.
        Dim pDomain As IDomain
        pDomain = pWSDomains.DomainByName(sDomainName)
        If pDomain Is Nothing Then
            Throw New ApplicationException("Geen coded value domein gevonden (key '" & sDomainKey & "').")
        End If

        ' Check if is a coded value domain.
        If pDomain.Type <> esriDomainType.esriDTCodedValue Then
            Throw New ApplicationException("Geen coded value domein gevonden (key '" & sDomainKey & "').")
        End If

        ' QI for the ICodedValueDomain interface.
        Dim pCodedValueDomain As ICodedValueDomain
        pCodedValueDomain = CType(pDomain, ICodedValueDomain)

        ' Store domain in local variable.
        m_domain = pCodedValueDomain

    End Sub

#End Region

#Region "Public methods"

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''     Display code-value pairs from domain in combo box control.
        ''' </summary>
        ''' <param name="combobox">The combobox control to add to.</param>
        ''' <param name="separator">Text to put between value and name in the combo box.</param>
        ''' <remarks>
        '''     Created to replace ModuleDomainAccess.PopulateCodes.
        '''     Combobox content is not cleared in this method.
        ''' </remarks>
        ''' <history>
        ''' 	[Kristof Vydt]	21/03/2007	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
    Public Sub PopulateCodes( _
            ByVal combobox As ComboBox, _
            Optional ByVal separator As String = ": ")

        ' Get the coded value domain.
        Dim pCodedValueDomain As ICodedValueDomain
        pCodedValueDomain = m_domain

        ' Get a count of the coded values.
        Dim lCodes As Integer
        lCodes = pCodedValueDomain.CodeCount

        ' Loop through the list of values and 
        ' add their names to the combo box.
        Dim i As Integer
        For i = 0 To lCodes - 1
            combobox.Items.Add(CType(pCodedValueDomain.Value(i), String) & _
                                separator & pCodedValueDomain.Name(i))
        Next i

    End Sub

#End Region

#Region "Public properties"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Retreive code value that matches the specified code name.
    ''' </summary>
    ''' <param name="CodeValue">Code value in coded value domain.</param>
    ''' <returns>Code name from coded value domain.</returns>
    ''' <remarks>
    '''     Created to replace ModuleDomainAccess.GetDomainCodeName.
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	21/03/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property CodeName(ByVal CodeValue As String) As String
        Get
            ' Get the coded value domain.
            Dim pCodedValueDomain As ICodedValueDomain
            pCodedValueDomain = m_domain

            ' Get a count of the coded values.
            Dim lCodes As Integer
            lCodes = pCodedValueDomain.CodeCount

            ' Loop through the list of values and return matching code.
            Dim i As Integer
            For i = 0 To lCodes - 1
                If Convert.ToString(pCodedValueDomain.Value(i)) = CodeValue Then
                    Return pCodedValueDomain.Name(i)
                    Exit For
                End If
            Next i

            ' Return nothing if code value was not found in domain.
            Return Nothing

        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Retreive code value that matches the specified code name.
    ''' </summary>
    ''' <param name="CodeName">Code name in coded value domain.</param>
    ''' <returns>Code value from coded value domain.</returns>
    ''' <remarks>
    '''     Created to replace ModuleDomainAccess.GetDomainCodeValue.
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	21/03/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property CodeValue(ByVal CodeName As String) As String
        Get
            ' Get the coded value domain.
            Dim pCodedValueDomain As ICodedValueDomain
            pCodedValueDomain = m_domain

            ' Get a count of the coded values.
            Dim lCodes As Integer
            lCodes = pCodedValueDomain.CodeCount

            ' Loop through the list of values and return matching code.
            Dim i As Integer
            For i = 0 To lCodes - 1
                If pCodedValueDomain.Name(i) = CodeName Then
                    Return Convert.ToString(pCodedValueDomain.Value(i))
                    Exit For
                End If
            Next i

            ' Return nothing if code name was not found in domain.
            Return Nothing

        End Get
    End Property

#End Region

End Class

Module ModuleDomainAccess

    '''' -----------------------------------------------------------------------------
    '''' <summary>
    ''''     Show domain code-value pairs in combo box control.
    '''' </summary>
    '''' <param name="pWorkspace">
    ''''     domain workspace object
    '''' </param>
    '''' <param name="sDomainName">
    ''''     string name of the domain
    '''' </param>
    '''' <param name="cboValues">
    ''''     combo box control
    '''' </param>
    '''' <param name="separator">
    ''''     text to separate code and value
    '''' </param>
    '''' <remarks>
    '''' </remarks>
    '''' <history>
    '''' 	[Kristof Vydt]	16/09/2005	Created
    ''''     [Kristof Vydt]  31/08/2006  Use private function GetCodedValueDomain().
    '''' 	[Kristof Vydt]	22/03/2007	Deprecated.
    '''' </history>
    '''' -----------------------------------------------------------------------------
    'Public Sub PopulateCodes( _
    '    ByVal pWorkspace As IWorkspace, _
    '    ByVal sDomainName As String, _
    '    ByVal cboValues As ComboBox, _
    '    Optional ByVal separator As String = ": ")

    '    Try
    '        ' +++ Get the domain with the name specified
    '        Dim pCodedValueDomain As ICodedValueDomain
    '        pCodedValueDomain = GetCodedValueDomain(pWorkspace, sDomainName)

    '        ' +++ Get a count of the coded values
    '        Dim lCodes As Integer
    '        lCodes = pCodedValueDomain.CodeCount

    '        ' +++ Loop through the list of values and add them
    '        ' +++ and their names to the combo box
    '        Dim i As Integer
    '        For i = 0 To lCodes - 1
    '            cboValues.Items.Add(CType(pCodedValueDomain.Value(i), String) & separator & pCodedValueDomain.Name(i))
    '        Next i

    '    Catch ex As Exception
    '        Throw ex
    '    End Try

    'End Sub

    '''' -----------------------------------------------------------------------------
    '''' <summary>
    ''''     Retreive code name from coded domain list that matches the specified code value.
    '''' </summary>
    '''' <param name="pWorkspace">
    ''''     domain workspace object
    '''' </param>
    '''' <param name="sDomain">
    ''''     domain name string
    '''' </param>
    '''' <param name="pCodeValue">
    ''''     code value object
    '''' </param>
    '''' <returns>code name string</returns>
    '''' <remarks>
    '''' </remarks>
    '''' <history>
    '''' 	[Kristof Vydt]	31/08/2006	Created
    '''' 	[Kristof Vydt]	22/03/2007	Deprecated.
    '''' </history>
    '''' -----------------------------------------------------------------------------
    'Public Function GetDomainCodeName( _
    '    ByVal pWorkspace As IWorkspace, _
    '    ByVal sDomainName As String, _
    '    ByVal pCodeValue As Object) As String

    '    Try
    '        ' +++ Get the domain with the name specified
    '        Dim pCodedValueDomain As ICodedValueDomain
    '        pCodedValueDomain = GetCodedValueDomain(pWorkspace, sDomainName)

    '        ' +++ Get a count of the coded values
    '        Dim lCodes As Integer
    '        lCodes = pCodedValueDomain.CodeCount

    '        ' +++ Loop through the list of values and return matching code.
    '        Dim i As Integer
    '        For i = 0 To lCodes - 1
    '            If pCodedValueDomain.Value(i) Is pCodeValue Then
    '                Return pCodedValueDomain.Name(i)
    '                Exit For
    '            End If
    '        Next i

    '    Catch ex As Exception
    '        Throw ex
    '    End Try

    'End Function

    '''' -----------------------------------------------------------------------------
    '''' <summary>
    ''''     Retreive code value from coded domain list that matches the specified code name.
    '''' </summary>
    '''' <param name="pWorkspace">
    ''''     domain workspace object
    '''' </param>
    '''' <param name="sDomain">
    ''''     domain name string
    '''' </param>
    '''' <param name="sCodeName">
    ''''     code name string
    '''' </param>
    '''' <returns>code value object</returns>
    '''' <remarks>
    '''' </remarks>
    '''' <history>
    '''' 	[Kristof Vydt]	31/08/2006	Created
    '''' 	[Kristof Vydt]	22/03/2007	Deprecated.
    '''' </history>
    '''' -----------------------------------------------------------------------------
    'Public Function GetDomainCodeValue( _
    '    ByVal pWorkspace As IWorkspace, _
    '    ByVal sDomainName As String, _
    '    ByVal sCodeName As String) As Object

    '    Try
    '        ' +++ Get the domain with the name specified
    '        Dim pCodedValueDomain As ICodedValueDomain
    '        pCodedValueDomain = GetCodedValueDomain(pWorkspace, sDomainName)

    '        ' +++ Get a count of the coded values
    '        Dim lCodes As Integer
    '        lCodes = pCodedValueDomain.CodeCount

    '        ' +++ Loop through the list of values and return matching code.
    '        Dim i As Integer
    '        For i = 0 To lCodes - 1
    '            If pCodedValueDomain.Name(i) = sCodeName Then
    '                Return pCodedValueDomain.Value(i)
    '                Exit For
    '            End If
    '        Next i

    '    Catch ex As Exception
    '        Throw ex
    '    End Try

    'End Function

    '''' -----------------------------------------------------------------------------
    '''' <summary>
    ''''     Retreive coded value domain from database.
    '''' </summary>
    '''' <param name="pWorkspace">
    ''''     domain workspace object
    '''' </param>
    '''' <param name="sDomain">
    ''''     domain name string
    '''' </param>
    '''' <returns>The coded value domain.</returns>
    '''' <remarks>
    '''' </remarks>
    '''' <history>
    '''' 	[Kristof Vydt]	31/08/2006	Created
    '''' 	[Kristof Vydt]	22/03/2007	Deprecated.
    '''' </history>
    '''' -----------------------------------------------------------------------------
    'Private Function GetCodedValueDomain( _
    '    ByVal pWorkspace As IWorkspace, _
    '    ByVal sDomainName As String) As ICodedValueDomain

    '    Try
    '        ' +++ QI for the IWorkspaceDomains interface
    '        Dim pWSDomains As IWorkspaceDomains
    '        pWSDomains = CType(pWorkspace, IWorkspaceDomains)

    '        ' +++ Get the domain with the name specified
    '        Dim pDomain As IDomain
    '        pDomain = pWSDomains.DomainByName(sDomainName)

    '        ' +++ Check if is a coded value domain
    '        If pDomain.Type <> esriDomainType.esriDTCodedValue Then
    '            Exit Function
    '        End If

    '        ' +++ QI for the ICodedValueDomain interface
    '        Dim pCodedValueDomain As ICodedValueDomain
    '        pCodedValueDomain = CType(pDomain, ICodedValueDomain)

    '        Return pCodedValueDomain

    '    Catch ex As Exception
    '        Throw ex
    '    End Try

    'End Function

End Module