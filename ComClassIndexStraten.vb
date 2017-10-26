Option Explicit On 
Option Strict On

#Region "Imports namespaces"

Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Framework

#End Region

''' -----------------------------------------------------------------------------
''' Project	 : Digipolis.Hydranten.BeheerHydranten
''' Class	 : Hydranten.BeheerHydranten.ComClassIndexStraten
''' 
''' -----------------------------------------------------------------------------
''' <summary>
'''     Command "Hydrantenboek Stratenindex printen" with COM interface.
''' </summary>
''' <remarks>
''' </remarks>
''' <history>
''' 	[Kristof Vydt]	16/09/2005	Created
'''     [Kristof Vydt]  22/02/2007  Adopt to XML configuration.
'''     [Koen Vermeer]  23/06/2008  Conversion to ArcGIS 9.2
''' </history>
''' -----------------------------------------------------------------------------
<ComClass(ComClassIndexStraten.ClassId, ComClassIndexStraten.InterfaceId, ComClassIndexStraten.EventsId)> _
    <CLSCompliant(False)> _
Public Class ComClassIndexStraten
    Inherits BaseCommand

#Region "Local variables"
    Dim mxApp As IMxApplication 'ArcMap
#End Region

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "0E6F61FE-2D89-4310-92D2-6E365C202671"
    Public Const InterfaceId As String = "B2CFABE8-D245-4C98-A294-7A17E18FC384"
    Public Const EventsId As String = "F28DE6E6-449E-48BC-99D4-9A3314431503"
#End Region

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()
        MyBase.m_category = "Hydrantenbeheer"
        MyBase.m_caption = "Stratenindex..."
        MyBase.m_message = "Afdrukken van stratenindex uit hydrantenboek"
        MyBase.m_toolTip = "Afdrukken van stratenindex uit hydrantenboek"
        MyBase.m_name = "Hydrantenbeheer_Stratenindex"
    End Sub

    Public Overrides Sub OnCreate(ByVal hook As Object)
        If Not (hook Is Nothing) Then
            If TypeOf (hook) Is IMxApplication Then 'ArcMap
                mxApp = CType(hook, IMxApplication)
            End If
        End If
    End Sub

    Public Overrides Sub OnClick()
        Try

            ' Load configuration.
            Config = New AppSettings(mxApp)

            ' Open GUI form "Stratenindex".
            Dim indexForm As FormIndexStraten = New FormIndexStraten(mxApp)
            indexForm.ShowDialog()

        Catch ex As Exception
            ErrorHandler(ex)
        End Try
    End Sub
End Class


