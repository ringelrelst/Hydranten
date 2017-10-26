Option Explicit On 
Option Strict On

#Region "Import namespaces"

Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.framework

#End Region

''' -----------------------------------------------------------------------------
''' Project	 : Digipolis.Hydranten.BeheerHydranten
''' Class	 : Hydranten.BeheerHydranten.ComClassUploadHydranten
''' 
''' -----------------------------------------------------------------------------
''' <summary>
'''     Command "Opladen hydranten" with COM interface.
''' </summary>
''' <remarks>
''' </remarks>
''' <history>
''' 	[Kristof Vydt]	16/09/2005	Created
'''     [Kristof Vydt]  22/02/2007  Adopt to XML configuration.
'''     [Koen Vermeer]  23/06/2008  Conversion to ArcGIS 9.2
''' </history>
''' -----------------------------------------------------------------------------
<ComClass(ComClassUploadHydranten.ClassId, ComClassUploadHydranten.InterfaceId, ComClassUploadHydranten.EventsId)> _
    <CLSCompliant(False)> _
Public Class ComClassUploadHydranten
    Inherits BaseCommand

#Region "Local variables"
    Dim mxApp As IMxApplication 'ArcMap
#End Region

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "CB07BEED-8411-4614-B425-F6A65F6609AE"
    Public Const InterfaceId As String = "1035E249-198E-4BF1-A24C-D6D62F9E295C"
    Public Const EventsId As String = "07D4C451-AB38-4CEC-A556-5540F2CF43CE"
#End Region

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()
        MyBase.m_category = "Hydrantenbeheer"
        MyBase.m_caption = "Opladen hydranten"
        MyBase.m_message = "Opladen en integreren van hydranten uit Excell"
        MyBase.m_toolTip = "Opladen en integreren van hydranten uit Excell"
        MyBase.m_name = "Hydrantenbeheer_Opladen hydranten"
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

            ' Open GUI form "Laden van hydranten".
            Dim uploadForm As FormUploadHydranten = New FormUploadHydranten(mxApp)
            uploadForm.ShowDialog()

        Catch ex As Exception
            ErrorHandler(ex)
        End Try
    End Sub
End Class


