Option Explicit On 
Option Strict On

Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.Utility.BaseClasses

''' -----------------------------------------------------------------------------
''' Project	 : Digipolis.Hydranten.BeheerHydranten
''' Class	 : Hydranten.BeheerHydranten.ComClassPrintMapBook
''' 
''' -----------------------------------------------------------------------------
''' <summary>
'''     Command "MapBook Series printen" with COM interface.
''' </summary>
''' <remarks>
''' </remarks>
''' <history>
''' 	[Kristof Vydt]	16/09/2005	Created
''' </history>
''' -----------------------------------------------------------------------------
<ComClass(ComClassPrintMapBook.ClassId, ComClassPrintMapBook.InterfaceId, ComClassPrintMapBook.EventsId)> _
Public Class ComClassPrintMapBook
    Inherits BaseCommand

    Dim m_application As IMxApplication 'ArcMap

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "064303C6-06B8-4D0A-890C-E8B47FE72FB2"
    Public Const InterfaceId As String = "216C993A-A30E-4358-AA7B-F1BD90B9ADAF"
    Public Const EventsId As String = "63974E15-2A0F-4B75-BCA3-7D8EFE1D97ED"
#End Region

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()
        MyBase.m_category = "Hydrantenbeheer"
        MyBase.m_caption = "Hydrantenpagina..."
        MyBase.m_message = "Afdrukken van pagina's uit hydrantenboek"
        MyBase.m_toolTip = "Afdrukken van pagina's uit hydrantenboek"
        MyBase.m_name = "Hydrantenbeheer_Hydrantenpagina"
    End Sub

    Public Overrides Sub OnCreate(ByVal hook As Object)
        If Not (hook Is Nothing) Then
            If TypeOf (hook) Is IMxApplication Then 'ArcMap
                m_application = CType(hook, IMxApplication)
            End If
        End If
    End Sub

    Public Overrides Sub OnClick()
        Try

            'Set configuration file path global variable.
            Dim pArcGisApplication As IApplication = CType(m_application, IApplication)
            g_FilePath_Config = GetConfigFilePath(pArcGisApplication)

            'TODO: Run Print pages from MapBook Series.
            MsgBox(MyBase.m_toolTip)

            'Dim pMapBook As DSMapBookPrj.IDSMapBook = GetMapBook(CType(m_application, IApplication))
            'If pMapBook.ContentCount > 0 Then
            '    'TODO
            'End If

        Catch ex As Exception
            ErrorHandler(ex)
        End Try
    End Sub


End Class


