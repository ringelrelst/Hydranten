Option Explicit On 
Option Strict On

#Region "Import namespaces"

Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Framework

Imports DSMapBookPrj
Imports DSMapBookUIPrj

#End Region

''' -----------------------------------------------------------------------------
''' Project	 : Digipolis.Hydranten.BeheerHydranten
''' Class	 : Hydranten.BeheerHydranten.ComClassBookExport
''' 
''' -----------------------------------------------------------------------------
''' <summary>
'''     Command "Export hydrantenboek" with COM interface.
''' </summary>
''' <remarks>
''' </remarks>
''' <history>
''' 	[Kristof Vydt]	28/09/2005	Created
''' 	[Kristof Vydt]	29/09/2005	OnClick reviewed.
''' 	[Kristof Vydt]	24/10/2005	Application exceptions using global constant messages.
''' 	[Kristof Vydt]	26/10/2005	Force layers not mentioned in the config file invisible.
'''     [Kristof Vydt]  22/02/2007  Adopt to XML configuration.
'''     [Kristof Vydt]  09/03/2007  Rewrite forcing visibility of map layers.
'''     [Koen Vermeer]  23/06/2008  Conversion to ArcGIS 9.2
''' </history>
''' -----------------------------------------------------------------------------
<ComClass(ComClassBookExport.ClassId, ComClassBookExport.InterfaceId, ComClassBookExport.EventsId)> _
    <CLSCompliant(False)> _
Public Class ComClassBookExport
    Inherits BaseCommand

#Region "Local variables"
    Dim mxApp As IMxApplication 'ArcMap
#End Region

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "72B3E9FE-2D15-4520-BB73-03463DBD4980"
    Public Const InterfaceId As String = "3AF69D3F-E9FB-45E3-98A2-3AE02E8A9BEF"
    Public Const EventsId As String = "6AE649B2-31A8-4DFC-8F75-25DBB3FD74B2"
#End Region

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()
        MyBase.m_category = "Hydrantenbeheer"
        MyBase.m_caption = "Hydrantenpagina's exporteren"
        MyBase.m_message = "Exporteren van pagina's uit hydrantenboek"
        MyBase.m_toolTip = "Exporteren van pagina's uit hydrantenboek"
        MyBase.m_name = "Hydrantenbeheer_BookExport"

        ' Icon for the command.
        Dim BitmapName As String
        Dim BitmapStream As System.IO.Stream
        BitmapName = "Digipolis.Hydranten.BeheerHydranten." & c_Bitmap_BookExport
        BitmapStream = GetType(ComClassBeheerHydranten).Assembly.GetManifestResourceStream(BitmapName)
        MyBase.m_bitmap = New System.Drawing.Bitmap(BitmapStream)
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

            ' Get the ArcMap document object.
            Dim gisApp As IApplication = CType(mxApp, IApplication)
            Dim mxDoc As IMxDocument = CType(gisApp.Document, IMxDocument)

            ' Force layers visibility based on configuration.
            Dim layerTable As Hashtable = Config.QueryLayerVisibility
            For Each layerName As String In layerTable.Keys
                Dim visibility As Boolean = Convert.ToBoolean(layerTable.Item(layerName))
                EnforceLayerVisibility(mxDoc, layerName, visibility)
            Next

            ' Get the MapSeries.
            Dim pMapBook As IDSMapBook = GetMapBook(gisApp)
            If pMapBook Is Nothing Then Throw New ApplicationException(c_Message_MapBookNotFound)
            If pMapBook.ContentCount < 1 Then Throw New ApplicationException(c_Message_MapBookIsEmpty)
            If Not TypeOf pMapBook.ContentItem(0) Is DSMapSeries Then Throw New ApplicationException(c_Message_MapSeriesNotFound)
            Dim pMapSeries As DSMapSeries = CType(pMapBook.ContentItem(0), DSMapSeries)

            ' Show Print form from MapBook Series.
            For i As Integer = 0 To mxDoc.ContentsViewCount - 1
                Dim pContView As IContentsView = mxDoc.ContentsView(i)
                If pContView.Name = "Map Book" Then
                    If TypeOf pContView Is DSMapBookTab Then
                        Dim pMapBookTab As DSMapBookTab = CType(pContView, DSMapBookTab)
                        'pMapBookTab.ShowExporterDialog(m_pMxApp, , CType(pMapBook, System.Object))
                        pMapBookTab.ShowExporterDialog(mxApp, CType(pMapSeries, _IDSMapSeries), Nothing)
                        Exit For
                    End If
                End If
            Next

        Catch ex As Exception
            ErrorHandler(ex)
        End Try
    End Sub

End Class


