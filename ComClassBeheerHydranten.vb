Option Explicit On 
Option Strict On

#Region "Imports namespaces"

Imports System.Runtime.InteropServices

Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.SystemUI

#End Region

''' -----------------------------------------------------------------------------
''' Project	 : Digipolis.Hydranten.BeheerHydranten
''' Class	 : Hydranten.BeheerHydranten.ComClassBeheerHydranten
''' 
''' -----------------------------------------------------------------------------
''' <summary>
'''     Command "Beheer van hydranten" with COM interface.
''' </summary>
''' <remarks>
''' </remarks>
''' <history>
''' 	[Kristof Vydt]	16/09/2005	Created
''' 	[Kristof Vydt]	23/09/2005	OnFormClose added and OnClick modified, to prevent opening of multiple forms.
'''     [Kristof Vydt]  22/02/2007  Adopt to XML configuration.
'''     [Kristof Vydt]  19/04/2007  Close form when exception occured during OnClick.
'''     [Koen Vermeer]  23/06/2008  Conversion to ArcGIS 9.2
''' </history>
''' -----------------------------------------------------------------------------
<ComClass(ComClassBeheerHydranten.ClassId, ComClassBeheerHydranten.InterfaceId, ComClassBeheerHydranten.EventsId)> _
    <CLSCompliant(False)> _
Public NotInheritable Class ComClassBeheerHydranten
    Inherits BaseCommand

#Region "Local variables"
    Dim mxApp As IMxApplication 'ArcMap application object
    Dim form As FormBeheerHydranten 'Management form
#End Region

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "EBEDEDB0-1380-49BF-8264-C7436F9A4BB3"
    Public Const InterfaceId As String = "F09D8729-D096-4833-8588-7FF01BAAAE77"
    Public Const EventsId As String = "17814031-1EBC-49D9-AA2D-4CDA5978CF58"
#End Region

#Region "Component Category Registration"
    <ComRegisterFunction()> _
    Public Shared Sub Reg(ByVal regKey As String)
        MxCommands.Register(regKey)
    End Sub

    <ComUnregisterFunction()> _
    Public Shared Sub Unreg(ByVal regKey As String)
        MxCommands.Unregister(regKey)
    End Sub
#End Region

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()
        MyBase.m_category = "Hydrantenbeheer"
        MyBase.m_caption = "Beheer hydranten"
        MyBase.m_message = "Beheer van hydranten"
        MyBase.m_toolTip = "Beheer van hydranten"
        MyBase.m_name = "Hydrantenbeheer_Beheer hydranten"

        ' Icon for the command.
        Dim BitmapName As String
        Dim BitmapStream As System.IO.Stream
        BitmapName = "Digipolis.Hydranten.BeheerHydranten." & c_Bitmap_BeheerHydranten
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

            ' Open GUI form "Beheer van hydranten".
            If form Is Nothing Then
                ' Create a new form.
                form = New FormBeheerHydranten(mxApp)
                AddHandler form.Closed, AddressOf OnFormClose
                form.Show()
                form.SetDesktopLocation(0, 20)
            Else
                ' Re-use the existing form. 
                form.WindowState = Windows.Forms.FormWindowState.Normal
                form.Focus()
                form.SetDesktopLocation(0, 20)
            End If

        Catch ex As Exception
            ErrorHandler(ex)
            form.Close()
        End Try

    End Sub

    Public Sub OnFormClose(ByVal sender As System.Object, ByVal e As System.EventArgs)
        ' Set a private variable to track that the form was closed.
        form = Nothing
    End Sub

End Class


