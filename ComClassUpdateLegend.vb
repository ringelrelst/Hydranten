Option Explicit On 
Option Strict On

#Region "Imports namespaces"

Imports System.Runtime.InteropServices

Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Framework

#End Region

''' -----------------------------------------------------------------------------
''' Project	 : Digipolis.Hydranten.BeheerHydranten
''' Class	 : Hydranten.BeheerHydranten.ComClassUpdateLegend
''' 
''' -----------------------------------------------------------------------------
''' <summary>
'''     Command "Update legende codes" with COM interface.
''' </summary>
''' <remarks>
''' </remarks>
''' <history>
''' 	[Kristof Vydt]	16/09/2005	Created
'''     [Kristof Vydt]  22/02/2007  Adopt to XML configuration.
'''     [Koen Vermeer]  23/06/2008  Conversion to ArcGIS 9.2
''' </history>
''' -----------------------------------------------------------------------------
<ComClass(ComClassUpdateLegend.ClassId, ComClassUpdateLegend.InterfaceId, ComClassUpdateLegend.EventsId)> _
    <CLSCompliant(False)> _
Public Class ComClassUpdateLegend
    Inherits BaseCommand

#Region "Local variables"
    Dim mxApp As IMxApplication 'ArcMap application object
#End Region

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "DE9506A3-C99A-4DDA-B5CE-D2CE87805C30"
    Public Const InterfaceId As String = "707BA6C5-BA22-4BD3-AD6C-8EAB9FB517F9"
    Public Const EventsId As String = "F20FFD8B-8C0C-4C87-B10A-BE53D8ECB980"
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
        MyBase.m_caption = "Herbereken legende"
        MyBase.m_message = "Herbereken legende codes van alle hydranten"
        MyBase.m_toolTip = "Herbereken legende codes van alle hydranten"
        MyBase.m_name = "Hydrantenbeheer_Herbereken legende"
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

            'Open GUI form "Update Legend Codes".
            Dim myForm As FormUpdateLegend = New FormUpdateLegend(mxApp)
            myForm.Show()
            'myForm.SetDesktopLocation(0, 20)
            myForm.UpdateLegendCodes() 'start the update procedure

        Catch ex As Exception
            ErrorHandler(ex)
        End Try
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
        If Not mxApp Is Nothing Then Marshal.ReleaseComObject(mxApp)
    End Sub

End Class


