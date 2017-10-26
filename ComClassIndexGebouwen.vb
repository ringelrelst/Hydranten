Option Explicit On 
Option Strict On

#Region "Imports namespaces"

Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Framework

#End Region

''' -----------------------------------------------------------------------------
''' Project	 : Digipolis.Hydranten.BeheerHydranten
''' Class	 : Hydranten.BeheerHydranten.ComClassIndexGebouwen
''' 
''' -----------------------------------------------------------------------------
''' <summary>
'''     Command "Hydrantenboek Gebouwenindex printen" with COM interface.
''' </summary>
''' <remarks>
''' </remarks>
''' <history>
''' 	[Kristof Vydt]	16/09/2005	Created
'''     [Kristof Vydt]  22/02/2007  Adopt to XML configuration.
'''     [Koen Vermeer]  23/06/2008  Conversion to ArcGIS 9.2
''' </history>
''' -----------------------------------------------------------------------------
<ComClass(ComClassIndexGebouwen.ClassId, ComClassIndexGebouwen.InterfaceId, ComClassIndexGebouwen.EventsId)> _
    <CLSCompliant(False)> _
Public Class ComClassIndexGebouwen
    Inherits BaseCommand

#Region "Local variables"
    Dim mxApp As IMxApplication 'ArcMap
#End Region

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "2253C71C-77D0-404A-9004-709B595309FC"
    Public Const InterfaceId As String = "18CB9477-3068-4B98-9AFC-6D2FB2D2E5F1"
    Public Const EventsId As String = "1231207C-D9A8-4628-95B1-D4D4DC10029D"
#End Region

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()
        MyBase.m_category = "Hydrantenbeheer"
        MyBase.m_caption = "Gebouwenindex..."
        MyBase.m_message = "Afdrukken van gebouwenindex uit hydrantenboek"
        MyBase.m_toolTip = "Afdrukken van gebouwenindex uit hydrantenboek"
        MyBase.m_name = "Hydrantenbeheer_Gebouwenindex"
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

            ' Open GUI form "Gebouwenindex".
            Dim indexForm As FormIndexGebouwen = New FormIndexGebouwen(mxApp)
            indexForm.ShowDialog()

        Catch ex As Exception
            ErrorHandler(ex)
        End Try
    End Sub

End Class


