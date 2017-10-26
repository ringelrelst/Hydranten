Option Explicit On 
Option Strict On

#Region " Imports namespaces "
Imports System.Runtime.InteropServices
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.SystemUI
Imports ESRI.ArcGIS.Framework
#End Region

''' -----------------------------------------------------------------------------
''' Project	 : Digipolis.Hydranten.BeheerHydranten
''' Class	 : Hydranten.BeheerHydranten.ComClassRootMenu
''' 
''' -----------------------------------------------------------------------------
''' <summary>
'''     Root-Menu "Hydrantenbeheer"
''' </summary>
''' <remarks>
''' </remarks>
''' <history>
''' 	[Kristof Vydt]	16/09/2005	Created
''' 	[Kristof Vydt]	28/09/2005	Add commands for HydrantBook
'''     [Koen Vermeer]  23/06/2008  Conversion to ArcGIS 9.2
''' </history>
''' -----------------------------------------------------------------------------
<ComClass(ComClassRootMenu.ClassId, ComClassRootMenu.InterfaceId, ComClassRootMenu.EventsId)> _
Public NotInheritable Class ComClassRootMenu
    Implements IMenuDef
    Implements IRootLevelMenu

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "7E6D4979-DB85-48D0-838C-96837355B59C"
    Public Const InterfaceId As String = "FA51E432-A0E0-4F93-A078-EC65B57C1DC2"
    Public Const EventsId As String = "EE395906-4718-4971-818C-5BA2A1B25521"
#End Region

#Region "Component Category Registration"
    <ComRegisterFunction()> _
    Public Shared Sub Reg(ByVal regKey As String)
        MxCommandBars.Register(regKey)
    End Sub

    <ComUnregisterFunction()> _
    Public Shared Sub Unreg(ByVal regKey As String)
        MxCommandBars.Unregister(regKey)
    End Sub
#End Region

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()
    End Sub

    Public ReadOnly Property Caption() As String Implements ESRI.ArcGIS.SystemUI.IMenuDef.Caption
        Get
            Return "Hydrantenbeheer"
        End Get
    End Property

    Public ReadOnly Property ItemCount() As Integer Implements ESRI.ArcGIS.SystemUI.IMenuDef.ItemCount
        Get
            Return 4
        End Get
    End Property

    Public ReadOnly Property Name() As String Implements ESRI.ArcGIS.SystemUI.IMenuDef.Name
        Get
            Return "Hydrantenbeheer"
        End Get
    End Property

    <CLSCompliant(False)> _
    Public Sub GetItemInfo(ByVal pos As Integer, ByVal itemDef As ESRI.ArcGIS.SystemUI.IItemDef) Implements ESRI.ArcGIS.SystemUI.IMenuDef.GetItemInfo
        Select Case pos
            Case 0 'Opladen uit Excell
                itemDef.ID = "Digipolis.Hydranten.BeheerHydranten.ComClassUploadHydranten"
                itemDef.Group = False
            Case 1 'Beheer...
                itemDef.ID = "Digipolis.Hydranten.BeheerHydranten.ComClassMenuBeheer"
                itemDef.Group = False
            Case 2 'Herbereken legendecode
                itemDef.ID = "Digipolis.Hydranten.BeheerHydranten.ComClassUpdateLegend"
                itemDef.Group = False
            Case 3 'Hydrantenboek...
                itemDef.ID = "Digipolis.Hydranten.BeheerHydranten.ComClassMenuHydrantenboek"
                itemDef.Group = False
        End Select
    End Sub
End Class

''' -----------------------------------------------------------------------------
''' Project	 : Digipolis.Hydranten.BeheerHydranten
''' Class	 : Hydranten.BeheerHydranten.ComClassMenuBeheer
''' 
''' -----------------------------------------------------------------------------
''' <summary>
'''     Sub-Menu "Beheer..."
''' </summary>
''' <remarks>
''' </remarks>
''' <history>
''' 	[Kristof Vydt]	16/09/2005	Created
''' </history>
''' -----------------------------------------------------------------------------
<ComClass(ComClassMenuBeheer.ClassId, ComClassMenuBeheer.InterfaceId, ComClassMenuBeheer.EventsId)> _
Public NotInheritable Class ComClassMenuBeheer
    Implements IMenuDef

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "2fdb987b-4f07-46c5-9ee2-6bb0924dfe19"
    Public Const InterfaceId As String = "b1837b07-b1f0-49d3-96ff-4d6ce1804f46"
    Public Const EventsId As String = "ae111de7-f3ae-4264-8d86-a28d3014efe9"
#End Region

#Region "Component Category Registration"
    <ComRegisterFunction()> _
    Public Shared Sub Reg(ByVal regKey As String)
        MxCommandBars.Register(regKey)
    End Sub

    <ComUnregisterFunction()> _
    Public Shared Sub Unreg(ByVal regKey As String)
        MxCommandBars.Unregister(regKey)
    End Sub
#End Region

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()
    End Sub

    Public ReadOnly Property Caption() As String Implements ESRI.ArcGIS.SystemUI.IMenuDef.Caption
        Get
            Return "Beheer..."
        End Get
    End Property

    Public ReadOnly Property ItemCount() As Integer Implements ESRI.ArcGIS.SystemUI.IMenuDef.ItemCount
        Get
            Return 3
        End Get
    End Property

    Public ReadOnly Property Name() As String Implements ESRI.ArcGIS.SystemUI.IMenuDef.Name
        Get
            Return "Beheer"
        End Get
    End Property

    <CLSCompliant(False)> _
    Public Sub GetItemInfo(ByVal pos As Integer, ByVal itemDef As ESRI.ArcGIS.SystemUI.IItemDef) Implements ESRI.ArcGIS.SystemUI.IMenuDef.GetItemInfo
        Select Case pos
            Case 0 'Beheer van hydranten
                itemDef.ID = "Digipolis.Hydranten.BeheerHydranten.ComClassBeheerHydranten"
                itemDef.Group = False
            Case 1 'Beheer van speciale gebouwen
                itemDef.ID = "Digipolis.Hydranten.BeheerHydranten.ComClassBeheerGebouwen"
                itemDef.Group = False
            Case 2 'Beheer van gevarenthema's
                itemDef.ID = "Digipolis.Hydranten.BeheerHydranten.ComClassBeheerGevaren"
                itemDef.Group = False
        End Select
    End Sub
End Class

''' -----------------------------------------------------------------------------
''' Project	 : Digipolis.Hydranten.BeheerHydranten
''' Class	 : Hydranten.BeheerHydranten.ComClassMenuHydrantenboek
''' 
''' -----------------------------------------------------------------------------
''' <summary>
'''     Sub-Menu "Hydrantenboek..."
''' </summary>
''' <remarks>
''' </remarks>
''' <history>
''' 	[Kristof Vydt]	16/09/2005	Created
''' 	[Kristof Vydt]	28/09/2005	Add commands for HydrantBook
''' </history>
''' -----------------------------------------------------------------------------
<ComClass(ComClassMenuHydrantenboek.ClassId, ComClassMenuHydrantenboek.InterfaceId, ComClassMenuHydrantenboek.EventsId)> _
Public NotInheritable Class ComClassMenuHydrantenboek
    Implements IMenuDef

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "614C509A-58CA-4167-8DBE-31A106F84505"
    Public Const InterfaceId As String = "1C7D468F-2DA1-4f9d-B1AC-C289BF557CCA"
    Public Const EventsId As String = "1449D330-14D5-4f0e-9B19-E60D381D4497"
#End Region

#Region "Component Category Registration"
    <ComRegisterFunction()> _
    Public Shared Sub Reg(ByVal regKey As String)
        MxCommandBars.Register(regKey)
    End Sub

    <ComUnregisterFunction()> _
    Public Shared Sub Unreg(ByVal regKey As String)
        MxCommandBars.Unregister(regKey)
    End Sub
#End Region

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()
    End Sub

    Public ReadOnly Property Caption() As String Implements ESRI.ArcGIS.SystemUI.IMenuDef.Caption
        Get
            Return "Hydrantenboek..."
        End Get
    End Property

    Public ReadOnly Property ItemCount() As Integer Implements ESRI.ArcGIS.SystemUI.IMenuDef.ItemCount
        Get
            Return 5
        End Get
    End Property

    Public ReadOnly Property Name() As String Implements ESRI.ArcGIS.SystemUI.IMenuDef.Name
        Get
            Return "Beheer"
        End Get
    End Property

    <CLSCompliant(False)> _
    Public Sub GetItemInfo(ByVal pos As Integer, ByVal itemDef As ESRI.ArcGIS.SystemUI.IItemDef) Implements ESRI.ArcGIS.SystemUI.IMenuDef.GetItemInfo
        Select Case pos
            Case 0 'Hydrantenpagina's aanmaken
                itemDef.ID = "DSMapBookUIPrj.CreateMapBook"
                itemDef.Group = False
            Case 1 'Hydrantenpagina's printen
                itemDef.ID = "Digipolis.Hydranten.BeheerHydranten.ComClassBookPrint"
                itemDef.Group = False
            Case 2 'Hydrantenpagina's exporteren
                itemDef.ID = "Digipolis.Hydranten.BeheerHydranten.ComClassBookExport"
                itemDef.Group = False
            Case 3 'Stratenindex printen
                itemDef.ID = "Digipolis.Hydranten.BeheerHydranten.ComClassIndexStraten"
                itemDef.Group = True
            Case 4 'Gebouwenindex printen
                itemDef.ID = "Digipolis.Hydranten.BeheerHydranten.ComClassIndexGebouwen"
                itemDef.Group = False
        End Select
    End Sub
End Class