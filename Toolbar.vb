Option Explicit On 
Option Strict On

Imports System.Runtime.InteropServices
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.SystemUI

''' -----------------------------------------------------------------------------
''' Project	 : Digipolis.Hydranten.BeheerHydranten
''' Class	 : Hydranten.BeheerHydranten.ComClassToolbar
''' 
''' -----------------------------------------------------------------------------
''' <summary>
'''     The one toolbar that this project contains.
''' </summary>
''' <remarks>
'''     The toolbar is registered for ArcGIS, but not autmatically visible.
'''     To display this toolbar, open ArcMap, go to the menu Tools > Customize.
'''     On the 'Toolbars' tab, check the checkbox in front of 'Hydrantenbeheer'.
''' </remarks>
''' <history>
''' 	[Kristof Vydt]	16/09/2005	Created
''' 	[Kristof Vydt]	28/09/2005	Add tools for HydrantBook
'''     [Kristof Vydt]  14/12/2006  Item 7 (search streetname) removed, that was already hidden to the user.
'''     [Koen Vermeer]  23/06/2008  Conversion to ArcGIS 9.2
''' </history>
''' -----------------------------------------------------------------------------
<ComClass(ComClassToolbar.ClassId, ComClassToolbar.InterfaceId, ComClassToolbar.EventsId)> _
Public NotInheritable Class ComClassToolbar
    Implements IToolBarDef

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "0886A63B-F3D3-488B-83B4-13C6DCE0FF01"
    Public Const InterfaceId As String = "3A5E5FD0-C559-4037-9DE3-8152BF5B8ECE"
    Public Const EventsId As String = "6C5DB2EF-7FF5-4D3F-8520-899A81CB1A32"
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

    Public ReadOnly Property Caption() As String Implements ESRI.ArcGIS.SystemUI.IToolBarDef.Caption
        Get
            ' Set the string that appears as the toolbar's title
            Return "Hydrantenbeheer"
        End Get
    End Property

    <CLSCompliant(False)> _
    Public Sub GetItemInfo(ByVal pos As Integer, ByVal itemDef As ESRI.ArcGIS.SystemUI.IItemDef) Implements ESRI.ArcGIS.SystemUI.IToolBarDef.GetItemInfo
        ' Define the commands that will be on the toolbar. The 1st command
        ' will be the custom command MyCustomTool. The 2nd and 3rd commands will
        ' be the builtin AddData commands and ZoomIn tool.
        ' ID is the ProgID of the command. Group determines whether the command
        ' begins a new group on the toolbar
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
            Case 3 'Create hydrantenboek
                itemDef.ID = "DSMapBookUIPrj.CreateMapBook"
                itemDef.Group = True
            Case 4 'Print hydrantenboek
                itemDef.ID = "Digipolis.Hydranten.BeheerHydranten.ComClassBookPrint"
                itemDef.Group = False
            Case 5 'Export hydrantenboek
                itemDef.ID = "Digipolis.Hydranten.BeheerHydranten.ComClassBookExport"
                itemDef.Group = False
            Case 6 'Edit Annotation Tool
                itemDef.ID = "esriEditor.AnnoEditTool"
                itemDef.Group = True
        End Select
    End Sub

    Public ReadOnly Property ItemCount() As Integer Implements ESRI.ArcGIS.SystemUI.IToolBarDef.ItemCount
        Get
            'Set how many commands will be on the toolbar
            Return 7
        End Get
    End Property

    Public ReadOnly Property Name() As String Implements ESRI.ArcGIS.SystemUI.IToolBarDef.Name
        Get
            ' Set the internal name of the toolbar.
            Return "Hydrantenbeheer"
        End Get
    End Property

End Class


