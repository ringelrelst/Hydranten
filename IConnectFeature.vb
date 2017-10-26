Option Explicit On 
Option Strict On

''' -----------------------------------------------------------------------------
''' Project	 : Digipolis.Hydranten.BeheerHydranten
''' Interface	 : Hydranten.BeheerHydranten.IConnectFeature
''' 
''' -----------------------------------------------------------------------------
''' <summary>
'''     Each form that uses the functionality "ConnectFeature"
'''     must implement this interface. This is necessary for the method 
'''     ReturnFeature in ModuleConnectFeature.
''' </summary>
''' <remarks>
''' </remarks>
''' <history>
''' 	[Kristof Vydt]	16/09/2005	Created
''' 	[Kristof Vydt]	05/10/2005	Add Toolbutton property.
''' </history>
''' -----------------------------------------------------------------------------
Public Interface IConnectFeature

    ' Each of these properties is implemented in the forms
    ' to refer to a textbox value.

    Property Straatnaam() As String
    Property Straatcode() As String
    Property Postcode() As String
    ReadOnly Property Toolbutton() As Windows.Forms.CheckBox

End Interface
