''' -----------------------------------------------------------------------------
''' <summary>
'''     FontConverter Class: Convert a Font to an OLE Font.
''' </summary>
''' <remarks>
'''     Usage:  Dim pFont As System.Drawing.Font = New System.Drawing.Font("ESRI Cartography", 18)
'''             Dim pFontDisp As stdole.IFontDisp = FontConverter.FontToOLEFont(pFont)
''' </remarks>
''' <history>
''' 	[Kristof Vydt]	16/09/2005	Created
'''     [Koen Vermeer]  23/06/2008  Conversion to ArcGIS 9.2
''' </history>
''' -----------------------------------------------------------------------------
''' 

Public Class FontConverter
    Inherits System.Windows.Forms.AxHost

    'Purpose: convert a Font to an OLE Font 
    'Usage:  Dim pFont As System.Drawing.Font = New System.Drawing.Font("ESRI Cartography", 18)
    '        Dim pFontDisp As stdole.IFontDisp = FontConverter.FontToOLEFont(pFont)

    Public Sub New()
        MyBase.New("")
    End Sub

    <CLSCompliant(False)> _
    Public Shared Function FontToOLEFont(ByVal font As System.Drawing.Font) As stdole.IFontDisp
        Return GetIFontFromFont(font)
    End Function

End Class
