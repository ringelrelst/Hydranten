Option Explicit On 
Option Strict On

''' -----------------------------------------------------------------------------
''' <summary>
'''     Define all global variables &amp; constants.
''' </summary>
''' <remarks>
''' </remarks>
''' <history>
''' 	[Kristof Vydt]	16/09/2005	Created
''' 	[Kristof Vydt]	23/09/2005	Layer names all in lowercase &amp; msgbox titles added.
''' 	[Kristof Vydt]	28/09/2005	Bitmap references for hydrantbook added.
''' 	[Kristof Vydt]	11/10/2005	Rephrase some error messages.
''' 	[Kristof Vydt]	24/10/2005	Messages added.
''' 	[Kristof Vydt]	25/10/2005	Messages added.
''' 	[Kristof Vydt]	13/07/2006	Messages added.
''' 	[Kristof Vydt]	17/07/2006	Messages added.
''' 	[Kristof Vydt]	02/08/2006	Replace some global constant by a function that reads the value from config file.
''' 	[Kristof Vydt]	11/08/2006	Title added.
'''     [Kristof Vydt]  18/08/2006  Add marker element name property as global constant.
'''     [Kristof Vydt]  31/08/2006  Modify c_Message_NoHydrantToCopyAttributes text.
'''     [Kristof Vydt]  15/09/2006  Add c_Message_UploadFinished and remove Imports.
'''     [Kristof Vydt]  21/09/2006  Add c_Message_NonUniqueColumn, c_Message_NoExportData.
'''     [Kristof Vydt]  28/09/2006  Add c_Message_DepricateHydrants.
'''     [Kristof Vydt]  20/02/2007  Add c_ZoomPolygonBuffer.
'''     [Kristof Vydt]  21/02/2007  Introducing the XML configuration file to replace the INI file.
'''     [Kristof Vydt]  08/03/2007  Add c_FileName_ConfigSchema.
'''     [Kristof Vydt]  14/03/2007  Add c_AllowedSortChars.
'''     [Kristof Vydt]  15/03/2007  Add c_Title_ExportStratenindex &amp; c_Title_ExportGebouwenindex.
'''     [Kristof Vydt]  21/03/2007  Add c_FilePath_IndexGebouwType.
'''     [Koen Vermeer]  23/06/2008  Conversion to ArcGIS 9.2
''' </history>
''' -----------------------------------------------------------------------------
Module ModuleGlobals

    ' Debug status value.
    Public Const c_DebugStatus As Boolean = False
    'Public Const c_DebugStatus As Boolean = True

    ' Global file references.
    Public Config As AppSettings 'Configuration object
    Public g_FilePath_Config As String 'Configuration-file resides next to the mxd-files, in the same folder.
    Public Const c_FileName_Config As String = "Config.xml" 'Filename of configuration file.
    Public Const c_FileName_ConfigSchema As String = "Config.xsd" 'Filename of configuration schema file.
    Public Const c_FileName_WordTemplateIndexStraten As String = "brandweer.dot" 'Filename of the Microsoft Word template (.dot) file in the WorkgroupTemplates folder, used for Index Straten.
    Public Const c_FileName_WordTemplateIndexGebouwen As String = "brandweer.dot" 'Filename of the Microsoft Word template (.dot) file in the WorkgroupTemplates folder, used for Index Gebouwen.
    Public Const c_FileDir_Icons As String = "Icons"
    Public Const c_FileDir_Output As String = "Output"
    Public Const c_FilePath_IndexStraten As String = "c:\bwindex\index.txt" 'Full path where street index export file must be stored.
    Public Const c_FilePath_IndexGebouwNummer As String = "c:\bwindex\gebouwnr.txt" 'Full path where building index export file (sorted on number) must be stored.
    Public Const c_FilePath_IndexGebouwNaam As String = "c:\bwindex\gebouwab.txt" 'Full path where building index export file (sorted on name) must be stored.
    Public Const c_FilePath_IndexGebouwType As String = "c:\bwindex\gebouwtype.txt" 'Full path where building index export file (sorted on type) must be stored.

    ' Macro names in Microsoft Word template files.
    Public Const c_MacroName_IndexStraten As String = "Statenindex.MAIN"
    Public Const c_MacroName_IndexGebouwen As String = "Gebouwenindex.MAIN"

    ' Fixed zoom buffer (in map coord units) when zooming to a point.
    Public Const c_ZoomPointBuffer As Double = 200D

    ' Relative zoom buffer (in decimal %) when zooming to a polygon/line.
    Public Const c_ZoomPolygonBuffer As Double = 5D

    ' Messages for validations, errors, warnings.
    Public Const c_Message_EmptyFeatureSet As String = _
                                "Lege feature set."
    Public Const c_Message_AanduidingIsEmpty As String = _
                                "Het attribuut AANDUIDING moet ingevuld zijn."
    Public Const c_Message_BrandweerNrIsEmpty As String = _
                                "Het attribuut BRANDWEERNR moet ingevuld zijn."
    Public Const c_Message_BrandweerNrIsAlreadyInUse As String = _
                                "Het huidige BRANDWEERNR is reeds in gebruik door een actieve, ondergrondse hydrant."
    Public Const c_Message_BronIsEmpty As String = _
                                "Het attribuut BRON moet ingevuld zijn."
    Public Const c_Message_CorrectBeforeContinue As String = _
                                "Corrigeer deze fout(en) alvorens verder te gaan."
    Public Const c_Message_DiameterIsEmpty As String = _
                                "Het attribuut DIAMETER moet ingevuld zijn."
    Public Const c_Message_EinddatumIsNotEmpty As String = _
                                "EINDDATUM moet leeg zijn zolang status niet op HISTORIEK staat."
    Public Const c_Message_HistoricWithoutEinddatum As String = _
                                "Een EINDDATUM moet gespecifieerd zijn indien status op HISTORIEK staat."
    Public Const c_Message_HydrantNotConnected As String = _
                                "De hydrant is niet geconnecteerd aan straat/dok/park."
    Public Const c_Message_HydrantNotLabelled As String = _
                                "De hydrant heeft geen label."
    Public Const c_Message_HistoricHydrantLabelled As String = _
                                "De historische hydrant mag geen label hebben."
    Public Const c_Message_InvalidCoords As String = _
                                "De attribuut-coördinaten zijn niet correct ingevuld."
    Public Const c_Message_LeidingtypeIsEmpty As String = _
                                "Het attribuut LEIDINGTYPE moet ingevuld zijn."
    Public Const c_Message_LiggingIsEmpty As String = _
                                "Het attribuut LIGGING moet ingevuld zijn."
    Public Const c_Message_SaveChanges As String = _
                                "Wenst u de wijzigingen te bewaren ?"
    Public Const c_Message_NoConnectFeatureSelection As String = _
                                "Er werd geen straat, dok of park geselecteerd." & vbNewLine & _
                                "Maak een nieuwe selectie om te connecteren."
    Public Const c_Message_ConfirmDeleteAnno As String = _
                                "Wenst u alle gerelateerde labels te verwijderen ?"
    Public Const c_Message_DeleteAnnoCount As String = _
                                "^0 gerelateerde label(s) werd(en) verwijderd." '<^0> is replace by the number
    Public Const c_Message_UpdateLegendProgress As String = _
                                "^0 feature(s) werd(en) gecontroleerd." '<^0> is replace by the number
    Public Const c_Message_UpdateLegendCount As String = _
                                "Van ^0 feature(s) werd(en) de legende code vernieuwd." '<^0> is replace by the number
    Public Const c_Message_NoFeaturesToCopyAddress As String = _
                                "Er is geen hoofdgebouw of straatas geselecteerd." & vbNewLine & _
                                "Selecteer één hoofdgebouw of één straatas om adres over te nemen."
    Public Const c_Message_CopyAddressFromMultipleObjects As String = _
                                "Er is meer dan één hoofdgebouw of straatas geselecteerd." & vbNewLine & _
                                "Selecteer één hoofdgebouw of één straatas om adres over te nemen."
    Public Const c_Message_NoHydrantToCopyAttributes As String = _
                                "Er is geen geldige hydrant van de huidige sector geselecteerd." & vbNewLine & _
                                "Selecteer één hydrant met status 'verwijderd'" & vbNewLine & _
                                "om attributen over te nemen."
    Public Const c_Message_WrongStatusToCopyAttributes As String = _
                                "De geselecteerde hydrant heeft niet status 'verwijderd'." & vbNewLine & _
                                "Selecteer één hydrant met status 'verwijderd'" & vbNewLine & _
                                "om attributen over te nemen."
    Public Const c_Message_MultipleHydrantsToCopyAttributes As String = _
                                "Er is meer dan één hydrant geselecteerd." & vbNewLine & _
                                "Selecteer één hydrant met status 'verwijderd'" & vbNewLine & _
                                "om attributen over te nemen."
    Public Const c_Message_MapBookNotFound As String = _
                                "MapBook is niet gevonden."
    Public Const c_Message_MapBookIsEmpty As String = _
                                "MapBook is leeg."
    Public Const c_Message_MapSeriesNotFound As String = _
                                "Eerste item in MapBook is geen MapSeries."
    Public Const c_Message_LeverancierNrNotUnique As String = _
                                "Het leveranciernummer '^0' is niet uniek in de databank." '<^0> is replace by the number
    Public Const c_Message_HydrantChangesNotStored As String = _
                                "De wijzigingen aan de hydrant (OBJECTID=^0)" & vbNewLine & _
                                "kunnen niet worden bewaard. De feature werd niet gevonden." '<^0> is replace by the number
    Public Const c_Message_GevaarAanduidingIsEmpty As String = _
                                "Aanduiding moet ingevuld zijn."
    Public Const c_Message_GevaarStraatnaamIsEmpty As String = _
                                "Straatnaam moet ingevuld zijn."
    Public Const c_Message_NoLinkedHydrants As String = _
                                "!!! geen hydranten"
    Public Const c_Message_SaveEdits As String = _
                                "Alvorens verder te gaan zal de huidige editeer sessie worden gesloten." & vbNewLine & _
                                "Wenst u de niet-bevestigde wijzigingen te bewaren?" & vbNewLine & vbNewLine & _
                                "Klik op 'Ja'/'Yes' om wijzigingen te bewaren." & vbNewLine & _
                                "Wijzigingen gaan verloren indien u op 'Nee'/'No' klikt."
    Public Const c_Message_ModifySourceHydrant As String = _
                                "Wenst u de einddatum en status van de geselecteerde hydrant nu bij te werken, en labels " & vbNewLine & _
                                "van de geselecteerde hydrant nu over te dragen naar de hydrant die u momenteel beheert?"
    Public Const c_Message_UploadFinished As String = _
                                "Opladen is beëindigd."
    Public Const c_Message_DepricateHydrants As String = _
                                "Status 'Verwijderd' toekennen aan geregistreerde hydranten?"
    Public Const c_Message_NonUniqueColumn As String = _
                                "Kolom '^0' is bevat niet-unieke waarden." '<^0> is replace by the column name
    Public Const c_Message_NoExportData As String = _
                                "Er zijn geen gegevens om te exporteren."

    ' Titles for messageboxes.
    Public Const c_Title_DeleteAnno As String = "Verwijder gerelateerde labels"
    Public Const c_Title_SaveEdits As String = "Wijzigingen bewaren"
    Public Const c_Title_BeheerHydranten As String = "Beheer van hydranten"
    Public Const c_Title_BeheerGebouwen As String = "Beheer van speciale gebouwen"
    Public Const c_Title_BeheerGevaren As String = "Beheer van gevarenthema's"
    Public Const c_Title_CopyAttributes As String = "Attributen overnemen"
    Public Const c_Title_OpladenHydranten As String = "Opladen hydranten"
    Public Const c_Title_ExportStratenindex As String = "Export stratenindex"
    Public Const c_Title_ExportGebouwenindex As String = "Export gebouwenindex"

    ' The text that is used to concatinate multiple strings into one string.
    ' Also to split a long string in substrings (like the postcodes from the configuration ini-file).
    Public Const c_ListSeparator As Char = ";"c

    ' Bitmap names for toolbutton icons.
    Public Const c_Bitmap_BeheerHydranten As String = "HYDRANT1.BMP"
    Public Const c_Bitmap_BeheerGebouwen As String = "GEBOUW1.BMP"
    Public Const c_Bitmap_BeheerGevaren As String = "GEVAAR1.BMP"
    Public Const c_Bitmap_BookPrint As String = "PRINT1.BMP"
    Public Const c_Bitmap_BookExport As String = "EXPORT1.BMP"

    ' Marker element name property.
    Public Const c_MarkerTag As String = "BRANDWEER"

    ' String with all characters that are allowed in the sorting column.
    ' Used when filling up the "straatlijst" lookup table.
    Public Const c_AllowedSortChars As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"

End Module
