Option Strict On
Option Explicit On 

Imports System.Runtime.InteropServices.Marshal
Imports Microsoft.Office.Interop

Module ModuleExcelAccess

    '''' -----------------------------------------------------------------------------
    '''' <summary>
    ''''     Return an open connection to the specified Excell file.
    '''' </summary>
    '''' <param name="XLSFilePath">
    ''''     The full path of an existing Excel file.
    '''' </param>
    '''' <returns>
    ''''     Return an open connection to the specified Excell file.
    '''' </returns>
    '''' <remarks>
    ''''     An exception might occur if the specified file does not match connection properties.
    '''' </remarks>
    '''' <history>
    '''' 	[Kristof Vydt]	16/09/2005	Created
    '''' 	[Kristof Vydt]	22/09/2006	Deprecated
    '''' 	[Kristof Vydt]	14/12/2006	Commented out
    '''' </history>
    '''' -----------------------------------------------------------------------------
    'Public Function ConnectXLS(ByVal XLSFilePath As String) As ADODB.Connection
    '    Dim pConnection As New ADODB.Connection
    '    Try
    '        'Online resources:
    '        '- http://msdn.microsoft.com/library/en-us/dnasdj01/html/asp0193.asp 
    '        '- http://msdn.microsoft.com/library/en-us/dv_vbcode/html/vbtskcodeexamplereadingexceldataintodataset.asp)
    '        With pConnection
    '            .Provider = "Microsoft.Jet.OLEDB.4.0"
    '            .Properties("Extended Properties").Value = "Excel 8.0"
    '            .Open(XLSFilePath)
    '        End With
    '        ConnectXLS = pConnection
    '        pConnection = Nothing
    '    Catch ex As Exception
    '        Throw ex
    '    End Try
    'End Function

    '''' -----------------------------------------------------------------------------
    '''' <summary>
    ''''     Return an Excel sheet as a recordset.
    '''' </summary>
    '''' <param name="XLSConnection">
    ''''     Connection to an existing Excel file.
    '''' </param>
    '''' <param name="XLSSheetName">
    ''''     The name of the sheet is optional.
    '''' </param>
    '''' <returns>
    ''''     The content of the first Worksheet as an ADO recordset.
    '''' </returns>
    '''' <remarks>
    '''' </remarks>
    '''' <history>
    '''' 	[Kristof Vydt]	16/09/2005	Created
    '''' 	[Kristof Vydt]	22/09/2006	Deprecated
    '''' 	[Kristof Vydt]	14/12/2006	Commented out
    '''' </history>
    '''' -----------------------------------------------------------------------------
    'Public Function ReadXLS(ByVal XLSConnection As ADODB.Connection, Optional ByVal XLSSheetName As String = "Sheet1") As ADODB.Recordset
    '    Dim pConnection As ADODB.Connection
    '    Dim pRecordSet As New ADODB.Recordset
    '    Dim closedConnection As Boolean = False
    '    Try
    '        'Open the connection to the XLS file.
    '        pConnection = XLSConnection
    '        If pConnection.State = ADODB.ObjectStateEnum.adStateClosed Then
    '            closedConnection = True
    '            pConnection.Open()
    '        End If

    '        'Open sheet as recordset.
    '        pRecordSet.Open( _
    '            "Select * from [" & XLSSheetName & "$]", _
    '            pConnection, _
    '            ADODB.CursorTypeEnum.adOpenDynamic, _
    '            ADODB.LockTypeEnum.adLockOptimistic)

    '        'Return the recordset
    '        ReadXLS = pRecordSet

    '    Catch ex As Exception
    '        Throw ex

    '    Finally
    '        'Close the XLS connection.
    '        If closedConnection Then _
    '            If pConnection.State = ADODB.ObjectStateEnum.adStateOpen Then _
    '                pConnection.Close()

    '        'Clean up pointers.
    '        pConnection = Nothing
    '        pRecordSet = Nothing
    '    End Try
    'End Function

    '''' -----------------------------------------------------------------------------
    '''' <summary>
    ''''     Open an existing Excel file for the user.
    '''' </summary>
    '''' <param name="XLSFilePath">
    ''''     The full path of an existing Excel file.
    '''' </param>
    '''' <returns>
    ''''     The Excel Workbook object.
    '''' </returns>
    '''' <remarks>
    '''' </remarks>
    '''' <history>
    '''' 	[Kristof Vydt]	16/09/2005	Created
    '''' 	[Kristof Vydt]	22/09/2006	Deprecated
    '''' 	[Kristof Vydt]	14/12/2006	Commented out
    '''' </history>
    '''' -----------------------------------------------------------------------------
    'Public Function OpenXLS(ByVal XLSFilePath As String) As Excel.Workbook
    '    Dim oExcel As Excel.ApplicationClass
    '    Dim oWorkbook As Excel.Workbook
    '    Try
    '        oExcel = New Excel.ApplicationClass
    '        oWorkbook = oExcel.Workbooks.Open(XLSFilePath)
    '        OpenXLS = oWorkbook
    '    Catch ex As Exception
    '        Throw ex
    '    Finally
    '        oExcel = Nothing
    '        oWorkbook = Nothing
    '    End Try
    'End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Return name of an Excel worksheet.
    ''' </summary>
    ''' <param name="XLSFilePath">Full path of Excel file.</param>
    ''' <param name="SheetIndex">Index of the Excel sheet.</param>
    ''' <returns>Excel worksheet name.</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	22/09/2006	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function GetXLSSheetName(ByVal XLSFilePath As String, ByVal SheetIndex As Integer) As String

        Dim oExcel As Excel.ApplicationClass = Nothing
        Dim oBooks As Excel.Workbooks = Nothing
        Dim oBook As Excel.Workbook = Nothing
        Dim oSheets As Excel.Sheets = Nothing
        Dim oSheet As Excel.Worksheet = Nothing
        Dim sheetName As String = ""

        Try
            oExcel = New Excel.ApplicationClass
            oExcel.Visible = False
            oBooks = oExcel.Workbooks
            oBooks.Open(XLSFilePath)
            oBook = oBooks.Item(1)
            oSheets = oBook.Worksheets
            oSheet = CType(oSheets.Item(SheetIndex), Excel.Worksheet)
            sheetName = oSheet.Name

        Catch ex As Exception
            ErrorHandler(ex)

        Finally
            oExcel.Quit()
            ReleaseComObject(oSheet)
            ReleaseComObject(oSheets)
            ReleaseComObject(oBook)
            ReleaseComObject(oBooks)
            ReleaseComObject(oExcel)
            oExcel = Nothing
            oBooks = Nothing
            oBook = Nothing
            oSheets = Nothing
            oSheet = Nothing
            System.GC.Collect()
        End Try

        Return sheetName

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Read data from an Excel worksheet into a new datatable.
    ''' </summary>
    ''' <param name="dataSet">The dataset where to add the new datatabel.</param>
    ''' <param name="tableName">The name of the new datatabel.</param>
    ''' <param name="xlsFilePath">The Excel file path.</param>
    ''' <param name="xlsSheetName">The Excel worksheet name.</param>
    ''' <param name="whereClause">Optional import condition.</param>
    ''' <remarks>
    '''     The first Excel row contains the column names.
    '''     The optional whereClause refers to existing Excel column names.
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	22/09/2006	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub ImportXLSWorksheet( _
            ByRef dataSet As DataSet, _
            ByVal tableName As String, _
            ByVal xlsFilePath As String, _
            ByVal xlsSheetName As String, _
            Optional ByVal whereClause As String = "")

        Dim con As New OleDb.OleDbConnection
        con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=""" & xlsFilePath & """;Extended Properties=""Excel 8.0;HDR=Yes;IMEX=1"""
        'Comment: For Microsoft Excel 8.0 (97), 9.0 (2000) and 10.0 (2002) workbooks, use Excel 8.0.
        con.Open()

        Try

            Dim cmd As New OleDb.OleDbCommand
            cmd.Connection = con
            cmd.CommandText = "SELECT * FROM [" & xlsSheetName & "$]"
            If whereClause <> "" Then cmd.CommandText = cmd.CommandText & " WHERE " & whereClause

            'Dim dataTable As dataTable = dataSet.Tables.Add("UPLOAD")
            Dim dataTable As DataTable = dataSet.Tables.Add(tableName)

            Dim adapt As New OleDb.OleDbDataAdapter(cmd)
            adapt.Fill(dataTable)

        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            con.Close()
        End Try

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Outputs a DataTable to an Excel Worksheet.
    ''' </summary>
    ''' <param name="dataTable">The data table to dump.</param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	22/09/2006	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub DataTableToXLS( _
            ByVal dataTable As DataTable)

        Dim oExcel As New Excel.Application
        Dim oBooks As Excel.Workbooks = Nothing
        Dim oBook As Excel.Workbook = Nothing
        Dim oSheets As Excel.Sheets = Nothing
        Dim oSheet As Excel.Worksheet = Nothing
        Dim oCells As Excel.Range = Nothing

        Try
            'oExcel.Visible = True
            'oExcel.DisplayAlerts = False

            'Start a new Excel workbook
            oBooks = oExcel.Workbooks
            oBooks.Add()
            oBook = oBooks.Item(1)
            oSheets = oBook.Worksheets
            oSheet = CType(oSheets.Item(1), Excel.Worksheet)
            oCells = oSheet.Cells

            ' Fill in the data
            DumpData(dataTable, oCells)
            oExcel.Visible = True

        Catch ex As Exception
            ErrorHandler(ex)

        Finally
            ' Deallocate everything
            'oExcel.Quit()
            ReleaseComObject(oCells)
            ReleaseComObject(oSheet)
            ReleaseComObject(oSheets)
            ReleaseComObject(oBook)
            ReleaseComObject(oBooks)
            ReleaseComObject(oExcel)
            oExcel = Nothing
            oBooks = Nothing
            oBook = Nothing
            oSheets = Nothing
            oSheet = Nothing
            oCells = Nothing
            System.GC.Collect()
        End Try
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Outputs a DataTable to an Excel Range.
    ''' </summary>
    ''' <param name="dt">Date table.</param>
    ''' <param name="oCells">Range of cells.</param>
    ''' <remarks>
    '''     Source: http://www.aspnetpro.com/NewsletterArticle/2003/09/asp200309so_l/asp200309so_l.asp
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	22/09/2006	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub DumpData( _
            ByVal dt As DataTable, _
            ByVal oCells As Excel.Range)

        Dim dr As DataRow, ary() As Object
        Dim iRow As Integer, iCol As Integer

        'Output Column Headers
        For iCol = 0 To dt.Columns.Count - 1
            oCells(1, iCol + 1) = dt.Columns(iCol).ToString
        Next

        'Output Data
        For iRow = 0 To dt.Rows.Count - 1
            dr = dt.Rows.Item(iRow)
            ary = dr.ItemArray
            For iCol = 0 To UBound(ary)
                oCells(iRow + 2, iCol + 1) = ary(iCol).ToString
            Next
        Next

    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Open and show the Excel file.
    ''' </summary>
    ''' <param name="xlsFilePath">The Excel file path.</param>
    ''' <param name="sheetIndex">Optional sheet index.</param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	22/09/2006	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub ShowXLS( _
                ByVal xlsFilePath As String, _
                Optional ByVal sheetIndex As Integer = 1)

        Dim oExcel As New Excel.Application
        Dim oBooks As Excel.Workbooks = Nothing
        Dim oBook As Excel.Workbook = Nothing
        Dim oSheets As Excel.Sheets = Nothing
        Dim oSheet As Excel.Worksheet = Nothing

        Try
            ' Open the XLS file.
            oExcel = New Excel.ApplicationClass
            oBooks = oExcel.Workbooks
            oBooks.Open(xlsFilePath)
            oBook = oBooks.Item(1)
            oSheets = oBook.Sheets
            oSheet = CType(oSheets.Item(sheetIndex), Excel.Worksheet)
            oSheet.Activate()
            oExcel.Visible = True

        Catch ex As Exception
            ErrorHandler(ex)

        Finally
            ' Deallocate everything
            'oExcel.Quit()
            ReleaseComObject(oSheet)
            ReleaseComObject(oSheets)
            ReleaseComObject(oBook)
            ReleaseComObject(oBooks)
            ReleaseComObject(oExcel)
            oExcel = Nothing
            oBooks = Nothing
            oBook = Nothing
            oSheets = Nothing
            oSheet = Nothing
            System.GC.Collect()
        End Try

    End Sub

End Module
