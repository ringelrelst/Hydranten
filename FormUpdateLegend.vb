Option Explicit On 
Option Strict On

#Region " Imports namespaces "
Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.Marshal
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.Geodatabase
#End Region

''' -----------------------------------------------------------------------------
''' Project	 : Digipolis.Hydranten.BeheerHydranten
''' Class	 : Hydranten.BeheerHydranten.FormUpdateLegend
''' 
''' -----------------------------------------------------------------------------
''' <summary>
'''     Update legend code attribute of all hydrants in bulk.
''' </summary>
''' <remarks>
''' </remarks>
''' <history>
''' 	[Kristof Vydt]	16/09/2005	Created
'''     [Kristof Vydt]	22/09/2005	Close button added.
'''     [Koen Vermeer]  23/06/2008  Conversion to ArcGIS 9.2
''' </history>
''' -----------------------------------------------------------------------------
Public Class FormUpdateLegend
    Inherits System.Windows.Forms.Form

#Region " Private variables "
    Private m_application As IMxApplication 'hold current ArcMap application
#End Region

#Region " Windows Form Designer generated code "

    'Public Sub New()
    '    MyBase.New()

    '    'This call is required by the Windows Form Designer.
    '    InitializeComponent()

    '    'Add any initialization after the InitializeComponent() call

    'End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents ProgressBarLegendUpdate As System.Windows.Forms.ProgressBar
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents ButtonClose As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.ProgressBarLegendUpdate = New System.Windows.Forms.ProgressBar
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.ButtonClose = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'ProgressBarLegendUpdate
        '
        Me.ProgressBarLegendUpdate.Enabled = False
        Me.ProgressBarLegendUpdate.Location = New System.Drawing.Point(8, 8)
        Me.ProgressBarLegendUpdate.Name = "ProgressBarLegendUpdate"
        Me.ProgressBarLegendUpdate.Size = New System.Drawing.Size(360, 16)
        Me.ProgressBarLegendUpdate.TabIndex = 8
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 32)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(328, 16)
        Me.Label1.TabIndex = 9
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(8, 48)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(312, 16)
        Me.Label2.TabIndex = 10
        '
        'ButtonClose
        '
        Me.ButtonClose.Location = New System.Drawing.Point(320, 40)
        Me.ButtonClose.Name = "ButtonClose"
        Me.ButtonClose.Size = New System.Drawing.Size(48, 23)
        Me.ButtonClose.TabIndex = 11
        Me.ButtonClose.Text = "Sluiten"
        '
        'FormUpdateLegend
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(378, 72)
        Me.Controls.Add(Me.ButtonClose)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.ProgressBarLegendUpdate)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.Name = "FormUpdateLegend"
        Me.Text = "Herbereken legende codes"
        Me.TopMost = True
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region " Form controls events "

    Private Sub ButtonClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonClose.Click
        Me.Close()
    End Sub

#End Region

#Region " Overloaded constructor "
    <CLSCompliant(False)> _
    Public Sub New(ByVal ArcMapApplication As IMxApplication)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        m_application = ArcMapApplication

    End Sub
#End Region

#Region " Utility procedures "
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     The actual procedure to update the legend code of all the hydrant features.
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Kristof Vydt]	16/09/2005	Created
    ''' 	[Kristof Vydt]	17/10/2005	Add progress monitor update after update loop.
    '''                                 Initialize values for the calculation of a legend code.
    ''' 	[Kristof Vydt]	02/08/2006	Replace some global constant by a function that reads the value from config file.
    ''' 	[Kristof Vydt]	22/09/2006	Use new GetAttributeValue() and UpdateLegendCode() methods.
    ''' 	[Kristof Vydt]	13/10/2006	Correct type error in hydrant attribute keyword.
    '''                                 Correct logic for updating number of updated legend codes.
    '''     [Kristof Vydt]  22/02/2007  Adopt to UpdateLegendCode that return a boolean now.
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub UpdateLegendCodes()

        'ArcGIS pointers.
        Dim arcMapDocument As IMxDocument = Nothing
        Dim hydrantLayer As IFeatureLayer = Nothing
        Dim hydrantCursor As IFeatureCursor = Nothing
        Dim hydrantFeature As IFeature = Nothing

        'Counter variables.
        Dim totalFeaturesCounter As Integer
        Dim updatedFeaturesCounter As Integer

        Try

            'Get all hydrant features in a cursor.
            arcMapDocument = CType(CType(m_application, IApplication).Document, IMxDocument)
            hydrantLayer = GetFeatureLayer(arcMapDocument.FocusMap, GetLayerName("Hydrant"))
            If hydrantLayer Is Nothing Then Throw New LayerNotFoundException(GetLayerName("Hydrant"))
            hydrantCursor = hydrantLayer.Search(Nothing, Nothing)

            'Initialize progress monitor.
            totalFeaturesCounter = 0
            hydrantFeature = hydrantCursor.NextFeature
            While Not hydrantFeature Is Nothing
                totalFeaturesCounter += 1
                hydrantFeature = hydrantCursor.NextFeature
            End While
            Me.ProgressBarLegendUpdate.Minimum = 0
            Me.ProgressBarLegendUpdate.Maximum = totalFeaturesCounter

            ' Loop through each hydrant feature in the cursor.
            totalFeaturesCounter = 0
            updatedFeaturesCounter = 0
            hydrantCursor = hydrantLayer.Search(Nothing, Nothing)
            hydrantFeature = hydrantCursor.NextFeature
            While Not hydrantFeature Is Nothing

                ' Progress monitor update.
                totalFeaturesCounter += 1
                Me.ProgressBarLegendUpdate.Value = totalFeaturesCounter
                Me.Label1.Text = Replace(c_Message_UpdateLegendProgress, "^0", CStr(totalFeaturesCounter))
                Me.Label2.Text = Replace(c_Message_UpdateLegendCount, "^0", CStr(updatedFeaturesCounter))
                Me.Refresh()

                ' Update the legend code attribute for current feature.
                Dim overwritten As Boolean = ModuleHydrant.UpdateLegendCode(hydrantFeature)
                If overwritten Then updatedFeaturesCounter += 1

                ' Next feature.
                hydrantFeature = hydrantCursor.NextFeature
            End While

            ' Progress monitor update.
            Me.ProgressBarLegendUpdate.Value = totalFeaturesCounter
            Me.Label1.Text = Replace(c_Message_UpdateLegendProgress, "^0", CStr(totalFeaturesCounter))
            Me.Label2.Text = Replace(c_Message_UpdateLegendCount, "^0", CStr(updatedFeaturesCounter))
            Me.Refresh()

            ' Partial refresh of the map.
            If updatedFeaturesCounter > 0 Then _
                arcMapDocument.ActiveView.PartialRefresh(esriViewDrawPhase.esriViewGeography, hydrantLayer, Nothing)

        Catch ex As Exception
            Throw ex

        Finally
            ' Release COM objects.
            If Not arcMapDocument Is Nothing Then
                ReleaseComObject(arcMapDocument)
                arcMapDocument = Nothing
            End If
            If Not hydrantLayer Is Nothing Then
                ReleaseComObject(hydrantLayer)
                hydrantLayer = Nothing
            End If
            If Not hydrantCursor Is Nothing Then
                ReleaseComObject(hydrantCursor)
                hydrantCursor = Nothing
            End If
            If Not hydrantFeature Is Nothing Then
                ReleaseComObject(hydrantFeature)
                hydrantFeature = Nothing
            End If
            GC.Collect()
        End Try
    End Sub
#End Region

End Class
