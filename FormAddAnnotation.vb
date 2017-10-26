Option Explicit On 
Option Strict On

#Region " Imports namespaces "
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geometry
#End Region

''' -----------------------------------------------------------------------------
''' Project	 : Digipolis.Hydranten.BeheerHydranten
''' Class	 : Hydranten.BeheerHydranten.FormAddAnnotation
''' 
''' -----------------------------------------------------------------------------
''' <summary>
'''     Form for creating new annotations.
'''     The user has to select an annotation class name from the list.
''' </summary>
''' <remarks>
''' </remarks>
''' <history>
''' 	[Kristof Vydt]	21/10/2005	Created
''' 	[Kristof Vydt]	24/10/2005	Annotation class name also used as symbol name.
'''     [Koen Vermeer]  23/06/2008  Conversion to ArcGIS 9.2
''' </history>
''' -----------------------------------------------------------------------------
Public Class FormAddAnnotation
    Inherits System.Windows.Forms.Form

#Region " Private variables "
    Dim m_annoLayer As IAnnotationLayer 'The layer to store the new annotation
    Dim m_annoClassName As String       'The name of the annotation class
    'Dim m_annoClassID As Integer        'The annotation class ID
    'Dim m_symbolID As Integer           'The annotation symbol ID (currently assumed to be the same as AnnotationClassID)
    Dim m_pointGeom As IPoint           'The geom location of the new annotation
    Dim m_displayText As String         'The text that the annotation must display on the map
    Dim m_linkFieldName As String       'The name of the field that stores the feature link
    Dim m_linkFieldValue As String      'The feature link value that must be stored
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
    Friend WithEvents ButtonOK As System.Windows.Forms.Button
    Friend WithEvents ButtonCancel As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TextBoxLayerName As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents LabelLayerName As System.Windows.Forms.Label
    Friend WithEvents ComboBoxClasses As System.Windows.Forms.ComboBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TextBoxDisplayText As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents TextBoxLocation As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents TextBoxLinkID As System.Windows.Forms.TextBox
    Friend WithEvents LabelLinkID As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.ButtonOK = New System.Windows.Forms.Button
        Me.ButtonCancel = New System.Windows.Forms.Button
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.TextBoxLayerName = New System.Windows.Forms.TextBox
        Me.ComboBoxClasses = New System.Windows.Forms.ComboBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.LabelLayerName = New System.Windows.Forms.Label
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.TextBoxLinkID = New System.Windows.Forms.TextBox
        Me.LabelLinkID = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.TextBoxLocation = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.TextBoxDisplayText = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'ButtonOK
        '
        Me.ButtonOK.Location = New System.Drawing.Point(128, 192)
        Me.ButtonOK.Name = "ButtonOK"
        Me.ButtonOK.TabIndex = 9
        Me.ButtonOK.Text = "OK"
        '
        'ButtonCancel
        '
        Me.ButtonCancel.Location = New System.Drawing.Point(208, 192)
        Me.ButtonCancel.Name = "ButtonCancel"
        Me.ButtonCancel.TabIndex = 10
        Me.ButtonCancel.Text = "Annuleren"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.TextBoxLayerName)
        Me.GroupBox1.Controls.Add(Me.ComboBoxClasses)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.LabelLayerName)
        Me.GroupBox1.Location = New System.Drawing.Point(8, 8)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(280, 72)
        Me.GroupBox1.TabIndex = 11
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Tekstlabel toevoegen in"
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(8, 24)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(8, 16)
        Me.Label2.TabIndex = 14
        Me.Label2.Text = ">"
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(20, 48)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(100, 16)
        Me.Label4.TabIndex = 13
        Me.Label4.Text = "Annotatie klasse :"
        '
        'TextBoxLayerName
        '
        Me.TextBoxLayerName.BackColor = System.Drawing.Color.White
        Me.TextBoxLayerName.Enabled = False
        Me.TextBoxLayerName.Location = New System.Drawing.Point(120, 16)
        Me.TextBoxLayerName.Name = "TextBoxLayerName"
        Me.TextBoxLayerName.Size = New System.Drawing.Size(152, 20)
        Me.TextBoxLayerName.TabIndex = 12
        Me.TextBoxLayerName.Text = ""
        '
        'ComboBoxClasses
        '
        Me.ComboBoxClasses.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBoxClasses.Location = New System.Drawing.Point(120, 40)
        Me.ComboBoxClasses.Name = "ComboBoxClasses"
        Me.ComboBoxClasses.Size = New System.Drawing.Size(152, 21)
        Me.ComboBoxClasses.TabIndex = 11
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(8, 48)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(8, 16)
        Me.Label3.TabIndex = 10
        Me.Label3.Text = ">"
        '
        'LabelLayerName
        '
        Me.LabelLayerName.Location = New System.Drawing.Point(20, 24)
        Me.LabelLayerName.Name = "LabelLayerName"
        Me.LabelLayerName.Size = New System.Drawing.Size(96, 16)
        Me.LabelLayerName.TabIndex = 9
        Me.LabelLayerName.Text = "Annotatie laag :"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.Label8)
        Me.GroupBox2.Controls.Add(Me.TextBoxLinkID)
        Me.GroupBox2.Controls.Add(Me.LabelLinkID)
        Me.GroupBox2.Controls.Add(Me.Label6)
        Me.GroupBox2.Controls.Add(Me.TextBoxLocation)
        Me.GroupBox2.Controls.Add(Me.Label7)
        Me.GroupBox2.Controls.Add(Me.Label1)
        Me.GroupBox2.Controls.Add(Me.TextBoxDisplayText)
        Me.GroupBox2.Controls.Add(Me.Label5)
        Me.GroupBox2.Location = New System.Drawing.Point(8, 88)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(280, 96)
        Me.GroupBox2.TabIndex = 12
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Gegevens"
        '
        'Label8
        '
        Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.Location = New System.Drawing.Point(8, 72)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(8, 16)
        Me.Label8.TabIndex = 23
        Me.Label8.Text = ">"
        '
        'TextBoxLinkID
        '
        Me.TextBoxLinkID.BackColor = System.Drawing.Color.White
        Me.TextBoxLinkID.Enabled = False
        Me.TextBoxLinkID.Location = New System.Drawing.Point(120, 64)
        Me.TextBoxLinkID.Name = "TextBoxLinkID"
        Me.TextBoxLinkID.Size = New System.Drawing.Size(152, 20)
        Me.TextBoxLinkID.TabIndex = 22
        Me.TextBoxLinkID.Text = ""
        '
        'LabelLinkID
        '
        Me.LabelLinkID.Location = New System.Drawing.Point(24, 72)
        Me.LabelLinkID.Name = "LabelLinkID"
        Me.LabelLinkID.Size = New System.Drawing.Size(96, 16)
        Me.LabelLinkID.TabIndex = 21
        Me.LabelLinkID.Text = "Link info :"
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(8, 48)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(8, 16)
        Me.Label6.TabIndex = 20
        Me.Label6.Text = ">"
        '
        'TextBoxLocation
        '
        Me.TextBoxLocation.BackColor = System.Drawing.Color.White
        Me.TextBoxLocation.Enabled = False
        Me.TextBoxLocation.Location = New System.Drawing.Point(120, 40)
        Me.TextBoxLocation.Name = "TextBoxLocation"
        Me.TextBoxLocation.Size = New System.Drawing.Size(152, 20)
        Me.TextBoxLocation.TabIndex = 19
        Me.TextBoxLocation.Text = ""
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(24, 48)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(96, 16)
        Me.Label7.TabIndex = 18
        Me.Label7.Text = "Plaatsing :"
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(8, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(8, 16)
        Me.Label1.TabIndex = 17
        Me.Label1.Text = ">"
        '
        'TextBoxDisplayText
        '
        Me.TextBoxDisplayText.BackColor = System.Drawing.Color.White
        Me.TextBoxDisplayText.Enabled = False
        Me.TextBoxDisplayText.Location = New System.Drawing.Point(120, 16)
        Me.TextBoxDisplayText.Name = "TextBoxDisplayText"
        Me.TextBoxDisplayText.Size = New System.Drawing.Size(152, 20)
        Me.TextBoxDisplayText.TabIndex = 16
        Me.TextBoxDisplayText.Text = ""
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(24, 24)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(96, 16)
        Me.Label5.TabIndex = 15
        Me.Label5.Text = "Toon tekst :"
        '
        'FormAddAnnotation
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(292, 222)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.ButtonCancel)
        Me.Controls.Add(Me.ButtonOK)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "FormAddAnnotation"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Label toevoegen"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region " Overloaded constructor "
    <CLSCompliant(False)> _
    Public Sub New( _
        ByVal annoLayer As IAnnotationLayer, _
        ByVal pointGeom As IPoint, _
        ByVal displayText As String, _
        ByVal linkFieldName As String, _
        ByVal linkFieldValue As String _
        )

        MyBase.New()
        m_annoLayer = annoLayer
        m_pointGeom = pointGeom
        m_displayText = displayText
        m_linkFieldName = linkFieldName
        m_linkFieldValue = linkFieldValue

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        InitializeForm()

    End Sub
#End Region

#Region " Form controls events "

    Private Sub ButtonCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonCancel.Click
        Me.Close()
    End Sub

    Private Sub ButtonOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonOK.Click
        Try
            'Analyse and store GUI settings.
            m_annoClassName = CStr(Me.ComboBoxClasses.Items(Me.ComboBoxClasses.SelectedIndex))
            'm_annoClassID = GetAnnoClassID(m_annoLayer, m_annoClassName)
            'm_symbolID = m_annoClassID

            'Create the new annotation.
            AddAnno( _
                annoLayer:=m_annoLayer, _
                annoClassName:=m_annoClassName, _
                symbolName:=m_annoClassName, _
                textString:=m_displayText, _
                pointGeom:=m_pointGeom, _
                linkField:=m_linkFieldName, _
                linkValue:=m_linkFieldValue)

            'Close this modal form.
            Me.Close()

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

#End Region

#Region " Utility procedures "

    Private Sub InitializeForm()
        Try
            'Show constructor info in the form.
            Me.TextBoxLayerName.Text = CType(m_annoLayer, IFeatureLayer).Name
            Me.TextBoxDisplayText.Text = m_displayText
            Me.TextBoxLinkID.Text = m_linkFieldName & " = " & m_linkFieldValue
            Me.TextBoxLocation.Text = m_pointGeom.X & " / " & m_pointGeom.Y

            'Show a list of classes in the form.
            Dim pGroupLayer As ICompositeLayer = CType(m_annoLayer, ICompositeLayer)
            Me.ComboBoxClasses.Items.Clear()
            For i As Integer = 0 To pGroupLayer.Count - 1
                Me.ComboBoxClasses.Items.Add(pGroupLayer.Layer(i).Name)
            Next i
            Me.ComboBoxClasses.SelectedIndex = 0

        Catch ex As Exception
            Throw ex
        End Try

    End Sub

#End Region

End Class
