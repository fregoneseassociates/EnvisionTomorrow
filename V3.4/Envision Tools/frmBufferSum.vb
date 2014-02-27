'MICROSOFT REFERENCES
Imports System
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Runtime.InteropServices
Imports System.Collections
Imports System.IO
Imports System.Drawing.Printing
Imports System.Data
Imports System.Math

'ESRI REFERENCES
Imports ESRI.ArcGIS.ADF
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.CartoUI
Imports ESRI.ArcGIS.Catalog
Imports ESRI.ArcGIS.CatalogUI
Imports ESRI.ArcGIS.DataSourcesGDB
Imports ESRI.ArcGIS.DataSourcesFile
Imports ESRI.ArcGIS.Display
Imports ESRI.ArcGIS.Editor
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.GeoDatabaseUI
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.Geoprocessor
Imports ESRI.ArcGIS.SystemUI

Public Class frmBufferSum
    Dim blnOpening As Boolean = True
    Public m_arrPolyFeatureLayers As New ArrayList

    Private Sub frmBufferSum_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'CLOSE THE 7Ds FORM
        m_arrPolyFeatureLayers = Nothing
        m_frmAccessBufferSummary = Nothing
    End Sub

    Private Sub btnNext1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNext1.Click
        Me.tabSteps.SelectTab(1)
    End Sub

    Private Sub btnPrevious1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrevious1.Click
        Me.tabSteps.SelectTab(0)
    End Sub

    Private Sub btnNext2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNext2.Click
        Me.tabSteps.SelectTab(2)
    End Sub

    Private Sub btnPrevious2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrevious2.Click
        Me.tabSteps.SelectTab(1)
    End Sub

    Private Sub btnPrevious3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrevious3.Click
        Me.tabSteps.SelectTab(1)
        Me.lblRunStatus.Visible = False
        Me.barStatusRun.Value = False
    End Sub

    Private Sub frmBufferSum_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'LOAD THE SELECTED ENVISION LAYER NUMERICAL FIELDS INTO THE STEP ONE TAB CONTROL
        Dim mxApplication As IMxApplication = Nothing
        Dim pMxDocument As IMxDocument = Nothing
        Dim mapCurrent As Map
        Dim pLyr As ILayer
        Dim pFeatLyr As IFeatureLayer
        Dim intLayer As Integer
        Dim intFeatCount As Integer = 0
        Dim pActiveView As IActiveView = Nothing
        Dim pFeatClass As IFeatureClass
        Dim pTable As ITable
        Dim strFields As String
        Dim intFld As Integer = 0
        Dim intFldCount As Integer = 20
        Dim pField As IField = Nothing
        Dim pFeatSelection As IFeatureSelection
        Dim pCursor As ICursor = Nothing

        blnOpening = True
        '********************************************************************
        'Populate the combo box with layer information
        '********************************************************************
        Try
            If Not TypeOf m_appEnvision Is IApplication Then
                GoTo CleanUp
            End If

            pMxDocument = DirectCast(m_appEnvision.Document, IMxDocument)
            mapCurrent = CType(pMxDocument.FocusMap, Map)
            pActiveView = CType(pMxDocument.FocusMap, IActiveView)
            If mapCurrent.LayerCount = 0 Then
                m_frmAccessBufferSummary.Close()
                GoTo CleanUp
            End If

            'BUILD LIST OF AVAILABLE FEATURE CLASSES
            m_arrPolyFeatureLayers = New ArrayList
            m_arrPolyFeatureLayers.Clear()
            m_frmAccessBufferSummary.cmbLayers.Items.Clear()
            If mapCurrent.LayerCount > 0 Then
                For intLayer = 0 To mapCurrent.LayerCount - 1
                    pLyr = CType(mapCurrent.Layer(intLayer), ILayer)
                    If TypeOf mapCurrent.Layer(intLayer) Is IFeatureLayer Then
                        pFeatLyr = CType(mapCurrent.Layer(intLayer), IFeatureLayer)
                        If pFeatLyr.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                            If Not m_pEditFeatureLyr.Name = pFeatLyr.Name Then
                                m_frmAccessBufferSummary.cmbLayers.Items.Add(pFeatLyr.Name)
                                m_arrPolyFeatureLayers.Add(pFeatLyr)
                                intFeatCount = intFeatCount + 1
                            End If
                        End If
                        pFeatLyr = Nothing
                    End If
                Next
                If m_arrPolyFeatureLayers.Count <= 0 Then
                    m_frmAccessBufferSummary.rdbNeighborhoodLayer.Enabled = False
                    m_frmAccessBufferSummary.cmbLayers.Enabled = False
                Else
                    m_frmAccessBufferSummary.rdbNeighborhoodLayer.Enabled = True
                    'SELECT NEIGHBORHOOD LAYER IF PREVIOUSLY SELECTED
                    If m_strNeighborhoodLayerName.Length > 0 Then
                        If m_frmAccessBufferSummary.cmbLayers.FindString(m_strNeighborhoodLayerName) >= 0 Then
                            m_frmAccessBufferSummary.cmbLayers.Text = m_strNeighborhoodLayerName
                        End If
                    End If
                End If
            Else
                m_frmAccessBufferSummary.cmbLayers.Enabled = False
                m_frmAccessBufferSummary.rdbNeighborhoodLayer.Enabled = False
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Envision Buffer Summary Form Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo CleanUp
        End Try

        'EXIT IF NO ENVISION LAYHER HAS BEEN SET
        If m_pEditFeatureLyr Is Nothing Then
            m_strProcessing = m_strProcessing & "No Envision layer has been defined: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
            GoTo CleanUp
        Else
            pFeatClass = m_pEditFeatureLyr.FeatureClass
            pTable = CType(pFeatClass, ITable)
        End If

        'CLEAR OUT LIST
        Me.chkNumericalFlds.Items.Clear()

        'CYCLE THROUGH FIELDS TO FIND NUMERICAL FIELD(S)
        For intFld = 0 To pFeatClass.Fields.FieldCount - 1
            pField = pFeatClass.Fields.Field(intFld)
            If pField.Type = esriFieldType.esriFieldTypeInteger Or pField.Type = esriFieldType.esriFieldTypeDouble Or pField.Type = esriFieldType.esriFieldTypeSmallInteger Then
                If Not pField.Name = "OBJECTID" Then
                    Me.chkNumericalFlds.Items.Add(pField.Name)
                End If
            End If
        Next

        'ENABLE THE SELECTION SET PROCESSING OPTION IF A SELECTION IS AVIALABLE
        pFeatSelection = CType(m_pEditFeatureLyr, IFeatureSelection)
        pFeatSelection.SelectionSet.Search(Nothing, False, pCursor)
        If pFeatSelection.SelectionSet.Count > 0 Then
            Me.rdbSelected.Enabled = True
            Me.rdbSelected.Text = "Selected Features (" & pFeatSelection.SelectionSet.Count.ToString & ")"
        End If

        'ADD THE BUFFER TYPES 
        Me.dgvBuffers.Rows.Clear()
        Me.dgvBuffers.Rows.Add()
        Me.dgvBuffers.Rows(0).Cells(1).Value = "1/4 Mile"
        Me.dgvBuffers.Rows(0).Cells(2).Value = "0.25"
        Me.dgvBuffers.Rows(0).Cells(3).Value = "_QRT_MI"
        Me.dgvBuffers.Rows.Add()
        Me.dgvBuffers.Rows(1).Cells(1).Value = "1/2 Mile"
        Me.dgvBuffers.Rows(1).Cells(2).Value = "0.50"
        Me.dgvBuffers.Rows(1).Cells(3).Value = "_HALF_MI"
        Me.dgvBuffers.Rows.Add()
        Me.dgvBuffers.Rows(2).Cells(1).Value = "1 Mile"
        Me.dgvBuffers.Rows(2).Cells(2).Value = "1.00"
        Me.dgvBuffers.Rows(2).Cells(3).Value = "_1MI"
        Me.dgvBuffers.Rows.Add()
        Me.dgvBuffers.Rows(3).Cells(1).Value = "User Defined 1"
        Me.dgvBuffers.Rows(3).Cells(2).Value = "0.00"
        Me.dgvBuffers.Rows(3).Cells(3).Value = "_UD1_MI"
        Me.dgvBuffers.Rows.Add()
        Me.dgvBuffers.Rows(4).Cells(1).Value = "User Defined 2"
        Me.dgvBuffers.Rows(4).Cells(2).Value = "0.00"
        Me.dgvBuffers.Rows(4).Cells(3).Value = "_UD2_MI"
        Me.dgvBuffers.Rows.Add()
        Me.dgvBuffers.Rows(5).Cells(1).Value = "User Defined 3"
        Me.dgvBuffers.Rows(5).Cells(2).Value = "0.00"
        Me.dgvBuffers.Rows(5).Cells(3).Value = "_UD3_MI"

CleanUp:
        blnOpening = False
        m_frmAccessBufferSummary.lblRunStatus.Text = ""
        m_frmAccessBufferSummary.barStatusRun.Value = 0
        mxApplication = Nothing
        pMxDocument = Nothing
        mapCurrent = Nothing
        pLyr = Nothing
        pFeatLyr = Nothing
        intLayer = Nothing
        intFeatCount = Nothing
        pActiveView = Nothing
        pFeatClass = Nothing
        pTable = Nothing
        strFields = Nothing
        intFld = Nothing
        intFldCount = Nothing
        pField = Nothing
        pFeatSelection = Nothing
        pCursor = Nothing
        GC.WaitForFullGCComplete()
        GC.Collect()
    End Sub

    Private Sub chkNumericalFlds_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkNumericalFlds.SelectedValueChanged, chkNumericalFlds.DoubleClick
        'CROSSWALK THE SELECTED NUMERICAL ATTRIBUTE FIELDS WITH THE SELECTED BUFFERS TO CREATE THE FIELD/BUFFER MATRIX
        Dim intFld As Integer
        Dim strFieldName As String
        Dim intMatrixRow As Integer
        Dim blnFound As Boolean

        If blnOpening Then
            GoTo CleanUp
        End If

        'CYCLE THROUGH FIELD LIST TO SEE WHICH ARE SELECTED 
        'ADD ROW TO MATRIX FOR NEW CHECKED FIELD OR REMOVE UNCHECKED FIELDS
        For intFld = 0 To Me.chkNumericalFlds.Items.Count - 1
            strFieldName = Me.chkNumericalFlds.Items.Item(intFld)
            If Me.chkNumericalFlds.CheckedItems.Contains(strFieldName) Then
                blnFound = False
                For intMatrixRow = 0 To Me.dgvMatrix.Rows.Count - 1
                    If strFieldName = Me.dgvMatrix.Rows(intMatrixRow).Cells(0).Value Then
                        blnFound = True
                        Exit For
                    End If
                Next
                If Not blnFound Then
                    Me.dgvMatrix.Rows.Add()
                    Me.dgvMatrix.Rows(Me.dgvMatrix.Rows.Count - 1).Cells(0).Value = strFieldName
                    '1/4 Mile Buffer
                    Me.dgvMatrix.Rows(Me.dgvMatrix.Rows.Count - 1).Cells(1).Value = Me.dgvBuffers.Rows(0).Cells(0).Value
                    Me.dgvMatrix.Rows(Me.dgvMatrix.Rows.Count - 1).Cells(2).Value = strFieldName & Me.dgvBuffers.Rows(0).Cells(3).Value
                    '1/2 Mile Buffer
                    Me.dgvMatrix.Rows(Me.dgvMatrix.Rows.Count - 1).Cells(3).Value = Me.dgvBuffers.Rows(1).Cells(0).Value
                    Me.dgvMatrix.Rows(Me.dgvMatrix.Rows.Count - 1).Cells(4).Value = strFieldName & Me.dgvBuffers.Rows(1).Cells(3).Value
                    '1 Mile Buffer
                    Me.dgvMatrix.Rows(Me.dgvMatrix.Rows.Count - 1).Cells(5).Value = Me.dgvBuffers.Rows(2).Cells(0).Value
                    Me.dgvMatrix.Rows(Me.dgvMatrix.Rows.Count - 1).Cells(6).Value = strFieldName & Me.dgvBuffers.Rows(2).Cells(3).Value
                    'User Defined 1 Buffer
                    Me.dgvMatrix.Rows(Me.dgvMatrix.Rows.Count - 1).Cells(7).Value = Me.dgvBuffers.Rows(3).Cells(0).Value
                    Me.dgvMatrix.Rows(Me.dgvMatrix.Rows.Count - 1).Cells(8).Value = strFieldName & Me.dgvBuffers.Rows(3).Cells(3).Value
                    'User Defined 2 Buffer
                    Me.dgvMatrix.Rows(Me.dgvMatrix.Rows.Count - 1).Cells(9).Value = Me.dgvBuffers.Rows(4).Cells(0).Value
                    Me.dgvMatrix.Rows(Me.dgvMatrix.Rows.Count - 1).Cells(10).Value = strFieldName & Me.dgvBuffers.Rows(4).Cells(3).Value
                    'User Defined 3 Buffer
                    Me.dgvMatrix.Rows(Me.dgvMatrix.Rows.Count - 1).Cells(11).Value = Me.dgvBuffers.Rows(5).Cells(0).Value
                    Me.dgvMatrix.Rows(Me.dgvMatrix.Rows.Count - 1).Cells(12).Value = strFieldName & Me.dgvBuffers.Rows(5).Cells(3).Value
                End If
            Else
                'REMOVE IF FOUND IN MATRIX LIST
                For intMatrixRow = 0 To Me.dgvMatrix.Rows.Count - 1
                    If strFieldName = Me.dgvMatrix.Rows(intMatrixRow).Cells(0).Value Then
                        Me.dgvMatrix.Rows.RemoveAt(intMatrixRow)
                        Exit For
                    End If
                Next
            End If
        Next

CleanUp:
        intFld = Nothing
        strFieldName = Nothing
        intMatrixRow = Nothing
        blnFound = Nothing
        GC.WaitForFullGCComplete()
        GC.Collect()
    End Sub

    Private Sub dgvBuffers_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvBuffers.CellContentClick, dgvBuffers.CellValueChanged
        'APPLY THE USER BUFFER SELECTIONS TO THE FIELD/BUFFER MATRIX
        Dim intFld As Integer
        Dim strFieldName As String
        Dim intMatrixRow As Integer
        Dim blnFound As Boolean

        If blnOpening Then
            GoTo CleanUp
        End If

        'CONTROL THE COLUMN VISUAL STATUS
        '1/4 Mile Buffer
        Me.dgvMatrix.Columns.Item(1).Visible = Me.dgvBuffers.Rows(0).Cells(0).Value
        Me.dgvMatrix.Columns.Item(2).Visible = Me.dgvBuffers.Rows(0).Cells(0).Value
        '1/2 Mile Buffer
        Me.dgvMatrix.Columns.Item(3).Visible = Me.dgvBuffers.Rows(1).Cells(0).Value
        Me.dgvMatrix.Columns.Item(4).Visible = Me.dgvBuffers.Rows(1).Cells(0).Value
        '1 Mile Buffer
        Me.dgvMatrix.Columns.Item(5).Visible = Me.dgvBuffers.Rows(2).Cells(0).Value
        Me.dgvMatrix.Columns.Item(6).Visible = Me.dgvBuffers.Rows(2).Cells(0).Value
        'User Defined 1 Buffer
        Me.dgvMatrix.Columns.Item(7).Visible = Me.dgvBuffers.Rows(3).Cells(0).Value
        Me.dgvMatrix.Columns.Item(8).Visible = Me.dgvBuffers.Rows(3).Cells(0).Value
        'User Defined 2 Buffer
        Me.dgvMatrix.Columns.Item(9).Visible = Me.dgvBuffers.Rows(4).Cells(0).Value
        Me.dgvMatrix.Columns.Item(10).Visible = Me.dgvBuffers.Rows(4).Cells(0).Value
        'User Defined 3 Buffer
        Me.dgvMatrix.Columns.Item(11).Visible = Me.dgvBuffers.Rows(5).Cells(0).Value
        Me.dgvMatrix.Columns.Item(12).Visible = Me.dgvBuffers.Rows(5).Cells(0).Value

        'CYCLE THROUGH MATRIX AND UPDATE THE DEFAULT OUTPUT FIELD NAMES
        For intMatrixRow = 0 To Me.dgvMatrix.Rows.Count - 1
            strFieldName = Me.dgvMatrix.Rows(intMatrixRow).Cells(0).Value
            '1/4 Mile Buffer
            Me.dgvMatrix.Rows(intMatrixRow).Cells(1).Value = Me.dgvBuffers.Rows(0).Cells(0).Value
            Me.dgvMatrix.Rows(intMatrixRow).Cells(2).Value = strFieldName & Me.dgvBuffers.Rows(0).Cells(3).Value
            '1/2 Mile Buffer
            Me.dgvMatrix.Rows(intMatrixRow).Cells(3).Value = Me.dgvBuffers.Rows(1).Cells(0).Value
            Me.dgvMatrix.Rows(intMatrixRow).Cells(4).Value = strFieldName & Me.dgvBuffers.Rows(1).Cells(3).Value
            '1 Mile Buffer
            Me.dgvMatrix.Rows(intMatrixRow).Cells(5).Value = Me.dgvBuffers.Rows(2).Cells(0).Value
            Me.dgvMatrix.Rows(intMatrixRow).Cells(6).Value = strFieldName & Me.dgvBuffers.Rows(2).Cells(3).Value
            'User Defined 1 Buffer
            Me.dgvMatrix.Rows(intMatrixRow).Cells(7).Value = Me.dgvBuffers.Rows(3).Cells(0).Value
            Me.dgvMatrix.Rows(intMatrixRow).Cells(8).Value = strFieldName & Me.dgvBuffers.Rows(3).Cells(3).Value
            'User Defined 2 Buffer
            Me.dgvMatrix.Rows(intMatrixRow).Cells(9).Value = Me.dgvBuffers.Rows(4).Cells(0).Value
            Me.dgvMatrix.Rows(intMatrixRow).Cells(10).Value = strFieldName & Me.dgvBuffers.Rows(4).Cells(3).Value
            'User Defined 3 Buffer
            Me.dgvMatrix.Rows(intMatrixRow).Cells(11).Value = Me.dgvBuffers.Rows(5).Cells(0).Value
            Me.dgvMatrix.Rows(intMatrixRow).Cells(12).Value = strFieldName & Me.dgvBuffers.Rows(5).Cells(3).Value
        Next

        'MAKE SURE THE 1/4, 1/2 and 1 MILE BUFFER VALUES REMAIN CONSTANT
        Me.dgvBuffers.Rows(0).Cells(2).Value = 0.25
        Me.dgvBuffers.Rows(1).Cells(2).Value = 0.5
        Me.dgvBuffers.Rows(2).Cells(2).Value = 1.0
CleanUp:
        intFld = Nothing
        strFieldName = Nothing
        intMatrixRow = Nothing
        blnFound = Nothing
        GC.WaitForFullGCComplete()
        GC.Collect()
    End Sub

    Private Sub btnRun_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRun.Click
        Dim pFeatClass As IFeatureClass
        Dim pTable As ITable
        Dim strFieldName As String
        Dim intMatrixRow As Integer
        Dim blnFound As Boolean = True
        Dim tblScLookUpTbl As ITable = Nothing
        Dim strAttributeFld As String = ""
        Dim strSummaryFld As String = ""
        Dim dblMiles As Double = 0
        Dim pField As IField
        Dim blnInteger As Boolean
        Dim pFeatLyr As IFeatureLayer

        'STOP EDIT TO DATA GRID VIEWS
        Me.dgvBuffers.EndEdit()
        Me.dgvMatrix.EndEdit()
        Me.dgvBuffers.RefreshEdit()
        Me.dgvMatrix.RefreshEdit()


        'EXECUTE THE ENVISION ACCESS BUFFER SUMMARIES 
        m_strProcessing = "--------------------------------------------------------------------------" & vbNewLine
        m_strProcessing = m_strProcessing & "Starting ACCESS | BUFFER SUMMARIES Processing: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
        m_strProcessing = m_strProcessing & "--------------------------------------------------------------------------" & vbNewLine
        m_frmAccessBufferSummary.lblRunStatus.Text = ""
        m_frmAccessBufferSummary.barStatusRun.Value = 0
        Me.lblRunStatus.Visible = True
        Me.barStatusRun.Visible = True
        Me.lblRunStatus.Text = ""
        Me.barStatusRun.Value = 0

        'CHECK IF THERE IS A FEATURE SET OPTIONS SELECTED
        If Not Me.rdbFull.Checked And Not Me.rdbSelected.Checked And Not Me.rdbPartial.Checked Then
            MessageBox.Show("Please selected a Processing Feature Selection Option in Step 1. ", "Processing Feature Selection Option Not Selected", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            Me.tabSteps.SelectTab(0)
            GoTo CleanUp
        End If

        'EXIT IF NO ENVISION LAYHER HAS BEEN SET
        If m_frmAccessBufferSummary.rdbNeighborhoodLayer.Checked Then
            Try
                pFeatLyr = m_arrPolyFeatureLayers.Item(m_frmAccessBufferSummary.cmbLayers.FindString(m_frmAccessBufferSummary.cmbLayers.Text))
                pFeatClass = pFeatLyr.FeatureClass
                pTable = CType(pFeatClass, ITable)
            Catch ex As Exception
                m_strProcessing = m_strProcessing & "Error retrieving Neighborhood layer: " & vbNewLine & ex.Message & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
                GoTo CleanUp
            End Try
        Else
            If m_pEditFeatureLyr Is Nothing Then
                m_strProcessing = m_strProcessing & "No Envision layer has been defined: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
                GoTo CleanUp
            Else
                pFeatClass = m_pEditFeatureLyr.FeatureClass
                pTable = CType(pFeatClass, ITable)
            End If
        End If

        If Me.dgvMatrix.RowCount <= 0 Then
            MessageBox.Show("No numerical attribute fields have been selected in the 'Step 1 - Numerical Field Selection' tab", "Numerical Field Selection Required", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo CleanUp
        End If

        'END ANY EDIT SESSIONS UNDER WAY
        If Not m_dockEnvisionWinForm.btnEditing.Text = "Start Edit" Then
            EditSession()
        End If

        '*********************************************************
        'REVIEW THE MATRIX TO INSURE ALL PARAMETERS ARE SET
        'CYCLE THROUGH MATRIX AND UPDATE THE DEFAULT OUTPUT FIELD NAMES, ADD IF MISSING
        '*********************************************************
        Me.lblRunStatus.Text = "Checking for required Output Summary field(s)"
        For intMatrixRow = 0 To Me.dgvMatrix.Rows.Count - 1
            '1/4 Mile Buffer
            If Me.dgvMatrix.Rows(intMatrixRow).Cells(1).Value Then
                strFieldName = Me.dgvMatrix.Rows(intMatrixRow).Cells(2).Value
                If strFieldName Is Nothing Then
                    m_strProcessing = m_strProcessing & "Output field name is missing for the selected summary field, " & strFieldName & ", for the 1/4 Mile buffer." & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
                    blnFound = False
                Else
                    If pTable.FindField(strFieldName) <= -1 Then
                        m_strProcessing = m_strProcessing & "Required output field not found, " & strFieldName & ".  Adding field to Envision layer." & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
                        Me.lblRunStatus.Text = "Adding missing field: " & strFieldName
                        AddEnvisionField(pTable, strFieldName, "DOUBLE", 16, 6)
                    End If
                End If
            End If
            '1/2 Mile Buffer
            If Me.dgvMatrix.Rows(intMatrixRow).Cells(3).Value Then
                strFieldName = Me.dgvMatrix.Rows(intMatrixRow).Cells(4).Value
                If strFieldName Is Nothing Then
                    m_strProcessing = m_strProcessing & "Output field name is missing for the selected summary field, " & strFieldName & ", for the 1/2 Mile buffer." & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
                    blnFound = False
                Else
                    If pTable.FindField(strFieldName) <= -1 Then
                        m_strProcessing = m_strProcessing & "Required output field not found, " & strFieldName & ".  Adding field to Envision layer." & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
                        Me.lblRunStatus.Text = "Adding missing field: " & strFieldName
                        AddEnvisionField(pTable, strFieldName, "DOUBLE", 16, 6)
                    End If
                End If
            End If
            '1 Mile Buffer
            If Me.dgvMatrix.Rows(intMatrixRow).Cells(5).Value Then
                strFieldName = Me.dgvMatrix.Rows(intMatrixRow).Cells(6).Value
                If strFieldName Is Nothing Then
                    m_strProcessing = m_strProcessing & "Output field name is missing for the selected summary field, " & strFieldName & ", for the 1 Mile buffer." & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
                    blnFound = False
                Else
                    If pTable.FindField(strFieldName) <= -1 Then
                        m_strProcessing = m_strProcessing & "Required output field not found, " & strFieldName & ".  Adding field to Envision layer." & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
                        Me.lblRunStatus.Text = "Adding missing field: " & strFieldName
                        AddEnvisionField(pTable, strFieldName, "DOUBLE", 16, 6)
                    End If
                End If
            End If
            'User Defined 1 Buffer
            If Me.dgvMatrix.Rows(intMatrixRow).Cells(7).Value Then
                strFieldName = Me.dgvMatrix.Rows(intMatrixRow).Cells(8).Value
                If strFieldName Is Nothing Then
                    m_strProcessing = m_strProcessing & "Output field name is missing for the selected summary field, " & strFieldName & ", for the User Defined 1 buffer." & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
                    blnFound = False
                Else
                    If pTable.FindField(strFieldName) <= -1 Then
                        m_strProcessing = m_strProcessing & "Required output field not found, " & strFieldName & ".  Adding field to Envision layer." & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
                        Me.lblRunStatus.Text = "Adding missing field: " & strFieldName
                        AddEnvisionField(pTable, strFieldName, "DOUBLE", 16, 6)
                    End If
                End If
            End If
            'User Defined 2 Buffer
            If Me.dgvMatrix.Rows(intMatrixRow).Cells(9).Value Then
                strFieldName = Me.dgvMatrix.Rows(intMatrixRow).Cells(10).Value
                If strFieldName Is Nothing Then
                    m_strProcessing = m_strProcessing & "Output field name is missing for the selected summary field, " & strFieldName & ", for the User Defined 2 buffer." & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
                    blnFound = False
                Else
                    If pTable.FindField(strFieldName) <= -1 Then
                        m_strProcessing = m_strProcessing & "Required output field not found, " & strFieldName & ".  Adding field to Envision layer." & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
                        Me.lblRunStatus.Text = "Adding missing field: " & strFieldName
                        AddEnvisionField(pTable, strFieldName, "DOUBLE", 16, 6)
                    End If
                End If
            End If
            'User Defined 3 Buffer
            If Me.dgvMatrix.Rows(intMatrixRow).Cells(11).Value Then
                strFieldName = Me.dgvMatrix.Rows(intMatrixRow).Cells(12).Value
                If strFieldName Is Nothing Then
                    m_strProcessing = m_strProcessing & "Output field name is missing for the selected summary field, " & strFieldName & ", for the User Defined 2 buffer." & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
                    blnFound = False
                Else
                    If pTable.FindField(strFieldName) <= -1 Then
                        m_strProcessing = m_strProcessing & "Required output field not found, " & strFieldName & ".  Adding field to Envision layer." & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
                        Me.lblRunStatus.Text = "Adding missing field: " & strFieldName
                        AddEnvisionField(pTable, strFieldName, "DOUBLE", 16, 6)
                    End If
                End If
            End If
        Next
        If Not blnFound Then
            MessageBox.Show("Please review the Output Field names provided in the matrix as it appears one or more of the field names is missing.", "Undefined Output Field(s)", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo CleanUp
        End If

        '*********************************************************
        'NOW CYCLE THROUGH THE BUFFER LIST AND EXECUTE THE SUB TO CONDUCT THE ATTRIBUTE SUMMARY
        '*********************************************************
        For intMatrixRow = 0 To Me.dgvMatrix.Rows.Count - 1
            strAttributeFld = Me.dgvMatrix.Rows(intMatrixRow).Cells(0).Value
            '1/4 Mile Buffer
            If Me.dgvMatrix.Rows(intMatrixRow).Cells(1).Value Then
                strSummaryFld = Me.dgvMatrix.Rows(intMatrixRow).Cells(2).Value
                dblMiles = 0.25
                pField = pFeatClass.Fields.Field((pFeatClass.FindField(strSummaryFld)))
                If pField.Type = esriFieldType.esriFieldTypeInteger Then
                    blnInteger = True
                Else
                    blnInteger = False
                End If
                Me.lblRunStatus.Text = "Processing 1/4 Mile for field: " & strAttributeFld
                BufferSummary(strAttributeFld, strSummaryFld, dblMiles, blnInteger)
            End If
            '1/2 Mile Buffer
            If Me.dgvMatrix.Rows(intMatrixRow).Cells(3).Value Then
                strSummaryFld = Me.dgvMatrix.Rows(intMatrixRow).Cells(4).Value
                dblMiles = 0.5
                pField = pFeatClass.Fields.Field((pFeatClass.FindField(strSummaryFld)))
                If pField.Type = esriFieldType.esriFieldTypeInteger Then
                    blnInteger = True
                Else
                    blnInteger = False
                End If
                Me.lblRunStatus.Text = "Processing 1/2 Mile for field: " & strAttributeFld
                BufferSummary(strAttributeFld, strSummaryFld, dblMiles, blnInteger)
            End If
            '1 Mile Buffer
            If Me.dgvMatrix.Rows(intMatrixRow).Cells(5).Value Then
                strSummaryFld = Me.dgvMatrix.Rows(intMatrixRow).Cells(6).Value
                dblMiles = 1.0
                pField = pFeatClass.Fields.Field((pFeatClass.FindField(strSummaryFld)))
                If pField.Type = esriFieldType.esriFieldTypeInteger Then
                    blnInteger = True
                Else
                    blnInteger = False
                End If
                Me.lblRunStatus.Text = "Processing 1 Mile for field: " & strAttributeFld
                BufferSummary(strAttributeFld, strSummaryFld, dblMiles, blnInteger)
            End If
            'User Defined 1 Buffer
            If Me.dgvMatrix.Rows(intMatrixRow).Cells(7).Value Then
                strSummaryFld = Me.dgvMatrix.Rows(intMatrixRow).Cells(8).Value
                dblMiles = Me.dgvBuffers.Rows(3).Cells(2).Value
                pField = pFeatClass.Fields.Field((pFeatClass.FindField(strSummaryFld)))
                If pField.Type = esriFieldType.esriFieldTypeInteger Then
                    blnInteger = True
                Else
                    blnInteger = False
                End If
                Me.lblRunStatus.Text = "Processing User Defined 1 for field: " & strAttributeFld
                BufferSummary(strAttributeFld, strSummaryFld, dblMiles, blnInteger)
            End If
            'User Defined 2 Buffer
            If Me.dgvMatrix.Rows(intMatrixRow).Cells(9).Value Then
                strSummaryFld = Me.dgvMatrix.Rows(intMatrixRow).Cells(10).Value
                dblMiles = Me.dgvBuffers.Rows(4).Cells(2).Value
                pField = pFeatClass.Fields.Field((pFeatClass.FindField(strSummaryFld)))
                If pField.Type = esriFieldType.esriFieldTypeInteger Then
                    blnInteger = True
                Else
                    blnInteger = False
                End If
                Me.lblRunStatus.Text = "Processing User Defined 2 for field: " & strAttributeFld
                BufferSummary(strAttributeFld, strSummaryFld, dblMiles, blnInteger)
            End If
            'User Defined 3 Buffer
            If Me.dgvMatrix.Rows(intMatrixRow).Cells(11).Value Then
                strSummaryFld = Me.dgvMatrix.Rows(intMatrixRow).Cells(12).Value
                dblMiles = Me.dgvBuffers.Rows(5).Cells(2).Value
                pField = pFeatClass.Fields.Field((pFeatClass.FindField(strSummaryFld)))
                If pField.Type = esriFieldType.esriFieldTypeInteger Then
                    blnInteger = True
                Else
                    blnInteger = False
                End If
                Me.lblRunStatus.Text = "Processing User Defined 3 for field: " & strAttributeFld
                BufferSummary(strAttributeFld, strSummaryFld, dblMiles, blnInteger)
            End If
        Next

        'UNSELECT THE SELECTED FEATURE OPTION AND DISABLE AS ALL SELECTIONS WILL BE CLEARED
        Me.rdbSelected.Enabled = False
        Me.rdbSelected.Text = "Selected Features (0)"
        If Me.rdbSelected.Checked Then
            Me.rdbSelected.Checked = False
        End If

CleanUp:
        Me.lblRunStatus.Text = "All Processing has been completed."
        Me.barStatusRun.Value = 100
        pFeatClass = Nothing
        pTable = Nothing
        strFieldName = Nothing
        intMatrixRow = Nothing
        blnFound = Nothing
        tblScLookUpTbl = Nothing
        strAttributeFld = Nothing
        strSummaryFld = Nothing
        dblMiles = Nothing
        pField = Nothing
        blnInteger = Nothing
        GC.WaitForFullGCComplete()
        GC.Collect()
    End Sub

    Private Sub BufferSummary(ByVal strAttribFld As String, ByVal strSummaryFld As String, ByVal dblMiles As Double, ByVal blnInteger As Boolean)

        Dim pFeatClass As IFeatureClass
        Dim pTable As ITable
        Dim strFieldName As String
        Dim intMatrixRow As Integer
        Dim blnFound As Boolean = True
        Dim pDef As IFeatureLayerDefinition2
        Dim pQFilter As IQueryFilter
        Dim strDefExpression As String = ""
        Dim strQString As String = Nothing
        Dim pFeatureCursor As IFeatureCursor = Nothing
        Dim pFeatureCursorSel As IFeatureCursor
        Dim intTotalCount As Integer
        Dim pQBLayer As IQueryByLayer
        Dim pFSel As IFeatureSelection = Nothing
        Dim pFeat As IFeature
        Dim intObjId As Integer = 0
        Dim intObjFld As Integer
        Dim intDevTypeFld As Integer = 0
        Dim intAttribFld As Integer = 0
        Dim intSummaryFld As Integer = 0
        Dim strDevType As String = ""
        Dim dblCurrentValue As Double = 0
        Dim intCurrentValue As Double = 0
        Dim dblValue As Double = 0
        Dim intValue As Double = 0
        Dim dblSummary As Double = 0
        Dim dblSummaryTotal As Double = 0
        Dim intSummaryTotal As Double = 0
        Dim pCursor As ICursor = Nothing
        Dim pFeatSel As IFeature
        Dim intCount As Integer = 0
        Dim intFailed As Integer = 0
        Dim mxApplication As IMxApplication
        Dim pMxDoc As IMxDocument
        Dim pFeatLyr As IFeatureLayer = Nothing

        Try
            'CHECK AND RETRIEVE THE PROCESSING OPTION
            If Me.rdbNeighborhoodLayer.Checked Then
                pFeatLyr = m_arrPolyFeatureLayers.Item(Me.cmbLayers.FindString(Me.cmbLayers.Text))
                pFeatClass = pFeatLyr.FeatureClass
                pDef = pFeatLyr
                pFSel = pFeatLyr
            Else
                pFeatClass = m_pEditFeatureLyr.FeatureClass
                pDef = m_pEditFeatureLyr
                pFSel = m_pEditFeatureLyr
            End If
            strDefExpression = pDef.DefinitionExpression
            pQFilter = New QueryFilter
            If Me.rdbPartial.Checked Then
                strQString = "NOT DEV_TYPE = ''"
                If strDefExpression.Length > 0 Then
                    pQFilter.WhereClause = "(" & strQString & ") AND (" & pDef.DefinitionExpression & ")"
                    pFeatureCursor = pFeatClass.Search(pQFilter, False)
                Else
                    pQFilter.WhereClause = strQString
                    pFeatureCursor = pFeatClass.Search(pQFilter, False)
                End If
                intTotalCount = pFeatClass.FeatureCount(pQFilter)
            ElseIf Me.rdbFull.Checked Then
                If strDefExpression.Length > 0 Then
                    pQFilter.WhereClause = pDef.DefinitionExpression
                    pFeatureCursor = pFeatClass.Search(pQFilter, False)
                    intTotalCount = pFeatClass.FeatureCount(pQFilter)
                Else
                    pFeatureCursor = pFeatClass.Search(Nothing, False)
                    intTotalCount = pFeatClass.FeatureCount(Nothing)
                End If
            ElseIf Me.rdbSelected.Checked Then
                pFSel.SelectionSet.Search(Nothing, False, pCursor)
                pFeatureCursor = DirectCast(pCursor, IFeatureCursor)
                intTotalCount = pFSel.SelectionSet.Count
            End If

            pQBLayer = New QueryByLayer
            intObjFld = pFeatClass.FindField("OBJECTID")
            'intDevTypeFld = pFeatClass.FindField("DEV_TYPE")
            intAttribFld = pFeatClass.FindField(strAttribFld)
            intSummaryFld = pFeatClass.FindField(strSummaryFld)
            pFeat = pFeatureCursor.NextFeature
            Do While Not pFeat Is Nothing
                intCount = intCount + 1
                Me.barStatusRun.Value = (intCount / intTotalCount) * 100
                Me.Refresh()
                'strDevType = ""
                'Try
                '    strDevType = pFeat.Value(intDevTypeFld)
                'Catch ex As Exception
                'End Try
                dblCurrentValue = 0
                intCurrentValue = 0
                If blnInteger Then
                    Try
                        intCurrentValue = pFeat.Value(intAttribFld)
                    Catch ex As Exception
                    End Try
                Else
                    Try
                        dblCurrentValue = pFeat.Value(intAttribFld)
                    Catch ex As Exception
                    End Try
                End If

                Try
                    'If strDevType.Length > 0 Then
                    intObjId = pFeat.Value(intObjFld)
                    pQFilter = New QueryFilter
                    strQString = "OBJECTID = " & intObjId.ToString
                    pQFilter.WhereClause = strQString
                    pFSel.Clear()
                    pFSel.SelectFeatures(pQFilter, esriSelectionResultEnum.esriSelectionResultNew, True)
                    pQBLayer = New QueryByLayer
                    With pQBLayer
                        If Me.rdbNeighborhoodLayer.Checked Then
                            .ByLayer = pFeatLyr
                        Else
                            .ByLayer = m_pEditFeatureLyr
                        End If
                        .FromLayer = pFSel
                        .BufferDistance = dblMiles
                        .BufferUnits = esriUnits.esriMiles
                        .ResultType = esriSelectionResultEnum.esriSelectionResultNew
                        .LayerSelectionMethod = esriLayerSelectionMethod.esriLayerSelectHaveTheirCenterIn
                        .UseSelectedFeatures = True
                        pFSel.SelectionSet = .Select
                    End With

                    pFSel.SelectionSet.Search(Nothing, False, pCursor)
                    pFeatureCursorSel = DirectCast(pCursor, IFeatureCursor)
                    dblSummaryTotal = 0
                    intSummaryTotal = 0
                    If pFSel.SelectionSet.Count > 0 Then
                        pFeatSel = pFeatureCursorSel.NextFeature
                        Do While Not pFeatSel Is Nothing
                            If blnInteger Then
                                Try
                                    intValue = pFeatSel.Value(intAttribFld)
                                Catch ex As Exception
                                    intValue = 0
                                End Try
                                intSummaryTotal = intSummaryTotal + intValue
                            Else
                                Try
                                    dblValue = pFeatSel.Value(intAttribFld)
                                Catch ex As Exception
                                    dblValue = 0
                                End Try
                                dblSummaryTotal = dblSummaryTotal + dblValue
                            End If
                            pFeatSel = pFeatureCursorSel.NextFeature
                        Loop
                    End If

                    If blnInteger Then
                        pFeat.Value(intSummaryFld) = intSummaryTotal
                    Else
                        pFeat.Value(intSummaryFld) = dblSummaryTotal
                    End If
                    pFeat.Store()
                    'End If
                Catch ex As Exception
                    'MessageBox.Show("Error: " & ex.Message & vbNewLine & "Count: " & intCount.ToString & vbNewLine & "OBJECTID = " & intObjId.ToString, "Buffer Summary Process Error", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
                    intFailed = intFailed + 1
                End Try

                pFeat = pFeatureCursor.NextFeature
            Loop

            'CLEAR ANY SELECTION ON THE ENVISION EDIT LAYER AND REFRESH THE VIEW DOCUMENT
            pFSel.Clear()
            mxApplication = CType(m_appEnvision, IMxApplication)
            pMxDoc = CType(m_appEnvision.Document, IMxDocument)
            pMxDoc.ActivatedView.Refresh()

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message & vbNewLine & "Count: " & intCount.ToString & vbNewLine & "OBJECTID = " & intObjId.ToString, "Buffer Summary Process Error", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo CleanUp
        End Try
CleanUp:
        If intFailed > 0 Then
            MessageBox.Show(intFailed.ToString & " feature(s) failed to processing for the buffering.", "Incomplete Processing", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
        End If
        pFeatClass = Nothing
        pTable = Nothing
        strFieldName = Nothing
        intMatrixRow = Nothing
        blnFound = Nothing
        pDef = Nothing
        pQFilter = Nothing
        strDefExpression = Nothing
        strQString = Nothing
        pFeatureCursor = Nothing
        pFeatureCursorSel = Nothing
        intTotalCount = Nothing
        pQBLayer = Nothing
        pFSel = Nothing
        pFeat = Nothing
        intObjId = Nothing
        intObjFld = Nothing
        intDevTypeFld = Nothing
        intAttribFld = Nothing
        intSummaryFld = Nothing
        strDevType = Nothing
        dblCurrentValue = Nothing
        intCurrentValue = Nothing
        dblValue = Nothing
        intValue = Nothing
        dblSummary = Nothing
        dblSummaryTotal = Nothing
        intSummaryTotal = Nothing
        pCursor = Nothing
        pFeatSel = Nothing
        intCount = Nothing
        intFailed = Nothing
        mxApplication = Nothing
        pMxDoc = Nothing
        pFeatLyr = Nothing
        GC.WaitForFullGCComplete()
        GC.Collect()
    End Sub

    Private Sub rdbNeighborhoodLayer_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdbNeighborhoodLayer.CheckedChanged
        NeighborhoodSelect()
    End Sub

    Private Sub cmbLayers_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbLayers.SelectedIndexChanged, rdbParcelLayer.CheckedChanged
        NeighborhoodSelect()
    End Sub

    Private Sub NeighborhoodSelect()
        Dim pFeatSelection As IFeatureSelection
        Dim pCursor As ICursor = Nothing
        Dim pFeatLyr As IFeatureLayer
        Try
            'PARCEL SELECTED
            If Me.rdbParcelLayer.Checked Then
                Me.cmbLayers.Enabled = False
                Me.rdbPartial.Enabled = True
                pFeatSelection = CType(m_pEditFeatureLyr, IFeatureSelection)
                pFeatSelection.SelectionSet.Search(Nothing, False, pCursor)
                If pFeatSelection.SelectionSet.Count > 0 Then
                    Me.rdbSelected.Enabled = True
                    Me.rdbSelected.Text = "Selected Features (" & pFeatSelection.SelectionSet.Count.ToString & ")"
                Else
                    Me.rdbSelected.Enabled = False
                    Me.rdbSelected.Text = "Selected Features"
                End If
            End If
            'NEIGHBORHOOD SELECTED
            If Me.rdbNeighborhoodLayer.Checked Then
                Me.cmbLayers.Enabled = True
                Me.rdbPartial.Enabled = False
                If Me.rdbPartial.Checked Then
                    Me.rdbFull.Checked = True
                End If
                If Me.cmbLayers.Text.Length > 0 Then
                    m_strNeighborhoodLayerName = Me.cmbLayers.Text
                    pFeatLyr = m_arrPolyFeatureLayers.Item(Me.cmbLayers.FindString(Me.cmbLayers.Text))
                    pFeatSelection = CType(pFeatLyr, IFeatureSelection)
                    pFeatSelection.SelectionSet.Search(Nothing, False, pCursor)
                    If pFeatSelection.SelectionSet.Count > 0 Then
                        Me.rdbSelected.Enabled = True
                        Me.rdbSelected.Text = "Selected Features (" & pFeatSelection.SelectionSet.Count.ToString & ")"
                    Else
                        Me.rdbSelected.Enabled = False
                        Me.rdbSelected.Text = "Selected Features"
                    End If
                Else
                    Me.rdbSelected.Enabled = False
                    Me.rdbSelected.Text = "Selected Features"
                End If
            End If
            GoTo CleanUp
        Catch ex As Exception
            GoTo CleanUp
        End Try

CleanUp:
        pFeatLyr = Nothing
        pFeatSelection = Nothing
        pCursor = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub

End Class

