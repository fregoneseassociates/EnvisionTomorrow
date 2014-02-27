Option Explicit On

'MIRCO SOFT REFERENCE
Imports System
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Runtime.InteropServices
Imports System.Collections
Imports System.IO
Imports System.Drawing.Printing
Imports System.Data
Imports Microsoft.Office
Imports System.Math

'ESRI REFERENCES
Imports ESRI.ArcGIS.AnalysisTools
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.CartoUI
Imports ESRI.ArcGIS.ConversionTools
Imports ESRI.ArcGIS.Catalog
Imports ESRI.ArcGIS.CatalogUI
Imports ESRI.ArcGIS.DataSourcesGDB
Imports ESRI.ArcGIS.DataSourcesFile
Imports ESRI.ArcGIS.Display
Imports ESRI.ArcGIS.Editor
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.GeoDatabaseUI
Imports ESRI.ArcGIS.Geoprocessor
Imports ESRI.ArcGIS.Geoprocessing
Imports ESRI.ArcGIS.SystemUI
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.ArcCatalog

Public Class frmFeatureSum

    Private m_arrPntFeatureLayers As New ArrayList
    Private m_arrArcFeatureLayers As New ArrayList
    Public m_arrPolyFeatureLayers As New ArrayList
    Private pGP As ESRI.ArcGIS.Geoprocessor.Geoprocessor
    Private m_pCS As ISpatialReference2
    Private m_intAcresFld As Integer = -1
    Private m_intSqMiFld As Integer = -1
    Private m_blnAcres As Boolean = False
    Private m_blnSqMi As Boolean = False


    Private Sub frmFeatureSum_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        m_arrPntFeatureLayers = Nothing
        m_arrArcFeatureLayers = Nothing
        m_arrPolyFeatureLayers = Nothing
        m_frmAccessFeatureSummary = Nothing
    End Sub

    Private Sub frmFeatureSum_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'POPULATE THE INPUTS WITH THE POINT AND LINE LAYERS AVAILABLE IN THE CURRENT VIEW DOCUMENT
        Dim mxApplication As IMxApplication = Nothing
        Dim pMxDocument As IMxDocument = Nothing
        Dim mapCurrent As Map
        Dim intCount As Integer
        Dim pFeatLyr As IFeatureLayer
        Dim pFeatClass As IFeatureClass
        Dim intLayer As Integer
        Dim intFeatCount As Integer = 0
        Dim pActiveView As IActiveView = Nothing
        Dim pDataset As IDataset
        Dim pFeatSelection As IFeatureSelection
        Dim pCursor As ICursor = Nothing
        Dim pGeoDataSet As IGeoDataset
        Dim pFeatureClass As IFeatureClass
        Dim pField As IField
        Dim intFld As Integer
        Dim pCS As ISpatialReference2

        'EXIT IF NO ENVISION LAYER
        If m_pEditFeatureLyr Is Nothing Then
            Me.Close()
            GoTo CleanUp
        End If

        'RETRIEVE CURRENT VIEW DOCUMENT TO OBTAIN LIST OF CURRENT LAYER(S)
        If Not TypeOf m_appEnvision Is IApplication Then
            GoTo CleanUp
        Else
            pMxDocument = DirectCast(m_appEnvision.Document, IMxDocument)
            mapCurrent = CType(pMxDocument.FocusMap, Map)
            pActiveView = CType(pMxDocument.FocusMap, IActiveView)
        End If

        'CHECK FOR AVAILABLE LAYER(S)
        If mapCurrent.LayerCount = 0 Then
            MessageBox.Show("Please add POINT and LINE layers to the current view document you would like to utilize in the process.", "No Layers Found", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            Me.Close()
            GoTo CleanUp
        End If

        'ENABLE THE SELECTION SET PROCESSING OPTION IF A SELECTION IS AVIALABLE
        pFeatSelection = CType(m_pEditFeatureLyr, IFeatureSelection)
        pFeatSelection.SelectionSet.Search(Nothing, False, pCursor)
        If pFeatSelection.SelectionSet.Count > 0 Then
            Me.rdbSelected.Enabled = True
            Me.rdbSelected.Text = "Selected Features (" & pFeatSelection.SelectionSet.Count.ToString & ")"
        End If

        'CLEAR OUT ANY EXISTING ROWS and ARRAYS
        Me.dgvARC.Rows.Clear()
        Me.dgvPOINT.Rows.Clear()
        m_arrPntFeatureLayers.Clear()
        m_arrPntFeatureLayers = New ArrayList
        m_arrArcFeatureLayers.Clear()
        m_arrArcFeatureLayers = New ArrayList

        'CYCLE THROUGH LAYERS TO POPULATE ARRAYS AND CONTROLS
        For intLayer = 0 To mapCurrent.LayerCount - 1
            If TypeOf mapCurrent.Layer(intLayer) Is IFeatureLayer Then
                pFeatLyr = CType(mapCurrent.Layer(intLayer), IFeatureLayer)
                pFeatClass = pFeatLyr.FeatureClass
                pDataset = CType(pFeatClass, IDataset)
                If pFeatClass.ShapeType = esriGeometryType.esriGeometryPolyline Then
                    Try
                        pGeoDataSet = pFeatLyr
                        pCS = pGeoDataSet.SpatialReference
                        If Not UCase(pCS.Name) = "UNKNOWN" Then
                            m_arrArcFeatureLayers.Add(pFeatLyr)
                            Me.dgvARC.Rows.Add()
                            Me.dgvARC.Rows(Me.dgvARC.RowCount - 1).Cells(0).Value = pFeatLyr.Name
                            Me.dgvARC.Rows(Me.dgvARC.RowCount - 1).Cells(2).Value = pDataset.Name & "_LENGTH"
                            Me.dgvARC.Rows(Me.dgvARC.RowCount - 1).Cells(3).Value = "Miles"
                            Me.dgvARC.Rows(Me.dgvARC.RowCount - 1).Cells(5).Value = pDataset.Name & "_LENGTH_ACRES"
                            Me.dgvARC.Rows(Me.dgvARC.RowCount - 1).Cells(7).Value = pDataset.Name & "_LENGTH_SQMI"
                            intFeatCount = intFeatCount + 1
                        End If
                    Catch ex As Exception
                    End Try
                ElseIf pFeatClass.ShapeType = esriGeometryType.esriGeometryPoint Then
                    Try
                        pGeoDataSet = pFeatLyr
                        pCS = pGeoDataSet.SpatialReference
                        If Not UCase(pCS.Name) = "UNKNOWN" Then
                            m_arrPntFeatureLayers.Add(pFeatLyr)
                            Me.dgvPOINT.Rows.Add()
                            Me.dgvPOINT.Rows(Me.dgvPOINT.RowCount - 1).Cells(0).Value = pFeatLyr.Name
                            Me.dgvPOINT.Rows(Me.dgvPOINT.RowCount - 1).Cells(2).Value = pDataset.Name & "_CNT"
                            Me.dgvPOINT.Rows(Me.dgvPOINT.RowCount - 1).Cells(4).Value = pDataset.Name & "_CNT_ACRES"
                            Me.dgvPOINT.Rows(Me.dgvPOINT.RowCount - 1).Cells(6).Value = pDataset.Name & "_CNT_SQMI"
                            intFeatCount = intFeatCount + 1
                        End If
                    Catch ex As Exception
                    End Try
                ElseIf pFeatClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                    If Not m_pEditFeatureLyr.Name = pFeatLyr.Name Then
                        m_arrPolyFeatureLayers.Add(pFeatLyr)
                        Me.cmbLayers.Items.Add(pFeatLyr.Name)
                        intFeatCount = intFeatCount + 1
                    End If
                End If
                    pFeatLyr = Nothing
                    pFeatClass = Nothing
            End If
        Next

        'SET ENABLE STATUS OF DATA GRID VIEWS BASED ON CORRESPONDING LIST CONTENT
        If m_arrPntFeatureLayers.Count > 0 Then
            Me.dgvPOINT.Enabled = True
        Else
            Me.dgvPOINT.Enabled = False
        End If
        If m_arrArcFeatureLayers.Count > 0 Then
            Me.dgvARC.Enabled = True
        Else
            Me.dgvARC.Enabled = False
        End If

        'RETREIVE THE SPATIAL REFERENCE FROM THE ENVISION LAYER IF PROJECTED
        Try
            pGeoDataSet = m_pEditFeatureLyr
            m_pCS = pGeoDataSet.SpatialReference
            Me.tbxProjection.Text = m_pCS.Name
        Catch ex1 As Exception
            Try
                m_pCS = pMxDocument.FocusMap.SpatialReference
                Me.tbxProjection.Text = m_pCS.Name
            Catch ex2 As Exception
                m_pCS = Nothing
            End Try
        End Try

        'LOAD IN ALL INTEGER AND DOUBLE FIELDS INTO THE ACRES AND SQ MI CONTROLS
        pFeatureClass = m_pEditFeatureLyr.FeatureClass
        For intFld = 0 To pFeatureClass.Fields.FieldCount - 1
            pField = pFeatureClass.Fields.Field(intFld)
            If pField.Type = esriFieldType.esriFieldTypeInteger Or pField.Type = esriFieldType.esriFieldTypeDouble Then
                If Not pField.Name = "OBJECTID" Then
                    Me.cmbAcresFld.Items.Add(pField.Name)
                    If UCase(pField.Name) = "ACRES" Then
                        Me.cmbAcresFld.Text = (pField.Name)
                    End If
                    Me.cmbSqMi.Items.Add(pField.Name)
                    If UCase(pField.Name) = "SQMI" Then
                        Me.cmbSqMi.Text = (pField.Name)
                    End If
                End If
            End If
        Next
        If Me.cmbAcresFld.Text.Length <= 0 Then
            If Me.cmbAcresFld.Items.Count > 0 Then
                Me.cmbAcresFld.Text = Me.cmbAcresFld.Items(0)
                Me.lblArc.Enabled = False
            Else
                Me.chkAcres.Checked = False
                Me.lblArc.Enabled = False
            End If
        End If
        If Me.cmbSqMi.Text.Length <= 0 Then
            If Me.cmbSqMi.Items.Count > 0 Then
                Me.cmbSqMi.Text = Me.cmbSqMi.Items(0)
                Me.lblPoint.Enabled = False
            Else
                Me.chkSqMi.Checked = False
                Me.lblPoint.Enabled = False
            End If
        End If

        'CLOSE FORM IF NO ARC OR POINT LAYERS WERE FOUND
        If intFeatCount <= 0 Then
            MessageBox.Show("No Point or Line feature layers were found in the current view document.  Please add layers to utilize these processes.", "No Feature Layers Found", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            Me.Close()
        End If

        'DISABLE OPTIONS THAT ARE NOT VALID
        Me.lblArc.Enabled = True
        Me.dgvARC.Enabled = True
        If Me.dgvARC.RowCount <= 0 Then
            Me.lblArc.Enabled = False
            Me.dgvARC.Enabled = False
        End If
        Me.lblPoint.Enabled = True
        Me.dgvPOINT.Enabled = True
        If Me.dgvPOINT.RowCount <= 0 Then
            Me.lblPoint.Enabled = False
            Me.dgvPOINT.Enabled = False
        End If

        'DISABLE NEIGHBOR IF NOT FEATURE LAYERS FOUND
        If intFeatCount <= 0 Then
            Me.cmbLayers.Enabled = False
            Me.rdbNeighborhoodLayer.Enabled = False
        Else
            Me.rdbNeighborhoodLayer.Enabled = True
        End If

        'PRESELECT THE NEIGHBORHOOD LAYER IF AVAILABLE
        If m_strNeighborhoodLayerName.Length > 0 Then
            If m_frmAccessFeatureSummary.cmbLayers.FindString(m_strNeighborhoodLayerName) Then
                m_frmAccessFeatureSummary.cmbLayers.Text = m_strNeighborhoodLayerName
            End If
        End If

        GoTo CleanUp

CleanUp:
        mxApplication = Nothing
        pMxDocument = Nothing
        mapCurrent = Nothing
        intCount = Nothing
        pFeatLyr = Nothing
        intLayer = Nothing
        intFeatCount = Nothing
        pActiveView = Nothing
        pDataset = Nothing
        pFeatSelection = Nothing
        pCursor = Nothing
        pGeoDataSet = Nothing
        pFeatureClass = Nothing
        pField = Nothing
        intFld = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub

    Private Sub btnRun_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRun.Click
        'RUN THE PROCESSING TO COUNT POINT FEATURES AND CLIP OUT LINE FEATURES FOR LENGTH SUMMARIES
        Dim pFeatClass As IFeatureClass
        Dim pTable As ITable
        Dim blnPointFound As Boolean = False
        Dim blnLineFound As Boolean = False
        Dim intRow As Integer = 0
        Dim strField As String = ""
        Dim blnExit As Boolean = False
        Dim intPointFailed As Integer = 0
        Dim intLineFailed As Integer = 0
        Dim intAcresFld As Integer = -1
        Dim intSqMiFld As Integer = -1
        Dim pFeatLyr As IFeatureLayer

        'REFRESH THE DATA GRIDS TO SHOW ANY UNACCEPTED INPUTS
        Me.dgvPOINT.EndEdit()
        Me.dgvPOINT.Refresh()
        Me.dgvPOINT.RefreshEdit()
        Me.dgvARC.EndEdit()
        Me.dgvARC.Refresh()
        Me.dgvARC.RefreshEdit()
        Me.Refresh()

        'SHOW STATUS CONTROLS
        Me.lblRunStatus.Text = ""
        Me.lblRunStatus.Visible = True
        Me.barStatusRun.Visible = True
        Me.barStatusRun.Value = 0

        'CHECK IF THERE IS A FEATURE SET OPTIONS SELECTED
        If Not Me.rdbFull.Checked And Not Me.rdbSelected.Checked And Not Me.rdbPartial.Checked Then
            MessageBox.Show("Please selected a Processing Feature Selection Option", "Processing Feature Selection Option Not Selected", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo CleanUp
        End If

        'EXIT IF NO ENVISION LAYHER HAS BEEN SET
        If me.rdbNeighborhoodLayer.Checked Then
            Try
                pFeatLyr = m_arrPolyFeatureLayers.Item(m_frmAccessFeatureSummary.cmbLayers.FindString(m_frmAccessFeatureSummary.cmbLayers.Text))
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

        'CHECK IF POINT LAYERS WILL BE PROCESSED
        If Me.dgvPOINT.Rows.Count > 0 Then
            For intRow = 0 To Me.dgvPOINT.Rows.Count - 1
                If Me.dgvPOINT.Rows(intRow).Cells(1).Value Or Me.dgvPOINT.Rows(intRow).Cells(3).Value Or _
                    Me.dgvPOINT.Rows(intRow).Cells(5).Value Then
                    blnPointFound = True
                    Exit For
                End If
            Next
        End If

        'CHECK IF A POINT TRACKING OPTION HAS AN OUTPUT FIELD DESIGNATED, ADD IF MISSING FROM ENVISION LAYER
        If Me.dgvPOINT.Rows.Count > 0 Then
            For intRow = 0 To Me.dgvPOINT.Rows.Count - 1
                If Me.dgvPOINT.Rows(intRow).Cells(1).Value Then
                    Try
                        strField = CStr(Me.dgvPOINT.Rows(intRow).Cells(2).Value)
                        If strField.Length <= 0 Then
                            blnExit = True
                        Else
                            Me.lblRunStatus.Text = "Adding Processing Field: " & strField
                            AddEnvisionField(pTable, strField, "DOUBLE", 16, 6)
                        End If
                    Catch ex As Exception
                        blnExit = True
                    End Try
                End If
                If Me.dgvPOINT.Rows(intRow).Cells(3).Value Then
                    m_blnSqMi = True
                    Try
                        strField = CStr(Me.dgvPOINT.Rows(intRow).Cells(4).Value)
                        If strField.Length <= 0 Then
                            blnExit = True
                        Else
                            Me.lblRunStatus.Text = "Adding Processing Field: " & strField
                            AddEnvisionField(pTable, strField, "DOUBLE", 16, 6)
                        End If
                    Catch ex As Exception
                        blnExit = True
                    End Try
                End If
                If Me.dgvPOINT.Rows(intRow).Cells(5).Value Then
                    m_blnAcres = True
                    Try
                        strField = CStr(Me.dgvPOINT.Rows(intRow).Cells(6).Value)
                        If strField.Length <= 0 Then
                            blnExit = True
                        Else
                            Me.lblRunStatus.Text = "Adding Processing Field: " & strField
                            AddEnvisionField(pTable, strField, "DOUBLE", 16, 6)
                        End If
                    Catch ex As Exception
                        blnExit = True
                    End Try
                End If
                If blnExit Then
                    MessageBox.Show("Please review the POINT layer processing options.  An output field has not been defined.", "Output Field Required", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
                    GoTo CleanUp
                End If
            Next
        End If

        'CHECK IF LINE LAYERS WILL BE PROCESSED
        If Me.dgvARC.Rows.Count > 0 Then
            For intRow = 0 To Me.dgvARC.Rows.Count - 1
                If Me.dgvARC.Rows(intRow).Cells(1).Value Or Me.dgvARC.Rows(intRow).Cells(4).Value Or _
                    Me.dgvARC.Rows(intRow).Cells(6).Value Then
                    blnLineFound = True
                    Exit For
                End If
            Next
        End If

        'CHECK IF A LINE TRACKING OPTION HAS AN OUTPUT FIELD DESIGNATED, ADD IF MISSING FROM ENVISION LAYER
        If Me.dgvARC.Rows.Count > 0 Then
            For intRow = 0 To Me.dgvPOINT.Rows.Count - 1
                If Me.dgvARC.Rows(intRow).Cells(1).Value Then
                    Try
                        strField = CStr(Me.dgvARC.Rows(intRow).Cells(2).Value)
                        If strField.Length <= 0 Then
                            blnExit = True
                        Else
                            Me.lblRunStatus.Text = "Adding Processing Field: " & strField
                            AddEnvisionField(pTable, strField, "DOUBLE", 16, 6)
                        End If
                    Catch ex As Exception
                        blnExit = True
                    End Try
                End If
                If Me.dgvARC.Rows(intRow).Cells(4).Value Then
                    m_blnSqMi = True
                    Try
                        strField = CStr(Me.dgvARC.Rows(intRow).Cells(5).Value)
                        If strField.Length <= 0 Then
                            blnExit = True
                        Else
                            Me.lblRunStatus.Text = "Adding Processing Field: " & strField
                            AddEnvisionField(pTable, strField, "DOUBLE", 16, 6)
                        End If
                    Catch ex As Exception
                        blnExit = True
                    End Try
                End If
                If Me.dgvARC.Rows(intRow).Cells(6).Value Then
                    m_blnAcres = True
                    Try
                        strField = CStr(Me.dgvARC.Rows(intRow).Cells(7).Value)
                        If strField.Length <= 0 Then
                            blnExit = True
                        Else
                            Me.lblRunStatus.Text = "Adding Processing Field: " & strField
                            AddEnvisionField(pTable, strField, "DOUBLE", 16, 6)
                        End If
                    Catch ex As Exception
                        blnExit = True
                    End Try
                End If
                If blnExit Then
                    MessageBox.Show("Please review the LINE layer processing options.  An output field has not been defined.", "Output Field Required", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
                    GoTo CleanUp
                End If
            Next
        End If

        'CHECK FOR AND SET ACRES AND SQMI FIELDS IF OPTIONS SELECTED
        If m_blnAcres Then
            If Me.cmbAcresFld.Text.Length <= 0 Then
                MessageBox.Show("Please select the field containing ACRE values.", "Missing Acres Field Input", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
                GoTo CleanUp
            End If
            m_intAcresFld = m_pEditFeatureLyr.FeatureClass.FindField(Me.cmbAcresFld.Text)
            If m_intAcresFld <= -1 Then
                MessageBox.Show("Please select the field containing ACRE values. The current selection is not valid.", "Missing Acres Field Input", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
                GoTo CleanUp
            End If
        End If
        If m_blnSqMi Then
            If Me.cmbSqMi.Text.Length <= 0 Then
                MessageBox.Show("Please select the field containing Sqare Mile values.", "Missing Sqare Mile Field Input", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
                GoTo CleanUp
            End If
            m_intSqMiFld = m_pEditFeatureLyr.FeatureClass.FindField(Me.cmbSqMi.Text)
            If m_intSqMiFld <= -1 Then
                MessageBox.Show("Please select the field containing Sqare Mile values. The current selection is not valid.", "Missing Sqare Mile Field Input", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
                GoTo CleanUp
            End If
        End If

        'CHECK IF A LINE TRACKING OPTION FOR LENGTH SUMMARY HAS A UNIT OPTION SELECTED
        If Me.dgvARC.Rows.Count > 0 Then
            For intRow = 0 To Me.dgvPOINT.Rows.Count - 1
                If Me.dgvARC.Rows(intRow).Cells(1).Value Then
                    Try
                        strField = CStr(Me.dgvARC.Rows(intRow).Cells(3).Value)
                        If strField.Length <= 0 Then
                            blnExit = True
                        End If
                    Catch ex As Exception
                        blnExit = True
                    End Try
                End If
                If blnExit Then
                    MessageBox.Show("Please review the LINE layer processing options.  Output units have not been defined for all selected line layers.", "Output Units Required", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
                    GoTo CleanUp
                End If
            Next
        End If

        'EXIT FOR NO PROCESSING OPTIONS ARE SELECTED
        If Not blnLineFound And Not blnPointFound Then
            MessageBox.Show("No POINT or LINE processing options have been selected.", "No Option Selected", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo CleanUp
        End If

        'EXECUTE POINT LAYER SUMMARIES IF SELECTED
        If blnPointFound Then
            intPointFailed = ProcessPointLayers()
        End If


        'EXECUTE LINE LAYER SUMMARIES IF SELECTED
        If blnLineFound Then
            intLineFailed = ProcessLineLayers()
        End If


        'REPORT LEVEL OF PROCESSING SUCCESS TO USER
        If intPointFailed > 0 Or intLineFailed > 0 Then
            If intPointFailed > 0 And intLineFailed > 0 Then
                MessageBox.Show("Processing has completed, but " & intPointFailed.ToString & " features failed the Point processes and " & intLineFailed.ToString & " features failed the Line processes", "Processing Completed with Exceptions", MessageBoxButtons.OK, MessageBoxIcon.None, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            ElseIf intPointFailed <= 0 And intLineFailed > 0 Then
                MessageBox.Show("Processing has completed, but " & intLineFailed.ToString & " features failed the Line processes", "Processing Completed with Exceptions", MessageBoxButtons.OK, MessageBoxIcon.None, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            ElseIf intPointFailed > 0 And intLineFailed <= 0 Then
                MessageBox.Show("Processing has completed, but " & intPointFailed.ToString & " features failed the Point processes", "Processing Completed with Exceptions", MessageBoxButtons.OK, MessageBoxIcon.None, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            End If
        Else
            MessageBox.Show("All features processed completely without error.", "Processing Completed", MessageBoxButtons.OK, MessageBoxIcon.None, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
        End If
CleanUp:
        pFeatClass = Nothing
        pTable = Nothing
        blnPointFound = Nothing
        blnLineFound = Nothing
        intRow = Nothing
        strField = Nothing
        blnExit = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub

    Private Function ProcessPointLayers() As Integer
        ProcessPointLayers = 0
        'CYCLE THROUGH THE SELECTED FEATURES FOR EACH SELECTED POINT LAYER
        Dim pFeatClass As IFeatureClass
        Dim pDef As IFeatureLayerDefinition2
        Dim pQFilter As IQueryFilter
        Dim strDefExpression As String = ""
        Dim pFSel As IFeatureSelection = Nothing
        Dim pCursor As ICursor = Nothing
        Dim strQString As String = Nothing
        Dim pFeatureCursor As IFeatureCursor = Nothing
        Dim intTotalCount As Integer
        Dim pQBLayer As IQueryByLayer
        Dim pFeat As IFeature
        Dim intObjId As Integer = 0
        Dim intCount As Integer = 0
        Dim intObjFld As Integer
        Dim intRow As Integer = 0
        Dim pLayer As ILayer = Nothing
        Dim pFSelOther As IFeatureSelection = Nothing
        Dim strField As String = ""
        Dim intFailed As Integer = 0
        Dim dblAcres As Double = 0
        Dim dblSqMi As Double = 0
        Dim dblAcresDensity As Double = 0
        Dim dblSqmiDensity As Double = 0
        Dim intAcresOutputFld As Integer = -1
        Dim intSqMiOutputFld As Integer = -1
        Dim pFeatLyr As IFeatureLayer

        Try
            'CHECK AND RETRIEVE THE PROCESSING OPTION
            If Me.rdbNeighborhoodLayer.Checked Then
                Try
                    pFeatLyr = m_arrPolyFeatureLayers.Item(m_frmAccessFeatureSummary.cmbLayers.FindString(m_frmAccessFeatureSummary.cmbLayers.Text))
                    pFeatClass = pFeatLyr.FeatureClass
                    pDef = pFeatLyr
                    pFSel = pFeatLyr
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
                    pDef = m_pEditFeatureLyr
                    pFSel = m_pEditFeatureLyr
                End If
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
            pFeat = pFeatureCursor.NextFeature
            Do While Not pFeat Is Nothing
                intCount = intCount + 1
                Me.barStatusRun.Value = (intCount / intTotalCount) * 100
                Me.Refresh()
                Try
                    intObjId = pFeat.Value(intObjFld)
                    pQFilter = New QueryFilter
                    strQString = "OBJECTID = " & intObjId.ToString
                    pQFilter.WhereClause = strQString
                    pFSel.Clear()
                    pFSel.SelectFeatures(pQFilter, esriSelectionResultEnum.esriSelectionResultNew, True)

                    'CYCLE THROUGH SELECTED POINT LAYERS
                    If pFSel.SelectionSet.Count() > 0 Then
                        For intRow = 0 To Me.dgvPOINT.Rows.Count - 1
                            'RETRIEVE THE POINT LAYER
                            If Me.dgvPOINT.Rows(intRow).Cells(1).Value Or Me.dgvPOINT.Rows(intRow).Cells(3).Value Or Me.dgvPOINT.Rows(intRow).Cells(5).Value Then
                                pLayer = m_arrPntFeatureLayers.Item(intRow)
                                pFSelOther = pLayer
                            End If

                            pQBLayer = New QueryByLayer
                            With pQBLayer
                                .ByLayer = pFSel
                                .FromLayer = pLayer
                                .ResultType = esriSelectionResultEnum.esriSelectionResultNew
                                .BufferDistance = 0
                                .BufferUnits = esriUnits.esriMiles
                                .LayerSelectionMethod = esriLayerSelectionMethod.esriLayerSelectHaveTheirCenterIn
                                .UseSelectedFeatures = True
                                pFSelOther.SelectionSet = .Select
                            End With

                            If Me.dgvPOINT.Rows(intRow).Cells(1).Value Then
                                strField = Me.dgvPOINT.Rows(intRow).Cells(2).Value
                                pFeat.Value(pFeatClass.FindField(strField)) = pFSelOther.SelectionSet.Count
                            End If

                            'RETRIEVE ACRES AND/OR SQMI FIELDS IF NEEDED
                            If m_blnAcres Then
                                intAcresOutputFld = -1
                                Try
                                    intAcresOutputFld = pFeatClass.FindField(Me.dgvPOINT.Rows(intRow).Cells(4).Value)
                                Catch ex As Exception

                                End Try
                            End If
                            If m_blnAcres And intAcresOutputFld > -1 Then
                                dblAcres = pFeat.Value(m_intAcresFld)
                                dblAcresDensity = 0
                                Try
                                    dblAcresDensity = pFSelOther.SelectionSet.Count / dblAcres
                                Catch ex As Exception
                                End Try
                                pFeat.Value(intAcresOutputFld) = dblAcresDensity
                            End If
                            If m_blnSqMi Then
                                intSqMiOutputFld = -1
                                Try
                                    intSqMiOutputFld = pFeatClass.FindField(Me.dgvPOINT.Rows(intRow).Cells(6).Value)
                                Catch ex As Exception

                                End Try
                            End If
                            If m_blnSqMi And intSqMiOutputFld > -1 Then
                                dblSqMi = pFeat.Value(m_intSqMiFld)
                                dblSqmiDensity = 0
                                Try
                                    dblSqmiDensity = pFSelOther.SelectionSet.Count / dblSqMi
                                Catch ex As Exception
                                End Try
                                pFeat.Value(intSqMiOutputFld) = dblSqmiDensity
                            End If

                            pFeat.Store()
                            pLayer = Nothing
                            pQBLayer = Nothing
                            pFSelOther.Clear()
                        Next
                    End If
                Catch ex As Exception
                    'MessageBox.Show("Error: " & ex.Message & vbNewLine & "Count: " & intCount.ToString & vbNewLine & "OBJECTID = " & intObjId.ToString, "Buffer Summary Process Error", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
                    intFailed = intFailed + 1
                End Try

                pFeat = pFeatureCursor.NextFeature
            Loop

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message & vbNewLine & "Count: " & intCount.ToString & vbNewLine & "OBJECTID = " & intObjId.ToString, "Buffer Summary Process Error", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo CleanUp
        End Try
CleanUp:
        ProcessPointLayers = intFailed
        If intFailed > 0 Then
            MessageBox.Show(intFailed.ToString & " feature(s) failed to processing for the buffering.", "Incomplete Processing", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
        End If
        pFeatClass = Nothing
        pFeatClass = Nothing
        pDef = Nothing
        pQFilter = Nothing
        strDefExpression = Nothing
        pFSel = Nothing
        pCursor = Nothing
        strQString = Nothing
        pFeatureCursor = Nothing
        intTotalCount = Nothing
        pQBLayer = Nothing
        pFeat = Nothing
        intObjId = Nothing
        intCount = Nothing
        intObjFld = Nothing
        intRow = Nothing
        pLayer = Nothing
        pFSelOther = Nothing
        strField = Nothing
        intFailed = Nothing
        dblAcres = Nothing
        dblSqMi = Nothing
        dblAcresDensity = Nothing
        dblSqmiDensity = Nothing
        intAcresOutputFld = Nothing
        intSqMiOutputFld = Nothing
        pFeatLyr = Nothing
        GC.WaitForFullGCComplete()
        GC.Collect()
    End Function

    Private Function ProcessLineLayers() As Integer
        'CLIP THE SELECTED PARCEL THEME BY THE SELECTED AOI
        'USED FOR PARCEL ONLY AND HYBRID OPTIONS

        Dim pFeatClass As IFeatureClass
        Dim pDef As IFeatureLayerDefinition2
        Dim pQFilter As IQueryFilter
        Dim strDefExpression As String = ""
        Dim pFSel As IFeatureSelection = Nothing
        Dim strQString As String = Nothing
        Dim pFeatureCursor As IFeatureCursor = Nothing
        Dim intTotalCount As Integer
        Dim pCursor As ICursor = Nothing
        Dim pDataset As IDataset = Nothing
        Dim objOutputSpatref As Object
        Dim objWorkspace As Object = Nothing
        Dim pFeat As IFeature
        Dim intObjId As Integer = 0
        Dim intCount As Integer = 0
        Dim intObjFld As Integer
        Dim intRow As Integer = 0
        Dim pLayer As ILayer = Nothing
        Dim pClip As ESRI.ArcGIS.AnalysisTools.Clip
        Dim pWksFactory As IWorkspaceFactory = Nothing
        Dim pFeatWks As IFeatureWorkspace
        Dim pClipFeatClass As IFeatureClass
        Dim strField As String = ""
        Dim intFailed As Integer = 0
        Dim pClipFeatureCursor As IFeatureCursor = Nothing
        Dim pClipFeat As IFeature
        Dim dblLength As Double
        Dim dblTotalLength As Double
        Dim intFldLength As Integer
        Dim dblAcres As Double = 0
        Dim dblSqMi As Double = 0
        Dim dblAcresDensity As Double = 0
        Dim dblSqmiDensity As Double = 0
        Dim intAcresOutputFld As Integer = -1
        Dim intSqMiOutputFld As Integer = -1
        Dim strUnits As String = "MILES"
        Dim pPCS As IProjectedCoordinateSystem4
        Dim pFeatLyr As IFeatureLayer = Nothing

        Try
            'CHECK AND RETRIEVE THE PROCESSING OPTION
            If Me.rdbNeighborhoodLayer.Checked Then
                Try
                    pFeatLyr = m_arrPolyFeatureLayers.Item(m_frmAccessFeatureSummary.cmbLayers.FindString(m_frmAccessFeatureSummary.cmbLayers.Text))
                    pFeatClass = pFeatLyr.FeatureClass
                    pDef = pFeatLyr
                    pFSel = pFeatLyr
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
                    pDef = m_pEditFeatureLyr
                    pFSel = m_pEditFeatureLyr
                End If
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

            'DEFINE THE GEOPROCESSOR FOR THE CLIPPING PROCESSES
            If pGP Is Nothing Then
                pDataset = CType(pFeatClass, IDataset)
                objWorkspace = pDataset.Workspace.PathName
                objOutputSpatref = m_pCS.FactoryCode
                pGP = New ESRI.ArcGIS.Geoprocessor.Geoprocessor
                pGP.SetEnvironmentValue("workspace", objWorkspace)
                pGP.SetEnvironmentValue("outputCoordinateSystem", objOutputSpatref)
                pGP.OverwriteOutput = True
                pGP.AddOutputsToMap = True
                pGP.TemporaryMapLayers = True
            End If
            pPCS = CType(m_pCS, IProjectedCoordinateSystem4)
            strUnits = CStr(pPCS.CoordinateUnit.Name)
            intObjFld = pFeatClass.FindField("OBJECTID")
            pFeat = pFeatureCursor.NextFeature
            Do While Not pFeat Is Nothing
                intCount = intCount + 1
                Me.barStatusRun.Value = (intCount / intTotalCount) * 100
                Me.Refresh()
                Try
                    intObjId = pFeat.Value(intObjFld)
                    pQFilter = New QueryFilter
                    strQString = "OBJECTID = " & intObjId.ToString
                    pQFilter.WhereClause = strQString
                    pFSel.Clear()
                    pFSel.SelectFeatures(pQFilter, esriSelectionResultEnum.esriSelectionResultNew, True)

                    'CYCLE THROUGH SELECTED LINE LAYERS
                    If pFSel.SelectionSet.Count() > 0 Then
                        For intRow = 0 To Me.dgvARC.Rows.Count - 1
                            'RETRIEVE THE LINE LAYER
                            If Me.dgvARC.Rows(intRow).Cells(1).Value Or Me.dgvARC.Rows(intRow).Cells(4).Value Or Me.dgvARC.Rows(intRow).Cells(6).Value Then
                                pLayer = m_arrArcFeatureLayers.Item(intRow)
                            End If
                            pClip = New ESRI.ArcGIS.AnalysisTools.Clip
                            pClip.in_features = pLayer
                            If Me.rdbNeighborhoodLayer.Checked Then
                                pClip.clip_features = pFeatLyr
                            Else
                                pClip.clip_features = m_pEditFeatureLyr
                            End If
                            pClip.out_feature_class = objWorkspace & "\ENVISION_ACCESS_LINE_FEAT_CLIP_TEMP"
                            RunTool(pGP, pClip)
                            pLayer = Nothing
                            pClip = Nothing
                            'RETRIEVE THE CLIPPED LAYER AND PROCESS THE FEATURES
                            pWksFactory = New FileGDBWorkspaceFactory
                            pFeatWks = pWksFactory.OpenFromFile(objWorkspace, 0)
                            pClipFeatClass = pFeatWks.OpenFeatureClass("ENVISION_ACCESS_LINE_FEAT_CLIP_TEMP")
                            intFldLength = pClipFeatClass.FindField("Shape_Length")

                            pClipFeatureCursor = pClipFeatClass.Search(Nothing, False)
                            pClipFeat = pClipFeatureCursor.NextFeature
                            If intFldLength > -1 Then
                                dblTotalLength = 0
                                Do While Not pClipFeat Is Nothing
                                    dblLength = 0
                                    Try
                                        dblLength = pClipFeat.Value(intFldLength)
                                    Catch ex As Exception
                                    End Try
                                    dblTotalLength = dblTotalLength + dblLength
                                    pClipFeat = pClipFeatureCursor.NextFeature
                                Loop
                            End If
                            pClipFeat = Nothing
                            pClipFeatureCursor = Nothing
                            pClipFeatClass = Nothing
                            GC.WaitForPendingFinalizers()
                            GC.Collect()

                            'SET THE LENGTH VALUE IF CHECKED
                            If Me.dgvARC.Rows(intRow).Cells(1).Value Then
                                intFldLength = -1
                                Try
                                    intFldLength = pFeatClass.FindField(Me.dgvARC.Rows(intRow).Cells(2).Value)
                                Catch ex As Exception
                                End Try
                                If UCase(CStr(Me.dgvARC.Rows(intRow).Cells(3).Value)) = "MILES" Then
                                    dblTotalLength = ReturnMiles(dblTotalLength, strUnits)
                                Else
                                    dblTotalLength = ReturnFeet(dblTotalLength, strUnits)
                                End If
                                pFeat.Value(intFldLength) = dblTotalLength
                            End If
                            'RETRIEVE ACRES AND/OR SQMI FIELDS IF NEEDED
                            If m_blnAcres Then
                                intAcresOutputFld = -1
                                Try
                                    intAcresOutputFld = pFeatClass.FindField(Me.dgvPOINT.Rows(intRow).Cells(4).Value)
                                Catch ex As Exception
                                End Try
                            End If
                            If m_blnAcres And intAcresOutputFld > -1 Then
                                dblAcres = pFeat.Value(m_intAcresFld)
                                dblAcresDensity = 0
                                Try
                                    dblAcresDensity = dblTotalLength / dblAcres
                                Catch ex As Exception
                                End Try
                                pFeat.Value(intAcresOutputFld) = dblAcresDensity
                            End If
                            If m_blnSqMi Then
                                intSqMiOutputFld = -1
                                Try
                                    intSqMiOutputFld = pFeatClass.FindField(Me.dgvPOINT.Rows(intRow).Cells(6).Value)
                                Catch ex As Exception
                                End Try
                            End If
                            If m_blnSqMi And intSqMiOutputFld > -1 Then
                                dblSqMi = pFeat.Value(m_intSqMiFld)
                                dblSqmiDensity = 0
                                Try
                                    dblSqmiDensity = dblTotalLength / dblSqMi
                                Catch ex As Exception
                                End Try
                                pFeat.Value(intSqMiOutputFld) = dblSqmiDensity
                            End If

                            pFeat.Store()
                            pLayer = Nothing
                        Next
                    End If
                Catch ex As Exception
                    'MessageBox.Show("Error: " & ex.Message & vbNewLine & "Count: " & intCount.ToString & vbNewLine & "OBJECTID = " & intObjId.ToString, "Buffer Summary Process Error", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
                    intFailed = intFailed + 1
                End Try

                pFeat = pFeatureCursor.NextFeature
            Loop

        Catch ex As Exception
            MessageBox.Show(ex.Message)
            GoTo CleanUp
        End Try

CleanUp:
        ProcessLineLayers = intFailed
        If intFailed > 0 Then
            MessageBox.Show(intFailed.ToString & " feature(s) failed to processing for the buffering.", "Incomplete Processing", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
        End If
        pFeatClass = Nothing
        pDef = Nothing
        pQFilter = Nothing
        strDefExpression = Nothing
        pFSel = Nothing
        pCursor = Nothing
        strQString = Nothing
        pFeatureCursor = Nothing
        intTotalCount = Nothing
        pFeat = Nothing
        intObjId = Nothing
        intCount = Nothing
        intObjFld = Nothing
        intRow = Nothing
        pLayer = Nothing
        strField = Nothing
        intFailed = Nothing
        objOutputSpatref = Nothing
        objWorkspace = Nothing
        pClip = Nothing
        pWksFactory = Nothing
        pFeatWks = Nothing
        pFeatClass = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Function

    Private Sub btnSelectPrj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSelectPrj.Click
        'PROVIDE THE SPATIAL REFERENCE PROPERTIES DIALOG TO THE USER TO DEFINE THE PROJECT PROJECTION
        Dim pSpaRefDlg As ISpatialReferenceDialog
        Dim pSpatRef As ISpatialReference
        Try
            pSpaRefDlg = New SpatialReferenceDialog
            Me.WindowState = FormWindowState.Minimized
            pSpatRef = pSpaRefDlg.DoModalCreate(False, False, False, 0)
            If pSpatRef Is Nothing Then
                GoTo CleanUp
            Else
                Me.tbxProjection.Text = pSpatRef.Name
                m_pCS = pSpatRef
            End If
        Catch ex As Exception
            MessageBox.Show("An error occured while defining the processing projection." & vbNewLine & ex.Message, "Projection Select Error")
            GoTo CleanUp
        End Try
CleanUp:
        Me.WindowState = FormWindowState.Normal
        pSpaRefDlg = Nothing
        pSpatRef = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub

    Private Sub chkAcres_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkAcres.CheckedChanged
        AccessStatusControlUpdate()
    End Sub

    Private Sub chkSqMi_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkSqMi.CheckedChanged
        AccessStatusControlUpdate()
    End Sub

    Private Sub AccessStatusControlUpdate()
        'CONTROL THE ENABLE AND VISUAL STATUS OF RELATED CONTROLS
        Me.lblAcres.Enabled = Me.chkAcres.Checked
        Me.cmbAcresFld.Enabled = Me.chkAcres.Checked
        Me.lblSqMi.Enabled = Me.chkSqMi.Checked
        Me.cmbSqMi.Enabled = Me.chkSqMi.Checked
        Me.dgvPOINT.Columns.Item(3).Visible = Me.chkSqMi.Checked
        Me.dgvPOINT.Columns.Item(4).Visible = Me.chkSqMi.Checked
        Me.dgvPOINT.Columns.Item(5).Visible = Me.chkAcres.Checked
        Me.dgvPOINT.Columns.Item(6).Visible = Me.chkAcres.Checked
        Me.dgvARC.Columns.Item(4).Visible = Me.chkSqMi.Checked
        Me.dgvARC.Columns.Item(5).Visible = Me.chkSqMi.Checked
        Me.dgvARC.Columns.Item(6).Visible = Me.chkAcres.Checked
        Me.dgvARC.Columns.Item(7).Visible = Me.chkAcres.Checked
    End Sub

    Private Sub cmbLayers_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbLayers.SelectedIndexChanged
        NeighborhoodSelect()
    End Sub

    Private Sub rdbNeighborhoodLayer_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdbNeighborhoodLayer.CheckedChanged
        NeighborhoodSelect()
    End Sub

    Private Sub NeighborhoodSelect()
        Dim pFeatSelection As IFeatureSelection
        Dim pCursor As ICursor = Nothing
        Dim pFeatLyr As IFeatureLayer
        Dim pFeatureClass As IFeatureClass = Nothing
        Dim intFld As Integer = 0
        Dim pField As IField

        Try
            'PARCEL SELECTED
            If Me.rdbParcelLayer.Checked Then
                Me.cmbLayers.Enabled = False
                Me.rdbPartial.Enabled = True
                pFeatureClass = m_pEditFeatureLyr.FeatureClass
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
                    pFeatureClass = pFeatLyr.FeatureClass
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
                    pFeatureClass = Nothing
                    Me.rdbSelected.Enabled = False
                    Me.rdbSelected.Text = "Selected Features"
                End If
            End If

            'FIELDS
            'LOAD IN ALL INTEGER AND DOUBLE FIELDS INTO THE ACRES AND SQ MI CONTROLS
            If pFeatureClass Is Nothing Then
                Me.chkAcres.Enabled = False
                Me.chkAcres.Checked = False
                Me.lblAcres.Enabled = False
                Me.cmbAcresFld.Enabled = False
                Me.cmbAcresFld.Items.Clear()
                Me.chkSqMi.Enabled = False
                Me.chkSqMi.Checked = False
                Me.lblSqMi.Enabled = False
                Me.cmbSqMi.Enabled = False
                Me.cmbSqMi.Items.Clear()
            Else
                Me.chkAcres.Enabled = True
                Me.chkSqMi.Enabled = True
                For intFld = 0 To pFeatureClass.Fields.FieldCount - 1
                    pField = pFeatureClass.Fields.Field(intFld)
                    If pField.Type = esriFieldType.esriFieldTypeInteger Or pField.Type = esriFieldType.esriFieldTypeDouble Then
                        If Not pField.Name = "OBJECTID" Then
                            Me.cmbAcresFld.Items.Add(pField.Name)
                            If UCase(pField.Name) = "ACRES" Then
                                Me.cmbAcresFld.Text = (pField.Name)
                            End If
                            Me.cmbSqMi.Items.Add(pField.Name)
                            If UCase(pField.Name) = "SQMI" Then
                                Me.cmbSqMi.Text = (pField.Name)
                            End If
                        End If
                    End If
                Next
                If Me.cmbAcresFld.Text.Length <= 0 Then
                    If Me.cmbAcresFld.Items.Count > 0 Then
                        Me.cmbAcresFld.Text = Me.cmbAcresFld.Items(0)
                        Me.lblAcres.Enabled = False
                    Else
                        Me.chkAcres.Checked = False
                        Me.lblAcres.Enabled = False
                    End If
                Else
                    If Me.cmbAcresFld.FindString(Me.cmbAcresFld.Text) <= -1 Then
                        If Me.cmbAcresFld.Items.Count > 0 Then
                            Me.cmbAcresFld.Text = Me.cmbAcresFld.Items(0)
                            Me.lblAcres.Enabled = False
                        Else
                            Me.chkAcres.Checked = False
                            Me.lblAcres.Enabled = False
                        End If
                    End If
                End If
                If Me.cmbSqMi.Text.Length <= 0 Then
                    If Me.cmbSqMi.Items.Count > 0 Then
                        Me.cmbSqMi.Text = Me.cmbSqMi.Items(0)
                        Me.lblSqMi.Enabled = False
                    Else
                        Me.chkSqMi.Checked = False
                        Me.lblSqMi.Enabled = False
                    End If
                Else
                    If Me.cmbSqMi.FindString(Me.cmbSqMi.Text) <= -1 Then
                        If Me.cmbSqMi.Items.Count > 0 Then
                            Me.cmbSqMi.Text = Me.cmbSqMi.Items(0)
                            Me.lblSqMi.Enabled = False
                        Else
                            Me.chkSqMi.Checked = False
                            Me.lblSqMi.Enabled = False
                        End If
                    End If
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
        pFeatureClass = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub

End Class