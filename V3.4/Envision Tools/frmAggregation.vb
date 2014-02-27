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
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.CartoUI
Imports ESRI.ArcGIS.DataSourcesGDB
Imports ESRI.ArcGIS.DataSourcesFile
Imports ESRI.ArcGIS.Display
Imports ESRI.ArcGIS.Editor
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.GeoDatabaseUI
Imports ESRI.ArcGIS.Geoprocessor
Imports ESRI.ArcGIS.SystemUI
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.ArcCatalog

Public Class frmAggregation
    Public m_arrPolyFeatureLayers As New ArrayList
    Dim blnOpening As Boolean = True
    Dim arrWeightedSum As ArrayList = New ArrayList
    Dim arrTotalSum As ArrayList = New ArrayList
    Dim arrDevTypes As ArrayList = New ArrayList
    Dim arrRedevRates As ArrayList = New ArrayList
    Dim arrAbadonRates As ArrayList = New ArrayList
    Dim arrNetAcrePercents As ArrayList = New ArrayList
    Dim pNeighborhoodLyr As IFeatureLayer
    Dim arrSelected As ArrayList = New ArrayList
    Dim gpAggregation As Geoprocessor = New Geoprocessor


    Private Sub frmAggregation_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        m_arrPolyFeatureLayers = Nothing
        m_frmFieldAggregation = Nothing
        blnOpening = Nothing
        pNeighborhoodLyr = Nothing
        arrWeightedSum = Nothing
        arrTotalSum = Nothing
        arrDevTypes = Nothing
        arrRedevRates = Nothing
        arrNetAcrePercents = Nothing
        arrAbadonRates = Nothing
        arrSelected = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub

    Private Sub frmAggregation_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        blnOpening = True
        Dim mxApplication As IMxApplication = Nothing
        Dim pMxDocument As IMxDocument = Nothing
        Dim mapCurrent As Map
        Dim uid As ESRI.ArcGIS.esriSystem.IUID = New ESRI.ArcGIS.esriSystem.UIDClass
        Dim enumLayer As ESRI.ArcGIS.Carto.IEnumLayer
        Dim pLyr As ILayer
        Dim pFeatLyr As IFeatureLayer
        Dim intLayer As Integer
        Dim intFeatCount As Integer = 0
        Dim pActiveView As IActiveView = Nothing
        Dim pFeatClass As IFeatureClass
        Dim intFld As Integer = 0
        Dim intFldCount As Integer = 20
        Dim pField As IField = Nothing
        Dim pFeatSelection As IFeatureSelection
        Dim pCursor As ICursor = Nothing
        Dim intCount As Integer
        Dim pDataset As IDataset
        Dim pFeatureCursor As IFeatureCursor = Nothing
        Dim pFeat As IFeature
        Dim intObjectId1 As Integer = 0
        Dim pQueryFilter As IQueryFilter
        Dim strQString As String = ""
        Dim GP As Geoprocessor
        Dim pCreateTable As ESRI.ArcGIS.DataManagementTools.CreateTable
        Dim pWksFactory As ESRI.ArcGIS.Geodatabase.IWorkspaceFactory = Nothing
        Dim pFeatWks As ESRI.ArcGIS.Geodatabase.IFeatureWorkspace = Nothing
        Dim rowTemp As IRow
        Dim strFieldName As String = ""
        Dim strFieldAlias As String = ""
        Dim intUseField As Integer = 0
        Dim intCalcAcres As Integer = 0
        Dim strCalcFieldName As String = ""
        Dim intCalcAcresOnly As Integer = 0

        'RETRIEVE CURRENT VIEW DOCUMENT TO OBTAIN LIST OF CURRENT LAYER(S)
        If Not TypeOf m_appEnvision Is IApplication Then
            GoTo CleanUp
        Else
            pMxDocument = DirectCast(m_appEnvision.Document, IMxDocument)
            mapCurrent = CType(pMxDocument.FocusMap, Map)
            pActiveView = CType(pMxDocument.FocusMap, IActiveView)
        End If


        'EXIT IF THE USER HAS NOT DEFINED A PARCEL LAYER
        If m_pEditFeatureLyr Is Nothing Then
            MessageBox.Show("Please select a Parcel layer to use this tool.", "Parcel Layer Not Defined", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo CleanUp
        Else
            'RETRIEVE SELECTED NEIGHBORHOOD LAYER
            pFeatSelection = CType(m_pEditFeatureLyr, IFeatureSelection)
            pFeatSelection.SelectionSet.Search(Nothing, False, pCursor)
            'Default to ALL features
            If pFeatSelection.SelectionSet.Count > 0 Then
                Me.chkParcelSelected.Visible = True
                Me.chkParcelSelected.Enabled = True
                Me.chkParcelSelected.Text = "Use " & pFeatSelection.SelectionSet.Count.ToString & " Selected Features"
            Else
                Me.chkParcelSelected.Visible = False
            End If
        End If

        'CHECK FOR AVAILABLE LAYER(S)
        If mapCurrent.LayerCount = 0 Then
            MessageBox.Show("Please a polygon layer to the current view document you would like to utilize in the process.", "No Layers Found", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            Me.Close()
            GoTo CleanUp
        End If

        m_arrPolyFeatureLayers = New ArrayList
        'CYCLE THROUGH LAYERS TO POPULATE ARRAYS AND CONTROLS
        'RETRIEVE THE FEATURE LAYERS TO POPULATE FEATURE LAYER OPTIONS
        uid.Value = "{E156D7E5-22AF-11D3-9F99-00C04F6BC78E}" '= IGeoFeatureLayer
        Try
            enumLayer = mapCurrent.Layers((CType(uid, ESRI.ArcGIS.esriSystem.UID)), True) ' Explicit Cast
            enumLayer.Reset()
            pLyr = enumLayer.Next
            Do While Not (pLyr Is Nothing)
                Try
                    pFeatLyr = CType(pLyr, IFeatureLayer)
                    pFeatClass = pFeatLyr.FeatureClass
                    If pFeatClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                        If Not pFeatLyr.Name = m_pEditFeatureLyr.Name Then
                            m_arrPolyFeatureLayers.Add(pFeatLyr)
                            Me.cmbLayers.Items.Add(pFeatLyr.Name)
                            intFeatCount = intFeatCount + 1
                        End If
                    End If
                Catch ex As Exception

                End Try
                pLyr = enumLayer.Next()
            Loop
        Catch ex As Exception
            GoTo CleanUp
        End Try

        'CLOSE IF NO POLYGON LAYERS FOUND
        If m_arrPolyFeatureLayers.Count <= 0 Then
            'MessageBox.Show("No polygon feature layer(s) could be found in the current view document.  Please load in a neighnorhood polygon layer to utilize these functions.", "No Polygon Layers Found", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            'Me.Close()
            'GoTo CleanUp
            Me.rdbParcel.Checked = True
            Me.rdbNeighborhood.Enabled = False
            Me.lblNeighborhoodLyr.Enabled = False
            Me.cmbLayers.Enabled = False
            Me.chkSelected.Visible = False
        End If

        'LOAD THE DEVELOPMENT TYPE ATTRIBUTES TABLE
        If m_tblAttribFields Is Nothing Then
            Try
                pWksFactory = New FileGDBWorkspaceFactory
                pFeatWks = pWksFactory.OpenFromFile(m_strFeaturePath, 0)
                m_tblAttribFields = pFeatWks.OpenTable("ENVISION_DEVTYPE_FIELD_TRACKING")
            Catch ex As Exception
                GP = New Geoprocessor
                pCreateTable = New ESRI.ArcGIS.DataManagementTools.CreateTable
                pCreateTable.out_name = "ENVISION_DEVTYPE_FIELD_TRACKING"
                pCreateTable.out_path = m_strFeaturePath
                RunTool(GP, pCreateTable)
                m_tblAttribFields = pFeatWks.OpenTable("ENVISION_DEVTYPE_FIELD_TRACKING")
            End Try
            If TypeOf m_tblAttribFields Is ITable Then
                LookUpTablesEnvisionAttributeFieldTrackingTblCheck(m_tblAttribFields)
            End If
        End If
        If m_tblAttribFields Is Nothing Then
            MessageBox.Show("The table, ENVISION_DEVTYPE_FIELD_TRACKING, could not be retrieved from the current Envision project.", "Required Table Not Found")
            GoTo CleanUp
        End If

        pCursor = m_tblAttribFields.Search(Nothing, False)
        rowTemp = pCursor.NextRow
        Do Until rowTemp Is Nothing
            strFieldName = ""
            strFieldAlias = ""
            intUseField = 0
            intCalcAcres = 0
            intCalcAcresOnly = 0
            strCalcFieldName = ""
            Try
                strFieldName = CStr(rowTemp.Value(rowTemp.Fields.FindField("FIELD_NAME")))
            Catch ex As Exception
            End Try
            Try
                strFieldAlias = CStr(rowTemp.Value(rowTemp.Fields.FindField("FIELD_ALIAS")))
            Catch ex As Exception
            End Try
            Try
                intUseField = CInt(rowTemp.Value(rowTemp.Fields.FindField("USE")))
            Catch ex As Exception
            End Try
            Try
                intCalcAcres = CInt(rowTemp.Value(rowTemp.Fields.FindField("CALC_BY_ACRES")))
            Catch ex As Exception
            End Try
            Try
                strCalcFieldName = CStr(rowTemp.Value(rowTemp.Fields.FindField("CALC_FIELD_NAME")))
            Catch ex As Exception
            End Try
            Try
                intCalcAcresOnly = CInt(rowTemp.Value(rowTemp.Fields.FindField("CALC_ONLY")))
            Catch ex As Exception
            End Try
            If intUseField = 1 Then
                If strFieldName.Length > 0 Then
                    If Not UCase(strFieldName) = "DEV_TYPE" And Not UCase(strFieldName) = "VAC_ACRES" And Not UCase(strFieldName) = "DEVD_ACRES" _
                        And Not UCase(strFieldName) = "CONSTRAINED_ACRE" And Not UCase(strFieldName) = "RED" And Not UCase(strFieldName) = "GREEN" _
                        And Not UCase(strFieldName) = "BLUE" Then
                        'CHECK FOR THE WEIGHTED AREA VARIABLES
                        If m_pEditFeatureLyr.FeatureClass.FindField(strFieldName) >= 0 Then
                            Me.dgvFields.Rows.Add()
                            If Not intUseField And intCalcAcresOnly Then
                                Me.dgvFields.Rows(Me.dgvFields.RowCount - 1).Cells(0).Value = 0
                            Else
                                If intCalcAcres = 1 Then
                                    Me.dgvFields.Rows(Me.dgvFields.RowCount - 1).Cells(0).Value = 0
                                Else
                                    Me.dgvFields.Rows(Me.dgvFields.RowCount - 1).Cells(0).Value = 1
                                End If
                            End If
                            Me.dgvFields.Rows(Me.dgvFields.RowCount - 1).Cells(1).Value = strFieldName
                            Me.dgvFields.Rows(Me.dgvFields.RowCount - 1).Cells(2).Value = strFieldAlias
                            Me.dgvFields.Rows(Me.dgvFields.RowCount - 1).Cells(3).Value = 1
                            Me.dgvFields.Rows(Me.dgvFields.RowCount - 1).Cells(4).Value = 0
                        End If
                        'CHECK FOR THE SUM VARIABLES
                        If intCalcAcres = 1 Then
                            If m_pEditFeatureLyr.FeatureClass.FindField(strCalcFieldName) >= 0 Then
                                Me.dgvFields.Rows.Add()
                                If intUseField Then
                                    'If intCalcAcresOnly Then
                                    Me.dgvFields.Rows(Me.dgvFields.RowCount - 1).Cells(0).Value = 1
                                    'Else
                                    '    Me.dgvFields.Rows(Me.dgvFields.RowCount - 1).Cells(0).Value = 0
                                    'End If
                                Else
                                    Me.dgvFields.Rows(Me.dgvFields.RowCount - 1).Cells(0).Value = 0
                                End If
                                Me.dgvFields.Rows(Me.dgvFields.RowCount - 1).Cells(1).Value = strCalcFieldName
                                Me.dgvFields.Rows(Me.dgvFields.RowCount - 1).Cells(2).Value = strFieldAlias & " (Calculate Sum)"
                                Me.dgvFields.Rows(Me.dgvFields.RowCount - 1).Cells(3).Value = 0
                                Me.dgvFields.Rows(Me.dgvFields.RowCount - 1).Cells(4).Value = 1
                            End If
                        End If
                        'CHECK FOR THE SUM VARIABLES
                        If intCalcAcres = 2 Then
                            If m_pEditFeatureLyr.FeatureClass.FindField(strCalcFieldName) >= 0 Then
                                Me.dgvFields.Rows.Add()
                                If intUseField Then
                                    'If intCalcAcresOnly Then
                                    Me.dgvFields.Rows(Me.dgvFields.RowCount - 1).Cells(0).Value = 1
                                    'Else
                                    '    Me.dgvFields.Rows(Me.dgvFields.RowCount - 1).Cells(0).Value = 0
                                    'End If
                                Else
                                    Me.dgvFields.Rows(Me.dgvFields.RowCount - 1).Cells(0).Value = 0
                                End If
                                Me.dgvFields.Rows(Me.dgvFields.RowCount - 1).Cells(1).Value = strCalcFieldName
                                Me.dgvFields.Rows(Me.dgvFields.RowCount - 1).Cells(2).Value = strFieldAlias & " (Dependent Var Sum)"
                                Me.dgvFields.Rows(Me.dgvFields.RowCount - 1).Cells(3).Value = 0
                                Me.dgvFields.Rows(Me.dgvFields.RowCount - 1).Cells(4).Value = 1
                            End If
                        End If
                    End If
                End If
            End If

            rowTemp = pCursor.NextRow
        Loop

        GoTo CleanUp

CleanUp:
        mxApplication = Nothing
        pMxDocument = Nothing
        mapCurrent = Nothing
        pFeatLyr = Nothing
        uid = Nothing
        enumLayer = Nothing
        pLyr = Nothing
        intLayer = Nothing
        intFeatCount = Nothing
        pActiveView = Nothing
        pFeatClass = Nothing
        intFld = Nothing
        intFldCount = Nothing
        pField = Nothing
        pFeatSelection = Nothing
        pCursor = Nothing
        intCount = Nothing
        pDataset = Nothing
        pFeatureCursor = Nothing
        pFeat = Nothing
        intObjectId1 = Nothing
        pQueryFilter = Nothing
        strQString = Nothing
        GP = Nothing
        pCreateTable = Nothing
        pWksFactory = Nothing
        pFeatWks = Nothing
        rowTemp = Nothing
        strFieldName = Nothing
        strFieldAlias = Nothing
        intUseField = Nothing
        intCalcAcres = Nothing
        strCalcFieldName = Nothing
        intCalcAcresOnly = Nothing
        GC.Collect()
        GC.WaitForPendingFinalizers()
    End Sub

    Private Sub cmbLayers_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbLayers.SelectedIndexChanged
        Dim pFeatLyr As IFeatureLayer
        Dim pFeatureClass As IFeatureClass = Nothing
        Dim pFeatSelection As IFeatureSelection
        Dim pCursor As ICursor = Nothing
        Dim intFld As Integer = 0
        Dim pField As IField

        'RETRIEVE SELECTED NEIGHBORHOOD LAYER
        If Me.cmbLayers.Text.Length > 0 Then
            m_strNeighborhoodLayerName = Me.cmbLayers.Text
            pFeatLyr = m_arrPolyFeatureLayers.Item(Me.cmbLayers.FindString(Me.cmbLayers.Text))
            pFeatureClass = pFeatLyr.FeatureClass
            pFeatSelection = CType(pFeatLyr, IFeatureSelection)
            pFeatSelection.SelectionSet.Search(Nothing, False, pCursor)
            'Default to ALL features
            If pFeatSelection.SelectionSet.Count > 0 Then
                Me.chkSelected.Visible = True
                Me.chkSelected.Enabled = True
                Me.chkSelected.Text = "Use " & pFeatSelection.SelectionSet.Count.ToString & " Selected Features"
            Else
                Me.chkSelected.Visible = False
            End If
        Else
            GoTo CleanUp
        End If

        LoadFields()
        GoTo CleanUp

CleanUp:
        pFeatLyr = Nothing
        pFeatureClass = Nothing
        pFeatSelection = Nothing
        pCursor = Nothing
        intFld = Nothing
        pField = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub

    Private Sub LoadFields()
        'POPULATE THE FIELD CONTROL WITH ALPHA NUMERIC FIELDS FROM THE SELECTED NEIGHBORHOOD LAYER
        Dim pFeatLyr As IFeatureLayer
        Dim pFeatureClass As IFeatureClass = Nothing
        Dim pFeatSelection As IFeatureSelection
        Dim pCursor As ICursor = Nothing
        Dim intFld As Integer = 0
        Dim pField As IField

        'RETRIEVE SELECTED NEIGHBORHOOD LAYER
        If Me.cmbLayers.Text.Length > 0 Then
            m_strNeighborhoodLayerName = Me.cmbLayers.Text
            pFeatLyr = m_arrPolyFeatureLayers.Item(Me.cmbLayers.FindString(Me.cmbLayers.Text))
            pFeatureClass = pFeatLyr.FeatureClass
            pFeatSelection = CType(pFeatLyr, IFeatureSelection)
            pFeatSelection.SelectionSet.Search(Nothing, False, pCursor)
        Else
            GoTo CleanUp
        End If

CleanUp:
        pFeatLyr = Nothing
        pFeatureClass = Nothing
        pFeatSelection = Nothing
        pCursor = Nothing
        intFld = Nothing
        pField = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub

    Private Sub cmbParcelFlds_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        SelectField()
    End Sub

    Private Sub SelectField()
        Dim pFeatSelection As IFeatureSelection
        Dim pCursor As ICursor = Nothing
        Dim pFeatLyr As IFeatureLayer
        Dim pFeatureClass As IFeatureClass = Nothing
        Dim intFld As Integer = 0
        Dim pField As IField
        Dim pFeatureCursor As IFeatureCursor = Nothing
        Dim pFeat As IFeature
        Dim strFieldName As String = ""

        'RETRIEVE SELECTED NEIGHBORHOOD LAYER
        If Me.cmbLayers.Text.Length > 0 Then
            m_strNeighborhoodLayerName = Me.cmbLayers.Text
            pFeatLyr = m_arrPolyFeatureLayers.Item(Me.cmbLayers.FindString(Me.cmbLayers.Text))
            pFeatureClass = pFeatLyr.FeatureClass
            pFeatSelection = CType(pFeatLyr, IFeatureSelection)
            pFeatSelection.SelectionSet.Search(Nothing, False, pCursor)
        Else
            GoTo CleanUp
        End If

CleanUp:
        pFeatSelection = Nothing
        pCursor = Nothing
        pFeatLyr = Nothing
        pFeatureClass = Nothing
        intFld = Nothing
        pField = Nothing
        pFeatureCursor = Nothing
        pFeat = Nothing
        strFieldName = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub

    Private Sub btnRun_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRun.Click
        Dim pFeatureClass As IFeatureClass = Nothing
        Dim pToTable As ITable
        Dim pFeatSelection As IFeatureSelection
        Dim pCursor As ICursor = Nothing
        Dim pFeatureCursor As IFeatureCursor = Nothing
        Dim pFeat As IFeature
        Dim arrFromFields As ArrayList = New ArrayList
        Dim arrToFields As ArrayList = New ArrayList
        Dim intFld As Integer
        Dim pFromField As IField
        Dim strFieldName As String = ""
        Dim strFieldAlias As String = ""
        Dim blnFromNameFound As Boolean
        Dim blnToNameFound As Boolean
        Dim pToField As IField
        Dim strPREname As String = ""
        Dim intNum As Integer = 0
        Dim intValue As Integer
        Dim strQString As String = ""
        Dim intTotalCount As Integer = 0
        Dim intRow As Integer = 0
        Dim intRedevFld As Integer = -1
        Dim rowTemp As IRow
        Dim strValue As String
        Dim dblValue As Double
        Dim strSearchFld As String = ""
        Dim intRec As Integer

        'EXIT IF THE USER HAS NOT DEFINED A PARCEL LAYER
        If m_pEditFeatureLyr Is Nothing Then
            MessageBox.Show("Please select a Parcel layer to use this tool.", "Parcel Layer Not Defined", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo CleanUp
        End If

        'BUILD LIST OF DEVELOPMENT TYPES AND CORRESPONDING REDEVELOPMENT RATES
        Try
            For intRow = 1 To 50
                rowTemp = m_tblDevelopmentTypes.GetRow(intRow)
                strValue = CStr(rowTemp.Value(m_tblDevelopmentTypes.FindField("DEVELOPMENT_TYPE")))
                If strValue.Length > 0 Then
                    Try
                        dblValue = CDbl(rowTemp.Value(m_tblDevelopmentTypes.FindField("REDEV_RATE")))
                    Catch ex As Exception
                        dblValue = 0
                    End Try
                    arrDevTypes.Add(strValue)
                    arrRedevRates.Add(dblValue)
                    Try
                        dblValue = CDbl(rowTemp.Value(m_tblDevelopmentTypes.FindField("ABANDON_RATE")))
                    Catch ex As Exception
                        dblValue = 0
                    End Try
                    arrAbadonRates.Add(dblValue)
                    Try
                        dblValue = CDbl(rowTemp.Value(m_tblDevelopmentTypes.FindField("NET_ACRE")))
                    Catch ex As Exception
                        dblValue = 0
                    End Try
                    arrNetAcrePercents.Add(dblValue)
                End If
            Next
        Catch ex As Exception
        End Try


        'RETIREVE THE SELECTED NEIGHBORHOOD LAYER
        If Me.rdbNeighborhood.Checked Then
            If Me.cmbLayers.Text.Length > 0 Then
                Try
                    m_strNeighborhoodLayerName = Me.cmbLayers.Text
                    pNeighborhoodLyr = m_arrPolyFeatureLayers.Item(Me.cmbLayers.FindString(Me.cmbLayers.Text))
                    pFeatureClass = pNeighborhoodLyr.FeatureClass
                    pToTable = CType(pFeatureClass, ITable)
                    pFeatSelection = CType(pNeighborhoodLyr, IFeatureSelection)
                    pFeatSelection.SelectionSet.Search(Nothing, False, pCursor)
                    pFeatureCursor = DirectCast(pCursor, IFeatureCursor)
                    pFeat = pFeatureCursor.NextFeature
                Catch ex As Exception
                    MessageBox.Show("Error in Neighborhood Aggregation:" & vbNewLine & ex.Message, "Neighborhood Error", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
                    GoTo CleanUp
                End Try
            Else
                MessageBox.Show("Please select a neighborhood layer.", "No Layer Selected", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
                GoTo CleanUp
            End If
        Else
            Try
                pNeighborhoodLyr = m_pEditFeatureLyr
                pFeatureClass = m_pEditFeatureLyr.FeatureClass
                pToTable = CType(pFeatureClass, ITable)
                pFeatSelection = CType(m_pEditFeatureLyr, IFeatureSelection)
                pFeatSelection.SelectionSet.Search(Nothing, False, pCursor)
                pFeatureCursor = DirectCast(pCursor, IFeatureCursor)
                pFeat = pFeatureCursor.NextFeature
            Catch ex As Exception
                MessageBox.Show("Error in Neighborhood Aggregation with Parcel Layer selection set:" & vbNewLine & ex.Message, "Neighborhood Error", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
                GoTo CleanUp
            End Try
        End If


        'CYCLE THROUGH THE SELECT FIELDS TO BE TRACKED, MAKING SURE THE INPUT FIELDS ARE PRESENT 
        'IN THE ENVISION PARCEL LAYER AS WELL AS THE OUTPUT FIELDS IN THE NEIGHBORHOOD LAYER
        '*******************************************************************************************************************
        'VAC_ACRES AND DEVD_ACRES ARE BY DEFAULT REQUIRED
        If m_pEditFeatureLyr.FeatureClass.Fields.FindField("VAC_ACRE") >= 0 Then
            If pFeatureClass.FindField("PRE_VAC_ACRE") = -1 Then
                AddEnvisionField(pToTable, "PRE_VAC_ACRE", "DOUBLE", 16, 6)
            End If
            If pFeatureClass.FindField("TOT_VAC_ACRE") = -1 Then
                AddEnvisionField(pToTable, "TOT_VAC_ACRE", "DOUBLE", 16, 6)
            End If
            If pFeatureClass.FindField("DEVD_VAC_ACRE") = -1 Then
                AddEnvisionField(pToTable, "DEVD_VAC_ACRE", "DOUBLE", 16, 6)
            End If
        End If
        If m_pEditFeatureLyr.FeatureClass.Fields.FindField("DEVD_ACRE") >= 0 Then
            If pFeatureClass.FindField("PRE_DEVD_ACRE") = -1 Then
                AddEnvisionField(pToTable, "PRE_DEVD_ACRE", "DOUBLE", 16, 6)
            End If
            If pFeatureClass.FindField("TOT_DEVD_ACRE") = -1 Then
                AddEnvisionField(pToTable, "TOT_DEVD_ACRE", "DOUBLE", 16, 6)
            End If
            If pFeatureClass.FindField("UNREDEVD_DEVD_ACRE") = -1 Then
                AddEnvisionField(pToTable, "UNREDEVD_DEVD_ACRE", "DOUBLE", 16, 6)
            End If
            If pFeatureClass.FindField("REDEVD_DEVD_ACRE") = -1 Then
                AddEnvisionField(pToTable, "REDEVD_DEVD_ACRE", "DOUBLE", 16, 6)
            End If
            If pFeatureClass.FindField("PRE_DEVD_SQ_MI") = -1 Then
                AddEnvisionField(pToTable, "PRE_DEVD_SQ_MI", "DOUBLE", 16, 6)
            End If
            If pFeatureClass.FindField("TOT_DEVD_SQ_MI") = -1 Then
                AddEnvisionField(pToTable, "TOT_DEVD_SQ_MI", "DOUBLE", 16, 6)
            End If
        End If

        '*******************************************************************************************************************
        'CYCLE THROUGH TRACKING LIST FOR REQUIRED INPUT AND OUTPUT FIELDS
        Try
            For intRow = 0 To Me.dgvFields.RowCount - 1
                If m_frmFieldAggregation.dgvFields.Rows(intRow).Cells(0).Value = 1 Then
                    strFieldName = Me.dgvFields.Rows(intRow).Cells(1).Value
                    If pFeatureClass.FindField("EX_PRE_" & strFieldName) = -1 Then
                        AddEnvisionField(pToTable, "EX_PRE_" & strFieldName, "DOUBLE", 16, 6)
                    End If
                    If pFeatureClass.FindField("EX_POST_" & strFieldName) = -1 Then
                        AddEnvisionField(pToTable, "EX_POST_" & strFieldName, "DOUBLE", 16, 6)
                    End If
                    If pFeatureClass.FindField("NEW_" & strFieldName) = -1 Then
                        AddEnvisionField(pToTable, "NEW_" & strFieldName, "DOUBLE", 16, 6)
                    End If
                    If pFeatureClass.FindField("TOT_" & strFieldName) = -1 Then
                        AddEnvisionField(pToTable, "TOT_" & strFieldName, "DOUBLE", 16, 6)
                    End If


                    If m_frmFieldAggregation.dgvFields.Rows(intRow).Cells(3).Value = 1 Then
                        arrWeightedSum.Add(strFieldName)
                    Else
                        arrTotalSum.Add(strFieldName)
                    End If
                End If
            Next
        Catch ex As Exception
            MessageBox.Show("Error in reviewing required fields:" & vbNewLine & ex.Message, "Processing Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo CleanUp
        End Try

        '*******************************************************************************************************************

        m_frmFieldAggregation.barStatus.Value = 0
        m_frmFieldAggregation.barStatus.Visible = True
        'RETRIEVE FIELD NAME FOR ID SEARCH
        strSearchFld = pFeatureClass.Fields.Field(0).Name


        'SELECT FEATURE(S) in PARCEL LAYER TO AGGREGATE
        If (Me.chkSelected.Checked Or Me.chkParcelSelected.Checked) And pFeatSelection.SelectionSet.Count > 0 Then
            Do While Not pFeat Is Nothing
                intValue = CInt(pFeat.Value(0))
                arrSelected.Add(intValue)
                pFeat = pFeatureCursor.NextFeature
            Loop
        Else
            intTotalCount = pFeatureClass.FeatureCount(Nothing)
            For intRec = 1 To intTotalCount
                arrSelected.Add(intRec)
            Next
        End If

        'ONLY RUN IF MORE THAN 1 FEATRURE IS FOUND
        Me.lblCount.Visible = True
        intValue = 0
        If arrSelected.Count > 0 Then
            For Each intRec In arrSelected
                intValue = intValue + 1
                strQString = strSearchFld & " = " & intRec.ToString
                Me.lblCount.Text = "Processing " & intValue.ToString & " of " & arrSelected.Count.ToString
                SelectedNeighborhoodFeature(strQString)
            Next
        End If

        MessageBox.Show("Field value aggregation has completed.", "Neighborhood Aggregation", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
        GoTo CloseForm

CloseForm:
        Me.barStatus.Value = 0
        Me.barStatus.Visible = False
        Me.Close()
        GoTo CleanUp

CleanUp:
        pFeatureClass = Nothing
        pToTable = Nothing
        pFeatSelection = Nothing
        pCursor = Nothing
        pFeatureCursor = Nothing
        pFeat = Nothing
        arrFromFields = Nothing
        arrToFields = Nothing
        intFld = Nothing
        pFromField = Nothing
        strFieldName = Nothing
        strFieldAlias = Nothing
        blnFromNameFound = Nothing
        blnToNameFound = Nothing
        pToField = Nothing
        strPREname = Nothing
        intNum = Nothing
        intValue = Nothing
        strQString = Nothing
        intTotalCount = Nothing
        intRec = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub

    Private Sub SelectedNeighborhoodFeature(ByVal strQString As String)
        Dim intRec As Integer = 1
        Dim strSearchFld As String = ""
        Dim pLocationSelect As ESRI.ArcGIS.DataManagementTools.SelectLayerByLocation
        Dim pNFeatSelection As IFeatureSelection = Nothing
        Dim pQueryFilter As IQueryFilter
        Dim pNFeatCursor As IFeatureCursor = Nothing
        Dim pNCursor As ICursor = Nothing
        Dim pFeatWrite As IFeature = Nothing
        Dim pEFeatSelection As IFeatureSelection = Nothing
        Dim pEFeatCursor As IFeatureCursor = Nothing
        Dim pECursor As ICursor = Nothing
        Dim pFeat As IFeature = Nothing
        Dim intTotalCount As Integer = 0
        Dim pParcelTable As ITable
        Dim intNum As Integer = 0
        Dim intRow As Integer = 0
        Dim dblPRE_VAC_ACRE As Double = 0
        Dim dblPRE_DEVD_ACRE As Double = 0
        Dim dblABANDON_DEVD_ACRE As Double = 0
        Dim dblDEVD_VAC_ACRE As Double = 0
        Dim dblUNREDEVD_DEVD_ACRE As Double = 0
        Dim dblREDEVD_DEVD_ACRE As Double = 0
        Dim dblTOT_VAC_ACRE As Double = 0
        Dim dblTOT_DEVD_ACRE As Double = 0
        Dim strDevType As String
        Dim dblDevdAcres As Double = 0
        Dim dblVacAcres As Double = 0
        Dim dblVar As Double
        Dim dblRedevRate As Double
        Dim dblAbandonRate As Double
        Dim dblNetAcre As Double
        Dim dblPRE As Double
        Dim dblPOST As Double
        Dim dblNEW As Double
        Dim dblTOT As Double
        Dim dblTemp As Double
        Dim intCount As Integer = 0
        Dim arrTotalDevdAcres As ArrayList = New ArrayList
        Dim arrTotalVacAcres As ArrayList = New ArrayList
        Dim arrTotalUnRedevdAcres As ArrayList = New ArrayList
        Dim arrTotalDevdVacAcres As ArrayList = New ArrayList
        Dim arrTotalReDevdAcres As ArrayList = New ArrayList

        Dim arrTempNumWeightedPreSum As ArrayList = New ArrayList
        Dim arrTempNumWeightedPostSum As ArrayList = New ArrayList
        Dim arrTempNumWeightedNewSum As ArrayList = New ArrayList
        Dim arrTempNumWeightedTotSum As ArrayList = New ArrayList
        Dim arrTempNumWeightedSum As ArrayList = New ArrayList
        Dim arrTempDomWeightedPreSum As ArrayList = New ArrayList
        Dim arrTempDomWeightedPostSum As ArrayList = New ArrayList
        Dim arrTempDomWeightedNewSum As ArrayList = New ArrayList
        Dim arrTempDomWeightedTotSum As ArrayList = New ArrayList
        Dim arrTempDomWeightedSum As ArrayList = New ArrayList

        Dim arrTempPreSum As ArrayList = New ArrayList
        Dim arrTempPostSum As ArrayList = New ArrayList
        Dim arrTempNewSum As ArrayList = New ArrayList
        Dim arrTempTotalSum As ArrayList = New ArrayList
        Dim strFieldName As String
        Dim intIndex As Integer
        Dim process As IGPProcess
        Dim refsLeft As Integer = 0
        Dim dblBaseAcres As Double = 0


        'RETRIEVE THE NEIGHBORHOOD LAYER AND SELECT THE CURRENT EDIT NEIGHBORHOOD 
        'ONLY EXECUTE FEATURE SELECTION IF A NEIGHBORHOOD LAYER IS DEFINED
        If Me.rdbNeighborhood.Checked Then
            pLocationSelect = New ESRI.ArcGIS.DataManagementTools.SelectLayerByLocation
            pLocationSelect.in_layer = m_pEditFeatureLyr
            pLocationSelect.overlap_type = "HAVE_THEIR_CENTER_IN"
            pLocationSelect.selection_type = "NEW_SELECTION"
            pLocationSelect.select_features = pNeighborhoodLyr
            process = pLocationSelect

            Try
                pNFeatSelection = CType(pNeighborhoodLyr, IFeatureSelection)
                pQueryFilter = New QueryFilter
                pQueryFilter.WhereClause = strQString
                pNFeatSelection.SelectFeatures(pQueryFilter, esriSelectionResultEnum.esriSelectionResultNew, False)
                pNFeatSelection.SelectionSet.Search(Nothing, False, pNCursor)
                pNFeatCursor = DirectCast(pNCursor, IFeatureCursor)
                pFeatWrite = pNFeatCursor.NextFeature
            Catch ex As Exception
                MessageBox.Show("Error in Neighborhood Aggregation:" & vbNewLine & ex.Message, "Neighborhood Error", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
                GoTo CleanUp
            End Try

            'RETRIEVE THE ENVISION EDIT LAYER AND EXECUTE OVERLAY SELECTION
            gpAggregation.Execute(process, Nothing)
            process = Nothing
            pLocationSelect = Nothing
            GC.WaitForFullGCComplete()
            GC.Collect()
            GC.WaitForPendingFinalizers()
            'Else
            '    Try
            '        pNFeatSelection = CType(m_pEditFeatureLyr, IFeatureSelection)
            '        pQueryFilter = New QueryFilter
            '        pQueryFilter.WhereClause = strQString

            '        pNFeatSelection.SelectionSet.Search(Nothing, False, pNCursor)
            '        pNFeatCursor = DirectCast(pNCursor, IFeatureCursor)
            '        pFeatWrite = pNFeatCursor.NextFeature
            '    Catch ex As Exception
            '        MessageBox.Show("Error in Neighborhood Aggregation:" & vbNewLine & ex.Message, "Neighborhood Error", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            '        GoTo CleanUp
            '    End Try
        End If

        Try
            pEFeatSelection = CType(m_pEditFeatureLyr, IFeatureSelection)
            If Me.rdbParcel.Checked Then
                pQueryFilter = New QueryFilter
                pQueryFilter.WhereClause = strQString
                pEFeatSelection.SelectFeatures(pQueryFilter, esriSelectionResultEnum.esriSelectionResultNew, False)
            End If
            pEFeatSelection.SelectionSet.Search(Nothing, False, pECursor)
            pEFeatCursor = DirectCast(pECursor, IFeatureCursor)
            pFeat = pEFeatCursor.NextFeature
            If Me.rdbParcel.Checked Then
                pFeatWrite = pFeat
            End If
        Catch ex As Exception
            MessageBox.Show("Error in Neighborhood Aggregation:" & vbNewLine & ex.Message, "Neighborhood Error", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo CleanUp
        End Try

        'EXIT IF NO FEATURES SELECTED
        If pEFeatSelection.SelectionSet.Count <= 0 Then
            'MessageBox.Show(pEFeatSelection.SelectionSet.Count.ToString)
        Else
            intTotalCount = pEFeatSelection.SelectionSet.Count
            pParcelTable = CType(m_pEditFeatureLyr.FeatureClass, ITable)
        End If

        'BUILD EMPTY FIELD VARIABLE LIST
        If arrWeightedSum.Count > 0 Then
            For intNum = 1 To arrWeightedSum.Count
                arrTempNumWeightedPreSum.Add(0)
                arrTempNumWeightedPostSum.Add(0)
                arrTempNumWeightedNewSum.Add(0)
                arrTempNumWeightedTotSum.Add(0)
                arrTempNumWeightedSum.Add(0)
                arrTempDomWeightedPreSum.Add(0)
                arrTempDomWeightedPostSum.Add(0)
                arrTempDomWeightedNewSum.Add(0)
                arrTempDomWeightedTotSum.Add(0)
                arrTempDomWeightedSum.Add(0)
                arrTotalDevdAcres.Add(0)
                arrTotalVacAcres.Add(0)
                arrTotalUnRedevdAcres.Add(0)
                arrTotalDevdVacAcres.Add(0)
                arrTotalReDevdAcres.Add(0)
            Next
        End If
        If arrTotalSum.Count > 0 Then
            For intNum = 1 To arrTotalSum.Count
                arrTempPreSum.Add(0)
                arrTempPostSum.Add(0)
                arrTempNewSum.Add(0)
                arrTempTotalSum.Add(0)
            Next
        End If

        'CYCLE THROUGH THE SELECTED FEATURE(S) AND SUM VARIABLES
        pEFeatSelection.SelectionSet.Search(Nothing, False, pECursor)
        pEFeatCursor = DirectCast(pECursor, IFeatureCursor)
        pFeat = pEFeatCursor.NextFeature
        intCount = 0
        Do While Not pFeat Is Nothing
            intCount = intCount + 1
            'RETRIEVE DEV TYPE
            Try
                strDevType = CStr(pFeat.Value(m_pEditFeatureLyr.FeatureClass.Fields.FindField("DEV_TYPE")))
            Catch ex As Exception
                strDevType = ""
            End Try
            'RETRIEVE THE REDEV RATE FROM THE DEVELOPMENT TYPES LOOKUP TABLE 
            Try
                'dblRedevRate = CDbl(pFeat.Value(m_pEditFeatureLyr.FeatureClass.Fields.FindField("REDEV_RATE")))
                intIndex = arrDevTypes.IndexOf(strDevType)
                dblRedevRate = CDbl(arrRedevRates.Item(intIndex))
                If dblRedevRate < 0 Then
                    dblRedevRate = 0
                End If
            Catch ex As Exception
                dblRedevRate = 0
            End Try
            'RETRIEVE THE ABANDOMENT RATE FROM THE DEVELOPMENT TYPES LOOKUP TABLE 
            Try
                dblAbandonRate = CDbl(arrAbadonRates.Item(intIndex))
                If dblAbandonRate < 0 Then
                    dblAbandonRate = 0
                End If
            Catch ex As Exception
                dblAbandonRate = 0
            End Try
            'RETRIEVE AND ADD VACANT ACRES VALUE
            Try
                dblVacAcres = CDbl(pFeat.Value(m_pEditFeatureLyr.FeatureClass.Fields.FindField("VAC_ACRE")))
            Catch ex As Exception
                dblVacAcres = 0
            End Try
            dblPRE_VAC_ACRE = dblPRE_VAC_ACRE + dblVacAcres
            'RETRIEVE AND ADD DEVELOPED ACRES VALUE
            Try
                dblDevdAcres = CDbl(pFeat.Value(m_pEditFeatureLyr.FeatureClass.Fields.FindField("DEVD_ACRE")))
            Catch ex As Exception
                dblDevdAcres = 0
            End Try
            dblPRE_DEVD_ACRE = dblPRE_DEVD_ACRE + dblDevdAcres

            'ADBANDONED DEVELOPED ACRES VALUE
            If dblAbandonRate > 0 Then
                dblABANDON_DEVD_ACRE = dblABANDON_DEVD_ACRE + (dblDevdAcres * dblAbandonRate)
            End If

            If strDevType.Length > 0 And dblAbandonRate = 0 Then
                'ADD DEVELOPED VACANT ACRES
                dblDEVD_VAC_ACRE = dblDEVD_VAC_ACRE + dblVacAcres
            End If

            'ADD TO UNREDEVELOPED DEVELOPED ACRES BUCKET
            dblUNREDEVD_DEVD_ACRE = dblUNREDEVD_DEVD_ACRE + ((dblDevdAcres * (1 - dblRedevRate)) - (dblDevdAcres * dblAbandonRate))

            'ADD TO REDEVELOPED DEVELOPED ACRES BUCKET
            dblREDEVD_DEVD_ACRE = dblREDEVD_DEVD_ACRE + (dblDevdAcres * dblRedevRate)

            'REVIEW THE WEIGHTED VALUES, IF GREATER THEN ZERO ADD TO WEIGHTED ACRES LIST
            If arrWeightedSum.Count > 0 Then
                For intNum = 0 To arrWeightedSum.Count - 1
                    strFieldName = arrWeightedSum.Item(intNum)
                    'RETRIEVE VARIABLE
                    Try
                        dblVar = CDbl(pFeat.Value(m_pEditFeatureLyr.FeatureClass.Fields.FindField(strFieldName)))
                    Catch ex As Exception
                        dblVar = 0
                    End Try
                    'RETRIEVE EXISTING VARIABLE
                    Try
                        dblPRE = CDbl(pFeat.Value(m_pEditFeatureLyr.FeatureClass.Fields.FindField("EX_" & strFieldName)))
                    Catch ex As Exception
                        dblPRE = 0
                    End Try


                    dblTemp = arrTotalVacAcres.Item(intNum)
                    arrTotalVacAcres.RemoveAt(intNum)
                    arrTotalVacAcres.Insert(intNum, (dblVacAcres + dblTemp))

                    dblTemp = arrTotalUnRedevdAcres.Item(intNum)
                    arrTotalUnRedevdAcres.RemoveAt(intNum)
                    arrTotalUnRedevdAcres.Insert(intNum, ((dblDevdAcres * (1 - (dblRedevRate + dblAbandonRate))) + dblTemp))


                    'ADD TO BASE ACRES IF GREATER THAN ZERO
                    If dblPRE > 0 Then
                        dblTemp = arrTotalDevdAcres.Item(intNum)
                        arrTotalDevdAcres.RemoveAt(intNum)
                        arrTotalDevdAcres.Insert(intNum, (dblDevdAcres + dblTemp))
                    End If
                    If dblVar > 0 Then
                        dblTemp = arrTotalDevdVacAcres.Item(intNum)
                        arrTotalDevdVacAcres.RemoveAt(intNum)
                        arrTotalDevdVacAcres.Insert(intNum, (dblVacAcres + dblTemp))

                        dblTemp = arrTotalReDevdAcres.Item(intNum)
                        arrTotalReDevdAcres.RemoveAt(intNum)
                        arrTotalReDevdAcres.Insert(intNum, ((dblDevdAcres * (dblRedevRate + dblAbandonRate)) + dblTemp))
                    End If
                Next
            End If

            Me.barStatus.Value = ((intCount / intTotalCount) * 100)
            Me.Refresh()

            pFeat = pEFeatCursor.NextFeature
        Loop

        'CALC AFTER BUCKETS CREATED
        dblTOT_VAC_ACRE = ((dblPRE_VAC_ACRE - dblDEVD_VAC_ACRE) + dblABANDON_DEVD_ACRE)
        dblTOT_DEVD_ACRE = (dblPRE_DEVD_ACRE - dblABANDON_DEVD_ACRE) + dblDEVD_VAC_ACRE


        'WRITE THE STORED VALUES
        Try
            pFeatWrite.Value(pFeatWrite.Fields.FindField("PRE_VAC_ACRE")) = dblPRE_VAC_ACRE
            pFeatWrite.Value(pFeatWrite.Fields.FindField("TOT_VAC_ACRE")) = dblTOT_VAC_ACRE 'dblTotalVacAcres - dblTotalDevdVacAcres
            pFeatWrite.Value(pFeatWrite.Fields.FindField("PRE_DEVD_ACRE")) = dblPRE_DEVD_ACRE 'dblTotalDevdAcres
            pFeatWrite.Value(pFeatWrite.Fields.FindField("UNREDEVD_DEVD_ACRE")) = dblUNREDEVD_DEVD_ACRE 'dblTotalUnDevdAcres
            pFeatWrite.Value(pFeatWrite.Fields.FindField("REDEVD_DEVD_ACRE")) = dblREDEVD_DEVD_ACRE 'dblTotalReDevdAcres
            pFeatWrite.Value(pFeatWrite.Fields.FindField("DEVD_VAC_ACRE")) = dblDEVD_VAC_ACRE 'dblTotalDevdVacAcres
            pFeatWrite.Value(pFeatWrite.Fields.FindField("TOT_DEVD_ACRE")) = dblTOT_DEVD_ACRE '(dblTotalDevdAcres + dblTotalDevdVacAcres)
            pFeatWrite.Value(pFeatWrite.Fields.FindField("PRE_DEVD_SQ_MI")) = (dblPRE_DEVD_ACRE * 0.0015625)
            pFeatWrite.Value(pFeatWrite.Fields.FindField("TOT_DEVD_SQ_MI")) = (dblTOT_DEVD_ACRE * 0.0015625)
            pFeatWrite.Store()
        Catch ex As Exception

        End Try

        'CYCLE THROUGH THE SELECTED FEATURE(S) AND SUM VARIABLES
        pEFeatSelection.SelectionSet.Search(Nothing, False, pECursor)
        pEFeatCursor = DirectCast(pECursor, IFeatureCursor)
        pFeat = pEFeatCursor.NextFeature
        intCount = 0
        Do While Not pFeat Is Nothing
            intCount = intCount + 1
            'RETRIEVE DEV TYPE
            Try
                strDevType = CStr(pFeat.Value(m_pEditFeatureLyr.FeatureClass.Fields.FindField("DEV_TYPE")))
            Catch ex As Exception
                strDevType = ""
            End Try
            'RETRIEVE THE REDEV RATE FROM THE DEVELOPMENT TYPES LOOKUP TABLE 
            intIndex = arrDevTypes.IndexOf(strDevType)
            Try
                'dblRedevRate = CDbl(pFeat.Value(m_pEditFeatureLyr.FeatureClass.Fields.FindField("REDEV_RATE")))
                dblRedevRate = CDbl(arrRedevRates.Item(intIndex))
                If dblRedevRate < 0 Then
                    dblRedevRate = 0
                End If
            Catch ex As Exception
                dblRedevRate = 0
            End Try
            'RETRIEVE THE ABANDOMENT RATE FROM THE DEVELOPMENT TYPES LOOKUP TABLE 
            Try
                dblAbandonRate = CDbl(arrAbadonRates.Item(intIndex))
                If dblAbandonRate < 0 Then
                    dblAbandonRate = 0
                End If
            Catch ex As Exception
                dblAbandonRate = 0
            End Try
            'RETRIEVE THE NET ACRES PRECENTAGE FROM THE DEVELOPMENT TYPES LOOKUP TABLE 
            Try
                dblNetAcre = CDbl(arrNetAcrePercents.Item(intIndex))
                If dblNetAcre < 0 Then
                    dblNetAcre = 0
                End If
            Catch ex As Exception
                dblNetAcre = 0
            End Try
            'RETRIEVE AND ADD VACANT ACRES VALUE
            Try
                dblVacAcres = CDbl(pFeat.Value(m_pEditFeatureLyr.FeatureClass.Fields.FindField("VAC_ACRE")))
            Catch ex As Exception
                dblVacAcres = 0
            End Try
            'dblTotalVacAcres = dblTotalVacAcres + dblVacAcres
            'RETRIEVE AND ADD DEVELOPED ACRES VALUE
            Try
                dblDevdAcres = CDbl(pFeat.Value(m_pEditFeatureLyr.FeatureClass.Fields.FindField("DEVD_ACRE")))
            Catch ex As Exception
                dblDevdAcres = 0
            End Try

            'WRITE THE STORED VALUES
            Try
                '====================================================
                'TOT_ Calculations
                If arrWeightedSum.Count > 0 Then
                    For intNum = 0 To arrWeightedSum.Count - 1
                        strFieldName = arrWeightedSum.Item(intNum)
                        'RETRIEVE VARIABLE
                        Try
                            dblVar = CDbl(pFeat.Value(m_pEditFeatureLyr.FeatureClass.Fields.FindField(strFieldName)))
                        Catch ex As Exception
                            dblVar = 0
                        End Try
                        'RETRIEVE EXISTING VARIABLE
                        Try
                            dblPRE = CDbl(pFeat.Value(m_pEditFeatureLyr.FeatureClass.Fields.FindField("EX_" & strFieldName)))
                        Catch ex As Exception
                            dblPRE = 0
                        End Try


                        'CALCULATE EX_PRE
                        If dblPRE > 0 Then
                            'NUMERATOR CALC
                            dblTemp = arrTempNumWeightedPreSum.Item(intNum)
                            arrTempNumWeightedPreSum.RemoveAt(intNum)
                            arrTempNumWeightedPreSum.Insert(intNum, ((dblDevdAcres * dblPRE) + dblTemp))
                            'NUMERATOR CALC
                            dblTemp = arrTempDomWeightedPreSum.Item(intNum)
                            arrTempDomWeightedPreSum.RemoveAt(intNum)
                            arrTempDomWeightedPreSum.Insert(intNum, (dblDevdAcres + dblTemp))
                        End If


                        'CALCULATE EX_POST
                        If dblPRE > 0 Then
                            If strDevType = "" Then
                                'NUMERATOR CALC 
                                dblTemp = arrTempNumWeightedPostSum.Item(intNum)
                                arrTempNumWeightedPostSum.RemoveAt(intNum)
                                arrTempNumWeightedPostSum.Insert(intNum, ((dblDevdAcres * dblPRE) + dblTemp))
                                'DENOMINATOR CALC 
                                dblTemp = arrTempDomWeightedPostSum.Item(intNum)
                                arrTempDomWeightedPostSum.RemoveAt(intNum)
                                arrTempDomWeightedPostSum.Insert(intNum, ((dblDevdAcres) + dblTemp))
                            Else
                                If Not (dblRedevRate + dblAbandonRate) = 0 And dblVar > 0 Then
                                    'NUMERATOR CALC 
                                    dblTemp = arrTempNumWeightedPostSum.Item(intNum)
                                    arrTempNumWeightedPostSum.RemoveAt(intNum)
                                    arrTempNumWeightedPostSum.Insert(intNum, (((dblDevdAcres * dblPRE) * (1 - (dblRedevRate + dblAbandonRate))) + dblTemp))
                                    'DENOMINATOR CALC 
                                    dblTemp = arrTempDomWeightedPostSum.Item(intNum)
                                    arrTempDomWeightedPostSum.RemoveAt(intNum)
                                    arrTempDomWeightedPostSum.Insert(intNum, ((dblDevdAcres * (1 - (dblRedevRate + dblAbandonRate))) + dblTemp))
                                ElseIf Not (dblRedevRate + dblAbandonRate) = 0 And dblVar = 0 Then
                                    'NUMERATOR CALC 
                                    dblTemp = arrTempNumWeightedPostSum.Item(intNum)
                                    arrTempNumWeightedPostSum.RemoveAt(intNum)
                                    arrTempNumWeightedPostSum.Insert(intNum, (((dblDevdAcres * dblPRE) * (1 - (dblRedevRate + dblAbandonRate))) + dblTemp))
                                    'DENOMINATOR CALC 
                                    dblTemp = arrTempDomWeightedPostSum.Item(intNum)
                                    arrTempDomWeightedPostSum.RemoveAt(intNum)
                                    arrTempDomWeightedPostSum.Insert(intNum, ((dblDevdAcres * (1 - (dblRedevRate + dblAbandonRate))) + dblTemp))
                                ElseIf (dblRedevRate + dblAbandonRate) = 0 And dblVar > 0 Then
                                    'NUMERATOR CALC 
                                    dblTemp = arrTempNumWeightedPostSum.Item(intNum)
                                    arrTempNumWeightedPostSum.RemoveAt(intNum)
                                    arrTempNumWeightedPostSum.Insert(intNum, ((dblDevdAcres * dblPRE) + dblTemp))
                                    'DENOMINATOR CALC 
                                    dblTemp = arrTempDomWeightedPostSum.Item(intNum)
                                    arrTempDomWeightedPostSum.RemoveAt(intNum)
                                    arrTempDomWeightedPostSum.Insert(intNum, ((dblDevdAcres) + dblTemp))
                                End If
                            End If
                        End If

                        'CALCULATE NEW_
                        If Not strDevType = "" And dblVar > 0 Then
                            If Not (dblRedevRate + dblAbandonRate) = 0 Then
                                'NUMERATOR CALC 
                                dblTemp = arrTempNumWeightedNewSum.Item(intNum)
                                arrTempNumWeightedNewSum.RemoveAt(intNum)
                                arrTempNumWeightedNewSum.Insert(intNum, (((dblDevdAcres * dblVar * (dblRedevRate + dblAbandonRate)) + (dblVacAcres * dblVar)) + dblTemp))
                                'DENOMINATOR CALC 
                                dblTemp = arrTempDomWeightedNewSum.Item(intNum)
                                arrTempDomWeightedNewSum.RemoveAt(intNum)
                                arrTempDomWeightedNewSum.Insert(intNum, (((dblDevdAcres * (dblRedevRate + dblAbandonRate)) + (dblVacAcres)) + dblTemp))
                            Else
                                'NUMERATOR CALC 
                                dblTemp = arrTempNumWeightedNewSum.Item(intNum)
                                arrTempNumWeightedNewSum.RemoveAt(intNum)
                                arrTempNumWeightedNewSum.Insert(intNum, ((dblVacAcres * dblVar) + dblTemp))
                                'DENOMINATOR CALC 
                                dblTemp = arrTempDomWeightedNewSum.Item(intNum)
                                arrTempDomWeightedNewSum.RemoveAt(intNum)
                                arrTempDomWeightedNewSum.Insert(intNum, (dblVacAcres + dblTemp))
                            End If
                        End If


                        'CALCULATE TOT_
                        If strDevType = "" Then
                            If dblPRE > 0 Then
                                'NUMERATOR CALC 
                                dblTemp = arrTempNumWeightedTotSum.Item(intNum)
                                arrTempNumWeightedTotSum.RemoveAt(intNum)
                                arrTempNumWeightedTotSum.Insert(intNum, ((dblDevdAcres * dblPRE) + dblTemp))
                                'DENOMINATOR CALC 
                                dblTemp = arrTempDomWeightedTotSum.Item(intNum)
                                arrTempDomWeightedTotSum.RemoveAt(intNum)
                                arrTempDomWeightedTotSum.Insert(intNum, (dblDevdAcres + dblTemp))
                            End If
                        Else
                            If Not (dblRedevRate + dblAbandonRate) = 0 Then
                                If dblVar > 0 Then
                                    If dblPRE > 0 Then
                                        'NUMERATOR CALC 
                                        dblTemp = arrTempNumWeightedTotSum.Item(intNum)
                                        arrTempNumWeightedTotSum.RemoveAt(intNum)
                                        arrTempNumWeightedTotSum.Insert(intNum, (((dblDevdAcres * dblVar * (dblRedevRate)) + (dblDevdAcres * dblPRE * (1 - (dblRedevRate + dblAbandonRate))) + (dblVacAcres * dblVar)) + dblTemp))

                                        'DENOMINATOR CALC 
                                        dblTemp = arrTempDomWeightedTotSum.Item(intNum)
                                        arrTempDomWeightedTotSum.RemoveAt(intNum)
                                        arrTempDomWeightedTotSum.Insert(intNum, ((dblDevdAcres * (dblRedevRate)) + (dblDevdAcres * (1 - (dblRedevRate + dblAbandonRate))) + (dblVacAcres) + dblTemp))
                                    End If
                                    If dblPRE = 0 Then
                                        'NUMERATOR CALC 
                                        dblTemp = arrTempNumWeightedTotSum.Item(intNum)
                                        arrTempNumWeightedTotSum.RemoveAt(intNum)
                                        arrTempNumWeightedTotSum.Insert(intNum, (((dblDevdAcres * dblVar * (dblRedevRate)) + (dblVacAcres * dblVar)) + dblTemp))

                                        'DENOMINATOR CALC 
                                        dblTemp = arrTempDomWeightedTotSum.Item(intNum)
                                        arrTempDomWeightedTotSum.RemoveAt(intNum)
                                        arrTempDomWeightedTotSum.Insert(intNum, ((dblDevdAcres * (dblRedevRate)) + (dblVacAcres) + dblTemp))
                                    End If
                                End If
                                If dblVar = 0 Then
                                    'NUMERATOR CALC 
                                    dblTemp = arrTempNumWeightedTotSum.Item(intNum)
                                    arrTempNumWeightedTotSum.RemoveAt(intNum)
                                    arrTempNumWeightedTotSum.Insert(intNum, ((dblDevdAcres * dblPRE * (1 - (dblRedevRate + dblAbandonRate))) + dblTemp))
                                    If dblPRE > 0 Then
                                        'DENOMINATOR CALC 
                                        dblTemp = arrTempDomWeightedTotSum.Item(intNum)
                                        arrTempDomWeightedTotSum.RemoveAt(intNum)
                                        arrTempDomWeightedTotSum.Insert(intNum, (((dblDevdAcres * (1 - (dblRedevRate + dblAbandonRate)))) + dblTemp))
                                    End If
                                End If
                            Else
                                If dblVar > 0 Then
                                    'NUMERATOR CALC 
                                    dblTemp = arrTempNumWeightedTotSum.Item(intNum)
                                    arrTempNumWeightedTotSum.RemoveAt(intNum)
                                    arrTempNumWeightedTotSum.Insert(intNum, (((dblDevdAcres * dblPRE) + (dblVacAcres * dblVar)) + dblTemp))
                                    'DENOMINATOR CALC 
                                    If dblPRE > 0 Then
                                        dblTemp = arrTempDomWeightedTotSum.Item(intNum)
                                        arrTempDomWeightedTotSum.RemoveAt(intNum)
                                        arrTempDomWeightedTotSum.Insert(intNum, ((dblDevdAcres + dblVacAcres) + dblTemp))
                                    Else
                                        dblTemp = arrTempDomWeightedTotSum.Item(intNum)
                                        arrTempDomWeightedTotSum.RemoveAt(intNum)
                                        arrTempDomWeightedTotSum.Insert(intNum, (dblVacAcres + dblTemp))
                                    End If
                                End If
                            End If
                        End If
                    Next
                End If
                '====================================================

                '====================================================
                'TOT_ Calculations
                If arrTotalSum.Count > 0 Then
                    For intNum = 1 To arrTotalSum.Count
                        strFieldName = arrTotalSum.Item(intNum - 1)
                        'RETRIEVE VARIABLE
                        Try
                            dblVar = CDbl(pFeat.Value(m_pEditFeatureLyr.FeatureClass.Fields.FindField(strFieldName)))
                        Catch ex As Exception
                            dblVar = 0
                        End Try
                        'RETRIEVE EXISTING VARIABLE
                        Try
                            dblPRE = CDbl(pFeat.Value(m_pEditFeatureLyr.FeatureClass.Fields.FindField("EX_" & strFieldName)))
                        Catch ex As Exception
                            dblPRE = 0
                        End Try

                        'CALCULATE THE POST VALUE
                        Try
                            dblPOST = dblPRE * (1 - (dblRedevRate + dblAbandonRate))
                        Catch ex As Exception
                            dblPOST = 0
                        End Try
                        ''CALCULATE THE NEW SUM
                        'Try
                        '    dblNEW = (dblVacAcres * dblNetAcre * dblVar) + (dblDevdAcres * (dblRedevRate + dblAbandonRate) * dblVar)
                        'Catch ex As Exception
                        '    dblNEW = 0
                        'End Try
                        'CALCULATE THE TOTAL SUM
                        Try
                            dblTOT = dblPOST + dblVar
                            ' dblTOT = dblPOST + dblNEW
                        Catch ex As Exception
                            dblTOT = 0
                        End Try

                        'CALCULATE THE AREA WEIGHTED SUM VALUE
                        If dblTOT > 0 Then
                            dblTemp = arrTempPreSum.Item(intNum - 1)
                            arrTempPreSum.RemoveAt(intNum - 1)
                            arrTempPreSum.Insert(intNum - 1, (dblPRE + dblTemp))
                            dblTemp = arrTempPostSum.Item(intNum - 1)
                            arrTempPostSum.RemoveAt(intNum - 1)
                            arrTempPostSum.Insert(intNum - 1, (dblPOST + dblTemp))
                            dblTemp = arrTempNewSum.Item(intNum - 1)
                            arrTempNewSum.RemoveAt(intNum - 1)
                            arrTempNewSum.Insert(intNum - 1, (dblVar + dblTemp))
                            dblTemp = arrTempTotalSum.Item(intNum - 1)
                            arrTempTotalSum.RemoveAt(intNum - 1)
                            arrTempTotalSum.Insert(intNum - 1, (dblTOT + dblTemp))
                        End If
                    Next
                End If
                '====================================================

            Catch ex As Exception
                MessageBox.Show("Error in writing to fields:" & vbNewLine & ex.Message, "Write Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
                GoTo CleanUp
            End Try


            Me.barStatus.Value = ((intCount / intTotalCount) * 100)
            Me.Refresh()

            pFeat = pEFeatCursor.NextFeature
        Loop

        'WRITE THE STORED VALUES
        Try
            If arrWeightedSum.Count > 0 Then
                For intNum = 1 To arrWeightedSum.Count
                    strFieldName = arrWeightedSum.Item(intNum - 1)
                    If pFeatWrite.Fields.FindField("EX_PRE_" & strFieldName) >= 0 Then
                        dblTemp = CDbl(arrTempNumWeightedPreSum.Item(intNum - 1)) / CDbl(arrTempDomWeightedPreSum.Item(intNum - 1))
                        'MessageBox.Show((arrTempNumWeightedPreSum.Item(intNum - 1)).ToString & vbNewLine & CDbl(arrTempDomWeightedPreSum.Item(intNum - 1).ToString), dblTemp.ToString)
                        If CDbl(arrTempDomWeightedPreSum.Item(intNum - 1)) = 0 Then
                            pFeatWrite.Value(pFeatWrite.Fields.FindField("EX_PRE_" & strFieldName)) = 0
                        Else
                            pFeatWrite.Value(pFeatWrite.Fields.FindField("EX_PRE_" & strFieldName)) = dblTemp
                        End If
                    End If
                    If pFeatWrite.Fields.FindField("EX_POST_" & strFieldName) >= 0 Then
                        dblTemp = CDbl(arrTempNumWeightedPostSum.Item(intNum - 1)) / CDbl(arrTempDomWeightedPostSum.Item(intNum - 1))
                        If CDbl(arrTempDomWeightedPostSum.Item(intNum - 1)) = 0 Then
                            pFeatWrite.Value(pFeatWrite.Fields.FindField("EX_POST_" & strFieldName)) = 0
                        Else
                            pFeatWrite.Value(pFeatWrite.Fields.FindField("EX_POST_" & strFieldName)) = dblTemp
                        End If
                    End If
                    If pFeatWrite.Fields.FindField("NEW_" & strFieldName) >= 0 Then
                        dblTemp = CDbl(arrTempNumWeightedNewSum.Item(intNum - 1)) / CDbl(arrTempDomWeightedNewSum.Item(intNum - 1))

                        If CDbl(arrTempDomWeightedNewSum.Item(intNum - 1)) = 0 Then
                            pFeatWrite.Value(pFeatWrite.Fields.FindField("NEW_" & strFieldName)) = 0
                        Else
                            pFeatWrite.Value(pFeatWrite.Fields.FindField("NEW_" & strFieldName)) = dblTemp
                        End If
                    End If
                    If pFeatWrite.Fields.FindField("TOT_" & strFieldName) >= 0 Then
                        'MessageBox.Show(CDbl(arrTempNumWeightedTotSum.Item(intNum - 1)).ToString, CDbl(arrTempDomWeightedTotSum.Item(intNum - 1)).ToString)
                        dblTemp = CDbl(arrTempNumWeightedTotSum.Item(intNum - 1)) / CDbl(arrTempDomWeightedTotSum.Item(intNum - 1))
                        If CDbl(arrTempDomWeightedTotSum.Item(intNum - 1)) = 0 Then
                            pFeatWrite.Value(pFeatWrite.Fields.FindField("TOT_" & strFieldName)) = 0
                        Else
                            pFeatWrite.Value(pFeatWrite.Fields.FindField("TOT_" & strFieldName)) = dblTemp
                        End If
                    End If
                Next
            End If

            If arrTotalSum.Count > 0 Then
                For intNum = 1 To arrTotalSum.Count
                    strFieldName = arrTotalSum.Item(intNum - 1)
                    If pFeatWrite.Fields.FindField("EX_PRE_" & strFieldName) >= 0 Then
                        pFeatWrite.Value(pFeatWrite.Fields.FindField("EX_PRE_" & strFieldName)) = CDbl(arrTempPreSum.Item(intNum - 1))
                    End If
                    If pFeatWrite.Fields.FindField("EX_POST_" & strFieldName) >= 0 Then
                        pFeatWrite.Value(pFeatWrite.Fields.FindField("EX_POST_" & strFieldName)) = CDbl(arrTempPostSum.Item(intNum - 1))
                    End If
                    If pFeatWrite.Fields.FindField("NEW_" & strFieldName) >= 0 Then
                        pFeatWrite.Value(pFeatWrite.Fields.FindField("NEW_" & strFieldName)) = CDbl(arrTempNewSum.Item(intNum - 1))
                    End If
                    If pFeatWrite.Fields.FindField("TOT_" & strFieldName) >= 0 Then
                        pFeatWrite.Value(pFeatWrite.Fields.FindField("TOT_" & strFieldName)) = CDbl(arrTempTotalSum.Item(intNum - 1))
                    End If
                Next
            End If

            pFeatWrite.Store()
            GoTo CleanUp
        Catch ex As Exception
            MessageBox.Show("Error in writing to fields:" & vbNewLine & ex.Message, "Write Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo CleanUp
        End Try

CleanUp:
        Try
            If Not Me.rdbParcel.Checked Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(pEFeatCursor)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(pNFeatCursor)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(pNCursor)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(pNFeatSelection)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(pECursor)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(pFeatWrite)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(pEFeatSelection)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        intRec = Nothing
        pLocationSelect = Nothing
        strQString = Nothing
        pNFeatSelection = Nothing
        pQueryFilter = Nothing
        pNFeatCursor = Nothing
        pNCursor = Nothing
        pFeatWrite = Nothing
        pEFeatSelection = Nothing
        pEFeatCursor = Nothing
        pECursor = Nothing
        pFeat = Nothing
        intTotalCount = Nothing
        pParcelTable = Nothing
        intNum = Nothing
        intRow = Nothing
        dblPRE_VAC_ACRE = Nothing
        dblPRE_DEVD_ACRE = Nothing
        dblABANDON_DEVD_ACRE = Nothing
        dblDEVD_VAC_ACRE = Nothing
        dblUNREDEVD_DEVD_ACRE = Nothing
        dblREDEVD_DEVD_ACRE = Nothing
        dblTOT_VAC_ACRE = Nothing
        dblTOT_DEVD_ACRE = Nothing
        strDevType = Nothing
        dblDevdAcres = Nothing
        dblVacAcres = Nothing
        dblVar = Nothing
        dblRedevRate = Nothing
        dblAbandonRate = Nothing
        dblNetAcre = Nothing
        dblPRE = Nothing
        dblPOST = Nothing
        dblNEW = Nothing
        dblTOT = Nothing
        dblTemp = Nothing
        intCount = Nothing
        arrTempPreSum = Nothing
        arrTempPostSum = Nothing
        arrTempNewSum = Nothing
        arrTempTotalSum = Nothing
        strFieldName = Nothing
        intIndex = Nothing
        GC.Collect()
        GC.WaitForFullGCComplete()
        GC.WaitForPendingFinalizers()
    End Sub

    Private Sub ProcessParcels(ByVal strQString As String)
        Dim intRec As Integer = 1
        Dim strSearchFld As String = ""
        Dim pLocationSelect As ESRI.ArcGIS.DataManagementTools.SelectLayerByLocation
        Dim pNFeatSelection As IFeatureSelection = Nothing
        Dim pQueryFilter As IQueryFilter
        Dim pNFeatCursor As IFeatureCursor = Nothing
        Dim pNCursor As ICursor = Nothing
        Dim pFeatWrite As IFeature = Nothing
        Dim pEFeatSelection As IFeatureSelection = Nothing
        Dim pEFeatCursor As IFeatureCursor = Nothing
        Dim pECursor As ICursor = Nothing
        Dim pFeat As IFeature = Nothing
        Dim intTotalCount As Integer = 0
        Dim pParcelTable As ITable
        Dim intNum As Integer = 0
        Dim intRow As Integer = 0
        Dim dblPRE_VAC_ACRE As Double = 0
        Dim dblPRE_DEVD_ACRE As Double = 0
        Dim dblABANDON_DEVD_ACRE As Double = 0
        Dim dblDEVD_VAC_ACRE As Double = 0
        Dim dblUNREDEVD_DEVD_ACRE As Double = 0
        Dim dblREDEVD_DEVD_ACRE As Double = 0
        Dim dblTOT_VAC_ACRE As Double = 0
        Dim dblTOT_DEVD_ACRE As Double = 0
        Dim strDevType As String
        Dim dblDevdAcres As Double = 0
        Dim dblVacAcres As Double = 0
        Dim dblVar As Double
        Dim dblRedevRate As Double
        Dim dblAbandonRate As Double
        Dim dblNetAcre As Double
        Dim dblPRE As Double
        Dim dblPOST As Double
        Dim dblNEW As Double
        Dim dblTOT As Double
        Dim dblTemp As Double
        Dim intCount As Integer = 0
        Dim arrTotalDevdAcres As ArrayList = New ArrayList
        Dim arrTotalVacAcres As ArrayList = New ArrayList
        Dim arrTotalUnRedevdAcres As ArrayList = New ArrayList
        Dim arrTotalDevdVacAcres As ArrayList = New ArrayList
        Dim arrTotalReDevdAcres As ArrayList = New ArrayList

        Dim arrTempNumWeightedPreSum As ArrayList = New ArrayList
        Dim arrTempNumWeightedPostSum As ArrayList = New ArrayList
        Dim arrTempNumWeightedNewSum As ArrayList = New ArrayList
        Dim arrTempNumWeightedTotSum As ArrayList = New ArrayList
        Dim arrTempNumWeightedSum As ArrayList = New ArrayList
        Dim arrTempDomWeightedPreSum As ArrayList = New ArrayList
        Dim arrTempDomWeightedPostSum As ArrayList = New ArrayList
        Dim arrTempDomWeightedNewSum As ArrayList = New ArrayList
        Dim arrTempDomWeightedTotSum As ArrayList = New ArrayList
        Dim arrTempDomWeightedSum As ArrayList = New ArrayList

        Dim arrTempPreSum As ArrayList = New ArrayList
        Dim arrTempPostSum As ArrayList = New ArrayList
        Dim arrTempNewSum As ArrayList = New ArrayList
        Dim arrTempTotalSum As ArrayList = New ArrayList
        Dim strFieldName As String
        Dim intIndex As Integer
        Dim process As IGPProcess
        Dim refsLeft As Integer = 0
        Dim dblBaseAcres As Double = 0

        If pNeighborhoodLyr Is Nothing Then
            GoTo CleanUp
        Else
            strSearchFld = pNeighborhoodLyr.FeatureClass.Fields.Field(0).Name
        End If

        pLocationSelect = New ESRI.ArcGIS.DataManagementTools.SelectLayerByLocation
        pLocationSelect.in_layer = m_pEditFeatureLyr
        pLocationSelect.overlap_type = "HAVE_THEIR_CENTER_IN"
        pLocationSelect.selection_type = "NEW_SELECTION"
        pLocationSelect.select_features = pNeighborhoodLyr
        process = pLocationSelect

        'RETRIEVE THE NEIGHBORHOOD LAYER AND SELECT THE CURRENT EDIT NEIGHBORHOOD 
        Try
            pNFeatSelection = CType(pNeighborhoodLyr, IFeatureSelection)
            pQueryFilter = New QueryFilter
            pQueryFilter.WhereClause = strQString
            pNFeatSelection.SelectFeatures(pQueryFilter, esriSelectionResultEnum.esriSelectionResultNew, False)
            pNFeatSelection.SelectionSet.Search(Nothing, False, pNCursor)
            pNFeatCursor = DirectCast(pNCursor, IFeatureCursor)
            pFeatWrite = pNFeatCursor.NextFeature
        Catch ex As Exception
            MessageBox.Show("Error in Neighborhood Aggregation:" & vbNewLine & ex.Message, "Neighborhood Error", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo CleanUp
        End Try

        If pNFeatSelection.SelectionSet.Count <= 0 Then
            MessageBox.Show(strQString)
        End If

        'RETRIEVE THE ENVISION EDIT LAYER AND EXECUTE OVERLAY SELECTION
        gpAggregation.Execute(process, Nothing)
        process = Nothing

        'RunAggregationTool(pLocationSelect)
        'pLocationSelect.in_layer = Nothing

        pLocationSelect.in_layer = Nothing
        GC.WaitForPendingFinalizers()
        GC.WaitForFullGCComplete()
        GC.Collect()
        Try
            pEFeatSelection = CType(m_pEditFeatureLyr, IFeatureSelection)
            pEFeatSelection.SelectionSet.Search(Nothing, False, pECursor)
            pEFeatCursor = DirectCast(pECursor, IFeatureCursor)
            pFeat = pEFeatCursor.NextFeature
        Catch ex As Exception
            MessageBox.Show("Error in Neighborhood Aggregation:" & vbNewLine & ex.Message, "Neighborhood Error", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo CleanUp
        End Try

        'EXIT IF NO FEATURES SELECTED
        If pEFeatSelection.SelectionSet.Count <= 0 Then
            'MessageBox.Show(pEFeatSelection.SelectionSet.Count.ToString)
        Else
            intTotalCount = pEFeatSelection.SelectionSet.Count
            pParcelTable = CType(m_pEditFeatureLyr.FeatureClass, ITable)
        End If

        'BUILD EMPTY FIELD VARIABLE LIST
        If arrWeightedSum.Count > 0 Then
            For intNum = 1 To arrWeightedSum.Count
                arrTempNumWeightedPreSum.Add(0)
                arrTempNumWeightedPostSum.Add(0)
                arrTempNumWeightedNewSum.Add(0)
                arrTempNumWeightedTotSum.Add(0)
                arrTempNumWeightedSum.Add(0)
                arrTempDomWeightedPreSum.Add(0)
                arrTempDomWeightedPostSum.Add(0)
                arrTempDomWeightedNewSum.Add(0)
                arrTempDomWeightedTotSum.Add(0)
                arrTempDomWeightedSum.Add(0)
                arrTotalDevdAcres.Add(0)
                arrTotalVacAcres.Add(0)
                arrTotalUnRedevdAcres.Add(0)
                arrTotalDevdVacAcres.Add(0)
                arrTotalReDevdAcres.Add(0)
            Next
        End If
        If arrTotalSum.Count > 0 Then
            For intNum = 1 To arrTotalSum.Count
                arrTempPreSum.Add(0)
                arrTempPostSum.Add(0)
                arrTempNewSum.Add(0)
                arrTempTotalSum.Add(0)
            Next
        End If

        'CYCLE THROUGH THE SELECTED FEATURE(S) AND SUM VARIABLES
        pEFeatSelection.SelectionSet.Search(Nothing, False, pECursor)
        pEFeatCursor = DirectCast(pECursor, IFeatureCursor)
        pFeat = pEFeatCursor.NextFeature
        intCount = 0
        Do While Not pFeat Is Nothing
            intCount = intCount + 1
            'RETRIEVE DEV TYPE
            Try
                strDevType = CStr(pFeat.Value(m_pEditFeatureLyr.FeatureClass.Fields.FindField("DEV_TYPE")))
            Catch ex As Exception
                strDevType = ""
            End Try
            'RETRIEVE THE REDEV RATE FROM THE DEVELOPMENT TYPES LOOKUP TABLE 
            Try
                'dblRedevRate = CDbl(pFeat.Value(m_pEditFeatureLyr.FeatureClass.Fields.FindField("REDEV_RATE")))
                intIndex = arrDevTypes.IndexOf(strDevType)
                dblRedevRate = CDbl(arrRedevRates.Item(intIndex))
                If dblRedevRate < 0 Then
                    dblRedevRate = 0
                End If
            Catch ex As Exception
                dblRedevRate = 0
            End Try
            'RETRIEVE THE ABANDOMENT RATE FROM THE DEVELOPMENT TYPES LOOKUP TABLE 
            Try
                dblAbandonRate = CDbl(arrAbadonRates.Item(intIndex))
                If dblAbandonRate < 0 Then
                    dblAbandonRate = 0
                End If
            Catch ex As Exception
                dblAbandonRate = 0
            End Try
            'RETRIEVE AND ADD VACANT ACRES VALUE
            Try
                dblVacAcres = CDbl(pFeat.Value(m_pEditFeatureLyr.FeatureClass.Fields.FindField("VAC_ACRE")))
            Catch ex As Exception
                dblVacAcres = 0
            End Try
            dblPRE_VAC_ACRE = dblPRE_VAC_ACRE + dblVacAcres
            'RETRIEVE AND ADD DEVELOPED ACRES VALUE
            Try
                dblDevdAcres = CDbl(pFeat.Value(m_pEditFeatureLyr.FeatureClass.Fields.FindField("DEVD_ACRE")))
            Catch ex As Exception
                dblDevdAcres = 0
            End Try
            dblPRE_DEVD_ACRE = dblPRE_DEVD_ACRE + dblDevdAcres

            'ADBANDONED DEVELOPED ACRES VALUE
            If dblAbandonRate > 0 Then
                dblABANDON_DEVD_ACRE = dblABANDON_DEVD_ACRE + (dblDevdAcres * dblAbandonRate)
            End If

            If strDevType.Length > 0 And dblAbandonRate = 0 Then
                'ADD DEVELOPED VACANT ACRES
                dblDEVD_VAC_ACRE = dblDEVD_VAC_ACRE + dblVacAcres
            End If

            'ADD TO UNREDEVELOPED DEVELOPED ACRES BUCKET
            dblUNREDEVD_DEVD_ACRE = dblUNREDEVD_DEVD_ACRE + ((dblDevdAcres * (1 - dblRedevRate)) - (dblDevdAcres * dblAbandonRate))

            'ADD TO REDEVELOPED DEVELOPED ACRES BUCKET
            dblREDEVD_DEVD_ACRE = dblREDEVD_DEVD_ACRE + (dblDevdAcres * dblRedevRate)

            'REVIEW THE WEIGHTED VALUES, IF GREATER THEN ZERO ADD TO WEIGHTED ACRES LIST
            If arrWeightedSum.Count > 0 Then
                For intNum = 0 To arrWeightedSum.Count - 1
                    strFieldName = arrWeightedSum.Item(intNum)
                    'RETRIEVE VARIABLE
                    Try
                        dblVar = CDbl(pFeat.Value(m_pEditFeatureLyr.FeatureClass.Fields.FindField(strFieldName)))
                    Catch ex As Exception
                        dblVar = 0
                    End Try
                    'RETRIEVE EXISTING VARIABLE
                    Try
                        dblPRE = CDbl(pFeat.Value(m_pEditFeatureLyr.FeatureClass.Fields.FindField("EX_" & strFieldName)))
                    Catch ex As Exception
                        dblPRE = 0
                    End Try


                    dblTemp = arrTotalVacAcres.Item(intNum)
                    arrTotalVacAcres.RemoveAt(intNum)
                    arrTotalVacAcres.Insert(intNum, (dblVacAcres + dblTemp))

                    dblTemp = arrTotalUnRedevdAcres.Item(intNum)
                    arrTotalUnRedevdAcres.RemoveAt(intNum)
                    arrTotalUnRedevdAcres.Insert(intNum, ((dblDevdAcres * (1 - (dblRedevRate + dblAbandonRate))) + dblTemp))


                    'ADD TO BASE ACRES IF GREATER THAN ZERO
                    If dblPRE > 0 Then
                        dblTemp = arrTotalDevdAcres.Item(intNum)
                        arrTotalDevdAcres.RemoveAt(intNum)
                        arrTotalDevdAcres.Insert(intNum, (dblDevdAcres + dblTemp))
                    End If
                    If dblVar > 0 Then
                        dblTemp = arrTotalDevdVacAcres.Item(intNum)
                        arrTotalDevdVacAcres.RemoveAt(intNum)
                        arrTotalDevdVacAcres.Insert(intNum, (dblVacAcres + dblTemp))

                        dblTemp = arrTotalReDevdAcres.Item(intNum)
                        arrTotalReDevdAcres.RemoveAt(intNum)
                        arrTotalReDevdAcres.Insert(intNum, ((dblDevdAcres * (dblRedevRate + dblAbandonRate)) + dblTemp))
                    End If
                Next
            End If

            Me.barStatus.Value = ((intCount / intTotalCount) * 100)
            Me.Refresh()

            pFeat = pEFeatCursor.NextFeature
        Loop

        'CALC AFTER BUCKETS CREATED
        dblTOT_VAC_ACRE = ((dblPRE_VAC_ACRE - dblDEVD_VAC_ACRE) + dblABANDON_DEVD_ACRE)
        dblTOT_DEVD_ACRE = (dblPRE_DEVD_ACRE - dblABANDON_DEVD_ACRE) + dblDEVD_VAC_ACRE

        'WRITE THE STORED VALUES
        Try
            pFeatWrite.Value(pFeatWrite.Fields.FindField("PRE_VAC_ACRE")) = dblPRE_VAC_ACRE
            pFeatWrite.Value(pFeatWrite.Fields.FindField("TOT_VAC_ACRE")) = dblTOT_VAC_ACRE 'dblTotalVacAcres - dblTotalDevdVacAcres
            pFeatWrite.Value(pFeatWrite.Fields.FindField("PRE_DEVD_ACRE")) = dblPRE_DEVD_ACRE 'dblTotalDevdAcres
            pFeatWrite.Value(pFeatWrite.Fields.FindField("UNREDEVD_DEVD_ACRE")) = dblUNREDEVD_DEVD_ACRE 'dblTotalUnDevdAcres
            pFeatWrite.Value(pFeatWrite.Fields.FindField("REDEVD_DEVD_ACRE")) = dblREDEVD_DEVD_ACRE 'dblTotalReDevdAcres
            pFeatWrite.Value(pFeatWrite.Fields.FindField("DEVD_VAC_ACRE")) = dblDEVD_VAC_ACRE 'dblTotalDevdVacAcres
            pFeatWrite.Value(pFeatWrite.Fields.FindField("TOT_DEVD_ACRE")) = dblTOT_DEVD_ACRE '(dblTotalDevdAcres + dblTotalDevdVacAcres)
            pFeatWrite.Value(pFeatWrite.Fields.FindField("PRE_DEVD_SQ_MI")) = (dblPRE_DEVD_ACRE * 0.0015625)
            pFeatWrite.Value(pFeatWrite.Fields.FindField("TOT_DEVD_SQ_MI")) = (dblTOT_DEVD_ACRE * 0.0015625)
            pFeatWrite.Store()
        Catch ex As Exception

        End Try


        'CYCLE THROUGH THE SELECTED FEATURE(S) AND SUM VARIABLES
        pEFeatSelection.SelectionSet.Search(Nothing, False, pECursor)
        pEFeatCursor = DirectCast(pECursor, IFeatureCursor)
        pFeat = pEFeatCursor.NextFeature
        intCount = 0
        Do While Not pFeat Is Nothing
            intCount = intCount + 1
            'RETRIEVE DEV TYPE
            Try
                strDevType = CStr(pFeat.Value(m_pEditFeatureLyr.FeatureClass.Fields.FindField("DEV_TYPE")))
            Catch ex As Exception
                strDevType = ""
            End Try
            'RETRIEVE THE REDEV RATE FROM THE DEVELOPMENT TYPES LOOKUP TABLE 
            intIndex = arrDevTypes.IndexOf(strDevType)
            Try
                'dblRedevRate = CDbl(pFeat.Value(m_pEditFeatureLyr.FeatureClass.Fields.FindField("REDEV_RATE")))
                dblRedevRate = CDbl(arrRedevRates.Item(intIndex))
                If dblRedevRate < 0 Then
                    dblRedevRate = 0
                End If
            Catch ex As Exception
                dblRedevRate = 0
            End Try
            'RETRIEVE THE ABANDOMENT RATE FROM THE DEVELOPMENT TYPES LOOKUP TABLE 
            Try
                dblAbandonRate = CDbl(arrAbadonRates.Item(intIndex))
                If dblAbandonRate < 0 Then
                    dblAbandonRate = 0
                End If
            Catch ex As Exception
                dblAbandonRate = 0
            End Try
            'RETRIEVE THE NET ACRES PRECENTAGE FROM THE DEVELOPMENT TYPES LOOKUP TABLE 
            Try
                dblNetAcre = CDbl(arrNetAcrePercents.Item(intIndex))
                If dblNetAcre < 0 Then
                    dblNetAcre = 0
                End If
            Catch ex As Exception
                dblNetAcre = 0
            End Try
            'RETRIEVE AND ADD VACANT ACRES VALUE
            Try
                dblVacAcres = CDbl(pFeat.Value(m_pEditFeatureLyr.FeatureClass.Fields.FindField("VAC_ACRE")))
            Catch ex As Exception
                dblVacAcres = 0
            End Try
            'dblTotalVacAcres = dblTotalVacAcres + dblVacAcres
            'RETRIEVE AND ADD DEVELOPED ACRES VALUE
            Try
                dblDevdAcres = CDbl(pFeat.Value(m_pEditFeatureLyr.FeatureClass.Fields.FindField("DEVD_ACRE")))
            Catch ex As Exception
                dblDevdAcres = 0
            End Try

            'WRITE THE STORED VALUES
            Try
                '====================================================
                'TOT_ Calculations
                If arrWeightedSum.Count > 0 Then
                    For intNum = 0 To arrWeightedSum.Count - 1
                        strFieldName = arrWeightedSum.Item(intNum)
                        'RETRIEVE VARIABLE
                        Try
                            dblVar = CDbl(pFeat.Value(m_pEditFeatureLyr.FeatureClass.Fields.FindField(strFieldName)))
                        Catch ex As Exception
                            dblVar = 0
                        End Try
                        'RETRIEVE EXISTING VARIABLE
                        Try
                            dblPRE = CDbl(pFeat.Value(m_pEditFeatureLyr.FeatureClass.Fields.FindField("EX_" & strFieldName)))
                        Catch ex As Exception
                            dblPRE = 0
                        End Try


                        'CALCULATE EX_PRE
                        If dblPRE > 0 Then
                            'NUMERATOR CALC
                            dblTemp = arrTempNumWeightedPreSum.Item(intNum)
                            arrTempNumWeightedPreSum.RemoveAt(intNum)
                            arrTempNumWeightedPreSum.Insert(intNum, ((dblDevdAcres * dblPRE) + dblTemp))
                            'NUMERATOR CALC
                            dblTemp = arrTempDomWeightedPreSum.Item(intNum)
                            arrTempDomWeightedPreSum.RemoveAt(intNum)
                            arrTempDomWeightedPreSum.Insert(intNum, (dblDevdAcres + dblTemp))
                        End If


                        'CALCULATE EX_POST
                        If dblPRE > 0 Then
                            If strDevType = "" Then
                                'NUMERATOR CALC 
                                dblTemp = arrTempNumWeightedPostSum.Item(intNum)
                                arrTempNumWeightedPostSum.RemoveAt(intNum)
                                arrTempNumWeightedPostSum.Insert(intNum, ((dblDevdAcres * dblPRE) + dblTemp))
                                'DENOMINATOR CALC 
                                dblTemp = arrTempDomWeightedPostSum.Item(intNum)
                                arrTempDomWeightedPostSum.RemoveAt(intNum)
                                arrTempDomWeightedPostSum.Insert(intNum, ((dblDevdAcres) + dblTemp))
                            Else
                                If Not (dblRedevRate + dblAbandonRate) = 0 And dblVar > 0 Then
                                    'NUMERATOR CALC 
                                    dblTemp = arrTempNumWeightedPostSum.Item(intNum)
                                    arrTempNumWeightedPostSum.RemoveAt(intNum)
                                    arrTempNumWeightedPostSum.Insert(intNum, (((dblDevdAcres * dblPRE) * (1 - (dblRedevRate + dblAbandonRate))) + dblTemp))
                                    'DENOMINATOR CALC 
                                    dblTemp = arrTempDomWeightedPostSum.Item(intNum)
                                    arrTempDomWeightedPostSum.RemoveAt(intNum)
                                    arrTempDomWeightedPostSum.Insert(intNum, ((dblDevdAcres * (1 - (dblRedevRate + dblAbandonRate))) + dblTemp))
                                ElseIf Not (dblRedevRate + dblAbandonRate) = 0 And dblVar = 0 Then
                                    'NUMERATOR CALC 
                                    dblTemp = arrTempNumWeightedPostSum.Item(intNum)
                                    arrTempNumWeightedPostSum.RemoveAt(intNum)
                                    arrTempNumWeightedPostSum.Insert(intNum, (((dblDevdAcres * dblPRE) * (1 - (dblRedevRate + dblAbandonRate))) + dblTemp))
                                    'DENOMINATOR CALC 
                                    dblTemp = arrTempDomWeightedPostSum.Item(intNum)
                                    arrTempDomWeightedPostSum.RemoveAt(intNum)
                                    arrTempDomWeightedPostSum.Insert(intNum, ((dblDevdAcres * (1 - (dblRedevRate + dblAbandonRate))) + dblTemp))
                                ElseIf (dblRedevRate + dblAbandonRate) = 0 And dblVar > 0 Then
                                    'NUMERATOR CALC 
                                    dblTemp = arrTempNumWeightedPostSum.Item(intNum)
                                    arrTempNumWeightedPostSum.RemoveAt(intNum)
                                    arrTempNumWeightedPostSum.Insert(intNum, ((dblDevdAcres * dblPRE) + dblTemp))
                                    'DENOMINATOR CALC 
                                    dblTemp = arrTempDomWeightedPostSum.Item(intNum)
                                    arrTempDomWeightedPostSum.RemoveAt(intNum)
                                    arrTempDomWeightedPostSum.Insert(intNum, ((dblDevdAcres) + dblTemp))
                                End If
                            End If
                        End If

                        'CALCULATE NEW_
                        If Not strDevType = "" And dblVar > 0 Then
                            If Not (dblRedevRate + dblAbandonRate) = 0 Then
                                'NUMERATOR CALC 
                                dblTemp = arrTempNumWeightedNewSum.Item(intNum)
                                arrTempNumWeightedNewSum.RemoveAt(intNum)
                                arrTempNumWeightedNewSum.Insert(intNum, (((dblDevdAcres * dblVar * (dblRedevRate + dblAbandonRate)) + (dblVacAcres * dblVar)) + dblTemp))
                                'DENOMINATOR CALC 
                                dblTemp = arrTempDomWeightedNewSum.Item(intNum)
                                arrTempDomWeightedNewSum.RemoveAt(intNum)
                                arrTempDomWeightedNewSum.Insert(intNum, (((dblDevdAcres * (dblRedevRate + dblAbandonRate)) + (dblVacAcres)) + dblTemp))
                            Else
                                'NUMERATOR CALC 
                                dblTemp = arrTempNumWeightedNewSum.Item(intNum)
                                arrTempNumWeightedNewSum.RemoveAt(intNum)
                                arrTempNumWeightedNewSum.Insert(intNum, ((dblVacAcres * dblVar) + dblTemp))
                                'DENOMINATOR CALC 
                                dblTemp = arrTempDomWeightedNewSum.Item(intNum)
                                arrTempDomWeightedNewSum.RemoveAt(intNum)
                                arrTempDomWeightedNewSum.Insert(intNum, (dblVacAcres + dblTemp))
                            End If
                        End If


                        'CALCULATE TOT_
                        If strDevType = "" Then
                            If dblPRE > 0 Then
                                'NUMERATOR CALC 
                                dblTemp = arrTempNumWeightedTotSum.Item(intNum)
                                arrTempNumWeightedTotSum.RemoveAt(intNum)
                                arrTempNumWeightedTotSum.Insert(intNum, ((dblDevdAcres * dblPRE) + dblTemp))
                                'DENOMINATOR CALC 
                                dblTemp = arrTempDomWeightedTotSum.Item(intNum)
                                arrTempDomWeightedTotSum.RemoveAt(intNum)
                                arrTempDomWeightedTotSum.Insert(intNum, (dblDevdAcres + dblTemp))
                            End If
                        Else
                            If Not (dblRedevRate + dblAbandonRate) = 0 Then
                                If dblVar > 0 Then
                                    If dblPRE > 0 Then
                                        'NUMERATOR CALC 
                                        dblTemp = arrTempNumWeightedTotSum.Item(intNum)
                                        arrTempNumWeightedTotSum.RemoveAt(intNum)
                                        arrTempNumWeightedTotSum.Insert(intNum, (((dblDevdAcres * dblVar * (dblRedevRate)) + (dblDevdAcres * dblPRE * (1 - (dblRedevRate + dblAbandonRate))) + (dblVacAcres * dblVar)) + dblTemp))

                                        'DENOMINATOR CALC 
                                        dblTemp = arrTempDomWeightedTotSum.Item(intNum)
                                        arrTempDomWeightedTotSum.RemoveAt(intNum)
                                        arrTempDomWeightedTotSum.Insert(intNum, ((dblDevdAcres * (dblRedevRate)) + (dblDevdAcres * (1 - (dblRedevRate + dblAbandonRate))) + (dblVacAcres) + dblTemp))
                                    End If
                                    If dblPRE = 0 Then
                                        'NUMERATOR CALC 
                                        dblTemp = arrTempNumWeightedTotSum.Item(intNum)
                                        arrTempNumWeightedTotSum.RemoveAt(intNum)
                                        arrTempNumWeightedTotSum.Insert(intNum, (((dblDevdAcres * dblVar * (dblRedevRate)) + (dblVacAcres * dblVar)) + dblTemp))

                                        'DENOMINATOR CALC 
                                        dblTemp = arrTempDomWeightedTotSum.Item(intNum)
                                        arrTempDomWeightedTotSum.RemoveAt(intNum)
                                        arrTempDomWeightedTotSum.Insert(intNum, ((dblDevdAcres * (dblRedevRate)) + (dblVacAcres) + dblTemp))
                                    End If
                                End If
                                If dblVar = 0 Then
                                    'NUMERATOR CALC 
                                    dblTemp = arrTempNumWeightedTotSum.Item(intNum)
                                    arrTempNumWeightedTotSum.RemoveAt(intNum)
                                    arrTempNumWeightedTotSum.Insert(intNum, ((dblDevdAcres * dblPRE * (1 - (dblRedevRate + dblAbandonRate))) + dblTemp))
                                    If dblPRE > 0 Then
                                        'DENOMINATOR CALC 
                                        dblTemp = arrTempDomWeightedTotSum.Item(intNum)
                                        arrTempDomWeightedTotSum.RemoveAt(intNum)
                                        arrTempDomWeightedTotSum.Insert(intNum, (((dblDevdAcres * (1 - (dblRedevRate + dblAbandonRate)))) + dblTemp))
                                    End If
                                End If
                            Else
                                If dblVar > 0 Then
                                    'NUMERATOR CALC 
                                    dblTemp = arrTempNumWeightedTotSum.Item(intNum)
                                    arrTempNumWeightedTotSum.RemoveAt(intNum)
                                    arrTempNumWeightedTotSum.Insert(intNum, (((dblDevdAcres * dblPRE) + (dblVacAcres * dblVar)) + dblTemp))
                                    'DENOMINATOR CALC 
                                    If dblPRE > 0 Then
                                        dblTemp = arrTempDomWeightedTotSum.Item(intNum)
                                        arrTempDomWeightedTotSum.RemoveAt(intNum)
                                        arrTempDomWeightedTotSum.Insert(intNum, ((dblDevdAcres + dblVacAcres) + dblTemp))
                                    Else
                                        dblTemp = arrTempDomWeightedTotSum.Item(intNum)
                                        arrTempDomWeightedTotSum.RemoveAt(intNum)
                                        arrTempDomWeightedTotSum.Insert(intNum, (dblVacAcres + dblTemp))
                                    End If
                                End If
                            End If
                        End If
                    Next
                End If
                '====================================================

                '====================================================
                'TOT_ Calculations
                If arrTotalSum.Count > 0 Then
                    For intNum = 1 To arrTotalSum.Count
                        strFieldName = arrTotalSum.Item(intNum - 1)
                        'RETRIEVE VARIABLE
                        Try
                            dblVar = CDbl(pFeat.Value(m_pEditFeatureLyr.FeatureClass.Fields.FindField(strFieldName)))
                        Catch ex As Exception
                            dblVar = 0
                        End Try
                        'RETRIEVE EXISTING VARIABLE
                        Try
                            dblPRE = CDbl(pFeat.Value(m_pEditFeatureLyr.FeatureClass.Fields.FindField("EX_" & strFieldName)))
                        Catch ex As Exception
                            dblPRE = 0
                        End Try

                        'CALCULATE THE POST VALUE
                        Try
                            dblPOST = dblPRE * (1 - (dblRedevRate + dblAbandonRate))
                        Catch ex As Exception
                            dblPOST = 0
                        End Try
                        ''CALCULATE THE NEW SUM
                        'Try
                        '    dblNEW = (dblVacAcres * dblNetAcre * dblVar) + (dblDevdAcres * (dblRedevRate + dblAbandonRate) * dblVar)
                        'Catch ex As Exception
                        '    dblNEW = 0
                        'End Try
                        'CALCULATE THE TOTAL SUM
                        Try
                            dblTOT = dblPOST + dblVar
                            ' dblTOT = dblPOST + dblNEW
                        Catch ex As Exception
                            dblTOT = 0
                        End Try

                        'CALCULATE THE AREA WEIGHTED SUM VALUE
                        If dblTOT > 0 Then
                            dblTemp = arrTempPreSum.Item(intNum - 1)
                            arrTempPreSum.RemoveAt(intNum - 1)
                            arrTempPreSum.Insert(intNum - 1, (dblPRE + dblTemp))
                            dblTemp = arrTempPostSum.Item(intNum - 1)
                            arrTempPostSum.RemoveAt(intNum - 1)
                            arrTempPostSum.Insert(intNum - 1, (dblPOST + dblTemp))
                            dblTemp = arrTempNewSum.Item(intNum - 1)
                            arrTempNewSum.RemoveAt(intNum - 1)
                            arrTempNewSum.Insert(intNum - 1, (dblVar + dblTemp))
                            dblTemp = arrTempTotalSum.Item(intNum - 1)
                            arrTempTotalSum.RemoveAt(intNum - 1)
                            arrTempTotalSum.Insert(intNum - 1, (dblTOT + dblTemp))
                        End If
                    Next
                End If
                '====================================================

            Catch ex As Exception
                MessageBox.Show("Error in writing to fields:" & vbNewLine & ex.Message, "Write Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
                GoTo CleanUp
            End Try


            Me.barStatus.Value = ((intCount / intTotalCount) * 100)
            Me.Refresh()

            pFeat = pEFeatCursor.NextFeature
        Loop

        'WRITE THE STORED VALUES
        Try
            If arrWeightedSum.Count > 0 Then
                For intNum = 1 To arrWeightedSum.Count
                    strFieldName = arrWeightedSum.Item(intNum - 1)
                    If pFeatWrite.Fields.FindField("EX_PRE_" & strFieldName) >= 0 Then
                        dblTemp = CDbl(arrTempNumWeightedPreSum.Item(intNum - 1)) / CDbl(arrTempDomWeightedPreSum.Item(intNum - 1))
                        'MessageBox.Show((arrTempNumWeightedPreSum.Item(intNum - 1)).ToString & vbNewLine & CDbl(arrTempDomWeightedPreSum.Item(intNum - 1).ToString), dblTemp.ToString)
                        If CDbl(arrTempDomWeightedPreSum.Item(intNum - 1)) = 0 Then
                            pFeatWrite.Value(pFeatWrite.Fields.FindField("EX_PRE_" & strFieldName)) = 0
                        Else
                            pFeatWrite.Value(pFeatWrite.Fields.FindField("EX_PRE_" & strFieldName)) = dblTemp
                        End If
                    End If
                    If pFeatWrite.Fields.FindField("EX_POST_" & strFieldName) >= 0 Then
                        dblTemp = CDbl(arrTempNumWeightedPostSum.Item(intNum - 1)) / CDbl(arrTempDomWeightedPostSum.Item(intNum - 1))
                        If CDbl(arrTempDomWeightedPostSum.Item(intNum - 1)) = 0 Then
                            pFeatWrite.Value(pFeatWrite.Fields.FindField("EX_POST_" & strFieldName)) = 0
                        Else
                            pFeatWrite.Value(pFeatWrite.Fields.FindField("EX_POST_" & strFieldName)) = dblTemp
                        End If
                    End If
                    If pFeatWrite.Fields.FindField("NEW_" & strFieldName) >= 0 Then
                        dblTemp = CDbl(arrTempNumWeightedNewSum.Item(intNum - 1)) / CDbl(arrTempDomWeightedNewSum.Item(intNum - 1))

                        If CDbl(arrTempDomWeightedNewSum.Item(intNum - 1)) = 0 Then
                            pFeatWrite.Value(pFeatWrite.Fields.FindField("NEW_" & strFieldName)) = 0
                        Else
                            pFeatWrite.Value(pFeatWrite.Fields.FindField("NEW_" & strFieldName)) = dblTemp
                        End If
                    End If
                    If pFeatWrite.Fields.FindField("TOT_" & strFieldName) >= 0 Then
                        'MessageBox.Show(CDbl(arrTempNumWeightedTotSum.Item(intNum - 1)).ToString, CDbl(arrTempDomWeightedTotSum.Item(intNum - 1)).ToString)
                        dblTemp = CDbl(arrTempNumWeightedTotSum.Item(intNum - 1)) / CDbl(arrTempDomWeightedTotSum.Item(intNum - 1))
                        If CDbl(arrTempDomWeightedTotSum.Item(intNum - 1)) = 0 Then
                            pFeatWrite.Value(pFeatWrite.Fields.FindField("TOT_" & strFieldName)) = 0
                        Else
                            pFeatWrite.Value(pFeatWrite.Fields.FindField("TOT_" & strFieldName)) = dblTemp
                        End If
                    End If
                Next
            End If

            If arrTotalSum.Count > 0 Then
                For intNum = 1 To arrTotalSum.Count
                    strFieldName = arrTotalSum.Item(intNum - 1)
                    If pFeatWrite.Fields.FindField("EX_PRE_" & strFieldName) >= 0 Then
                        pFeatWrite.Value(pFeatWrite.Fields.FindField("EX_PRE_" & strFieldName)) = CDbl(arrTempPreSum.Item(intNum - 1))
                    End If
                    If pFeatWrite.Fields.FindField("EX_POST_" & strFieldName) >= 0 Then
                        pFeatWrite.Value(pFeatWrite.Fields.FindField("EX_POST_" & strFieldName)) = CDbl(arrTempPostSum.Item(intNum - 1))
                    End If
                    If pFeatWrite.Fields.FindField("NEW_" & strFieldName) >= 0 Then
                        pFeatWrite.Value(pFeatWrite.Fields.FindField("NEW_" & strFieldName)) = CDbl(arrTempNewSum.Item(intNum - 1))
                    End If
                    If pFeatWrite.Fields.FindField("TOT_" & strFieldName) >= 0 Then
                        pFeatWrite.Value(pFeatWrite.Fields.FindField("TOT_" & strFieldName)) = CDbl(arrTempTotalSum.Item(intNum - 1))
                    End If
                Next
            End If

            pFeatWrite.Store()
            GoTo CleanUp
        Catch ex As Exception
            MessageBox.Show("Error in writing to fields:" & vbNewLine & ex.Message, "Write Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo CleanUp
        End Try

CleanUp:
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(pNFeatCursor)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(pEFeatCursor)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(pNCursor)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(pECursor)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(pNFeatSelection)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(pEFeatSelection)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(pFeatWrite)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        intRec = Nothing
        pLocationSelect = Nothing
        strQString = Nothing
        pNFeatSelection = Nothing
        pQueryFilter = Nothing
        pNFeatCursor = Nothing
        pNCursor = Nothing
        pFeatWrite = Nothing
        pEFeatSelection = Nothing
        pEFeatCursor = Nothing
        pECursor = Nothing
        pFeat = Nothing
        intTotalCount = Nothing
        pParcelTable = Nothing
        intNum = Nothing
        intRow = Nothing
        dblPRE_VAC_ACRE = Nothing
        dblPRE_DEVD_ACRE = Nothing
        dblABANDON_DEVD_ACRE = Nothing
        dblDEVD_VAC_ACRE = Nothing
        dblUNREDEVD_DEVD_ACRE = Nothing
        dblREDEVD_DEVD_ACRE = Nothing
        dblTOT_VAC_ACRE = Nothing
        dblTOT_DEVD_ACRE = Nothing
        strDevType = Nothing
        dblDevdAcres = Nothing
        dblVacAcres = Nothing
        dblVar = Nothing
        dblRedevRate = Nothing
        dblAbandonRate = Nothing
        dblNetAcre = Nothing
        dblPRE = Nothing
        dblPOST = Nothing
        dblNEW = Nothing
        dblTOT = Nothing
        dblTemp = Nothing
        intCount = Nothing
        arrTempPreSum = Nothing
        arrTempPostSum = Nothing
        arrTempNewSum = Nothing
        arrTempTotalSum = Nothing
        strFieldName = Nothing
        intIndex = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
        GC.WaitForFullGCComplete()
    End Sub

    Public Sub RunAggregationTool(ByVal process As IGPProcess)
        '*******************************************************
        ' Set the overwrite output option to true
        '*******************************************************
        Try
            gpAggregation.OverwriteOutput = True
            gpAggregation.Execute(process, Nothing)
            GoTo CleanUp
        Catch err As Exception
            GoTo CleanUp
        End Try
CleanUp:
        process = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
        GC.WaitForFullGCComplete()
    End Sub

    Private Sub btnCheckAllWeighted_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCheckAllWeighted.Click
        WeightedCheckStatus(1)
    End Sub

    Private Sub btnUncheckAllWeighted_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUncheckAllWeighted.Click
        WeightedCheckStatus(0)
    End Sub

    Private Sub WeightedCheckStatus(ByVal intCheck As Integer)
        'SET STATUS OF THE OPTIONS
        Dim iRow As Integer
        For iRow = 0 To Me.dgvFields.RowCount - 1
            m_frmFieldAggregation.dgvFields.Rows(iRow).Cells(3).Value = intCheck
        Next
    End Sub

    Private Sub btnCheckAllSum_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCheckAllSum.Click
        SumCheckStatus(1)
    End Sub

    Private Sub btnUncheckAllSum_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUncheckAllSum.Click
        SumCheckStatus(0)
    End Sub

    Private Sub SumCheckStatus(ByVal intCheck As Integer)
        'SET STATUS OF THE OPTIONS
        Dim iRow As Integer
        For iRow = 0 To Me.dgvFields.RowCount - 1
            m_frmFieldAggregation.dgvFields.Rows(iRow).Cells(4).Value = intCheck
        Next
    End Sub

    Private Sub btnTrackAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTrackAll.Click
        TrackingCheckStatus(1)
    End Sub

    Private Sub btnTrackNone_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTrackNone.Click
        TrackingCheckStatus(0)
    End Sub

    Private Sub TrackingCheckStatus(ByVal intCheck As Integer)
        'SET STATUS OF THE OPTIONS
        Dim iRow As Integer
        For iRow = 0 To Me.dgvFields.RowCount - 1
            m_frmFieldAggregation.dgvFields.Rows(iRow).Cells(0).Value = intCheck
        Next
    End Sub

    Private Sub rdbNeighborhood_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdbNeighborhood.CheckedChanged, rdbParcel.CheckedChanged
        Me.chkParcelSelected.Enabled = Me.rdbParcel.Checked
        Me.chkSelected.Enabled = Me.rdbNeighborhood.Checked
        Me.lblNeighborhoodLyr.Enabled = Me.rdbNeighborhood.Checked
        Me.cmbLayers.Enabled = Me.rdbNeighborhood.Checked
    End Sub
End Class