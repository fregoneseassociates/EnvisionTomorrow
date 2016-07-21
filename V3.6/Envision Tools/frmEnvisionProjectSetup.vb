Imports System
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Runtime.InteropServices
Imports System.Collections
Imports System.IO
Imports System.Drawing.Printing
Imports System.Data
Imports System.Math
Imports Microsoft.Office

Imports ESRI.ArcGIS.AnalysisTools
Imports ESRI.ArcGIS.ConversionTools
Imports ESRI.ArcGIS.DataManagementTools
Imports ESRI.ArcGIS.DataSourcesRaster
Imports ESRI.ArcGIS.Geoprocessor
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.ADF
Imports ESRI.ArcGIS.Cartoui
Imports ESRI.ArcGIS.Catalog
Imports ESRI.ArcGIS.CatalogUI
Imports ESRI.ArcGIS.DataSourcesGDB
Imports ESRI.ArcGIS.DataSourcesFile
Imports ESRI.ArcGIS.Display
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.GeoDatabaseUI
Imports ESRI.ArcGIS.SystemUI
'Imports ESRI.ArcGIS.SpatialAnalyst
'Imports ESRI.ArcGIS.SpatialAnalystUI
'Imports ESRI.ArcGIS.Geoprocessing
Imports ESRI.ArcGIS.Editor
Imports ESRI.ArcGIS.NetworkAnalystTools

Public Class frmEnvisionProjectSetup

    '***************************************************************************************************************************************
    'MODULAR LEVEL VARIABLES
    '***************************************************************************************************************************************
    'VARIABLES DEFINED AND POPULATED UPON FORM LOAD
    Private pGPSetup As ESRI.ArcGIS.Geoprocessor.Geoprocessor
    Private m_arrETSetup_FLayers As ArrayList = New ArrayList
    Private m_arrPolyFeatureLayers As ArrayList = New ArrayList
    Private m_arrETSetup_RLayers As ArrayList = New ArrayList

    'TAB 1 INPUT CHECKS AND MODULAR VARIABLES
    Private blnWorkspace As Boolean = False
    Private strWorkspacePath As String = ""
    Private blnFGD_Exists As Boolean = False
    Private strProjectName As String = ""
    Private blnProjectName As Boolean = False

    Private blnSourceInputs As Boolean = True
    Private blnGridCellSize As Boolean = True
    Private blnMaxParcelSize As Boolean = True
    Private blnExtent As Boolean = False
    Private blnProjection As Boolean = False


    'TAB 2 INPUT CHECKS
    Private blnLineBuffer As Boolean = False
    Private blnSlope As Boolean = False
    'TAB 3 
    'Public blnLandUseLoad As Boolean = True
    'ENVISION PROJECT SETUP
    Public m_pETSpatRefProject As ISpatialReference
    Public m_blnETSAFound As Boolean = True
    Public m_blnETEnvisonSetupIsOpening As Boolean = True
    Public m_arrETFeatureLayers As New ArrayList
    Public m_arrETPolyFeatureLayers As New ArrayList
    Public m_arrETRasterLayers As New ArrayList
    Public m_strETLandUseLyrName As String
    Public m_strETLandUseFldName As String
    Public m_strSubareaLyrName As String
    Public m_strSubareaFldName As String
    Public m_pExtentEnv As IEnvelope
    Public m_pEnvisionParcelsFLayer As IFeatureLayer = Nothing
    'Public m_pEnvisionGridCellsFLayer As IFeatureLayer = Nothing
    'Public m_pEnvisionHybridFLayer As IFeatureLayer = Nothing
    'Public m_pEnvisionCDLandsFLayer As IFeatureLayer = Nothing
    Public m_blnBufferLines As Boolean = False
    Public m_blnExtentChange As Boolean = False
    Public m_dblCustomExtentUpperX As Double = 0
    Public m_dblCustomExtentUpperY As Double = 0
    Public m_dblCustomExtentLowerX As Double = 0
    Public m_dblCustomExtentLowerY As Double = 0

    'SET PROCESS FEATURE LAYERS 
    Dim m_lyrMainProcessingLayer As IFeatureLayer = Nothing
    Dim m_lyrAOI As IFeatureLayer = Nothing
    Dim m_lyrGrid As IFeatureLayer = Nothing
    Dim m_lyrOriginalParcels As IFeatureLayer = Nothing
    Dim m_lyrParcelClip As IFeatureLayer = Nothing
    Dim m_lyrEnvisionParcel As IFeatureLayer = Nothing
    Dim m_lyrMaxAcreParcels As IFeatureLayer = Nothing
    Dim m_lyrMinAcreParcels As IFeatureLayer = Nothing
    Dim m_lyrParcelsUnionWithGrid As IFeatureLayer = Nothing
    Dim m_lyrConstraints As IFeatureLayer = Nothing
    Dim m_lyrTempSlope As IFeatureLayer = Nothing

    'SETUP OBJECT USED DURING PROCESSING
    Private dtStartTime_ETSetup As Date

    Dim m_lyrParcelConstraints As IFeatureLayer = Nothing
    Dim m_lyrLandUse As IFeatureLayer = Nothing
    Dim m_lyrConstrainedParcels As IFeatureLayer = Nothing
    Dim m_lyrUnConstrainedParcels As IFeatureLayer = Nothing
    Dim m_lyrEnvision As IFeatureLayer = Nothing
    Dim m_lyrMashup As IFeatureLayer = Nothing
    Dim m_lyrParcelCentroids As IFeatureLayer = Nothing
    Dim m_lyrLUCentroids As IFeatureLayer = Nothing
    Dim m_lyrEnvisionFinalParcels As IFeatureLayer = Nothing


    Dim intBuffLayerCount As Integer = 0
    Dim m_blnConstraints As Boolean = False
    Dim m_blnVacant As Boolean = False
    Dim m_blnDeveloped As Boolean = False
    Dim m_blnLandUse As Boolean = False
    Dim arrMatchFrom As ArrayList = New ArrayList
    Dim arrMatchTo As ArrayList = New ArrayList

    Dim m_strLandUseLyrName As String
    '***************************************************************************************************************************************

    '***************************************************************************************************************************************
    'FORM LEVEL SUBS AND FUNCTIONS
    '***************************************************************************************************************************************
    Private Sub frmEnvisionProjectSetup_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim pUID As New UID
        Try
            'CLEAR ALL THE FEATURE LAYER OBJECTS
            m_lyrAOI = Nothing
            m_lyrGrid = Nothing
            m_lyrConstrainedParcels = Nothing
            m_lyrEnvisionParcel = Nothing
            m_lyrMaxAcreParcels = Nothing
            m_lyrMinAcreParcels = Nothing
            m_lyrParcelClip = Nothing
            m_lyrParcelsUnionWithGrid = Nothing
            m_lyrConstraints = Nothing
            m_lyrTempSlope = Nothing
            m_lyrParcelConstraints = Nothing
            m_lyrLandUse = Nothing
            m_lyrConstrainedParcels = Nothing
            m_lyrEnvision = Nothing
            pGPSetup = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()

            Me.WindowState = FormWindowState.Minimized

            pUID.Value = "esriCore.SelectTool"
            If Not m_appEnvision Is Nothing Then
                m_appEnvision.CurrentTool = m_appEnvision.Document.CommandBars.Find(pUID)
            End If
            GoTo CleanUp
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Closing Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo CleanUp
        End Try

CleanUp:
        pUID = Nothing
        m_frmEnvisionProjectSetup = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub

    Private Sub frmEnvisionProjectSetup_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        m_blnETSetupFormIsOpening = True
        If Not EnvisionProjectSetup_LoadData(sender, e) Then
            m_frmEnvisionProjectSetup.Close()
            m_blnETSetupFormIsOpening = False
        Else
            m_blnETSetupFormIsOpening = False
            If Not m_pETSpatRefProject Is Nothing Then
                Me.rdbExtentView.Checked = True
                blnProjection = True
            End If
        End If
    End Sub

    Public Function EnvisionProjectSetup_LoadData(ByVal sender As Object, ByVal e As System.EventArgs) As Boolean
        EnvisionProjectSetup_LoadData = True
        Dim mxApplication As IMxApplication = Nothing
        Dim pMxDocument As IMxDocument = Nothing
        Dim mapCurrent As Map
        Dim uid As ESRI.ArcGIS.esriSystem.IUID = New ESRI.ArcGIS.esriSystem.UIDClass
        Dim enumLayer As ESRI.ArcGIS.Carto.IEnumLayer
        Dim pSpatRef As ISpatialReference
        Dim intCount As Integer
        Dim pLyr As ILayer
        Dim pFeatLyr As IFeatureLayer
        Dim pRasterLyr As IRasterLayer
        Dim intLayer As Integer
        Dim intFeatCount As Integer = 0
        Dim pActiveView As IActiveView = Nothing
        Dim pPCS As IProjectedCoordinateSystem
        Dim intLayerCount As Integer = 0

        Try
            '********************************************************************
            'Populate the combo boxes with layer information
            '********************************************************************
            If Not TypeOf m_appEnvision Is IApplication Then
                GoTo CloseForm
            Else
                m_appEnvision.StatusBar.Message(0) = "Loading Envsion Setup from"
                pMxDocument = DirectCast(m_appEnvision.Document, IMxDocument)
                mapCurrent = CType(pMxDocument.FocusMap, Map)
                pActiveView = CType(pMxDocument.FocusMap, IActiveView)
            End If

            'RETRIEVE THE FEATURE LAYERS TO POPULATE FEATURE LAYER OPTIONS
            Try
                m_appEnvision.StatusBar.Message(0) = "Building list of feature layers"
                m_arrETSetup_FLayers = New ArrayList
                m_arrETSetup_FLayers = New ArrayList
                uid.Value = "{E156D7E5-22AF-11D3-9F99-00C04F6BC78E}" '= IGeoFeatureLayer
                enumLayer = mapCurrent.Layers((CType(uid, ESRI.ArcGIS.esriSystem.UID)), True) ' Explicit Cast
                enumLayer.Reset()
                pLyr = enumLayer.Next
                Do While Not (pLyr Is Nothing)
                    pFeatLyr = CType(pLyr, IFeatureLayer)
                    If pFeatLyr.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                        m_arrPolyFeatureLayers.Add(pFeatLyr)
                        m_frmEnvisionProjectSetup.cmbParcelLayers.Items.Add(pFeatLyr.Name)
                        m_frmEnvisionProjectSetup.cmbExtentLayers.Items.Add(pFeatLyr.Name)
                        m_frmEnvisionProjectSetup.cmbLandUseLayers.Items.Add(pFeatLyr.Name)
                        m_frmEnvisionProjectSetup.cmbSubaraLayers.Items.Add(pFeatLyr.Name)
                    End If
                    m_appEnvision.StatusBar.Message(0) = "Building list of feature layers : " & pFeatLyr.Name
                    m_frmEnvisionProjectSetup.dgvConstraints.Rows.Add()
                    m_frmEnvisionProjectSetup.dgvConstraints.Rows(m_frmEnvisionProjectSetup.dgvConstraints.RowCount - 1).Cells(0).Value = "FALSE"
                    m_frmEnvisionProjectSetup.dgvConstraints.Rows(m_frmEnvisionProjectSetup.dgvConstraints.RowCount - 1).Cells(1).Value = pFeatLyr.Name
                    m_frmEnvisionProjectSetup.dgvConstraints.Rows(m_frmEnvisionProjectSetup.dgvConstraints.RowCount - 1).Cells(2).Value = "0"
                    m_frmEnvisionProjectSetup.dgvConstraints.Rows(m_frmEnvisionProjectSetup.dgvConstraints.RowCount - 1).Cells(3).Value = "Miles"
                    m_arrETSetup_FLayers.Add(pFeatLyr)
                    pLyr = enumLayer.Next()
                Loop
                If m_frmEnvisionProjectSetup.cmbExtentLayers.Items.Count > 0 Then
                    m_frmEnvisionProjectSetup.cmbExtentLayers.Text = m_frmEnvisionProjectSetup.cmbExtentLayers.Items.Item(0)
                End If
            Catch ex As System.Exception
                m_arrETSetup_FLayers = New ArrayList
                m_arrPolyFeatureLayers = New ArrayList
            End Try

            'DISABLE THE CONTROLS REQUIRING A POLYGON FEATURE LAYER
            If m_arrPolyFeatureLayers.Count = 0 Then
                'DISABLE PARCEL CONTROLS
                m_frmEnvisionProjectSetup.rdbSourceParcels.Enabled = False
                m_frmEnvisionProjectSetup.rdbSourceHybrid.Enabled = False
                m_frmEnvisionProjectSetup.cmbParcelLayers.Enabled = False
                m_frmEnvisionProjectSetup.rdbExtentParcel.Enabled = False
                'DISABLE LAND USE CONTROLS
                m_frmEnvisionProjectSetup.InfoMenu_LandUse.Enabled = False
                m_frmEnvisionProjectSetup.lblLandUseAttributes.Enabled = False
                m_frmEnvisionProjectSetup.dgvLandUseAttributes.Enabled = False
                'DISABLE FIELD MAP CONTROLS
                m_frmEnvisionProjectSetup.lblFieldMap.Enabled = False
                m_frmEnvisionProjectSetup.dgvFieldMatch.Enabled = False
                'DISABLE SUBAREA CONTROLS
                m_frmEnvisionProjectSetup.lblSubaraLayers.Enabled = False
                m_frmEnvisionProjectSetup.cmbSubaraLayers.Enabled = False
                m_frmEnvisionProjectSetup.lblSubareaFields.Enabled = False
                m_frmEnvisionProjectSetup.cmbSubareaFields.Enabled = False
                m_frmEnvisionProjectSetup.btnAddSubarea.Enabled = False
                m_frmEnvisionProjectSetup.dgvSubAreas.Enabled = False
            End If

            'DISABLE THE CONTROLS REQUIRING A FEATURE LAYER
            If m_arrETSetup_FLayers.Count = 0 Then
                'DISABLE PARCEL CONTROLS
                m_frmEnvisionProjectSetup.rdbExtentLayer.Enabled = False
                m_frmEnvisionProjectSetup.cmbExtentLayers.Enabled = False
                'DISABLE CONSTRAINTS CONTROLS
                m_frmEnvisionProjectSetup.InfoMenu_Constraints.Enabled = False
                m_frmEnvisionProjectSetup.lblConstraintsLyrs.Enabled = False
                m_frmEnvisionProjectSetup.dgvConstraints.Enabled = False
            End If

            'RETRIEVE THE FEATURE LAYERS TO POPULATE FEATURE LAYER OPTIONS
            Try
                m_appEnvision.StatusBar.Message(0) = "Building list of raster layers : "
                uid.Value = "{D02371C7-35F7-11D2-B1F2-00C04F8EDEFF}" '= IRasterLayer
                enumLayer = mapCurrent.Layers((CType(uid, ESRI.ArcGIS.esriSystem.UID)), True) ' Explicit Cast
                enumLayer.Reset()
                pLyr = enumLayer.Next
                Do While Not (pLyr Is Nothing)
                    pRasterLyr = CType(pLyr, IRasterLayer)
                    m_arrETSetup_RLayers.Add(pRasterLyr)
                    m_appEnvision.StatusBar.Message(0) = "Building list of raster layers : " & pRasterLyr.Name
                    m_frmEnvisionProjectSetup.cmbGridLayers.Items.Add(pRasterLyr.Name)
                    pLyr = enumLayer.Next()
                Loop
            Catch ex As System.Exception
                m_arrETSetup_RLayers = New ArrayList
            End Try

            'DISABLE THE CONTROLS REQUIRING A RASTER LAYER
            If m_arrETSetup_RLayers.Count = 0 Then
                'DISABLE SLOPE CONSTRAINTS CONTROLS
                m_frmEnvisionProjectSetup.pnlSlope.Enabled = False
            End If

            'BY DEFUALT RETRIEVE AND LOAD THE VIEW DOCUMENT SPATIAL REFERENCE PROJECTION
            pSpatRef = pMxDocument.FocusMap.SpatialReference
            Try
                pPCS = pSpatRef
                If pPCS.FactoryCode > 0 Then
                    m_frmEnvisionProjectSetup.tbxProjection.Text = pSpatRef.Name
                    m_pETSpatRefProject = pSpatRef
                    m_frmEnvisionProjectSetup.lblProjectUnits.Text = pPCS.CoordinateUnit.Name
                End If
            Catch ex As Exception
                m_frmEnvisionProjectSetup.gpbExtentCoor.Enabled = True
                m_frmEnvisionProjectSetup.rdbExtentCustom.Checked = True
            End Try

            'CHECK THE CURRENT VIEW 
            m_frmEnvisionProjectSetup.rdbExtentView.Checked = True

            GoTo CleanUp
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Envision Project Setup Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo CloseForm
        End Try
        GoTo CleanUp

CloseForm:
        EnvisionProjectSetup_LoadData = False
        GoTo CleanUp

CleanUp:
        m_appEnvision.StatusBar.Message(0) = ""
        mxApplication = Nothing
        pMxDocument = Nothing
        mapCurrent = Nothing
        pSpatRef = Nothing
        intCount = Nothing
        pLyr = Nothing
        pFeatLyr = Nothing
        pRasterLyr = Nothing
        intLayer = Nothing
        intFeatCount = Nothing
        pActiveView = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Function
    '***************************************************************************************************************************************



    '***************************************************************************************************************************************
    'SUBS AND FUNCTIONS ASSOCIATED WITH CONTROLS ON TAB ONE
    '***************************************************************************************************************************************
    Private Sub btnWorkspace_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnWorkspace.Click
        'WORKSPACE DIRECTORY
        'THIS IS THE DIRECTORY WHERE A FILE GEODATABASE WILL BE CREATED FOR PROCESSING TO OCCUR AND FINAL OUTPUT STORED
        Dim MyDialog As New FolderBrowserDialog
        Dim intCount As Integer
        Dim strFGeo As String = ""
        MyDialog.Description = "Select the Envision Project Workspace Directory:"
        If (MyDialog.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            m_frmEnvisionProjectSetup.tbxWorkspace.Text = MyDialog.SelectedPath.ToString
            strWorkspacePath = MyDialog.SelectedPath.ToString
            blnWorkspace = True
            'CYCLE THROUGH 200 FOR DEFAULT FILE GEODATABASE 
            For intCount = 1 To 200
                strFGeo = "ENVISION_PROJECT_" & CStr(intCount)
                If Not Directory.Exists(Me.tbxWorkspace.Text & "\" & strFGeo & ".gdb") Then
                    Me.tbxProjectName.Text = strFGeo
                    Exit For
                End If
            Next
        End If


        'blnWorkspace = ProjectDirCheck()
        'tbxProjectName_TextChanged(sender, e)
        GoTo CleanUp
CleanUp:
        RunButtonStatus()
        MyDialog = Nothing
        intCount = Nothing
        strFGeo = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub

    Private Sub tbxProjectName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbxProjectName.TextChanged
        If Not m_frmEnvisionProjectSetup Is Nothing Then
            blnProjectName = EvisionFGBNameCheck()
            If blnProjectName Then
                strProjectName = m_frmEnvisionProjectSetup.tbxProjectName.Text
            Else
                strProjectName = ""
            End If
        End If
        RunButtonStatus()
    End Sub

    Private Function EvisionFGBNameCheck() As Boolean
        'CHECK IF THE FILE GEODATABASE PROJECT ALREADY EXISTS
        'THIS DIRECTORY SHOULD NOT EXIST AS ENVISION SETUP WILL CREATE IT.  THE ASSUMPTION IS E.T. SETUP WILL CREATE A NEW
        'PROCESSING FILE GEODATABASE EACH TIME THE PROCESS IS EXECUTED
        If m_frmEnvisionProjectSetup.tbxWorkspace.Text.Length > 0 Then
            If Directory.Exists(m_frmEnvisionProjectSetup.tbxWorkspace.Text) Then
                If Directory.Exists(m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb") Then
                    EvisionFGBNameCheck = False
                    m_frmEnvisionProjectSetup.tbxProjectName.ForeColor = Color.Red
                Else
                    EvisionFGBNameCheck = True
                    m_frmEnvisionProjectSetup.tbxProjectName.ForeColor = Color.Black
                End If
            End If
        End If
    End Function

    Private Sub FeatSourceRB_OptionSelect_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdbSourceParcels.CheckedChanged, rdbSourceGridCells.CheckedChanged, rdbSourceHybrid.CheckedChanged
        If Not m_blnETSetupFormIsOpening Then
            FeatSourceRB_OptionSelect()
            blnSourceInputs = True
            If Me.rdbSourceHybrid.Checked Or Me.rdbSourceParcels.Checked Then
                If Me.cmbParcelLayers.Text.Length = 0 Or Me.cmbParcelLayers.ForeColor = Color.Red Then
                    blnSourceInputs = False
                End If
            End If
        End If
        RunButtonStatus()
    End Sub

    Public Function FeatSourceRB_OptionSelect() As Boolean
        'SET ENABLE STATUS TO CONTROLS ASSOCIATED TO THE FEATURE SOURCE OPTIONS
        FeatSourceRB_OptionSelect = True
        Dim blnParcel As Boolean = True
        Dim blnGrid As Boolean = True
        If m_frmEnvisionProjectSetup.rdbSourceParcels.Checked Then
            blnParcel = True
            blnGrid = False
        ElseIf m_frmEnvisionProjectSetup.rdbSourceGridCells.Checked Then
            blnParcel = False
            blnGrid = True
        ElseIf m_frmEnvisionProjectSetup.rdbSourceHybrid.Checked Then
            blnParcel = True
            blnGrid = True
        End If

        m_frmEnvisionProjectSetup.lblParcelLyr.Enabled = blnParcel
        m_frmEnvisionProjectSetup.cmbParcelLayers.Enabled = blnParcel
        m_frmEnvisionProjectSetup.lblGridCellSize.Enabled = blnGrid
        m_frmEnvisionProjectSetup.tbxGridCellSize.Enabled = blnGrid
        m_frmEnvisionProjectSetup.lblMaxParcelSize.Enabled = blnGrid
        m_frmEnvisionProjectSetup.tbxMaxParcelSize.Enabled = blnGrid
        m_frmEnvisionProjectSetup.lblAcres.Enabled = blnParcel
        m_frmEnvisionProjectSetup.rdbExtentParcel.Enabled = blnParcel
        'RELATED CONTROL IN THE EXTENT OPTIONS
        If m_frmEnvisionProjectSetup.rdbExtentParcel.Checked And Not blnParcel Then
            m_frmEnvisionProjectSetup.rdbExtentView.Checked = True
        End If

        'FIELD MAP TAB CONTROLS
        m_frmEnvisionProjectSetup.dgvFieldMatch.Enabled = blnParcel


    End Function

    Private Sub cmbParcelLayers_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbParcelLayers.SelectedIndexChanged, cmbParcelLayers.TextChanged
        If Not m_blnETSetupFormIsOpening Then
            ParcelLyrCheck()
        End If

        Dim intCount As Integer
        Dim blnLayerFound As Boolean = False
        Dim pFeatureClass As IFeatureClass
        Dim pField As IField
        Dim intFld As Integer

        'LOAD THE FIELDS FROM THE SELECTED PARCEL LAYER
        If Me.cmbParcelLayers.Text.Length <= 0 Or m_blnETSetupFormIsOpening Or m_arrPolyFeatureLayers.Count <= 0 Then
            GoTo CleanUp
        Else
            'MAKE SURE THE SELECTED LAYER IS VALID
            If Me.cmbParcelLayers.Items.Contains(Me.cmbParcelLayers.Text) Then
                blnLayerFound = True
            End If
            'If Me.rdbExtentParcel.Checked And Not blnLayerFound Then
            '    Me.rdbExtentParcel.Checked = True
            'End If
            'Me.rdbExtentParcel.Enabled = blnLayerFound
        End If

        If m_arrPolyFeatureLayers.Count > 0 And blnLayerFound Then
            Try
                m_lyrOriginalParcels = CType(m_arrPolyFeatureLayers.Item(Me.cmbParcelLayers.SelectedIndex), IFeatureLayer)
                pFeatureClass = m_lyrOriginalParcels.FeatureClass
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Return Parcel Layer Fields Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
                GoTo CleanUp
            End Try

            'CLEAR THE FIELD MATCH DATA GRID AND POPULATE DOUBLE FIELDS FROM THE SELECGTED PARCEL LAYER
            'ALSO CLEAR ANDPOPULATE THE FIELD MAP CONTROLS
            Me.dgvFieldMatch.Rows.Clear()
            Me.Refresh()
            For intFld = 0 To pFeatureClass.Fields.FieldCount - 1
                pField = pFeatureClass.Fields.Field(intFld)
                If pField.Type = esriFieldType.esriFieldTypeInteger Or pField.Type = esriFieldType.esriFieldTypeDouble Or pField.Type = esriFieldType.esriFieldTypeSmallInteger Then
                    If Not pField.Name = "OBJECTID" Then
                        Me.dgvFieldMatch.Rows.Add()
                        Me.dgvFieldMatch.Rows(Me.dgvFieldMatch.RowCount - 1).Cells(0).Value = pField.Name
                        Me.dgvFieldMatch.Rows(Me.dgvFieldMatch.RowCount - 1).Cells(1).Value = 0
                        Me.dgvFieldMatch.Rows(Me.dgvFieldMatch.RowCount - 1).Cells(2).Value = ""
                    End If
                End If
            Next
        End If

        blnSourceInputs = True
        If Me.rdbSourceHybrid.Checked Or Me.rdbSourceParcels.Checked Then
            If Me.cmbParcelLayers.Text.Length = 0 Or Me.cmbParcelLayers.ForeColor = Color.Red Then
                blnSourceInputs = False
            End If
        End If

CleanUp:
        RunButtonStatus()
        intCount = Nothing
        blnLayerFound = Nothing
        pFeatureClass = Nothing
        pField = Nothing
        intFld = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub

    Private Sub tbxGridCellSize_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbxGridCellSize.TextChanged
        If Not m_blnETSetupFormIsOpening Then
            blnGridCellSize = GridCellParametersReview()
        End If
        RunButtonStatus()
    End Sub

    Private Sub tbxMaxParcelSize_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbxMaxParcelSize.TextChanged
        If Not m_blnETSetupFormIsOpening Then
            blnMaxParcelSize = ParcelSizeParametersReview()
        End If
        RunButtonStatus()
    End Sub

    Private Sub rdbExtentView_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdbExtentView.CheckedChanged
        'If m_blnETSetupFormIsOpening Then
        '    Exit Sub
        'End If

        'RETRIEVE THE EXTENT FOR THE SELECTED PROCESSING EXTENT OPTION
        Dim mxApplication As IMxApplication
        Dim mxDoc As IMxDocument
        Dim mapCurrent As Map
        Dim pActiveView As IActiveView
        Dim pLayer As ILayer
        Dim pExtentEnv As IEnvelope = Nothing
        Dim pClone1 As IClone
        Dim pClone2 As IClone
        Dim pSR As ISpatialReference = Nothing
        Dim pGeometry As IGeometry5


        'RETRIEVE THE EXTENT FROM THE SELECTED OPTION
        Try
            blnExtent = False
            mxApplication = CType(m_appEnvision, IMxApplication)
            mxDoc = CType(m_appEnvision.Document, IMxDocument)
            mapCurrent = CType(mxDoc.FocusMap, Map)

            'SET ENABLED STATUS TO COORDINATE INPUT CONTROLS
            m_frmEnvisionProjectSetup.gpbExtentCoor.Enabled = m_frmEnvisionProjectSetup.rdbExtentCustom.Checked

            'CLEAR THE EXTENT COORDINATES AND SET FONT COLOR TO BLACK
            If Not m_frmEnvisionProjectSetup.rdbExtentView.Checked Then
                m_frmEnvisionProjectSetup.tbxExtentLowerX.Text = ""
                m_frmEnvisionProjectSetup.tbxExtentLowerY.Text = ""
                m_frmEnvisionProjectSetup.tbxExtentUpperX.Text = ""
                m_frmEnvisionProjectSetup.tbxExtentUpperY.Text = ""
                m_frmEnvisionProjectSetup.tbxExtentLowerX.ForeColor = Color.Black
                m_frmEnvisionProjectSetup.tbxExtentLowerY.ForeColor = Color.Black
                m_frmEnvisionProjectSetup.tbxExtentUpperX.ForeColor = Color.Black
                m_frmEnvisionProjectSetup.tbxExtentUpperY.ForeColor = Color.Black
                GoTo CleanUp
            End If

            'RETRIEVE VIEW EXTENT ENVELOPE
            pActiveView = CType(mxDoc.FocusMap, IActiveView)
            pExtentEnv = pActiveView.Extent
            pSR = mxDoc.FocusMap.SpatialReference

            'CHECK THE SPATIAL REFERENCES TO MAKE SURE THEY MATCH, REPROJECT IF THEY DO NOT MATCH
            Try
                If Not m_pETSpatRefProject Is Nothing And Not pSR Is Nothing Then
                    pClone1 = m_pETSpatRefProject
                    pClone2 = pSR
                    If Not pClone1.IsEqual(pClone2) And Not pExtentEnv Is Nothing Then
                        pGeometry = pExtentEnv
                        pGeometry.Project(m_pETSpatRefProject)
                        pExtentEnv = pGeometry
                    End If
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Extent Retrieval Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
                GoTo CleanUp
            End Try

            'WRITE EXTENT COORDINATES TO THE REPRESENTATIVE CONTROLS
            If Not pExtentEnv Is Nothing Then
                m_frmEnvisionProjectSetup.tbxExtentLowerX.Text = pExtentEnv.XMin.ToString
                m_frmEnvisionProjectSetup.tbxExtentLowerY.Text = pExtentEnv.YMin.ToString
                m_frmEnvisionProjectSetup.tbxExtentUpperX.Text = pExtentEnv.XMax.ToString
                m_frmEnvisionProjectSetup.tbxExtentUpperY.Text = pExtentEnv.YMax.ToString
            End If
            m_pExtentEnv = pExtentEnv
            blnExtent = True

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Extent Retrieval Error - View Extent", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            m_pExtentEnv = Nothing
            GoTo CleanUp
        End Try

        GoTo CleanUp
CleanUp:
        RunButtonStatus()
        mxApplication = Nothing
        mxDoc = Nothing
        mapCurrent = Nothing
        pActiveView = Nothing
        pLayer = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub

    Private Sub rdbExtentParcel_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdbExtentParcel.CheckedChanged, cmbParcelLayers.SelectedIndexChanged
        If m_blnETSetupFormIsOpening Or Me.rdbExtentParcel.Checked = False Then
            Exit Sub
        End If

        'RETRIEVE THE EXTENT FOR THE SELECTED PROCESSING EXTENT OPTION
        Dim mxApplication As IMxApplication
        Dim mxDoc As IMxDocument
        Dim mapCurrent As Map
        Dim pActiveView As IActiveView
        Dim pLayer As ILayer
        Dim pExtentEnv As IEnvelope = Nothing
        Dim pGeodataset As IGeoDataset
        Dim pClone1 As IClone
        Dim pClone2 As IClone
        Dim pSR As ISpatialReference = Nothing
        Dim pGeometry As IGeometry5


        'RETRIEVE THE EXTENT FROM THE SELECTED OPTION
        Try
            blnExtent = False
            mxApplication = CType(m_appEnvision, IMxApplication)
            mxDoc = CType(m_appEnvision.Document, IMxDocument)
            mapCurrent = CType(mxDoc.FocusMap, Map)

            'SET ENABLED STATUS TO COORDINATE INPUT CONTROLS
            Me.gpbExtentCoor.Enabled = Me.rdbExtentCustom.Checked

            'CLEAR THE EXTENT COORDINATES AND SET FONT COLOR TO BLACK
            If Not m_frmEnvisionProjectSetup.rdbExtentParcel.Checked Then
                Me.tbxExtentLowerX.Text = ""
                Me.tbxExtentLowerY.Text = ""
                Me.tbxExtentUpperX.Text = ""
                Me.tbxExtentUpperY.Text = ""
                Me.tbxExtentLowerX.ForeColor = Color.Black
                Me.tbxExtentLowerY.ForeColor = Color.Black
                Me.tbxExtentUpperX.ForeColor = Color.Black
                Me.tbxExtentUpperY.ForeColor = Color.Black
                GoTo CleanUp
            End If

            'RETRIEVE EXTENT ENVELOPE FROM SELECTED FEATURE LAYER 
            If Me.cmbParcelLayers.SelectedIndex >= 0 Then
                pLayer = CType(m_arrETSetup_FLayers.Item(Me.cmbParcelLayers.SelectedIndex), ILayer)
                pGeodataset = pLayer
                pExtentEnv = pLayer.AreaOfInterest
                pSR = pGeodataset.SpatialReference
            Else
                GoTo CleanUp
            End If

            'CHECK THE SPATIAL REFERENCES TO MAKE SURE THEY MATCH, REPROJECT IF THEY DO NOT MATCH
            Try
                If Not m_pETSpatRefProject Is Nothing And Not pSR Is Nothing Then
                    pClone1 = m_pETSpatRefProject
                    pClone2 = pSR
                    If Not pClone1.IsEqual(pClone2) And Not pExtentEnv Is Nothing Then
                        pGeometry = pExtentEnv
                        pGeometry.Project(m_pETSpatRefProject)
                        pExtentEnv = pGeometry
                    End If
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Extent Retrieval Error - Parcel Layer Select", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
                GoTo CleanUp
            End Try

            'WRITE EXTENT COORDINATES TO THE REPRESENTATIVE CONTROLS
            If Not pExtentEnv Is Nothing Then
                Me.tbxExtentLowerX.Text = pExtentEnv.XMin.ToString
                Me.tbxExtentLowerY.Text = pExtentEnv.YMin.ToString
                Me.tbxExtentUpperX.Text = pExtentEnv.XMax.ToString
                Me.tbxExtentUpperY.Text = pExtentEnv.YMax.ToString
            End If
            m_pExtentEnv = pExtentEnv
            blnExtent = True

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Extent Retrieval Error - Parcel Layer Select", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            m_pExtentEnv = Nothing
            GoTo CleanUp
        End Try

        GoTo CleanUp
CleanUp:
        RunButtonStatus()
        mxApplication = Nothing
        mxDoc = Nothing
        mapCurrent = Nothing
        pActiveView = Nothing
        pLayer = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub

    Private Sub rbExtentCustom_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdbExtentCustom.CheckedChanged

        'LOAD EXISTING USER DEFINED COORDINATES IF AVAIALABLE
        If Me.rdbExtentCustom.Checked Then
            Try
                Me.tbxExtentLowerX.Text = CStr(m_dblCustomExtentLowerX)
                Me.tbxExtentLowerY.Text = CStr(m_dblCustomExtentLowerY)
                Me.tbxExtentUpperX.Text = CStr(m_dblCustomExtentUpperX)
                Me.tbxExtentUpperY.Text = CStr(m_dblCustomExtentUpperY)
                Me.tbxExtentLowerX.ForeColor = Color.Black
                Me.tbxExtentLowerY.ForeColor = Color.Black
                Me.tbxExtentUpperX.ForeColor = Color.Black
                Me.tbxExtentUpperY.ForeColor = Color.Black
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Load Error")
            End Try
        End If
        RunButtonStatus()

    End Sub

    Private Sub rdbExtentLayer_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdbExtentLayer.CheckedChanged, cmbExtentLayers.SelectedIndexChanged
        'RETRIEVE THE EXTENT FOR THE SELECTED PROCESSING EXTENT OPTION
        Dim mxApplication As IMxApplication
        Dim mxDoc As IMxDocument
        Dim mapCurrent As Map
        Dim pActiveView As IActiveView
        Dim pLayer As ILayer
        Dim pExtentEnv As IEnvelope = Nothing
        Dim pGeodataset As IGeoDataset
        Dim pClone1 As IClone
        Dim pClone2 As IClone
        Dim pSR As ISpatialReference = Nothing
        Dim pGeometry As IGeometry5


        'RETRIEVE THE EXTENT FROM THE SELECTED OPTION
        Try
            blnExtent = False
            mxApplication = CType(m_appEnvision, IMxApplication)
            mxDoc = CType(m_appEnvision.Document, IMxDocument)
            mapCurrent = CType(mxDoc.FocusMap, Map)

            'SET ENABLED STATUS TO COORDINATE INPUT CONTROLS
            Me.gpbExtentCoor.Enabled = Me.rdbExtentCustom.Checked
            Me.cmbExtentLayers.Enabled = Me.rdbExtentLayer.Checked

            'CLEAR THE EXTENT COORDINATES AND SET FONT COLOR TO BLACK
            If Not m_frmEnvisionProjectSetup.rdbExtentLayer.Checked Then
                Me.tbxExtentLowerX.Text = ""
                Me.tbxExtentLowerY.Text = ""
                Me.tbxExtentUpperX.Text = ""
                Me.tbxExtentUpperY.Text = ""
                Me.tbxExtentLowerX.ForeColor = Color.Black
                Me.tbxExtentLowerY.ForeColor = Color.Black
                Me.tbxExtentUpperX.ForeColor = Color.Black
                Me.tbxExtentUpperY.ForeColor = Color.Black
                GoTo CleanUp
            End If

            'RETRIEVE EXTENT ENVELOPE FROM SELECTED FEATURE LAYER 
            If m_frmEnvisionProjectSetup.cmbExtentLayers.SelectedIndex >= 0 Then
                pLayer = CType(mapCurrent.Layer(Me.cmbExtentLayers.SelectedIndex), ILayer)
                pGeodataset = pLayer
                pExtentEnv = pLayer.AreaOfInterest
                pSR = pGeodataset.SpatialReference
            Else
                GoTo CleanUp
            End If

            'CHECK THE SPATIAL REFERENCES TO MAKE SURE THEY MATCH, REPROJECT IF THEY DO NOT MATCH
            Try
                If Not m_pETSpatRefProject Is Nothing And Not pSR Is Nothing Then
                    pClone1 = m_pETSpatRefProject
                    pClone2 = pSR
                    If Not pClone1.IsEqual(pClone2) And Not pExtentEnv Is Nothing Then
                        pGeometry = pExtentEnv
                        pGeometry.Project(m_pETSpatRefProject)
                        pExtentEnv = pGeometry
                    End If
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Extent Retrieval Error - Layer Select", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
                GoTo CleanUp
            End Try

            'WRITE EXTENT COORDINATES TO THE REPRESENTATIVE CONTROLS
            If Not pExtentEnv Is Nothing Then
                Me.tbxExtentLowerX.Text = pExtentEnv.XMin.ToString
                Me.tbxExtentLowerY.Text = pExtentEnv.YMin.ToString
                Me.tbxExtentUpperX.Text = pExtentEnv.XMax.ToString
                Me.tbxExtentUpperY.Text = pExtentEnv.YMax.ToString
            End If
            m_pExtentEnv = pExtentEnv
            blnExtent = True

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Extent Retrieval Error - Layer Select", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            m_pExtentEnv = Nothing
            GoTo CleanUp
        End Try

        GoTo CleanUp
CleanUp:
        RunButtonStatus()
        mxApplication = Nothing
        mxDoc = Nothing
        mapCurrent = Nothing
        pActiveView = Nothing
        pLayer = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub

    Private Sub CoordinateInputs_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbxExtentUpperX.TextChanged, tbxExtentUpperY.TextChanged, tbxExtentLowerX.TextChanged, tbxExtentLowerY.TextChanged
        'RETRIEVE THE EXTENT FOR THE SELECTED PROCESSING EXTENT OPTION
        Dim mxApplication As IMxApplication
        Dim mxDoc As IMxDocument
        Dim mapCurrent As Map
        Dim pActiveView As IActiveView
        Dim pLayer As ILayer
        Dim pExtentEnv As IEnvelope = Nothing
        Dim pClone1 As IClone
        Dim pClone2 As IClone
        Dim pSR As ISpatialReference = Nothing
        Dim pGeometry As IGeometry5
        blnExtent = False

        'TRACK AND SAVE THE CUSTOM COORDIANTES THE USER ENTERS 
        If Not IsNumeric(Me.tbxExtentUpperX.Text) Then
            Me.tbxExtentUpperX.ForeColor = Color.Red
        End If
        If Not IsNumeric(Me.tbxExtentUpperY.Text) Then
            Me.tbxExtentUpperY.ForeColor = Color.Red
        End If
        If Not IsNumeric(Me.tbxExtentLowerX.Text) Then
            Me.tbxExtentLowerX.ForeColor = Color.Red
        End If
        If Not IsNumeric(Me.tbxExtentLowerY.Text) Then
            Me.tbxExtentLowerY.ForeColor = Color.Red
        End If
        If Me.rdbExtentCustom.Checked And IsNumeric(Me.tbxExtentLowerX.Text) And IsNumeric(Me.tbxExtentLowerY.Text) And _
                   IsNumeric(Me.tbxExtentUpperX.Text) And IsNumeric(Me.tbxExtentUpperY.Text) Then
            m_blnExtentChange = True
            Me.tbxExtentLowerX.ForeColor = Color.Black
            Me.tbxExtentLowerY.ForeColor = Color.Black
            Me.tbxExtentUpperX.ForeColor = Color.Black
            Me.tbxExtentUpperY.ForeColor = Color.Black
            If IsNumeric(Me.tbxExtentUpperX.Text) Then
                m_dblCustomExtentUpperX = CDbl(Me.tbxExtentUpperX.Text)
            End If
            If IsNumeric(Me.tbxExtentUpperY.Text) Then
                m_dblCustomExtentUpperY = CDbl(Me.tbxExtentUpperY.Text)
            End If
            If IsNumeric(Me.tbxExtentLowerX.Text) Then
                m_dblCustomExtentLowerX = CDbl(Me.tbxExtentLowerX.Text)
            End If
            If IsNumeric(Me.tbxExtentLowerY.Text) Then
                m_dblCustomExtentLowerY = CDbl(Me.tbxExtentLowerY.Text)
            End If
            'RETRIEVE THE EXTENT FROM THE SELECTED OPTION
            Try
                mxApplication = CType(m_appEnvision, IMxApplication)
                mxDoc = CType(m_appEnvision.Document, IMxDocument)
                mapCurrent = CType(mxDoc.FocusMap, Map)

                'SET ENABLED STATUS TO COORDINATE INPUT CONTROLS
                Me.gpbExtentCoor.Enabled = Me.rdbExtentCustom.Checked

                'RETRIEVE EXTENT ENVELOPE FROM SELECTED FEATURE LAYER 
                If IsNumeric(Me.tbxExtentLowerX.Text) And IsNumeric(Me.tbxExtentLowerY.Text) And _
                   IsNumeric(Me.tbxExtentUpperX.Text) And IsNumeric(Me.tbxExtentUpperY.Text) Then
                    pExtentEnv = New Envelope
                    pExtentEnv.XMin = CDbl(Me.tbxExtentLowerX.Text)
                    pExtentEnv.YMin = CDbl(Me.tbxExtentLowerY.Text)
                    pExtentEnv.XMax = CDbl(Me.tbxExtentUpperX.Text)
                    pExtentEnv.YMax = CDbl(Me.tbxExtentUpperY.Text)
                Else
                    GoTo CleanUp
                End If

                'CHECK THE SPATIAL REFERENCES TO MAKE SURE THEY MATCH, REPROJECT IF THEY DO NOT MATCH
                Try
                    If Not m_pETSpatRefProject Is Nothing And Not pSR Is Nothing Then
                        pClone1 = m_pETSpatRefProject
                        pClone2 = pSR
                        If Not pClone1.IsEqual(pClone2) And Not pExtentEnv Is Nothing Then
                            pGeometry = pExtentEnv
                            pGeometry.Project(m_pETSpatRefProject)
                            pExtentEnv = pGeometry
                        End If
                    End If
                Catch ex As Exception
                    MessageBox.Show(ex.Message, "Extent Retrieval Error - Text Change", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
                    GoTo CleanUp
                End Try

                'WRITE EXTENT COORDINATES TO THE REPRESENTATIVE CONTROLS
                m_pExtentEnv = pExtentEnv
                blnExtent = True

            Catch ex As Exception
                MessageBox.Show(ex.Message, "Extent Retrieval Error - Text Change", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
                m_pExtentEnv = Nothing
                GoTo CleanUp
            End Try

            GoTo CleanUp
        End If


CleanUp:
        RunButtonStatus()
        mxApplication = Nothing
        mxDoc = Nothing
        mapCurrent = Nothing
        pActiveView = Nothing
        pLayer = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub

    Private Sub btnExcelFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExcelFile.Click
        '****************************************************************
        'Provide the user with a Open File dialog to select the Envision Excel file
        '****************************************************************
        Dim FileDialog1 As New OpenFileDialog

        Try
            FileDialog1.Title = "Select an Envision Excel File"
            FileDialog1.Filter = "Excel Files (*.xls,*.xlsx,*.xlsm)|*.xls;*.xlsx;*.xlsm"
            If FileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                m_frmEnvisionProjectSetup.tbxExcel.Text = FileDialog1.FileName.ToString
                Me.WindowState = FormWindowState.Minimized
                LoadEnvisionExcelFile()
                Me.WindowState = FormWindowState.Normal
            End If
            GoTo CleanUp
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Envision Excel Load Error")
            GoTo CleanUp
        End Try

CleanUp:
        RunButtonStatus()
        m_frmEnvisionProjectSetup.lblStatus1.Visible = False
        m_frmEnvisionProjectSetup.barStatus1.Visible = False
        m_frmEnvisionProjectSetup.lblStatus1.Text = ""
        m_frmEnvisionProjectSetup.barStatus1.Value = 0
        m_frmEnvisionProjectSetup.Refresh()
        FileDialog1 = Nothing
        GC.Collect()
        GC.WaitForPendingFinalizers()
    End Sub

    Private Sub btnSelectPrj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSelectPrj.Click
        'PROVIDE THE SPATIAL REFERENCE PROPERTIES DIALOG TO THE USER TO DEFINE THE PROJECT PROJECTION
        Dim pSpaRefDlg As ISpatialReferenceDialog
        Dim pSpatRef As ISpatialReference
        Dim pPCS As IProjectedCoordinateSystem
        Try
            pSpaRefDlg = New SpatialReferenceDialog
            Me.WindowState = FormWindowState.Minimized
            pSpatRef = pSpaRefDlg.DoModalCreate(False, False, False, 0)
            If pSpatRef Is Nothing Then
                GoTo CleanUp
            Else
                m_frmEnvisionProjectSetup.tbxProjection.Text = pSpatRef.Name
                m_pETSpatRefProject = pSpatRef
                pPCS = m_pETSpatRefProject
                m_frmEnvisionProjectSetup.lblProjectUnits.Text = pPCS.CoordinateUnit.Name
                blnProjection = True
            End If
        Catch ex As Exception
            MessageBox.Show("An error occured while defining the project projection." & vbNewLine & ex.Message, "Projection Select Error")
            GoTo CleanUp
        End Try
CleanUp:
        RunButtonStatus()
        Me.WindowState = FormWindowState.Normal
        pSpaRefDlg = Nothing
        pSpatRef = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub



    '***************************************************************************************************************************************
    'SUBS AND FUNCTIONS ASSOCIATED WITH CONTROLS ON TAB TWO
    '***************************************************************************************************************************************
    Private Sub btnCheckAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCheckAll.Click
        'CHECK ALL LAYERS AS CONSTRAINTS
        ConstraintsCheckStatus(True)
    End Sub

    Private Sub btnUncheckAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUncheckAll.Click
        'UNCHECK ALL LAYERS AS CONSTRAINTS
        ConstraintsCheckStatus(False)
    End Sub

    Private Sub itmAttribtueConstraints_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles itmAttribtueConstraints.Click
        Me.itmRetainConstraints.Checked = False
        Me.itmDeleteConstraints.Checked = False
        Me.btnConstraintedOption.Text = Me.itmAttribtueConstraints.Text
    End Sub

    Private Sub itmRetainConstraints_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles itmRetainConstraints.Click
        Me.itmAttribtueConstraints.Checked = False
        Me.itmDeleteConstraints.Checked = False
        Me.btnConstraintedOption.Text = Me.itmRetainConstraints.Text
    End Sub

    Private Sub itmDeleteConstraints_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles itmDeleteConstraints.Click
        Me.itmAttribtueConstraints.Checked = False
        Me.itmRetainConstraints.Checked = False
        Me.btnConstraintedOption.Text = Me.itmDeleteConstraints.Text
    End Sub

    Public Sub ConstraintsCheckStatus(ByVal blnCheckStatus As Boolean)
        'SET CHECK STATUS TO ALL LAYERS AS DEFINED BY INUT VARIABLE
        Dim intRow As Integer
        Me.dgvConstraints.ClearSelection()
        For intRow = 0 To Me.dgvConstraints.Rows.Count - 1
            If blnCheckStatus Then
                Me.dgvConstraints.Rows(intRow).Cells(0).Value = "TRUE"
            Else
                Me.dgvConstraints.Rows(intRow).Cells(0).Value = "FALSE"
            End If
        Next
    End Sub

    Private Sub ConstraintsOptionsCheck(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbGridLayers.SelectedIndexChanged, cmbGridLayers.TextChanged, chkCreateSlope15to25.CheckedChanged, chkCreateSlope15to25.TextChanged, chkCreateSlope25Plus.CheckedChanged, chkCreateSlope25Plus.TextChanged, tbxConstraintsCellSize.TextChanged
        'REVIEW SLOPE INPUTS 
        Me.lblBndFeatClass.Enabled = False
        Me.cmbGridLayers.Enabled = False
        Me.tbxConstraintsCellSize.Enabled = False
        Me.lblConstraintsCellSize.Enabled = False
        If Me.chkCreateSlope25Plus.Checked Or Me.chkCreateSlope15to25.Checked Then
            Me.lblBndFeatClass.Enabled = True
            Me.cmbGridLayers.Enabled = True
            Me.tbxConstraintsCellSize.Enabled = True
            Me.lblConstraintsCellSize.Enabled = True
        End If
        If Me.cmbGridLayers.Text.Length > 0 Then
            If Me.cmbGridLayers.Items.IndexOf(Me.cmbGridLayers.Text) >= 0 Then
                Me.cmbGridLayers.ForeColor = Color.Black
            Else
                Me.cmbGridLayers.ForeColor = Color.Red
            End If
        End If
        If Me.tbxConstraintsCellSize.Text.Length > 0 Then
            Me.tbxConstraintsCellSize.ForeColor = Color.Red
            If IsNumeric(Me.tbxConstraintsCellSize.Text) Then
                If CInt(Me.tbxConstraintsCellSize.Text) > 0 Then
                    Me.tbxConstraintsCellSize.ForeColor = Color.Black
                End If
            End If
        End If
        RunButtonStatus()
    End Sub
    '***************************************************************************************************************************************



    '***************************************************************************************************************************************
    'SUBS AND FUNCTIONS ASSOCIATED WITH CONTROLS ON TAB THREE
    '***************************************************************************************************************************************
    Private Sub cmbLandUseLayers_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbLandUseLayers.TextChanged
        If m_blnETSetupFormIsOpening Then
            Exit Sub
        End If

        Dim intCount As Integer
        Dim blnLayerFound As Boolean = False
        Dim pFeatureClass As IFeatureClass
        Dim pField As IField
        Dim intFld As Integer

        'LOAD THE FIELDS FROM THE SELECTED LAND USE LAYER
        If Me.cmbLandUseLayers.Text = "Select the Land Use Feature Layer" Or m_blnETSetupFormIsOpening Or m_arrPolyFeatureLayers.Count <= 0 Then
            GoTo CleanUp
        Else
            'MAKE SURE THE SELECTED LAYER IS VALID
            For intCount = 0 To Me.cmbLandUseLayers.Items.Count - 1
                If Me.cmbLandUseLayers.Items.Item(intCount).ToString = Me.cmbLandUseLayers.Text Then
                    blnLayerFound = True
                    Exit For
                End If
            Next
            If Not blnLayerFound Then
                GoTo CleanUp
            End If
        End If

        Me.cmbLandUseField.Items.Clear()
        Me.dgvLandUseAttributes.Rows.Clear()
        If m_arrETSetup_FLayers.Count > 0 Then
            Try
                m_lyrLandUse = CType(m_arrPolyFeatureLayers.Item(Me.cmbLandUseLayers.SelectedIndex), IFeatureLayer)
                pFeatureClass = m_lyrLandUse.FeatureClass
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Return Land Use Layer Fields Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
                GoTo CleanUp
            End Try

            Me.cmbLandUseField.Items.Clear()
            For intFld = 0 To pFeatureClass.Fields.FieldCount - 1
                pField = pFeatureClass.Fields.Field(intFld)
                If pField.Type = esriFieldType.esriFieldTypeString Or pField.Type = esriFieldType.esriFieldTypeInteger Or pField.Type = esriFieldType.esriFieldTypeDouble Or pField.Type = esriFieldType.esriFieldTypeSmallInteger Then
                    Me.cmbLandUseField.Items.Add(pField.Name)
                End If
            Next
            Me.cmbLandUseField.Text = "Select a Field"
        End If

        m_strLandUseLyrName = Me.cmbLandUseLayers.Text
CleanUp:
        intCount = Nothing
        blnLayerFound = Nothing
        pFeatureClass = Nothing
        pField = Nothing
        intFld = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub

    Private Sub cmbLandUseFld_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbLandUseField.SelectedIndexChanged
        '    LoadLandUseFieldValues()
        'End Sub

        'Private Sub LoadLandUseFieldValues()
        'LOAD UNIQUE LAND USE VALUES FROM THE SELECTED LAND USE LAYER AND SELECTED FIELD
        ' blnLandUseLoad = True
        If Me.cmbLandUseLayers.Text = "Select the Land Use Feature Layer" Or m_blnETSetupFormIsOpening Or m_arrPolyFeatureLayers.Count <= 0 Then
            GoTo CleanUp
        Else
            Me.dgvLandUseAttributes.Rows.Clear()
            Me.dgvLandUseAttributes.Refresh()
        End If

        Dim pDataStatistics As IDataStatistics
        Dim pCursor As ICursor = Nothing
        Dim pEnumVar As System.Collections.IEnumerator = Nothing
        Dim intUniqueCount As Integer
        Dim pFeatLyr As IFeatureLayer
        Dim pTable As ITable
        Dim blnMoved As Boolean
        Dim objValue As Object
        Dim strValue As String
        'Dim arrValues As ArrayList
        Dim pRow As Row
        Dim dic As New SortedList(Of Object, String)
        Dim intField As Integer
        Dim intCount As Integer

        Try
            pFeatLyr = CType(m_arrPolyFeatureLayers.Item(Me.cmbLandUseLayers.SelectedIndex), IFeatureLayer)
            pTable = CType(pFeatLyr.FeatureClass, ITable)
            pDataStatistics = New DataStatistics
            pCursor = CType(pTable.Search(Nothing, False), ICursor)
            pDataStatistics.Field = Me.cmbLandUseField.Text
            pDataStatistics.Cursor = pCursor
            'pEnumVar = CType(pDataStatistics.UniqueValues, System.Collections.IEnumerator)
            intUniqueCount = pDataStatistics.UniqueValueCount()
            If intUniqueCount > 50 Then
                If MessageBox.Show("There are " + intUniqueCount.ToString + " unique values in the selected field.  Would you like to continue to load them all?", "Significant Number of Values", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2, MessageBoxOptions.DefaultDesktopOnly, False) = Windows.Forms.DialogResult.No Then
                    GoTo CleanUp
                End If
            End If
        Catch ex As Exception
        End Try

        '***************************************************************************
        'PUT THIS CODE BACK IN FOR 10.2 IF FIX IS IN
        '***************************************************************************
        'Try
        '    If intUniqueCount > 0 And Not pEnumVar Is Nothing Then
        '        blnMoved = pEnumVar.MoveNext
        '        Do Until blnMoved = False
        '            objValue = pEnumVar.Current
        '            strValue = objValue.ToString
        '            Me.dgvLandUseAttributes.Rows.Add()
        '            Me.dgvLandUseAttributes.Rows(Me.dgvLandUseAttributes.RowCount - 1).Cells(0).Value = strValue
        '            Me.dgvLandUseAttributes.Rows(Me.dgvLandUseAttributes.RowCount - 1).Cells(2).Value = ""
        '            Me.dgvLandUseAttributes.Rows(Me.dgvLandUseAttributes.RowCount - 1).Cells(3).Value = "TRUE"
        '            blnMoved = pEnumVar.MoveNext
        '        Loop
        '        Me.dgvLandUseAttributes.Refresh()
        '    End If
        'Catch ex As Exception
        'End Try
        '***************************************************************************
        'WORKAROUND
        '***************************************************************************
        Try
            If Not pCursor Is Nothing Then
                pRow = CType(pCursor.NextRow(), Row)
                intField = pRow.Fields.FindField(Me.cmbLandUseField.Text)

                'loop through the data and get the unique values
                Do Until pRow Is Nothing
                    Try
                        strValue = pRow.Value(intField)
                        If dic.ContainsKey(strValue) = False Then
                            dic.Add(strValue, "")
                            If dic.Count = intUniqueCount Then
                                Exit Do
                            End If
                        End If
                    Catch ex As Exception
                    End Try
                    pRow = CType(pCursor.NextRow(), Row)
                Loop

                If dic.Count > 0 Then
                    For intCount = 0 To dic.Count - 1
                        strValue = dic.Keys.Item(intCount)
                        Me.dgvLandUseAttributes.Rows.Add()
                        Me.dgvLandUseAttributes.Rows(Me.dgvLandUseAttributes.RowCount - 1).Cells(0).Value = strValue
                        Me.dgvLandUseAttributes.Rows(Me.dgvLandUseAttributes.RowCount - 1).Cells(2).Value = ""
                        Me.dgvLandUseAttributes.Rows(Me.dgvLandUseAttributes.RowCount - 1).Cells(3).Value = "TRUE"
                    Next
                End If
            End If
        Catch ex As Exception
            MessageBox.Show("Error is building a list of unique land use types." & vbNewLine & ex.Message, "Land Use Value Error")
        End Try


        GoTo CleanUp

CleanUp:
        pDataStatistics = Nothing
        pCursor = Nothing
        pEnumVar = Nothing
        intUniqueCount = Nothing
        pFeatLyr = Nothing
        pTable = Nothing
        blnMoved = Nothing
        objValue = Nothing
        pRow = Nothing
        dic = Nothing
        intField = Nothing
        strValue = Nothing
        GC.Collect()
        GC.WaitForPendingFinalizers()
    End Sub

    Private Sub mnuSetDeveloped_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSetDeveloped.Click
        'SET ALL LANDUSES TO DEVELOPED
        Dim intRow As Integer
        For intRow = 0 To Me.dgvLandUseAttributes.Rows.Count - 1
            Me.dgvLandUseAttributes.Rows(intRow).Cells(2).Value = "TRUE"
            Me.dgvLandUseAttributes.Rows(intRow).Cells(3).Value = ""
        Next
    End Sub

    Private Sub mnuSetVacant_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSetVacant.Click
        'SET ALL LANDUSES TO VACANT
        Dim intRow As Integer
        For intRow = 0 To Me.dgvLandUseAttributes.Rows.Count - 1
            Me.dgvLandUseAttributes.Rows(intRow).Cells(2).Value = ""
            Me.dgvLandUseAttributes.Rows(intRow).Cells(3).Value = "TRUE"
        Next
    End Sub

    Private Sub mnuAllConstrained_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAllConstrained.Click
        'SET ALL LAND USE TYPES TO HAVE CONSTRAINTS APPLIED
        Dim intRow As Integer
        For intRow = 0 To Me.dgvLandUseAttributes.Rows.Count - 1
            Me.dgvLandUseAttributes.Rows(intRow).Cells(8).Value = 1
        Next
    End Sub

    Private Sub mnuAllUnconstrained_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAllUnconstrained.Click
        'SET ALL LAND USE TYPES TO NOT HAVE CONSTRAINTS APPLIED
        Dim intRow As Integer
        For intRow = 0 To Me.dgvLandUseAttributes.Rows.Count - 1
            Me.dgvLandUseAttributes.Rows(intRow).Cells(8).Value = 0
        Next
    End Sub

    Private Sub mnuClearAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuClearAll.Click
        'CLEAR ALL LANDUSES
        Dim intRow As Integer
        For intRow = 0 To Me.dgvLandUseAttributes.Rows.Count - 1
            Me.dgvLandUseAttributes.Rows(intRow).Cells(2).Value = ""
            Me.dgvLandUseAttributes.Rows(intRow).Cells(3).Value = ""
        Next
    End Sub

    Private Sub mnuClearDeveloped_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuClearDeveloped.Click
        'CLEAR ALL DEVELOPED LANDUSES
        Dim intRow As Integer
        For intRow = 0 To Me.dgvLandUseAttributes.Rows.Count - 1
            Me.dgvLandUseAttributes.Rows(intRow).Cells(2).Value = ""
        Next
    End Sub

    Private Sub mnuClearVacant_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuClearVacant.Click
        'CLEAR ALL VACANT LANDUSES
        Dim intRow As Integer
        For intRow = 0 To Me.dgvLandUseAttributes.Rows.Count - 1
            Me.dgvLandUseAttributes.Rows(intRow).Cells(3).Value = ""
        Next
    End Sub

    Private Sub dgvLandUseAttributes_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dgvLandUseAttributes.MouseDown
        Dim hit As DataGridView.HitTestInfo = Me.dgvLandUseAttributes.HitTest(e.X, e.Y)
        Dim strDev As String
        Dim strVac As String

        If e.Button = Windows.Forms.MouseButtons.Right Then   'right mouse button was clicked
            GoTo CleanUp
        End If

        Try
            If hit.Type = DataGridViewHitTestType.Cell Then
                If hit.ColumnIndex <= 1 Or hit.ColumnIndex >= 4 Then
                    GoTo CleanUp
                End If
                strDev = CStr(Me.dgvLandUseAttributes.Rows(hit.RowIndex).Cells(2).Value)
                strVac = CStr(Me.dgvLandUseAttributes.Rows(hit.RowIndex).Cells(3).Value)
                If hit.ColumnIndex = 2 Then
                    If strDev = "" Then
                        Me.dgvLandUseAttributes.Rows(hit.RowIndex).Cells(2).Value = "TRUE"
                        If Me.dgvLandUseAttributes.Rows(hit.RowIndex).Cells(3).Value = "TRUE" Then
                            Me.dgvLandUseAttributes.Rows(hit.RowIndex).Cells(3).Value = ""
                        End If
                    Else
                        Me.dgvLandUseAttributes.Rows(hit.RowIndex).Cells(2).Value = ""
                    End If
                Else
                    If strVac = "" Then
                        Me.dgvLandUseAttributes.Rows(hit.RowIndex).Cells(3).Value = "TRUE"
                        If Me.dgvLandUseAttributes.Rows(hit.RowIndex).Cells(2).Value = "TRUE" Then
                            Me.dgvLandUseAttributes.Rows(hit.RowIndex).Cells(2).Value = ""
                        End If
                    Else
                        Me.dgvLandUseAttributes.Rows(hit.RowIndex).Cells(3).Value = ""
                    End If
                End If
            End If

            GoTo CleanUp
        Catch ex As Exception
            GoTo CleanUp
        End Try

CleanUp:
        hit = Nothing
        strDev = Nothing
        strVac = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub
    '***************************************************************************************************************************************



    '***************************************************************************************************************************************
    'SUBS AND FUNCTIONS ASSOCIATED WITH CONTROLS ON TAB FIVE
    '***************************************************************************************************************************************
    Private Sub cmbSubaraLayers_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbSubaraLayers.SelectedIndexChanged
        'LOAD A LIST OF STRING | INTEGER | DOUBLE FIELDS INTO FIELD COMBO BOX
        If m_blnETSetupFormIsOpening Or m_arrPolyFeatureLayers.Count <= 0 Then
            Exit Sub
        End If

        Dim pField As IField
        Dim intFldCount As Integer
        Dim pFeatLayer As IFeatureLayer

        pFeatLayer = CType(m_arrPolyFeatureLayers.Item(Me.cmbSubaraLayers.SelectedIndex), IFeatureLayer)
        If pFeatLayer Is Nothing Then
            GoTo cleanup
        Else
            'CLEAR OUT PREVIOUS 
            Me.cmbSubareaFields.Items.Clear()
            Me.cmbSubareaFields.Text = ""

            For intFldCount = 1 To pFeatLayer.FeatureClass.Fields.FieldCount - 1
                pField = pFeatLayer.FeatureClass.Fields.Field(intFldCount)
                If pField.Type = esriFieldType.esriFieldTypeInteger Or pField.Type = esriFieldType.esriFieldTypeString Or pField.Type = esriFieldType.esriFieldTypeDouble Then
                    Me.cmbSubareaFields.Items.Add(pField.Name)
                End If
            Next
        End If

CleanUp:
        pField = Nothing
        intFldCount = Nothing
        pFeatLayer = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub

    Private Sub btnAddSubarea_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddSubarea.Click
        'ADD SELECTED SUBAREA LAYER INFORMATION TO DATAGRID
        Dim strLayer As String = Me.cmbSubaraLayers.Text
        Dim strField As String = Me.cmbSubareaFields.Text
        Dim strOutFieldName As String = "SUBAREA_" & (Me.dgvSubAreas.RowCount + 1).ToString
        Dim blnFound As Boolean = False
        Dim intRow As Integer
        Dim intCount As Integer

        If m_blnETSetupFormIsOpening Then
            GoTo CleanUp
        End If

        If Not Me.cmbSubaraLayers.Items.IndexOf(strLayer) >= 0 Or strLayer.Length <= 0 Then
            MessageBox.Show("Please select a Subarea layer.", "Subarea Layer", MessageBoxButtons.OK, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo CleanUp
        End If
        If Not Me.cmbSubareaFields.Items.IndexOf(strField) >= 0 Or strField.Length <= 0 Then
            MessageBox.Show("Please select a Subarea field.", "Subarea Field", MessageBoxButtons.OK, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo CleanUp
        End If
        Try
            'ENSURE LAYER AND FIELD COMBINATION DOES NOT ALREADY EXIST
            For intRow = 0 To Me.dgvSubAreas.Rows.Count - 1
                If strOutFieldName = CStr(Me.dgvSubAreas.Rows(intRow).Cells(3).Value) Then
                    blnFound = True
                End If
            Next
            'CREATE ALTERNATIVE SUBAREA FIELD NAME IF ONE IS ALREADY TAKEN
            If blnFound Then
                For intRow = 0 To Me.dgvSubAreas.Rows.Count - 1
                    If strOutFieldName = CStr(Me.dgvSubAreas.Rows(intRow).Cells(3).Value) Then
                        blnFound = True
                        Exit For
                    End If
                Next
            End If
            'CHECK THE FIELD NAME DOES NOT ALREADY EXIST
            For intCount = 1 To 50
                strOutFieldName = "SUBAREA_" & intCount.ToString
                blnFound = False
                For intRow = 0 To Me.dgvSubAreas.Rows.Count - 1
                    If strOutFieldName = CStr(Me.dgvSubAreas.Rows(intRow).Cells(3).Value) Then
                        blnFound = True
                        Exit For
                    End If
                Next
                If Not blnFound Then
                    Exit For
                End If
            Next

            'REVIEW CURRENT ROWS FOR A MATCH OF SELECTED SUBAREA INPUTS
            Me.dgvSubAreas.Rows.Add()
            Me.dgvSubAreas.Rows(Me.dgvSubAreas.RowCount - 1).Cells(1).Value = "Remove"
            Me.dgvSubAreas.Rows(Me.dgvSubAreas.RowCount - 1).Cells(1).Value = strLayer
            Me.dgvSubAreas.Rows(Me.dgvSubAreas.RowCount - 1).Cells(2).Value = strField
            Me.dgvSubAreas.Rows(Me.dgvSubAreas.RowCount - 1).Cells(3).Value = strOutFieldName
            Me.dgvSubAreas.Refresh()
        Catch ex As Exception

        End Try

CleanUp:
        strLayer = Nothing
        strField = Nothing
        intRow = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub

    Private Sub dgvSubAreas_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dgvSubAreas.MouseDown
        'REMOVE THE SELECTED LINE FROM THE DATA GRID
        Dim hit As DataGridView.HitTestInfo = Me.dgvSubAreas.HitTest(e.X, e.Y)
        Dim intDeleteRow As Integer = -1


        If e.Button = Windows.Forms.MouseButtons.Right Then   'right mouse button was clicked
            GoTo CleanUp
        End If

        Try
            If hit.Type = DataGridViewHitTestType.Cell Then
                If Not hit.ColumnIndex = 0 Then
                    GoTo CleanUp
                End If
                intDeleteRow = hit.RowIndex
                Me.dgvSubAreas.Rows.RemoveAt(intDeleteRow)
            End If
            GoTo CleanUp
        Catch ex As Exception
            GoTo CleanUp
        End Try

CleanUp:
        hit = Nothing
        intDeleteRow = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub
    '***************************************************************************************************************************************



    '***************************************************************************************************************************************
    'CONTROLS THE DISPLAY OF TABS
    '***************************************************************************************************************************************
    Private Sub btnNext1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNext1.Click
        Me.InfoProjectSettings.SelectTab(1)
        Step1Check()
    End Sub

    Private Sub btnPrevious1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrevious1.Click
        Me.InfoProjectSettings.SelectTab(0)
    End Sub

    Private Sub btnNext2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNext2.Click
        Me.InfoProjectSettings.SelectTab(2)
    End Sub

    Private Sub btnPrevious2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrevious2.Click
        Me.InfoProjectSettings.SelectTab(1)
    End Sub

    Private Sub btnNext3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNext3.Click
        Me.InfoProjectSettings.SelectTab(3)
    End Sub

    Private Sub btnPrevious3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrevious3.Click
        Me.InfoProjectSettings.SelectTab(2)
    End Sub

    Private Sub btnNext4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNext4.Click
        Me.InfoProjectSettings.SelectTab(4)
    End Sub

    Private Sub btnPrevious4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrevious4.Click
        Me.InfoProjectSettings.SelectTab(3)
    End Sub

    Private Sub btnNext5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNext5.Click
        Me.InfoProjectSettings.SelectTab(5)
    End Sub

    Private Sub btnPrevious5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrevious5.Click
        Me.InfoProjectSettings.SelectTab(4)
    End Sub
    '***************************************************************************************************************************************



    '***************************************************************************************************************************************
    'FUNCTIONS WHICH REVIEW THE USER SELECTIONS FOR EACH TAB TO CONTROL THE STATUS OF THE RUN BUTTON
    '***************************************************************************************************************************************
    Public Function Step1Check() As Boolean
        'SET THE STATUS TO THE INPUTS FOR TAB 1
        Step1Check = False
        If blnWorkspace And blnProjectName And blnSourceInputs And blnExtent And blnProjection Then
            Step1Check = True
        End If
    End Function

    Public Function Step2Check() As Boolean
        'SET THE STATUS TO THE INPUTS FOR TAB 2
        Step2Check = True

        'CHECK THE SLOPE INPUTS
        blnSlope = True
        If Not Me.chkCreateSlope15to25.Checked And Not Me.chkCreateSlope25Plus.Checked Then
            blnSlope = False
        Else
            If Me.cmbGridLayers.Text.Length > 0 Then
                If Not Me.cmbGridLayers.Items.IndexOf(Me.cmbGridLayers.Text) >= 0 Then
                    blnSlope = False
                End If
            Else
                blnSlope = False
            End If
            If Me.tbxConstraintsCellSize.Text.Length > 0 Then
                If Not IsNumeric(Me.tbxConstraintsCellSize.Text) Then
                    blnSlope = False
                End If
            Else
                blnSlope = False
            End If
        End If

    End Function

    Public Function Step3Check() As Boolean
        'SET THE STATUS TO THE INPUTS FOR TAB 3
        Step3Check = True

    End Function

    Public Function Step4Check() As Boolean
        'SET THE STATUS TO THE INPUTS FOR TAB 4
        Step4Check = True
    End Function

    Public Sub RunButtonStatus()
        'CONTROLS THE ENABLED STATUS OF THE RUN BUTTON
        'EXECUTES A FUNCTION TO REVIEW USER INPUTS ON EACH TAB
        Dim blnTab1 As Boolean = False
        Dim blnTab2 As Boolean = True
        Dim blnTab3 As Boolean = True
        Dim blnTab4 As Boolean = True

        blnTab1 = Step1Check()
        blnTab2 = Step2Check()
        blnTab3 = Step3Check()
        blnTab4 = Step4Check()
        If blnTab1 And blnTab2 And blnTab3 And blnTab4 Then
            Me.btnFinalRun.Enabled = True
        Else
            Me.btnFinalRun.Enabled = False
        End If
    End Sub
    '***************************************************************************************************************************************



    '***************************************************************************************************************************************
    'EXECUTE SUBS AND FUNCTIONS TO CREATE THE ENVISION PROJECT 
    '***************************************************************************************************************************************
    Public Sub CalCulateGridCellCount()
        'CALCULATE THE TOTAL NUMBER OF GRID CELLS WILL RESULTS FROM THE CURRENTLY SELECTED OPTIONS
        Dim intNumRow As Integer = 0
        Dim intNumCol As Integer = 0
        Dim intTotalCount As Double = 0
        Dim dblGridCellSize As Double = 1
        If m_pExtentEnv Is Nothing Or m_blnETSetupFormIsOpening Then
            GoTo cleanup
        Else
            'RETRIEVE CALCULATION INPUTS
            If IsNumeric(m_frmEnvisionProjectSetup.tbxGridCellSize.Text) Then
                intNumRow = CInt(m_pExtentEnv.Width / dblGridCellSize)
                intNumCol = CInt(m_pExtentEnv.Height / dblGridCellSize)
                dblGridCellSize = CDbl(m_frmEnvisionProjectSetup.tbxGridCellSize.Text)
                MessageBox.Show(dblGridCellSize.ToString)
            Else
                'm_frmEnvisionProjectSetup.lblTotalGridCells.Text = "Numeric Value Required"
                GoTo CleanUp
            End If
            'RUN AND ASSIGN CALCULATION
            Try
                intTotalCount = CInt(intNumRow * intNumCol)
                ' m_frmEnvisionProjectSetup.lblTotalGridCells.Text = "= " & intTotalCount.ToString("#,#") & " Total Grid Cell Count"
                GoTo CleanUp
            Catch ex As Exception
                'm_frmEnvisionProjectSetup.lblTotalGridCells.Text = "Value to small"
                GoTo CleanUp
            End Try
        End If
CleanUp:
        intNumRow = Nothing
        intNumCol = Nothing
        intTotalCount = Nothing
        dblGridCellSize = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub

    Private Sub btnGenerateSlope_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'OPEN THE FORM TO CREATE THE SLOPE GRID
        Try
            If m_frmEnvisionSlope Is Nothing Then
                m_frmEnvisionSlope = New frmEnvisionSlope
            End If
            m_frmEnvisionSlope.ShowDialog()
            GoTo cleanup
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Envision Slope/Hillshade Form Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo cleanup
        End Try

CleanUp:
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub

    Private Sub btnRun_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFinalRun.Click
        'RESET THE PROCESSING TEXT, WHICH WILL BE WRITTEN TO A TEXT FILE WITHIN THE FILE GEODATABASE
        dtStartTime_ETSetup = DateTime.Now
        m_appEnvision.StatusBar.Message(0) = "Project Setup: Start"

        'CHECK THE GRID CELL SIZE AND EXTENT AREA
        'CELL SIZE AREA SHOULD BE SMALLER
        Dim pArea As IArea
        Dim blnLanduse As Boolean = False
        Dim blnConstraintsNLanduse As Boolean = False
        Dim blnSubarea As Boolean = False
        Dim intRow As Integer
        Dim pWksFactory As IWorkspaceFactory = Nothing
        Dim pFeatWks As IFeatureWorkspace
        Dim pFeatClass As IFeatureClass
        Dim pTable As ITable
        Dim oFile As System.IO.File
        Dim oWrite As System.IO.StreamWriter
        Dim blnCloseForm As Boolean = True
        Dim pLocationSel As ESRI.ArcGIS.DataManagementTools.SelectLayerByLocation
        Dim pFeatSel As IFeatureSelection
        Dim pCursor As ICursor = Nothing
        Dim arrLUfrom As ArrayList = New ArrayList
        Dim arrLUto As ArrayList = New ArrayList
        Dim arrIsVac As ArrayList = New ArrayList
        Dim arrIsDevd As ArrayList = New ArrayList
        Dim arrIsConstrained As ArrayList = New ArrayList
        Dim pDataset1 As DataSet = Nothing
        Dim strInput1 As String = ""
        Dim strInput2 As String = ""
        Dim dblMaxParcelSize As Double = 0

        'PROMPT THE USER TO CREATE PROCEED WITH THE ENVISION PROJECT SETUP
        If MessageBox.Show("Would you like to proceed with processing to create the Envision project?", "Envision Project Setup", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2, MessageBoxOptions.DefaultDesktopOnly, False) = Windows.Forms.DialogResult.No Then
            blnCloseForm = False
            GoTo CleanUp
        Else
            m_strProcessing = "PROCESSING START TIME: " & Date.Now.ToString("M/d/yyyy hh:mm:ss tt") & "---------------------------------------------------------------------------------------------" & vbNewLine
            'HIDE ALL CONTROLS EXCEPT STATUS CONTROLS
            Me.lblRunStatus.Text = "Starting Processes:"
            m_appEnvision.StatusBar.Message(0) = "Starting Processes:"
            Me.lblRunStatus.Visible = True
            Me.barStatusRun.Visible = False
            Me.btnPrevious4.Visible = False
            Me.btnFinalRun.Visible = False
            Me.itmSeparatorRun1.Visible = False
            Me.itmSeparatorRun2.Visible = False
            Me.Refresh()
        End If

        'REVIEW CELL SIZE TO ENSURE PROPER INPUT
        pArea = m_pExtentEnv
        If Me.rdbSourceGridCells.Checked Or Me.rdbSourceHybrid.Checked Then
            m_strProcessing = m_strProcessing & vbNewLine & "Review Grid Cell Size"
            m_appEnvision.StatusBar.Message(0) = "Review Grid Cell Size"
            If IsNumeric(Me.rdbSourceGridCells.Checked) Then
                If (CInt(Me.tbxGridCellSize.Text) * CInt(Me.tbxGridCellSize.Text)) > pArea.Area Then   'ASSUMPTION IS CELL SIZE AND AREA ARE IN THE SAME UNITS OF MEASURE
                    m_strProcessing = m_strProcessing & vbNewLine & "The designated cell size is larger than the currently selected extent.  Please review the inputs."
                    GoTo ProblemExit
                End If
            Else
                Me.InfoProjectSettings.SelectTab(0)
                m_strProcessing = m_strProcessing & vbNewLine & "The required grid cell size provided is not a numeric value.  Please review the inputs."
                GoTo ProblemExit
            End If
        End If

        'DEFINE A NEW GEOPROCESSOR FOR USE IN THE SETUP 
        If pGPSetup Is Nothing Then
            If Not CreateSetupGP() Then
                GoTo ProblemExit
            End If
        End If

        'CREATE THE ENVISION FILE GEODATABASE IN THE SELECTED WORKSPACE DIRECTORY
        Me.lblRunStatus.Text = "Creating Envision File Geodatabase"
        If Not CreateEnvisionFileGDB() Then
            GoTo ProblemExit
        End If

        'CREATE AREA OF INTEREST (AOI) FEAGTURE CLASS
        Me.lblRunStatus.Text = "Creating Area of Interest (AOI) Layer"
        If CreateAOIFeatClass() Then
            If Not CreateAOIFeature() Then
                GoTo CleanUp
            End If
        Else
            GoTo ProblemExit
        End If

        'CHECK PARCEL LAYER IS WITHIN AOI IF SELECTED
        If (Me.rdbSourceParcels.Checked Or Me.rdbSourceHybrid.Checked) And Not Me.rdbExtentParcel.Checked Then
            m_strProcessing = m_strProcessing & vbNewLine & "Reviewing selected parcel layer for feature(s) that are within AOI."
            pLocationSel = New ESRI.ArcGIS.DataManagementTools.SelectLayerByLocation
            pLocationSel.in_layer = m_lyrOriginalParcels
            pLocationSel.select_features = m_lyrAOI
            pLocationSel.selection_type = "NEW_SELECTION"
            pLocationSel.overlap_type = "INTERSECT"
            RunTool(pGPSetup, pLocationSel)
            pFeatSel = Nothing
            pFeatSel = CType(m_lyrOriginalParcels, IFeatureSelection)
            pFeatSel.SelectionSet.Search(Nothing, False, pCursor)
            If pFeatSel.SelectionSet.Count <= 0 Then
                m_strProcessing = m_strProcessing & vbNewLine & "***** No Parcel Features were found within the selected AOI."
                GoTo ProblemExit
            Else
                m_strProcessing = m_strProcessing & vbNewLine & pFeatSel.SelectionSet.Count.ToString & " Parcel Feature(s) were found within the selected AOI."
                Me.lblRunStatus.Text = pFeatSel.SelectionSet.Count.ToString & " Parcel Feature(s) were found within the selected AOI."
            End If
        Else
            m_strProcessing = m_strProcessing & vbNewLine & "AOI set to Parcel layer extent, all parcels set to process."
        End If

        'CREATE GRID FEATURE CLASS IF REQUIRED
        If Me.rdbSourceGridCells.Checked Or Me.rdbSourceHybrid.Checked Then
            Me.lblRunStatus.Text = "Creating Grid Cells Layer"
            If CreateGridFeatClass() Then
                If Not CreateGridCellFeatures() Then
                    GoTo ProblemExit
                End If
            Else
                GoTo ProblemExit
            End If
        End If

        'CREATE PARCELS FEATURE CLASS IF REQUIRED
        If Me.rdbSourceParcels.Checked Or Me.rdbSourceHybrid.Checked Then
            'ADD A TEMP ID FIELD TO THE ORIGINAL PARCEL LAYER TO TRACT FEATURES THROUGHT PRCOESS
            'ADD A TEMP ET ID FIELD FOR PARCEL OBJECT ID TO REFERENCE IN LATER PROCESSES
            pFeatClass = m_lyrOriginalParcels.FeatureClass
            pTable = CType(pFeatClass, ITable)

            'ADD PROCESSING FIELDS TO THE INPUT PARCEL
            Me.lblRunStatus.Text = "Adding temp field, ET_TEMP_OBJ_ID1, to the input parcel layer"
            AddEnvisionField(pTable, "ET_TEMP_OBJ_ID1", "INTEGER", 16, 0)
            Me.lblRunStatus.Text = "Calculating Temporary Parcel Ids #1"
            CalcTempObjId(pTable, "ET_TEMP_OBJ_ID1")
            Me.lblRunStatus.Text = "Adding temp field, ET_ACRES, to the input parcel layer"
            AddEnvisionField(pTable, "ET_ACRES", "DOUBLE", 16, 6)
            Me.lblRunStatus.Text = "Calculating Acres for input parcel feature(s)"
            CalAcres(pTable, "ET_ACRES")

            If Not Me.rdbExtentParcel.Checked Then
                Me.lblRunStatus.Text = "Clipping Parcel Layer to AOI"
                'EXECUTE FUNCTION TO CLIP INPUT PARCEL LAYER TO SELECTED AOI
                If Not CreateAOIClipOfParcels() Then
                    GoTo ProblemExit
                Else
                    m_lyrEnvisionParcel = m_lyrParcelClip
                End If
            Else
                m_lyrEnvisionParcel = m_lyrOriginalParcels
            End If
            pTable = Nothing
        End If

        'CREATE HYBRID GRID/PARCELS FEAGTURE CLASS IF REQUIRED
        If Me.rdbSourceHybrid.Checked And blnMaxParcelSize Then
            If IsNumeric(Me.tbxMaxParcelSize.Text) Then
                dblMaxParcelSize = CDbl(Me.tbxMaxParcelSize.Text)
                If dblMaxParcelSize < 0 Then
                    dblMaxParcelSize = 0
                    m_strProcessing = m_strProcessing & vbNewLine & "The Maximum Parcel size is less than zero and well be set to zero."
                End If
            Else
                dblMaxParcelSize = 0
                m_strProcessing = m_strProcessing & vbNewLine & "The Maximum Parcel size is NOT a numeric value and well be set to zero."
            End If
            If dblMaxParcelSize > 0 Then
                'LOOK FOR ENVISION_ACRES OR ACRES FIELD
                pFeatClass = m_lyrParcelClip.FeatureClass
                pTable = CType(pFeatClass, ITable)


                Me.lblRunStatus.Text = "Minimum parcel selection"
                'SELECT THOSE PARCELS WHICH ARE LESS THAN DESIGNATED ACRE SIZE
                If CreateMinAcresParcelFeatClass() Then
                    Me.lblRunStatus.Text = "Maximum parcel selection"
                    'SELECT THOSE PARCELS WHICH EQUAL OR GREATER MAXIMUM DESIGNATED ACRE SIZE
                    If CreateMaxAcresParcelFeatClass() Then
                        'UNION THE MAX PARCELS AND GRIDS, THEN UNION PREVIOUS UNION TO MIN PARCELS
                        If Not CreateUnionOfGridAndParcelsFeatClass() Then
                            GoTo ProblemExit
                        End If
                    End If
                Else
                    'INTERSECT GRID AND PARCEL FEATURE CLASSES
                    Me.lblRunStatus.Text = "Intersect Grid and Parcel layers"
                    If Not CreateUnionOfGridAndOrgParcelsFeatClass() Then
                        GoTo ProblemExit
                    End If
                End If
            End If
        End If

        'DEFINE m_lyrEnvisionParcel TO EQUAL THE RELAVENT FEATURE LAYER FROM PREVIOUS PROCESSES
        If Me.rdbSourceGridCells.Checked Then
            If Not m_lyrGrid Is Nothing Then
                m_lyrMainProcessingLayer = m_lyrGrid
            Else
                m_strProcessing = m_strProcessing & vbNewLine & "EXITING:  Feature Source set to Grid Cells Only, but m_lyrGrid feature layer is NULL"
                Me.lblRunStatus.Text = "EXITING:  Feature Source set to Grid Cells Only, but m_lyrGrid feature layer is NULL"
                GoTo ProblemExit
            End If
        End If

        If Me.rdbSourceParcels.Checked Then
            If Not m_lyrEnvisionParcel Is Nothing Then
                m_lyrMainProcessingLayer = m_lyrEnvisionParcel
            Else
                m_strProcessing = m_strProcessing & vbNewLine & "EXITING:  Feature Source set to Parcel Only, but m_lyrEnvisionParcel feature layer is NULL"
                Me.lblRunStatus.Text = "EXITING:  Feature Source set to Grid Cells Only, but m_lyrGrid feature layer is NULL"
                GoTo ProblemExit
            End If
        End If

        If Me.rdbSourceHybrid.Checked Then
            If Not m_lyrParcelsUnionWithGrid Is Nothing Then
                m_lyrMainProcessingLayer = m_lyrParcelsUnionWithGrid
            Else
                m_strProcessing = m_strProcessing & vbNewLine & "EXITING:  Feature Source set to Hybrid, but m_lyrParcelsUnionWithGrid feature layer is NULL"
                Me.lblRunStatus.Text = "EXITING:  Feature Source set to Grid Cells Only, but m_lyrGrid feature layer is NULL"
                GoTo ProblemExit
            End If
        End If

        'ADD REQUIRED FIELD(S)
        Try
            pTable = CType(m_lyrMainProcessingLayer.FeatureClass, IFeatureClass)
            AddEnvisionSetupRequiredFeilds(pTable)
        Catch ex As Exception
            m_strProcessing = m_strProcessing & vbNewLine & "EXITING:  Unable to retrive the table from the feature layer (m_lyrMainProcessingLayer)."
            GoTo ProblemExit
        End Try


        'BUILD ARRAY LISTS OF LAND USE INPUTS
        m_strProcessing = m_strProcessing & vbNewLine & vbNewLine & "LAND USE LAYER PROCESSING:" & vbNewLine & Date.Now.ToString("M/d/yyyy hh:mm:ss tt") & vbNewLine
        Me.lblRunStatus.Text = "LAND USE LAYER PROCESSING:" & vbNewLine & Date.Now.ToString("M/d/yyyy hh:mm:ss tt")
        For intRow = 0 To Me.dgvLandUseAttributes.Rows.Count - 1
            arrLUfrom.Add(Me.dgvLandUseAttributes.Rows(intRow).Cells(0).Value)
            arrLUto.Add(Me.dgvLandUseAttributes.Rows(intRow).Cells(1).Value)
            arrIsDevd.Add(Me.dgvLandUseAttributes.Rows(intRow).Cells(2).Value)
            arrIsVac.Add(Me.dgvLandUseAttributes.Rows(intRow).Cells(3).Value)
            arrIsConstrained.Add(Me.dgvLandUseAttributes.Rows(intRow).Cells(4).Value)
            ' MessageBox.Show(Me.dgvLandUseAttributes.Rows(intRow).Cells(0).Value.ToString & vbNewLine & Me.dgvLandUseAttributes.Rows(intRow).Cells(1).Value & vbNewLine & Me.dgvLandUseAttributes.Rows(intRow).Cells(2).Value & vbNewLine & Me.dgvLandUseAttributes.Rows(intRow).Cells(3).Value)
            If Me.dgvLandUseAttributes.Rows(intRow).Cells(2).Value = "TRUE" Or Me.dgvLandUseAttributes.Rows(intRow).Cells(3).Value = "TRUE" Then
                blnLanduse = True
            End If
            
        Next
        If blnLanduse Then
            CreateLandUseFeatClass(arrLUfrom, arrLUto, arrIsDevd, arrIsVac, arrIsConstrained)
        Else
            m_strProcessing = m_strProcessing & vbNewLine & vbNewLine & "No Land Use(s) Defined."
        End If

        Exit Sub

        'COMPRESS ALL THE CONSTRAINT LAYERS INTO A SINGLE CONSTRAINTS LAYER
        m_strProcessing = m_strProcessing & vbNewLine & vbNewLine & "CONSTRAINTS LAYER PROCESSING:" & vbNewLine & Date.Now.ToString("M/d/yyyy hh:mm:ss tt")
        m_blnConstraints = False
        For intRow = 0 To Me.dgvConstraints.Rows.Count - 1
            If Me.dgvConstraints.Rows(intRow).Cells(0).Value = "TRUE" Then
                m_blnConstraints = True
            End If
        Next
        'OUTPUT LAYER VARIABLE IS m_lyrConstraints
        If m_blnConstraints Then
            Me.lblRunStatus.Text = "Union Constraint Layer(s)"
            m_blnConstraints = CreateUnionConstraintsFeatClass()
        Else
            m_strProcessing = m_strProcessing & vbNewLine & vbNewLine & "No Constraints Defined."
        End If




        'CLEANUP THE FEATURE LAYERS NOT UTILIZED IN FOLLOWING SCRPTING
        m_lyrGrid = Nothing
        m_lyrParcelClip = Nothing
        m_lyrEnvisionParcel = Nothing
        m_lyrMaxAcreParcels = Nothing
        m_lyrMinAcreParcels = Nothing
        m_lyrParcelsUnionWithGrid = Nothing
        m_lyrTempSlope = Nothing
        m_lyrParcelConstraints = Nothing
        m_lyrLandUse = Nothing
        m_lyrConstrainedParcels = Nothing
        m_lyrUnConstrainedParcels = Nothing
        m_lyrEnvision = Nothing
        m_lyrMashup = Nothing
        m_lyrParcelCentroids = Nothing
        m_lyrLUCentroids = Nothing
        m_lyrEnvisionFinalParcels = Nothing
        GC.Collect()
        GC.WaitForPendingFinalizers()

        'CREATE THE FINAL MASH-UP LAYER OF DEFINED PARCELS AND CONSTRATINTS
        Me.lblRunStatus.Text = "Create Mash-up Layer"
        If Not m_lyrConstraints Is Nothing Then
            CreateConstraintsMashup()
        End If

        'CREATE THE FINAL PARCEL LAYER - CUT OUT THE CONSTRAINED AREAS IF REQUIRED
        ParcelLastStep()


        'CREATE THE FINAL PARCEL LAYER, WHICH MAY REQUIRE THE CONSTRAINTS CUT OUT
        CreateParcelCentroidLayer()
        Me.lblRunStatus.Text = "Creating Final Envision Parcel Layer"
        Me.Refresh()

        'RETRIEVE SUBAREA LAYER ATTRIBTUES


        'COPY EXCEL FILE TO THE ENVISION FILE GEODATABASE DIRECTORY
        Me.lblRunStatus.Text = "Copying Envision Excel"
        m_strProcessing = m_strProcessing & vbNewLine & "COPY Selected excel file to selected workspace directory:" & vbNewLine & Date.Now.ToString("M/d/yyyy hh:mm:ss tt")
        If Me.tbxExcel.Text.Length > 0 Then
            m_strProcessing = m_strProcessing & vbNewLine & "Source Envision Excel File: " & Me.tbxExcel.Text
            If CopyEnvisionExcelfile() Then
            End If
        Else
            m_strProcessing = m_strProcessing & vbNewLine & "NO ENVSION EXCEL FILE DESIGNATED" & vbNewLine & Date.Now.ToString("M/d/yyyy hh:mm:ss tt")
        End If
        GoTo CleanUp


ProblemExit:
        MessageBox.Show("A problem has occured.  Please review the following processing file for details:" & vbNewLine & (m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\EnvisionSetup_Processing.txt"), "Envision Setup Failed to Complete", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
        blnCloseForm = False
        GoTo CleanUp

CleanUp:
        'WRITE THE PROCESSING TEXT TO FILE FOR REVIEW 
        oWrite = oFile.CreateText(m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\EnvisionSetup_Processing_" & Me.tbxProjectName.Text & ".txt")
        oWrite.Write(m_strProcessing)
        oWrite.WriteLine("Processing End Time: " & Date.Now.ToString("M/d/yyyy hh:mm:ss tt"))
        oWrite.Close()

        Me.lblRunStatus.Visible = False
        Me.barStatusRun.Visible = False
        Me.btnPrevious4.Visible = True
        Me.btnFinalRun.Visible = True
        Me.itmSeparatorRun1.Visible = True
        Me.itmSeparatorRun2.Visible = True
        Me.Refresh()

        If blnCloseForm Then
            MessageBox.Show("Envision setup has completed its processing.  Please review the following processing file for details:" & vbNewLine & (m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\EnvisionSetup_Processing.txt"), "Envision Setup", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            Me.Close()
        End If

        'CLEAR ALL FUNCTION OBJECTS
        pArea = Nothing
        blnLanduse = Nothing
        blnSubarea = Nothing
        intRow = Nothing
        pWksFactory = Nothing
        pFeatWks = Nothing
        pFeatClass = Nothing
        pTable = Nothing
        oFile = Nothing
        oWrite = Nothing
        pGPSetup = Nothing
        pLocationSel = Nothing
        pFeatSel = Nothing
        pCursor = Nothing
        'CLEAR ALL MODULAR OBJECTS USED BY SETUP SCRIPTING
        m_lyrAOI = Nothing
        m_lyrGrid = Nothing
        m_lyrConstrainedParcels = Nothing
        m_lyrMaxAcreParcels = Nothing
        m_lyrMinAcreParcels = Nothing
        m_lyrParcelClip = Nothing
        m_lyrParcelsUnionWithGrid = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub

    Private Function AddEnvisionSetupRequiredFeilds(ByVal pTable As ITable) As Boolean
        'ADD THE FIELDS THAT WILL BE USED TO DURING THE ENVISION SETUP PROCESSES 
        AddEnvisionSetupRequiredFeilds = True

        Try
            'CALCUALTE AN NEW TEMPORARY ID FIELD FOR WHAT MAY BE NEW FEATURES
            AddEnvisionField(pTable, "ET_TEMP_OBJ_ID2", "INTEGER", 16, 0)
            CalcTempObjId(pTable, "ET_TEMP_OBJ_ID2")
            If pTable.FindField("ET_ACRES") <= -1 Then
                AddEnvisionField(pTable, "ET_ACRES", "DOUBLE", 16, 6)
            End If
            If pTable.FindField("EX_LU") <= -1 Then
                AddEnvisionField(pTable, "EX_LU", "STRING", 50, 0)
                CalcFldValues(pTable, "EX_LU", """""", "STRING")
            End If
            If pTable.FindField("VAC_ACRE") <= -1 Then
                AddEnvisionField(pTable, "VAC_ACRE", "DOUBLE", 16, 6)
                CalcFldValues(pTable, "VAC_ACRE", "0", "DOUBLE")
            End If
            If pTable.FindField("DEVD_ACRE") <= -1 Then
                AddEnvisionField(pTable, "DEVD_ACRE", "DOUBLE", 16, 6)
                CalcFldValues(pTable, "DEVD_ACRE", "0", "DOUBLE")
            End If
            If pTable.FindField("CONSTRAINED_ACRE") <= -1 Then
                AddEnvisionField(pTable, "CONSTRAINED_ACRE", "DOUBLE", 16, 6)
                CalcFldValues(pTable, "CONSTRAINED_ACRE", "0", "DOUBLE")
            End If
            If pTable.FindField("ET_CENTROID_X") <= -1 Then
                AddEnvisionField(pTable, "ET_CENTROID_X", "DOUBLE", 16, 6)
            End If
            If pTable.FindField("ET_CENTROID_Y") <= -1 Then
                AddEnvisionField(pTable, "ET_CENTROID_Y", "DOUBLE", 16, 6)
            End If
            If pTable.FindField("APPLY_CONSTRAINTS") <= -1 Then
                AddEnvisionField(pTable, "APPLY_CONSTRAINTS", "INTEGER", 1, 0)
                CalcFldValues(pTable, "APPLY_CONSTRAINTS", "1", "INTEGER")
            End If
        Catch ex As Exception
            AddEnvisionSetupRequiredFeilds = False
        End Try
Cleanup:

    End Function

    Private Sub DeleteNonParcelFields(ByVal pFeatclass As IFeatureClass)
        'DELETE ANY FIELDS THAT MIGHT HAVE BEEN ADDED FROM SETUP PROCESS, KEEPING NEW REQUIRED FIELD(S)
        'NEW Required Field(s): "EX_LU"
        Dim intCount As Integer
        Dim strFieldName As String
        Dim arrFieldNames As ArrayList = New ArrayList
        Dim pField As IField

        If Not pFeatclass Is Nothing Then
            If pFeatclass.Fields.FieldCount > 2 Then
                For intCount = 1 To pFeatclass.Fields.count - 1
                    If Not arrFieldNames.Add(pFeatclass.Fields.Field(intCount).Name) = "EX_LU" Then
                        arrFieldNames.Add(pFeatclass.Fields.Field(intCount).Name)
                    End If
                Next
            End If
        End If
        If arrFieldNames.Count > 0 Then
            For Each strFieldName In arrFieldNames
                Try
                    pField = pFeatclass.Fields.Field(pFeatclass.FindField(strFieldName))
                    If Not pField Is Nothing Then
                        pFeatclass.DeleteField(pField)
                    End If
                Catch ex As Exception
                End Try
            Next
        End If
        GoTo CleanUp

CleanUp:
        intCount = Nothing
        strFieldName = Nothing
        arrFieldNames = Nothing
        pField = Nothing
        GC.Collect()
        GC.WaitForPendingFinalizers()

    End Sub

    Private Sub BuildMatchFieldLists()
        'POPULATES THE FROM AND TO ARRAY LISTS FROM USER INPUTS IN FIELD MATCH TAB
        Dim intRow As Integer = 0
        Dim strFromFldName As String = ""
        Dim strToFldName As String = ""
        Dim pTable As ITable

        'RETREIVE PARCEL LAYER
        If Not m_lyrConstrainedParcels Is Nothing Then
            Try
                pTable = CType(m_lyrConstrainedParcels.FeatureClass, ITable)
            Catch ex As Exception
                m_strProcessing = m_strProcessing & vbNewLine & "BuildMatchFieldLists sub error: " & ex.Message
                GoTo CleanUp
            End Try
        Else
            GoTo CleanUp
        End If

        If Me.dgvFieldMatch.RowCount > 0 Then
            For intRow = 0 To Me.dgvFieldMatch.RowCount - 1
                strFromFldName = Me.dgvFieldMatch.Rows(Me.dgvFieldMatch.RowCount - 1).Cells(0).Value
                strToFldName = Me.dgvFieldMatch.Rows(Me.dgvFieldMatch.RowCount - 1).Cells(1).Value
                If strFromFldName.Length > 0 And strToFldName.Length > 0 Then
                    'ENSURE THE FROM FIELD IS VALID AND IN THE PARCEL LAYER
                    If pTable.FindField(strFromFldName) >= 0 Then
                        'ADD THE TO FIELD IF NOT PRESENT IN THE PARCEL LAYER
                        If pTable.FindField(strFromFldName) = -1 Then
                            AddEnvisionField(pTable, strToFldName, "DOUBLE", 16, 6)
                        End If
                        'IF BOTH FROM AND TO FIELDS ARE PRESENT, THEN ADD TO LISTS
                        If pTable.FindField(strToFldName) >= 0 And pTable.FindField(strFromFldName) >= 0 Then
                            arrMatchFrom.Add(strFromFldName)
                            arrMatchTo.Add(strToFldName)
                        End If
                    End If
                End If
            Next
        End If

CleanUp:

        GC.WaitForPendingFinalizers()
        GC.Collect()

    End Sub

    Private Function CreateEnvisionFileGDB() As Boolean
        'CREATES THE ENVISION FILE GEODATABASE
        m_strProcessing = m_strProcessing & vbNewLine & "Starting function CreateEnvisionFileGDB: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
        CreateEnvisionFileGDB = True
        Dim pCreateFileGDB As ESRI.ArcGIS.DataManagementTools.CreateFileGDB
        Try
            If pGPSetup Is Nothing Then
                CreateSetupGP()
            End If

            'VARIFY PROJECT DIRECTORY
            If m_frmEnvisionProjectSetup.tbxWorkspace.Text.Length <= 0 Then
                m_strProcessing = m_strProcessing & vbNewLine & "No workspace directory has been designated."
                CreateEnvisionFileGDB = False
                GoTo CleanUp
            Else
                If Not Directory.Exists(m_frmEnvisionProjectSetup.tbxWorkspace.Text) Then
                    m_strProcessing = m_strProcessing & vbNewLine & "The desingated workspace directory does not exist.  Please review inputs.  Workspace directory, " & m_frmEnvisionProjectSetup.tbxWorkspace.Text
                    CreateEnvisionFileGDB = False
                    GoTo CleanUp
                End If
            End If

            'CREATE FILE GEODATABASE
            If Not Directory.Exists((m_frmEnvisionProjectSetup.tbxWorkspace.Text + "\" + Me.tbxProjectName.Text)) Then
                pCreateFileGDB = New ESRI.ArcGIS.DataManagementTools.CreateFileGDB
                pCreateFileGDB.out_folder_path = m_frmEnvisionProjectSetup.tbxWorkspace.Text
                pCreateFileGDB.out_name = Me.tbxProjectName.Text
                RunTool(pGPSetup, pCreateFileGDB)
                m_strProcessing = m_strProcessing & vbNewLine & "The file geodatabase, " & (m_frmEnvisionProjectSetup.tbxWorkspace.Text + "\" & Me.tbxProjectName.Text) & ", has been created."
                pCreateFileGDB = Nothing
            Else
                m_strProcessing = m_strProcessing & vbNewLine & "The file geodatabase, " & (m_frmEnvisionProjectSetup.tbxWorkspace.Text + "\" & Me.tbxProjectName.Text) & ", already exists."
            End If
        Catch ex As Exception
            CreateEnvisionFileGDB = False
            m_strProcessing = m_strProcessing & vbNewLine & "Error is creating the Envision File geodatabase." & vbNewLine & ex.Message & vbNewLine & vbNewLine
            GoTo CleanUp
        End Try

CleanUp:
        m_strProcessing = m_strProcessing & vbNewLine & "Ending function CreateEnvisionFileGDB: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
        pCreateFileGDB = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Function

    Private Function CreateAOIFeatClass() As Boolean
        CreateAOIFeatClass = True
        'CREATE THE AOI FEATURE CLASS, WHICH CONTAIN THE AOI RECTANGLE FEATURE
        'WITH JUST ONE FEATURE REPRESENTING THE SELECT PROCESSING EXTENT FOR THE ENVISION PROJECT
        m_strProcessing = m_strProcessing & vbNewLine & "Starting function CreateAOIFeatClass: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
        Dim pCreateFeatClass As ESRI.ArcGIS.DataManagementTools.CreateFeatureclass
        Dim pWksFactory As IWorkspaceFactory = Nothing
        Dim pFeatWks As IFeatureWorkspace
        Dim pFeatClass As IFeatureClass
        Try
            If pGPSetup Is Nothing Then
                CreateSetupGP()
            End If
            pCreateFeatClass = New ESRI.ArcGIS.DataManagementTools.CreateFeatureclass
            pCreateFeatClass.out_path = m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb"
            pCreateFeatClass.out_name = "PROJECT_AOI"
            pCreateFeatClass.spatial_reference = m_pETSpatRefProject
            RunTool(pGPSetup, pCreateFeatClass)
            pCreateFeatClass = Nothing

            'DEFINE THE AOI LAYER
            pWksFactory = New FileGDBWorkspaceFactory
            pFeatWks = pWksFactory.OpenFromFile(m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb", 0)
            pFeatClass = pFeatWks.OpenFeatureClass("PROJECT_AOI")
            If pFeatClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                m_lyrAOI = New FeatureLayer
                m_lyrAOI.FeatureClass = pFeatClass
                m_lyrAOI.Name = m_lyrAOI.Name & "<SETUP> " & "PROJECT_AOI"
            End If
        Catch ex As Exception
            m_strProcessing = m_strProcessing & vbNewLine & "Error in creating AOI featureclass (PROJECT_AOI): " & vbNewLine & ex.Message
            CreateAOIFeatClass = False
            GoTo CleanUp
        End Try

CleanUp:
        m_strProcessing = m_strProcessing & vbNewLine & "End function CreateAOIFeatClass: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
        pCreateFeatClass = Nothing
        pWksFactory = Nothing
        pFeatWks = Nothing
        pFeatClass = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Function

    Private Function CreateAOIFeature() As Boolean
        'GENERATE THE GRID CELL POLYGONS
        m_strProcessing = m_strProcessing & vbNewLine & "Starting function CreateAOIFeature: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
        CreateAOIFeature = True

        Dim pntLowerLeft As ESRI.ArcGIS.Geometry.IPoint = New ESRI.ArcGIS.Geometry.Point
        Dim pntUpperRight As ESRI.ArcGIS.Geometry.IPoint = New ESRI.ArcGIS.Geometry.Point
        Dim intCellSize As Integer = 0
        Dim pInsertFeatureBuffer As IFeatureBuffer
        Dim pInsertFeatureCursor As IFeatureCursor
        Dim intNewFeatureCount As Integer = 0
        Dim pEnvisionFClass As IFeatureClass
        Dim pPolyTemp As IPolygon
        Dim pPointCollection As IPointCollection
        Dim pPointTemp As ESRI.ArcGIS.Geometry.Point
        Dim pNewFeat As IFeature
        Dim pArea As IArea
        Dim intRecCount As Integer = 0

        Try
            'RETRIEVE GRID CELL LAYER FEATURE CLASS
            If m_lyrAOI Is Nothing Then
                CreateAOIFeature = False
                GoTo CleanUp
            Else
                pEnvisionFClass = m_lyrAOI.FeatureClass
            End If

            'CHECK THE SELECTED EXTENT
            If m_pExtentEnv Is Nothing Then
                CreateAOIFeature = False
                GoTo CleanUp
            End If

            'FEATURE BUFFER SETUP
            pInsertFeatureCursor = pEnvisionFClass.Insert(True)
            pInsertFeatureBuffer = pEnvisionFClass.CreateFeatureBuffer

            Me.barStatusRun.Visible = False

            pPolyTemp = New Polygon
            pPointCollection = CType(pPolyTemp, IPointCollection)

            pPointTemp = New ESRI.ArcGIS.Geometry.Point
            pPointTemp.X = m_pExtentEnv.XMin
            pPointTemp.Y = m_pExtentEnv.YMax
            pPointCollection.AddPoint(pPointTemp)

            pPointTemp = New ESRI.ArcGIS.Geometry.Point
            pPointTemp.X = m_pExtentEnv.XMax
            pPointTemp.Y = m_pExtentEnv.YMax
            pPointCollection.AddPoint(pPointTemp)

            pPointTemp = New ESRI.ArcGIS.Geometry.Point
            pPointTemp.X = m_pExtentEnv.XMax
            pPointTemp.Y = m_pExtentEnv.YMin
            pPointCollection.AddPoint(pPointTemp)

            pPointTemp = New ESRI.ArcGIS.Geometry.Point
            pPointTemp.X = m_pExtentEnv.XMin
            pPointTemp.Y = m_pExtentEnv.YMin
            pPointCollection.AddPoint(pPointTemp)

            pPolyTemp = CType(pPointCollection, IPolygon)
            pPolyTemp.Close()

            intRecCount = intRecCount + 1
            If Not pPolyTemp Is Nothing Then
                pInsertFeatureBuffer.Shape = pPolyTemp
                pInsertFeatureCursor.InsertFeature(pInsertFeatureBuffer)
            End If
            pPointTemp = Nothing
            pPointCollection = Nothing
            pPolyTemp = Nothing
            pNewFeat = Nothing
            pArea = Nothing
            GC.Collect()
            GC.WaitForPendingFinalizers()
            pInsertFeatureCursor.Flush()

        Catch ex As Exception
            m_strProcessing = m_strProcessing & vbNewLine & "Error in creating the AOI feature: " & ex.Message
            CreateAOIFeature = False
            GoTo CleanUp
        End Try

CleanUp:
        m_strProcessing = m_strProcessing & vbNewLine & "End function CreateAOIFeature: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
        pntLowerLeft = Nothing
        pntUpperRight = Nothing
        intCellSize = Nothing
        pInsertFeatureBuffer = Nothing
        pInsertFeatureCursor = Nothing
        intNewFeatureCount = Nothing
        pEnvisionFClass = Nothing
        pPolyTemp = Nothing
        pPointCollection = Nothing
        pPointTemp = Nothing
        pNewFeat = Nothing
        pArea = Nothing
        intRecCount = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Function

    Private Function CreateGridFeatClass() As Boolean
        CreateGridFeatClass = True
        m_strProcessing = m_strProcessing & vbNewLine & "Starting function CreateGridFeatClass: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
        'CREATE THE GRID FEATURE CLASS, WHICH CONTAIN THE GRID FEATURE(S)
        Dim pCreateFeatClass As ESRI.ArcGIS.DataManagementTools.CreateFeatureclass
        Dim pWksFactory As IWorkspaceFactory = Nothing
        Dim pFeatWks As IFeatureWorkspace
        Dim pFeatClass As IFeatureClass
        Try
            If pGPSetup Is Nothing Then
                CreateSetupGP()
            End If

            pCreateFeatClass = New ESRI.ArcGIS.DataManagementTools.CreateFeatureclass
            pCreateFeatClass.out_path = m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb"
            pCreateFeatClass.out_name = "PROJECT_GRID"
            pCreateFeatClass.spatial_reference = m_pETSpatRefProject
            RunTool(pGPSetup, pCreateFeatClass)
            pCreateFeatClass = Nothing

            'DEFINE THE GRID LAYER
            pWksFactory = New FileGDBWorkspaceFactory
            pFeatWks = pWksFactory.OpenFromFile(m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb", 0)
            pFeatClass = pFeatWks.OpenFeatureClass("PROJECT_GRID")
            If pFeatClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                m_lyrGrid = New FeatureLayer
                m_lyrGrid.FeatureClass = pFeatClass
                m_lyrGrid.Name = m_lyrGrid.Name & "<SETUP> " & "GRID LAYER"
            End If
        Catch ex As Exception
            m_strProcessing = m_strProcessing & vbNewLine & "Error in creating grid feature class (PROJECT_GRID): " & ex.Message
            CreateGridFeatClass = False
            GoTo CleanUp
        End Try
CleanUp:
        m_strProcessing = m_strProcessing & vbNewLine & "End function CreateGridFeatClass: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
        pCreateFeatClass = Nothing
        pWksFactory = Nothing
        pFeatWks = Nothing
        pFeatClass = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Function

    Private Function CreateGridCellFeatures() As Boolean
        'GENERATE THE GRID CELL POLYGONS
        m_strProcessing = m_strProcessing & vbNewLine & "Starting function CreateGridCellFeatures: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
        CreateGridCellFeatures = True

        Dim pEnvisionFClass As IFeatureClass
        Dim intNumRow As Integer = 0
        Dim intNumCol As Integer = 0
        Dim intCellSize As Integer = 0
        Dim pntLowerLeft As ESRI.ArcGIS.Geometry.IPoint = New ESRI.ArcGIS.Geometry.Point
        Dim pntUpperRight As ESRI.ArcGIS.Geometry.IPoint = New ESRI.ArcGIS.Geometry.Point
        Dim dblX As Double = 0
        Dim dblY As Double = 0
        Dim dblBaseX As Double = 0
        Dim pGridFeatClass As IFeatureClass
        Dim intTotalCells As Integer = (intNumRow * intNumCol)
        Dim intTotalCount As Integer = -1
        Dim pInsertFeatureBuffer As IFeatureBuffer
        Dim pInsertFeatureCursor As IFeatureCursor = Nothing
        Dim intNewFeatureCount As Integer = 0
        Dim intRow As Integer = 0
        Dim intCol As Integer = 0
        Dim pEnvelope As IEnvelope
        Dim pPolyTemp As IPolygon
        Dim pPointCollection As IPointCollection
        Dim pPointTemp As ESRI.ArcGIS.Geometry.Point
        'Dim pArea As IArea
        Dim intRecCount As Integer = 0

        Try
            'RETRIEVE GRID CELL LAYER FEATURE CLASS
            If m_lyrGrid Is Nothing Then
                m_strProcessing = m_strProcessing & vbNewLine & "Unable to create grid feature(s) as grid layer is null."
                CreateGridCellFeatures = False
                GoTo CleanUp
            Else
                pEnvisionFClass = m_lyrGrid.FeatureClass
            End If

            'CHECK THE SELECTED EXTENT
            If m_pExtentEnv Is Nothing Then
                m_strProcessing = m_strProcessing & vbNewLine & "Unable to create grid feature(s) as AOI extent is null."
                CreateGridCellFeatures = False
                GoTo CleanUp
            End If

            'SET THE UPPER AND LOWER POINTS FROM THE EXTENT
            intCellSize = CInt(Me.tbxGridCellSize.Text)
            pntLowerLeft.X = m_pExtentEnv.XMin
            pntLowerLeft.Y = m_pExtentEnv.YMin
            pntUpperRight.X = m_pExtentEnv.XMax
            pntUpperRight.Y = m_pExtentEnv.YMax
            dblX = pntLowerLeft.X
            dblY = pntLowerLeft.Y
            dblBaseX = pntLowerLeft.Y
            intNumRow = CInt(m_pExtentEnv.Width / intCellSize)
            intNumCol = CInt(m_pExtentEnv.Height / intCellSize)

            'RETRIEVE GRID FEATURE CLASS
            pGridFeatClass = m_lyrGrid.FeatureClass

            intTotalCells = (intNumRow * intNumCol)
            intTotalCount = -1
            Me.barStatusRun.Visible = False
            For intRow = 1 To intNumRow
                For intCol = 1 To intNumCol
                    intTotalCount = intTotalCount + 1

                    pEnvelope = New Envelope
                    pEnvelope.XMin = dblX
                    pEnvelope.YMin = dblY
                    pEnvelope.XMax = dblX + intCellSize
                    pEnvelope.YMax = dblY + intCellSize

                    pPolyTemp = New Polygon
                    pPointCollection = CType(pPolyTemp, IPointCollection)

                    pPointTemp = New ESRI.ArcGIS.Geometry.Point
                    pPointTemp.X = pEnvelope.XMin
                    pPointTemp.Y = pEnvelope.YMax
                    pPointCollection.AddPoint(pPointTemp)

                    pPointTemp = New ESRI.ArcGIS.Geometry.Point
                    pPointTemp.X = pEnvelope.XMax
                    pPointTemp.Y = pEnvelope.YMax
                    pPointCollection.AddPoint(pPointTemp)

                    pPointTemp = New ESRI.ArcGIS.Geometry.Point
                    pPointTemp.X = pEnvelope.XMax
                    pPointTemp.Y = pEnvelope.YMin
                    pPointCollection.AddPoint(pPointTemp)

                    pPointTemp = New ESRI.ArcGIS.Geometry.Point
                    pPointTemp.X = pEnvelope.XMin
                    pPointTemp.Y = pEnvelope.YMin
                    pPointCollection.AddPoint(pPointTemp)

                    pPolyTemp = CType(pPointCollection, IPolygon)
                    pPolyTemp.Close()

                    'FEATURE BUFFER SETUP
                    pInsertFeatureCursor = pGridFeatClass.Insert(True)
                    pInsertFeatureBuffer = pGridFeatClass.CreateFeatureBuffer

                    intRecCount = intRecCount + 1
                    If Not pPolyTemp Is Nothing Then
                        pInsertFeatureBuffer.Shape = pPolyTemp
                        pInsertFeatureCursor.InsertFeature(pInsertFeatureBuffer)
                        dblY = dblY + intCellSize
                        intNewFeatureCount = intNewFeatureCount + 1
                        If intNewFeatureCount = 100 Then
                            pInsertFeatureCursor.Flush()
                            Me.lblRunStatus.Text = "Cell " + CStr(intTotalCount + 1) + " of " + intTotalCells.ToString
                            Me.Refresh()
                            intNewFeatureCount = 0
                        End If
                    End If
                    pEnvelope = Nothing
                    pPointTemp = Nothing
                    pPointCollection = Nothing
                    pPolyTemp = Nothing
                    GC.Collect()
                    GC.WaitForPendingFinalizers()
                Next
                dblY = dblBaseX
                dblX = dblX + intCellSize
            Next
            If Not pInsertFeatureCursor Is Nothing Then
                pInsertFeatureCursor.Flush()
            End If

        Catch ex As Exception
            m_strProcessing = m_strProcessing & vbNewLine & "Error in creating grid features.  Failed at row " & intNumRow.ToString & ", column " & intNumCol.ToString & "." & vbNewLine & ex.Message
            CreateGridCellFeatures = False
            GoTo CleanUp
        End Try

CleanUp:
        m_strProcessing = m_strProcessing & vbNewLine & "End function CreateGridCellFeatures: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
        pntLowerLeft = Nothing
        pntUpperRight = Nothing
        dblX = Nothing
        dblY = Nothing
        dblBaseX = Nothing
        intNumRow = Nothing
        intNumCol = Nothing
        intCellSize = Nothing
        pInsertFeatureBuffer = Nothing
        pInsertFeatureCursor = Nothing
        intNewFeatureCount = Nothing
        pEnvisionFClass = Nothing
        intRow = Nothing
        intCol = Nothing
        pEnvelope = Nothing
        pPolyTemp = Nothing
        pPointCollection = Nothing
        pPointTemp = Nothing
        intRecCount = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Function

    Private Function CreateAOIClipOfParcels() As Boolean
        'CLIP THE SELECTED PARCEL THEME BY THE SELECTED AOI
        'USED FOR PARCEL ONLY AND HYBRID OPTIONS
        m_strProcessing = m_strProcessing & vbNewLine & "Starting function CreateAOIClipOfParcels: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
        CreateAOIClipOfParcels = True
        Dim pClip As ESRI.ArcGIS.AnalysisTools.Clip
        Dim pWksFactory As IWorkspaceFactory = Nothing
        Dim pFeatWks As IFeatureWorkspace
        Dim pFeatClass As IFeatureClass
        Dim pTable As ITable

        Try
            If pGPSetup Is Nothing Then
                CreateSetupGP()
            End If
            pClip = New ESRI.ArcGIS.AnalysisTools.Clip
            pClip.in_features = m_lyrOriginalParcels
            pClip.clip_features = m_lyrAOI
            pClip.out_feature_class = m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb" & "\PROJECT_PARCEL_AOI_CLIP"
            RunTool(pGPSetup, pClip)
            pClip = Nothing

            'DEFINE THE PARCEL LAYER
            pWksFactory = New FileGDBWorkspaceFactory
            pFeatWks = pWksFactory.OpenFromFile(m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb", 0)
            pFeatClass = pFeatWks.OpenFeatureClass("PROJECT_PARCEL_AOI_CLIP")
            pTable = CType(pFeatClass, ITable)
            If pFeatClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                m_lyrParcelClip = New FeatureLayer
                m_lyrParcelClip.FeatureClass = pFeatClass
                m_lyrParcelClip.Name = m_lyrParcelClip.Name & "<SETUP> " & "PROJECT PARCELS CIPPED TO AOI"
            End If
            If m_lyrParcelClip.FeatureClass.FindField("ET_ACRES") <= -1 Then
                Me.lblRunStatus.Text = "Adding ET_ACRES Field:"
                AddEnvisionField(pTable, "ET_ACRES", "DOUBLE", 16, 6)
            End If
            'CALC ACRES 
            Me.lblRunStatus.Text = "Calculating Acres"
            CalAcres(pTable, "ET_ACRES")

        Catch ex As Exception
            m_strProcessing = m_strProcessing & vbNewLine & "Failure to clip innput parcel layer to AOI." & vbNewLine & ex.Message
            CreateAOIClipOfParcels = False
            GoTo CleanUp
        End Try
CleanUp:
        m_strProcessing = m_strProcessing & vbNewLine & "End function CreateAOIClipOfParcels: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
        pClip = Nothing
        pWksFactory = Nothing
        pFeatWks = Nothing
        pFeatClass = Nothing
        pTable = Nothing
        GC.Collect()
        GC.WaitForPendingFinalizers()
    End Function

    Private Function CreateMinAcresParcelFeatClass() As Boolean
        'CREATE THE MINIMUM ACRE FEATURE LAYER
        'USED FOR HYBRID OPTION
        m_strProcessing = m_strProcessing & vbNewLine & "Starting function CreateMinAcresParcelFeatClass: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
        CreateMinAcresParcelFeatClass = True
        Dim pSelect As ESRI.ArcGIS.AnalysisTools.Select
        Dim pWksFactory As IWorkspaceFactory = Nothing
        Dim pFeatWks As IFeatureWorkspace
        Dim pFeatClass As IFeatureClass
        Dim pDataset As IDataset

        Try
            If Not m_lyrEnvisionParcel Is Nothing Then
                pDataset = CType(m_lyrEnvisionParcel.FeatureClass, IDataset)
                pSelect = New ESRI.ArcGIS.AnalysisTools.Select
                pSelect.in_features = pDataset.Workspace.PathName & "\" & pDataset.BrowseName
                pSelect.where_clause = """" & "ET_ACRES" & """" & " < " & CStr(Me.tbxMaxParcelSize.Text)
                pSelect.out_feature_class = m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb" & "\PROJECT_MIN_ACRES_PARCELS"
                RunTool(pGPSetup, pSelect)
                pSelect = Nothing
            End If
            'DEFINE THE PARCEL/GRID INTERSECT LAYER
            pWksFactory = New FileGDBWorkspaceFactory
            pFeatWks = pWksFactory.OpenFromFile(m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb", 0)
            pFeatClass = pFeatWks.OpenFeatureClass("PROJECT_MIN_ACRES_PARCELS")
            If pFeatClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                m_lyrMinAcreParcels = New FeatureLayer
                m_lyrMinAcreParcels.FeatureClass = pFeatClass
                m_lyrMinAcreParcels.Name = m_lyrMinAcreParcels.Name & "<SETUP> " & "PROJECT MIN ACRES PARCELS"
            End If
        Catch ex As Exception
            m_strProcessing = m_strProcessing & vbNewLine & "An error occured attempting to select out minimum acre parcels less than " & Me.tbxMaxParcelSize.Text & " acres from the ACRES field." & vbNewLine & ex.Message
            CreateMinAcresParcelFeatClass = False
            GoTo CleanUp
        End Try
CleanUp:
        m_strProcessing = m_strProcessing & vbNewLine & "End function CreateMinAcresParcelFeatClass: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
        pSelect = Nothing
        pWksFactory = Nothing
        pFeatWks = Nothing
        pFeatClass = Nothing
        pDataset = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Function

    Private Function CreateMaxAcresParcelFeatClass() As Boolean
        'CREATE THE MAXIMUM ACRE FEATURE LAYER
        'USED FOR HYBRID OPTION
        m_strProcessing = m_strProcessing & vbNewLine & "Starting function CreateMaxAcresParcelFeatClass: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
        CreateMaxAcresParcelFeatClass = True
        Dim pSelect As ESRI.ArcGIS.AnalysisTools.Select
        Dim pWksFactory As IWorkspaceFactory = Nothing
        Dim pFeatWks As IFeatureWorkspace
        Dim pFeatClass As IFeatureClass
        Dim pDataset As IDataset
        Try
            If pGPSetup Is Nothing Then
                CreateSetupGP()
            End If
            If Not m_lyrEnvisionParcel Is Nothing Then
                pDataset = CType(m_lyrEnvisionParcel.FeatureClass, IDataset)
                pSelect = New ESRI.ArcGIS.AnalysisTools.Select
                pSelect.in_features = pDataset.Workspace.PathName & "\" & pDataset.BrowseName
                pSelect.where_clause = """" & "ET_ACRES" & """" & " >= " & CStr(Me.tbxMaxParcelSize.Text)
                pSelect.out_feature_class = m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb" & "\PROJECT_MAX_ACRES_PARCELS"
                RunTool(pGPSetup, pSelect)
                pSelect = Nothing
            End If
            'DEFINE THE PARCEL/GRID INTERSECT LAYER
            pWksFactory = New FileGDBWorkspaceFactory
            pFeatWks = pWksFactory.OpenFromFile(m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb", 0)
            pFeatClass = pFeatWks.OpenFeatureClass("PROJECT_MAX_ACRES_PARCELS")
            If pFeatClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                m_lyrMaxAcreParcels = New FeatureLayer
                m_lyrMaxAcreParcels.FeatureClass = pFeatClass
                m_lyrMaxAcreParcels.Name = m_lyrMaxAcreParcels.Name & "<SETUP> " & "PROJECT MAX ACRES PARCELS"
            End If
        Catch ex As Exception
            m_strProcessing = m_strProcessing & vbNewLine & "An error occured attempting to select out maximum acre parcels less than " & Me.tbxMaxParcelSize.Text & " acres from the ACRES field." & vbNewLine & ex.Message
            CreateMaxAcresParcelFeatClass = False
            GoTo CleanUp
        End Try
CleanUp:
        m_strProcessing = m_strProcessing & vbNewLine & "Starting function CreateMaxAcresParcelFeatClass: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
        pSelect = Nothing
        pWksFactory = Nothing
        pFeatWks = Nothing
        pFeatClass = Nothing
        pDataset = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Function

    Private Function CreateUnionOfGridAndParcelsFeatClass() As Boolean
        'USED FOR HYBRID OPTION
        'UNION THE MAX ACRE PARCELS AND GRID CELLS, THEN MERGE WITH MIN ACRE PARCELS OR UNION ALL PARCELS WITH GRID CELLS

        m_strProcessing = m_strProcessing & vbNewLine & "Starting function CreateUnionOfGridAndParcelsFeatClass: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
        CreateUnionOfGridandParcelsFeatClass = True
        Dim pIntersect As ESRI.ArcGIS.AnalysisTools.Intersect
        Dim pMerge As ESRI.ArcGIS.DataManagementTools.Merge
        Dim pExplode As ESRI.ArcGIS.DataManagementTools.MultipartToSinglepart
        Dim pWksFactory As IWorkspaceFactory = Nothing
        Dim pFeatWks As IFeatureWorkspace
        Dim pFeatClass As IFeatureClass
        Dim strInFeatures As String = ""
        Dim pDataset As IDataset
        Dim sb As System.Text.StringBuilder = New System.Text.StringBuilder
        Dim pTable As ITable
        Dim pTempLayer As IFeatureLayer = Nothing

        Try
            If pGPSetup Is Nothing Then
                CreateSetupGP()
            End If
            If Not m_lyrMaxAcreParcels Is Nothing And Not m_lyrGrid Is Nothing Then
                pDataset = CType(m_lyrMaxAcreParcels.FeatureClass, IDataset)
                strInFeatures = pDataset.Workspace.PathName & "\" & pDataset.BrowseName
                sb.AppendFormat("{0};", strInFeatures)
                pDataset = CType(m_lyrGrid.FeatureClass, IDataset)
                strInFeatures = pDataset.Workspace.PathName & "\" & pDataset.BrowseName
                sb.AppendFormat("{0};", strInFeatures)

                pIntersect = New ESRI.ArcGIS.AnalysisTools.Intersect
                pIntersect.in_features = sb.ToString
                pIntersect.join_attributes = "NO_FID"
                pIntersect.cluster_tolerance = 1
                If Not m_lyrMinAcreParcels Is Nothing Then
                    pIntersect.out_feature_class = m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb" & "\PROJECT_UNION_MAX_PARCELS_WITH_GRID"
                Else
                    pIntersect.out_feature_class.output = m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb" & "\ENVISION_PARCELS_WITH_GRIDCELLS"
                End If
                RunTool(pGPSetup, pIntersect)
                pIntersect = Nothing

                'OPEN THE FEATURE LAYER TO SEE IF THE LAYER WAS CREATED
                pWksFactory = New FileGDBWorkspaceFactory
                pFeatWks = pWksFactory.OpenFromFile(m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb", 0)
                pFeatClass = pFeatWks.OpenFeatureClass("PROJECT_UNION_MAX_PARCELS_WITH_GRID")
                If pFeatClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                    pTempLayer = New FeatureLayer
                    pTempLayer.FeatureClass = pFeatClass
                    pTempLayer.Name = "<SETUP> " & "PROJECT PARCEL AND MAX GRID CELLS MERGE"
                End If
                If Not m_lyrMinAcreParcels Is Nothing And Not pTempLayer Is Nothing Then
                    sb = New System.Text.StringBuilder
                    pDataset = CType(m_lyrMinAcreParcels.FeatureClass, IDataset)
                    strInFeatures = pDataset.Workspace.PathName & "\" & pDataset.BrowseName
                    sb.AppendFormat("{0};", strInFeatures)
                    pDataset = CType(pTempLayer.FeatureClass, IDataset)
                    strInFeatures = pDataset.Workspace.PathName & "\" & pDataset.BrowseName
                    sb.AppendFormat("{0};", strInFeatures)

                    pMerge = New ESRI.ArcGIS.DataManagementTools.Merge
                    pMerge.inputs = sb.ToString
                    pMerge.output = m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb" & "\ENVISION_PARCELS_WITH_GRIDCELLS"
                    RunTool(pGPSetup, pMerge)
                    pMerge = Nothing
                End If
            End If

            'DEFINE THE MIN PARCEL/UNION INTERSECT LAYER
            pWksFactory = New FileGDBWorkspaceFactory
            pFeatWks = pWksFactory.OpenFromFile(m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb", 0)
            pFeatClass = pFeatWks.OpenFeatureClass("ENVISION_PARCELS_WITH_GRIDCELLS")
            If pFeatClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                pTempLayer = New FeatureLayer
                pTempLayer.FeatureClass = pFeatClass
                pTempLayer.Name = "<SETUP> " & "ENVISION PARCELS WITH GRID CELLS"
            End If

            'CREATRE A MULTIPART VERSION OF THE PARCEL/GRIDD CELL LAYER 
            Try
                m_strProcessing = m_strProcessing & vbNewLine & "Multipart to Single part features from the Mash-up:  " & Date.Now.ToString("hh:mm:ss tt")
                pExplode = New ESRI.ArcGIS.DataManagementTools.MultipartToSinglepart
                pExplode.in_features = m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb" & "\ENVISION_PARCELS_WITH_GRIDCELLS"
                pExplode.out_feature_class = m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb" & "\ENVISION_PARCELS_WITH_GRIDCELLS_MULTIPART"
                RunTool(pGPSetup, pExplode)
                pExplode = Nothing

                'DEFINE THE PARCEL/GRID INTERSECT LAYER
                pWksFactory = New FileGDBWorkspaceFactory
                pFeatWks = pWksFactory.OpenFromFile(m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb", 0)
                pFeatClass = pFeatWks.OpenFeatureClass("ENVISION_PARCELS_WITH_GRIDCELLS_MULTIPART")
                If pFeatClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                    m_lyrParcelsUnionWithGrid = New FeatureLayer
                    m_lyrParcelsUnionWithGrid.FeatureClass = pFeatClass
                    m_lyrParcelsUnionWithGrid.Name = m_lyrParcelsUnionWithGrid.Name & "<SETUP> " & "PROJECT PARCELS UNIONED WITH GRID CELLS"
                End If
            Catch ex As Exception
                m_strProcessing = m_strProcessing & vbNewLine & "Error in process to create multipart layer from layer, ENVISION_PARCELS_WITH_GRIDCELLS" & vbNewLine & ex.Message
                CreateUnionOfGridAndParcelsFeatClass = False
                GoTo CleanUp
            End Try
        Catch ex As Exception
            m_strProcessing = m_strProcessing & vbNewLine & "Error in process to union max parcels to grid cells" & vbNewLine & ex.Message
            CreateUnionOfGridAndParcelsFeatClass = False
            GoTo CleanUp
        End Try

CleanUp:
        m_strProcessing = m_strProcessing & vbNewLine & "End function CreateUnionOfGridAndParcelsFeatClass: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
        pIntersect = Nothing
        pMerge = Nothing
        pExplode = Nothing
        pWksFactory = Nothing
        pFeatWks = Nothing
        pFeatClass = Nothing
        pDataset = Nothing
        sb = Nothing
        pTable = Nothing
        GC.Collect()
        GC.WaitForPendingFinalizers()
    End Function

    Private Function CreateUnionOfGridAndOrgParcelsFeatClass() As Boolean
        'INTERSECT THE PARCEL AND GRID LAYERS
        'USED FOR HYBRID OPTION
        m_strProcessing = m_strProcessing & vbNewLine & "Starting function CreateParcelGridIntersectFeatClass: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
        CreateUnionOfGridAndOrgParcelsFeatClass = True
        Dim pUnion As ESRI.ArcGIS.AnalysisTools.Union
        Dim pExplode As ESRI.ArcGIS.DataManagementTools.MultipartToSinglepart
        Dim pWksFactory As IWorkspaceFactory = Nothing
        Dim pFeatWks As IFeatureWorkspace
        Dim pFeatClass As IFeatureClass
        Dim strInFeatures As String
        Dim pDataset As IDataset = Nothing
        Dim sb As System.Text.StringBuilder = New System.Text.StringBuilder
        Dim pTempLayer As IFeatureLayer

        If pGPSetup Is Nothing Then
            CreateSetupGP()
        End If

        'UNION ALL PARCELS WITH ALL GRID CELLS IF NO MAX PARCEL SIZE IS SET AND HYBRID OPTION IS SELECTED
        Try
            sb = New System.Text.StringBuilder
            If Not m_lyrEnvisionParcel Is Nothing Then
                pDataset = CType(m_lyrEnvisionParcel.FeatureClass, IDataset)
            Else
                CreateUnionOfGridAndOrgParcelsFeatClass = False
                GoTo CleanUp
            End If
            strInFeatures = pDataset.Workspace.PathName & "\" & pDataset.BrowseName
            sb.AppendFormat("{0};", strInFeatures)
            If m_lyrGrid Is Nothing Then
                m_strProcessing = m_strProcessing & vbNewLine & "NO grid cell layer could be found."
                CreateUnionOfGridAndOrgParcelsFeatClass = False
                GoTo CleanUp
            End If
            pDataset = CType(m_lyrGrid.FeatureClass, IDataset)
            strInFeatures = pDataset.Workspace.PathName & "\" & pDataset.BrowseName
            sb.AppendFormat("{0};", strInFeatures)
            pUnion = New ESRI.ArcGIS.AnalysisTools.Union
            pUnion.in_features = sb.ToString
            pUnion.join_attributes = "ALL"
            pUnion.gaps = "GAPS"
            pUnion.cluster_tolerance = 1
            pUnion.out_feature_class = m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb" & "\ENVISION_PARCELS_WITH_GRIDCELLS"
            RunTool(pGPSetup, pUnion)
            pUnion = Nothing

            'DEFINE THE PARCEL/GRID INTERSECT LAYER
            pWksFactory = New FileGDBWorkspaceFactory
            pFeatWks = pWksFactory.OpenFromFile(m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb", 0)
            pFeatClass = pFeatWks.OpenFeatureClass("ENVISION_PARCELS_WITH_GRIDCELLS")
            If pFeatClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                pTempLayer = New FeatureLayer
                pTempLayer.FeatureClass = pFeatClass
                pTempLayer.Name = m_lyrParcelsUnionWithGrid.Name & "<SETUP> " & "PROJECT PARCELS UNIONED WITH GRID CELLS"
            End If

            'CREATRE A MULTIPART VERSION OF THE PARCEL/GRIDD CELL LAYER 
            Try
                m_strProcessing = m_strProcessing & vbNewLine & "Multipart to Single part features from the Mash-up:  " & Date.Now.ToString("hh:mm:ss tt")
                pExplode = New ESRI.ArcGIS.DataManagementTools.MultipartToSinglepart
                pExplode.in_features = m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb" & "\ENVISION_PARCELS_WITH_GRIDCELLS"
                pExplode.out_feature_class = m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb" & "\ENVISION_PARCELS_WITH_GRIDCELLS_MULTIPART"
                RunTool(pGPSetup, pExplode)
                pExplode = Nothing

                'DEFINE THE PARCEL/GRID INTERSECT LAYER
                pWksFactory = New FileGDBWorkspaceFactory
                pFeatWks = pWksFactory.OpenFromFile(m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb", 0)
                pFeatClass = pFeatWks.OpenFeatureClass("ENVISION_PARCELS_WITH_GRIDCELLS_MULTIPART")
                If pFeatClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                    m_lyrParcelsUnionWithGrid = New FeatureLayer
                    m_lyrParcelsUnionWithGrid.FeatureClass = pFeatClass
                    m_lyrParcelsUnionWithGrid.Name = m_lyrParcelsUnionWithGrid.Name & "<SETUP> " & "PROJECT PARCELS UNIONED WITH GRID CELLS"
                End If
            Catch ex As Exception
                m_strProcessing = m_strProcessing & vbNewLine & "Error in process to create multipart layer from layer, ENVISION_PARCELS_WITH_GRIDCELLS" & vbNewLine & ex.Message
                CreateUnionOfGridAndOrgParcelsFeatClass = False
                GoTo CleanUp
            End Try
            GoTo CleanUp
        Catch ex As Exception
            m_strProcessing = m_strProcessing & vbNewLine & "Error in UnionParcelToConstraints:" & vbNewLine & ex.Message
            CreateUnionOfGridAndOrgParcelsFeatClass = False
            GoTo CleanUp
        End Try

CleanUp:
        m_strProcessing = m_strProcessing & vbNewLine & "End function CreateUnionOfGridAndOrgParcelsFeatClass: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
        pUnion = Nothing
        pExplode = Nothing
        pWksFactory = Nothing
        pFeatWks = Nothing
        pFeatClass = Nothing
        pDataset = Nothing
        sb = Nothing
        pTempLayer = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Function

    Private Function EvisionIdField(ByVal pTable As ITable) As Boolean
        'CHECK FOR AND CALCULATE AN ENVISION TEMPORARY ID FIELD BASED ON OBJECTID FIELD
        m_strProcessing = m_strProcessing & vbNewLine & "Starting function EvisionIdField: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
        Try
            'DEFINE THE PARCEL/GRID LAYER
            If pTable.FindField("ET_ENVISION_ID") <= -1 Then
                AddEnvisionField(pTable, "ET_ENVISION_ID", "INTEGER", 25, 0)
            End If
            'CALCULATE THE FIELD TO OBJECT ID
            If Not pTable Is Nothing Then
                CalcFldValues(pTable, "ET_ENVISION_ID", "[" & pTable.Fields.Field(0).Name & "]", "INTEGER")
            End If
            GoTo CleanUp
        Catch ex As Exception
            m_strProcessing = m_strProcessing & vbNewLine & "Error" & vbNewLine & ex.Message
            EvisionIdField = False
            GoTo CleanUp
        End Try
CleanUp:
        m_strProcessing = m_strProcessing & vbNewLine & "End function EvisionIdField: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Function

    Private Function CreateUnionConstraintsFeatClass() As Boolean
        'INTERSECT ALL THE LAYERS DEISGNATED AS CONSTRAINTS LAYER(S)
        m_strProcessing = m_strProcessing & vbNewLine & "Starting function CreateUnionConstraintsFeatClass: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
        CreateUnionConstraintsFeatClass = True
        Dim pUnion As ESRI.ArcGIS.AnalysisTools.Union
        Dim pClip As ESRI.ArcGIS.AnalysisTools.Clip
        Dim pWksFactory As IWorkspaceFactory = Nothing
        Dim pFeatWks As IFeatureWorkspace
        Dim pFeatClass As IFeatureClass
        Dim strInFeatures As String
        Dim pDataset As IDataset = Nothing
        Dim sb As System.Text.StringBuilder = New System.Text.StringBuilder
        Dim pFeatLayer As IFeatureLayer
        Dim intCount As Integer
        Dim intUnionCount As Integer = 0
        Dim intUnionTotalCount As Integer = 0
        Dim strUnionName As String
        Dim pTable As ITable
        Dim strFieldName As String
        Dim pDissolve As ESRI.ArcGIS.DataManagementTools.Dissolve
        Dim pFeatSel As IFeatureSelection
        Dim pLocationSel As ESRI.ArcGIS.DataManagementTools.SelectLayerByLocation
        Dim pCursor As ICursor = Nothing
        Dim strBuffDist As String = ""
        Dim strBuffUnits As String = ""

        Try
            m_blnConstraints = False
            If pGPSetup Is Nothing Then
                CreateSetupGP()
            End If

            'CREATE SLOPE CONSTRAINT GRIDS AND FEATURE CLASS(ES)
            If Me.chkCreateSlope15to25.Checked Or Me.chkCreateSlope25Plus.Checked Then
                CreateSlopeFeatureClasses()
                If Me.chkCreateSlope15to25.Checked Then
                    Try
                        pWksFactory = New FileGDBWorkspaceFactory
                        pFeatWks = pWksFactory.OpenFromFile(m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb", 0)
                        pFeatClass = pFeatWks.OpenFeatureClass("SLOPE_15to25")
                        strInFeatures = m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb\SLOPE_15to25"
                        sb.AppendFormat("{0};", strInFeatures)
                        pWksFactory = Nothing
                        pFeatWks = Nothing
                        pFeatClass = Nothing
                        intUnionCount = intUnionCount + 1
                        GC.WaitForPendingFinalizers()
                        GC.Collect()
                    Catch ex As Exception
                    End Try
                End If

                If Me.chkCreateSlope25Plus.Checked Then
                    Try
                        pWksFactory = New FileGDBWorkspaceFactory
                        pFeatWks = pWksFactory.OpenFromFile(m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb", 0)
                        pFeatClass = pFeatWks.OpenFeatureClass("SLOPE_25PLUS")
                        GC.Collect()
                        GC.WaitForPendingFinalizers()
                        strInFeatures = m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb\SLOPE_25PLUS"
                        sb.AppendFormat("{0};", strInFeatures)
                        pWksFactory = Nothing
                        pFeatWks = Nothing
                        pFeatClass = Nothing
                        intUnionCount = intUnionCount + 1
                        GC.Collect()
                        GC.WaitForPendingFinalizers()
                    Catch ex As Exception
                    End Try
                End If

                If Me.chkCreateSlope15to25.Checked And Me.chkCreateSlope25Plus.Checked Then
                    Try
                        If intUnionCount = 2 Then
                            intUnionTotalCount = intUnionTotalCount + 1
                            pUnion = New ESRI.ArcGIS.AnalysisTools.Union
                            pUnion.in_features = sb.ToString
                            pUnion.cluster_tolerance = 1
                            pUnion.join_attributes = "ONLY_FID"
                            pUnion.gaps = "GAPS"
                            If m_blnConstraints Then
                                strUnionName = m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb" & "\ENVISION_CONSTRAINTS_FULL"
                            Else
                                strUnionName = m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb" & "\ENVISION_CONSTRAINTS" & intUnionTotalCount.ToString
                            End If
                            pUnion.out_feature_class = strUnionName
                            RunTool(pGPSetup, pUnion)
                            pUnion = Nothing
                            sb = New System.Text.StringBuilder
                            sb.AppendFormat("{0};", strUnionName)
                            intUnionCount = 1
                            intUnionTotalCount = intUnionTotalCount + 1
                        End If
                    Catch ex As Exception
                        intUnionCount = 0
                        sb = New System.Text.StringBuilder
                    End Try
                End If
            End If

            'CYCLE THROUGH LAYER LIST TO INTERSECT ALL CONSTRAINT LAYERS INTO A SINGLE LAYER
            For intCount = 0 To Me.dgvConstraints.Rows.Count - 1
                If Me.dgvConstraints.Rows(intCount).Cells(0).Value = "TRUE" Then
                    pFeatLayer = m_arrETSetup_FLayers.Item(intCount)
                    Try
                        strBuffDist = Me.dgvConstraints.Rows(intCount).Cells(2).Value
                    Catch ex As Exception
                        strBuffDist = "0"
                    End Try
                    Try
                        strBuffUnits = Me.dgvConstraints.Rows(intCount).Cells(3).Value
                    Catch ex As Exception
                        strBuffUnits = "Miles"
                    End Try

                    'If pFeatLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                    '    pDataset = CType(pFeatLayer.FeatureClass, IDataset)
                    '    If pDataset.Category.ToString.Contains("Shapefile") Then
                    '        strInFeatures = pDataset.Workspace.PathName & "\" & pDataset.BrowseName & ".shp"
                    '    Else
                    '        strInFeatures = pDataset.Workspace.PathName & "\" & pDataset.BrowseName
                    '    End If
                    '    pLocationSel = New ESRI.ArcGIS.DataManagementTools.SelectLayerByLocation
                    '    pLocationSel.in_layer = pFeatLayer
                    '    pLocationSel.select_features = m_lyrAOI
                    '    pLocationSel.selection_type = "NEW_SELECTION"
                    '    pLocationSel.overlap_type = "INTERSECT"
                    '    RunTool(pGPSetup, pLocationSel)
                    '    pFeatSel = Nothing
                    '    pFeatSel = CType(pFeatLayer, IFeatureSelection)
                    '    pFeatSel.SelectionSet.Search(Nothing, False, pCursor)
                    '    If pFeatSel.SelectionSet.Count >= 1 Then
                    '        sb.AppendFormat("{0};", strInFeatures)
                    '        pFeatLayer = Nothing
                    '        pDataset = Nothing
                    '        intUnionCount = intUnionCount + 1
                    '    Else
                    '        m_strProcessing = m_strProcessing & vbNewLine & "***** Constratints layer not within AOI: " & strInFeatures
                    '    End If
                    'Else
                    '    If pFeatLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryLine Or pFeatLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline Or pFeatLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPoint Then
                    pLocationSel = New ESRI.ArcGIS.DataManagementTools.SelectLayerByLocation
                    pLocationSel.in_layer = pFeatLayer
                    pLocationSel.select_features = m_lyrAOI
                    pLocationSel.selection_type = "NEW_SELECTION"
                    pLocationSel.overlap_type = "INTERSECT"
                    RunTool(pGPSetup, pLocationSel)
                    pFeatSel = Nothing
                    pFeatSel = CType(pFeatLayer, IFeatureSelection)
                    pFeatSel.SelectionSet.Search(Nothing, False, pCursor)
                    If pFeatSel.SelectionSet.Count >= 1 Then
                        strInFeatures = ""
                        If IsNumeric(strBuffDist) Then
                            strInFeatures = BufferConstraint(pFeatLayer, strBuffDist, strBuffUnits)
                        End If
                        If Not strInFeatures = "" Then
                            sb.AppendFormat("{0};", strInFeatures)
                            intUnionCount = intUnionCount + 1
                        End If
                    Else
                        pDataset = CType(pFeatLayer.FeatureClass, IDataset)
                        If pDataset.Category.ToString.Contains("Shapefile") Then
                            strInFeatures = pDataset.Workspace.PathName & "\" & pDataset.BrowseName & ".shp"
                        Else
                            strInFeatures = pDataset.Workspace.PathName & "\" & pDataset.BrowseName
                        End If
                        m_strProcessing = m_strProcessing & vbNewLine & "***** Constratints layer not within AOI: " & strInFeatures
                    End If
                    ' End If
                    'End If

                    If intUnionCount = 2 Then
                        pUnion = New ESRI.ArcGIS.AnalysisTools.Union
                        pUnion.in_features = sb.ToString
                        pUnion.join_attributes = "ONLY_FID"
                        pUnion.gaps = "GAPS"
                        strUnionName = m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb" & "\ENVISION_CONSTRAINTS" & intUnionTotalCount.ToString
                        pUnion.out_feature_class = strUnionName
                        RunTool(pGPSetup, pUnion)
                        pUnion = Nothing
                        sb = New System.Text.StringBuilder
                        sb.AppendFormat("{0};", strUnionName)
                        intUnionTotalCount = intUnionTotalCount + 1
                        intUnionCount = 1
                    End If
                End If
            Next

            'RUN ONE LAST UNION 
            pUnion = New ESRI.ArcGIS.AnalysisTools.Union
            pUnion.in_features = sb.ToString
            pUnion.join_attributes = "ONLY_FID"
            pUnion.gaps = "GAPS"
            strUnionName = m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb" & "\ENVISION_CONSTRAINTS_FULL"
            pUnion.out_feature_class = strUnionName
            RunTool(pGPSetup, pUnion)

            'DEFINE CONSTRAINTS LAYER
            pWksFactory = New FileGDBWorkspaceFactory
            pFeatWks = pWksFactory.OpenFromFile(m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb", 0)
            pFeatClass = pFeatWks.OpenFeatureClass("ENVISION_CONSTRAINTS_FULL")
            pTable = CType(pFeatClass, ITable)
            If pFeatClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                m_blnConstraints = True
                m_lyrConstraints = New FeatureLayer
                m_lyrConstraints.FeatureClass = pFeatClass
                m_lyrConstraints.Name = "<SETUP> " & "PROJECT CONSTRAINTS"
                'MAKE SURE THE CONSTRAINTS LAYER IS NOT EMPTY
                pTable = CType(pFeatClass, ITable)
                If pTable.RowCount(Nothing) <= 0 Then
                    m_lyrConstraints = Nothing
                    m_strProcessing = m_strProcessing & vbNewLine & "**** Empty constraints feature layer ******"
                    GoTo CleanUp
                End If
            End If
            'ADD AND CALCULATE CONSTRAINTS FIELD
            AddEnvisionField(pTable, "ET_CONSTRAINED", "INTEGER", 1, 0)
            CalcFldValues(pTable, "ET_CONSTRAINED", "1", "INTEGER")


            'DISSOLVE THE FINAL CONSTRAINTS LAYER
            pDissolve = New ESRI.ArcGIS.DataManagementTools.Dissolve
            pDissolve.in_features = m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb" & "\ENVISION_CONSTRAINTS_FULL"
            pDissolve.dissolve_field = "ET_CONSTRAINED"
            strUnionName = m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb" & "\ENVISION_CONSTRAINTS_FULL_DISSOLVED"

            pDissolve.out_feature_class = strUnionName
            RunTool(pGPSetup, pDissolve)
            pDissolve = Nothing

            'CLIP THE CONSTRAINTS TO AOI
            pClip = New ESRI.ArcGIS.AnalysisTools.Clip
            pClip.in_features = m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb" & "\ENVISION_CONSTRAINTS_FULL_DISSOLVED"
            pClip.clip_features = m_lyrAOI
            pClip.out_feature_class = m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb" & "\ENVISION_CONSTRAINTS_FULL_DISSOLVED_CLIPPED"
            RunTool(pGPSetup, pClip)
            pClip = Nothing

            'DEFINE DISSOLVED CONSTRAINTS LAYER
            m_lyrConstraints = Nothing
            GC.Collect()
            GC.WaitForPendingFinalizers()
            pWksFactory = New FileGDBWorkspaceFactory
            pFeatWks = pWksFactory.OpenFromFile(m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb", 0)
            pFeatClass = pFeatWks.OpenFeatureClass("ENVISION_CONSTRAINTS_FULL_DISSOLVED_CLIPPED")
            If pFeatClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                m_blnConstraints = True
                m_lyrConstraints = New FeatureLayer
                m_lyrConstraints.FeatureClass = pFeatClass
                m_lyrConstraints.Name = "<SETUP> " & "PROJECT CONSTRAINTS"
                'MAKE SURE THE CONSTRAINTS LAYER IS NOT EMPTY
                pTable = CType(pFeatClass, ITable)
                If pTable.RowCount(Nothing) <= 0 Then
                    m_lyrConstraints = Nothing
                    m_strProcessing = m_strProcessing & vbNewLine & "**** Empty constraints feature layer ******"
                    GoTo CleanUp
                End If

                'DELETE ALL FIELDS
                sb = New System.Text.StringBuilder
                pTable = CType(pFeatClass, Table)
                For intCount = 0 To pFeatClass.Fields.FieldCount - 1
                    sb.AppendFormat("{0};", pFeatClass.Fields.Field(intCount).Name)
                Next
                For Each strFieldName In sb.ToString.Split(";")
                    Try
                        If strFieldName.Length > 0 And Not UCase(strFieldName) = "ET_CONSTRAINED" And Not UCase(strFieldName) = "OBJECTID" And Not UCase(strFieldName) = "SHAPE" And Not UCase(strFieldName) = "SHAPE_LENGTH" And Not UCase(strFieldName) = "SHAPE_AREA" Then
                            If strFieldName.Length > 0 Then
                                pTable.DeleteField(pFeatClass.Fields.Field(pTable.FindField(strFieldName)))
                            End If
                        End If
                    Catch ex As Exception
                        m_strProcessing = m_strProcessing & vbNewLine & "Processing Error in deleting a field: " & strFieldName & vbNewLine & ex.Message
                    End Try
                Next
            End If

            GoTo CleanUp

        Catch ex As Exception
            m_strProcessing = m_strProcessing & vbNewLine & "Processing Contraints Error:" & vbNewLine & ex.Message
            CreateUnionConstraintsFeatClass = False
            GoTo CleanUp
        End Try

CleanUp:
        m_strProcessing = m_strProcessing & vbNewLine & "End function CreateUnionConstraintsFeatClass: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
        pUnion = Nothing
        pClip = Nothing
        pFeatWks = Nothing
        pFeatClass = Nothing
        strInFeatures = Nothing
        pDataset = Nothing
        sb = Nothing
        pFeatLayer = Nothing
        intCount = Nothing
        intUnionCount = Nothing
        intUnionTotalCount = Nothing
        strUnionName = Nothing
        pTable = Nothing
        strFieldName = Nothing
        pDissolve = Nothing
        pFeatSel = Nothing
        pLocationSel = Nothing
        pCursor = Nothing
        GC.Collect()
        GC.WaitForPendingFinalizers()

    End Function

    Private Function UnionParcelToConstraints() As Boolean
        'UNION THE CONSTRAINTS LAYER WITH PARCEL/GRID LAYER
        m_strProcessing = m_strProcessing & vbNewLine & "Starting function UnionParcelToConstraints: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
        Dim pUnion As ESRI.ArcGIS.AnalysisTools.Union
        Dim sb As System.Text.StringBuilder = New System.Text.StringBuilder
        Dim strUnionName As String = ""
        Dim pWksFactory As IWorkspaceFactory = Nothing
        Dim pFeatWks As IFeatureWorkspace
        Dim pFeatClass As IFeatureClass
        Try
            sb.AppendFormat("{0};", m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb" & "\ENVISION_PARCELS")
            sb.AppendFormat("{0};", m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb" & "\ENVISION_CONSTRAINTS_FULL_CLIPPED")
            pUnion = New ESRI.ArcGIS.AnalysisTools.Union
            pUnion.in_features = sb.ToString
            pUnion.join_attributes = "ALL"
            pUnion.gaps = "GAPS"
            pUnion.cluster_tolerance = 1
            strUnionName = m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb" & "\ENVISION_PARCELGRID_CONSTRAINTS"
            pUnion.out_feature_class = strUnionName
            RunTool(pGPSetup, pUnion)
            pUnion = Nothing
            sb = New System.Text.StringBuilder
            'DEFINE THE PARCEL/GRID/CONSTRAINTS INTERSECT LAYER
            pWksFactory = New FileGDBWorkspaceFactory
            pFeatWks = pWksFactory.OpenFromFile(m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb", 0)
            pFeatClass = pFeatWks.OpenFeatureClass("ENVISION_CONSTRAINTS_FULL")
            If pFeatClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                m_lyrConstraints = New FeatureLayer
                m_lyrConstraints.FeatureClass = pFeatClass
                m_lyrConstraints.Name = m_lyrParcelsUnionWithGrid.Name & "<SETUP> " & "PROJECT CONSTRAINTS"
            End If
            GoTo CleanUp
        Catch ex As Exception
            m_strProcessing = m_strProcessing & vbNewLine & "Error in UnionParcelToConstraints:" & vbNewLine & ex.Message
            GoTo CleanUp
        End Try
CleanUp:
        m_strProcessing = m_strProcessing & vbNewLine & "End function UnionParcelToConstraints: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
        pUnion = Nothing
        sb = Nothing
        strUnionName = Nothing
        pWksFactory = Nothing
        pFeatWks = Nothing
        pFeatClass = Nothing
        sb = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Function

    Private Function BufferConstraint(ByVal pFeatLayer As IFeatureLayer, ByVal strDistance As String, ByVal strUnits As String) As String
        'BUFFER THE SELECT LAYER WITH THE DEFINED BUFFER DISTANCE
        BufferConstraint = ""
        Dim pBuffer As ESRI.ArcGIS.AnalysisTools.Buffer
        Dim pWksFactory As IWorkspaceFactory = Nothing
        Dim pFeatWks As IFeatureWorkspace
        Dim pFeatClass As IFeatureClass
        Dim pDataset As IDataset

        Try
            'TRACK NUMBER OF BUFFER LAYERS CREATED
            intBuffLayerCount = intBuffLayerCount + 1

            pDataset = CType(pFeatLayer.FeatureClass, IDataset)
            pBuffer = New ESRI.ArcGIS.AnalysisTools.Buffer
            If pDataset.Category.ToString.Contains("Shapefile") Then
                pBuffer.in_features = pDataset.Workspace.PathName & "\" & pDataset.BrowseName & ".shp"
            Else
                pBuffer.in_features = pDataset.Workspace.PathName & "\" & pDataset.BrowseName
            End If
            pBuffer.buffer_distance_or_field = strDistance & " " & strUnits
            pBuffer.out_feature_class = m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb" & "\" & pDataset.BrowseName & "_Buffer" & intBuffLayerCount.ToString & "_" '& Me.tbxConstraintsBufferDistance.Text
            If pFeatLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryLine Or pFeatLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline Then
                pBuffer.line_end_type = "ROUND"
                pBuffer.line_side = "FULL"
            End If
            RunTool(pGPSetup, pBuffer)
            pBuffer = Nothing

            'DEFINE BUFFER LAYER
            pWksFactory = New FileGDBWorkspaceFactory
            pFeatWks = pWksFactory.OpenFromFile(m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb", 0)
            ' pFeatClass = pFeatWks.OpenFeatureClass(pDataset.BrowseName & "_Buffer" & intBuffLayerCount.ToString & "_" & Me.tbxConstraintsBufferDistance.Text)
            BufferConstraint = m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb" & "\" & pDataset.BrowseName & "_Buffer" & intBuffLayerCount.ToString & "_" '& Me.tbxConstraintsBufferDistance.Text
        Catch ex As Exception
            GoTo CleanUp
        End Try

CleanUp:
        pBuffer = Nothing
        pWksFactory = Nothing
        pFeatWks = Nothing
        pFeatClass = Nothing
        pDataset = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Function

    Private Function CreateLandUseFeatClass(ByVal arrFrom As ArrayList, ByVal arrTo As ArrayList, ByVal arrDeveloped As ArrayList, ByVal arrVacant As ArrayList, ByVal arrConstrained As ArrayList) As Boolean
        'CHECK TO SEE IF THE USER DEFINED LANDUSE INPUTS TO DEVELOP VACANT AND/OR DEVELOPED LANDUSE FEATURE CLASSES
        m_strProcessing = m_strProcessing & vbNewLine & "Starting function CreateLandUseFeatClass: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
        Me.lblRunStatus.Text = "Processing Land Use "
        CreateLandUseFeatClass = True
        Dim pDataset As IDataset
        Dim strParcelFilename As String = ""
        Dim strLanduseFilename As String = ""
        Dim pClip As ESRI.ArcGIS.AnalysisTools.Clip
        Dim pWksFactory As IWorkspaceFactory = Nothing
        Dim pFeatWks As IFeatureWorkspace
        Dim pFeatClass As IFeatureClass
        Dim strLandUseField As String = Me.cmbLandUseField.Text
        Dim pTable As ITable = Nothing
        Dim intCount As Integer
        Dim pField As IField
        Dim strFieldName As String
        Dim pFeatureCursor As IFeatureCursor
        Dim pFeat As IFeature
        Dim intFomLU As Integer = 0
        Dim intToLU As Integer = 0
        Dim strFomLU As String = ""
        Dim strToLU As String = ""
        Dim intLU As Integer = 0
        Dim arrFieldNames As ArrayList = New ArrayList
        Dim pUnion As ESRI.ArcGIS.AnalysisTools.Union
        Dim strInFeatures As String
        Dim strUnionName As String
        Dim sb As System.Text.StringBuilder = New System.Text.StringBuilder
        'Dim pDissolve As ESRI.ArcGIS.DataManagementTools.Dissolve
        'Dim pSummarize As ESRI.ArcGIS.AnalysisTools.Statistics
        Dim pJoin As ESRI.ArcGIS.DataManagementTools.JoinField
        'Dim pCalc As ESRI.ArcGIS.DataManagementTools.CalculateField
        'Dim pSummaryTable As ITable
        'Dim pRemoveJoin As ESRI.ArcGIS.DataManagementTools.RemoveJoin
        Dim blnDeleteFields As Boolean = True
        Dim strVacantStatus As String
        Dim strDevelopedStatus As String
        Dim strConstrained As String
        Dim pIntersect As ESRI.ArcGIS.AnalysisTools.Intersect


        If pGPSetup Is Nothing Then
            CreateSetupGP()
        End If

        'IF MATCH, THEN MAKE THE PARCEL OUTPUT THE LANDUSE LAYER
        If Me.cmbParcelLayers.SelectedIndex = Me.cmbLandUseLayers.SelectedIndex Then
            blnDeleteFields = False
            m_lyrLandUse = Nothing
            m_lyrLandUse = m_lyrMainProcessingLayer
        Else
            'CLIP THE SELECTED LAND USE LAYER TO THE AOI IF DIFFERENT THAN PARCEL LAYER
            Me.lblRunStatus.Text = "Processing Land Use - Clipping to AOI"
            Try
                pClip = New ESRI.ArcGIS.AnalysisTools.Clip
                pClip.in_features = m_lyrLandUse
                pClip.clip_features = m_lyrAOI
                pClip.out_feature_class = m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb" & "\ENVISION_LANDUSE_CLIP"
                RunTool(pGPSetup, pClip)
                pClip = Nothing
            Catch ex As Exception
                m_lyrLandUse = Nothing
                Me.lblRunStatus.Text = "Sub(CreateLandUseFeatClass) Error in clipping input land use layer: Error: "
                m_strProcessing = m_strProcessing & vbNewLine & " Sub(CreateLandUseFeatClass) Error in clipping input land use layer: Error: " & ex.Message
                CreateLandUseFeatClass = False
                GoTo CleanUp
            End Try

            'DEFINE THE LANDUSE LAYER
            Try
                pWksFactory = New FileGDBWorkspaceFactory
                pFeatWks = pWksFactory.OpenFromFile(m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb", 0)
                pFeatClass = pFeatWks.OpenFeatureClass("ENVISION_LANDUSE_CLIP")
                If pFeatClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                    m_lyrLandUse = Nothing
                    m_lyrLandUse = New FeatureLayer
                    m_lyrLandUse.FeatureClass = pFeatClass
                    m_lyrLandUse.Name = "<SETUP> " & "PROJECT LAND USE"
                End If

                'DELETE ANY EXISTING EX LU FIELDS THAT MAY EXIST
                If m_lyrLandUse.FeatureClass.FindField("EX_LU") >= 0 Then

                End If

            Catch ex As Exception
                m_strProcessing = m_strProcessing & vbNewLine & " Sub(CreateLandUseFeatClass) Error in defining land use layer: Error: " & ex.Message
                CreateLandUseFeatClass = False
                GoTo CleanUp
            End Try
        End If

        'POPULATE THE LAND USE OUTPUT WITH THE EXISITNG LAND USE ABBREVIATIONS
        If Not m_lyrLandUse Is Nothing Then
            pTable = CType(m_lyrLandUse.FeatureClass, Table)
            If pTable.RowCount(Nothing) <= 0 Then
                m_lyrLandUse = Nothing
                m_strProcessing = m_strProcessing & vbNewLine & "**** Empty Land Use feature class *****"
            Else

                Me.lblRunStatus.Text = "Processing Land Use - Add EX_LU fields"

                If pTable.FindField("EX_LU") <= -1 Then
                    Me.lblRunStatus.Text = "Processing Land Use - Adding field EX_LU"
                    AddEnvisionField(pTable, "EX_LU", "STRING", 50, 0)
                Else
                    Me.lblRunStatus.Text = "Processing Land Use - Reset Calc field EX_LU"
                    CalcFldValues(pTable, "EX_LU", """""", "STRING")
                End If
                If pTable.FindField("EX_LU_CONSTRAINED") <= -1 Then
                    Me.lblRunStatus.Text = "Processing Land Use - Adding field EX_LU_CONSTRAINED"
                    AddEnvisionField(pTable, "EX_LU_CONSTRAINED", "STRING", 5, 0)
                Else
                    Me.lblRunStatus.Text = "Processing Land Use - Reset Calc field EX_LU"
                    CalcFldValues(pTable, "EX_LU_CONSTRAINED", """""", "STRING")
                End If
                If pTable.FindField("EX_LU_DEVELOPED") <= -1 Then
                    Me.lblRunStatus.Text = "Processing Land Use - Adding field EX_LU_DEVELOPED"
                    AddEnvisionField(pTable, "EX_LU_DEVELOPED", "STRING", 5, 0)
                Else
                    Me.lblRunStatus.Text = "Processing Land Use - Reset Calc field EX_LU_DEVELOPED"
                    CalcFldValues(pTable, "EX_LU_DEVELOPED", """""", "STRING")
                End If
                If pTable.FindField("EX_LU_VACANT") <= -1 Then
                    Me.lblRunStatus.Text = "Processing Land Use - Adding field EX_LU_VACANT"
                    AddEnvisionField(pTable, "EX_LU_VACANT", "STRING", 5, 0)
                Else
                    Me.lblRunStatus.Text = "Processing Land Use - Reset Calc field EX_LU_VACANT"
                    CalcFldValues(pTable, "EX_LU_VACANT", """""", "STRING")
                End If

                'CALCULATE THE LANDUSE ABBREVIATIONS
                pTable = CType(m_lyrLandUse.FeatureClass, Table)
                intFomLU = m_lyrLandUse.FeatureClass.FindField(strLandUseField)
                intToLU = m_lyrLandUse.FeatureClass.FindField("EX_LU")
                pFeatureCursor = m_lyrLandUse.FeatureClass.Search(Nothing, False)
                pFeat = pFeatureCursor.NextFeature
                intCount = -1
                Do While Not pFeat Is Nothing
                    'WRITE ACRE VALUES TO FEATURE
                    strFomLU = ""
                    strToLU = ""
                    strVacantStatus = "FALSE"
                    strDevelopedStatus = "FALSE"
                    strConstrained = "FALSE"
                    intCount = intCount + 1
                    Me.lblRunStatus.Text = "Applying LU atrributes, feature: " & intCount.ToString
                    Try
                        strFomLU = pFeat.Value(intFomLU)
                    Catch ex As Exception
                        strFomLU = ""
                    End Try
                    Try
                        strToLU = ""
                        intLU = arrFrom.IndexOf(strFomLU)
                        If intLU >= 0 Then
                            strToLU = arrTo.Item(intLU)
                        End If
                    Catch ex As Exception
                        strToLU = ""
                    End Try
                    Try
                        strDevelopedStatus = arrDeveloped.Item(intLU).ToString
                        If strDevelopedStatus = "" Then
                            strDevelopedStatus = "FALSE"
                        End If
                    Catch ex As Exception
                    End Try
                    Try
                        strVacantStatus = arrVacant.Item(intLU).ToString
                        If strVacantStatus = "" Then
                            strVacantStatus = "FALSE"
                        End If
                    Catch ex As Exception
                    End Try
                    Try
                        If arrConstrained.Item(intLU) = 1 Then
                            strConstrained = "TRUE"
                        End If
                    Catch ex As Exception
                    End Try
                    pFeat.Value(intToLU) = strToLU
                    pFeat.Value(pFeat.Fields.FindField("EX_LU_DEVELOPED")) = strVacantStatus
                    pFeat.Value(pFeat.Fields.FindField("EX_LU_VACANT")) = strDevelopedStatus
                    pFeat.Value(pFeat.Fields.FindField("EX_LU_CONSTRAINED")) = strConstrained
                    pFeat.Store()
                    pFeat = pFeatureCursor.NextFeature
                Loop
            End If
        Else
            m_strProcessing = m_strProcessing & vbNewLine & "**** Land Use layer was equal to nothing *****"
            GoTo CleanUp
        End If

        '*******************************************************************************************************************************************
        If Not Me.cmbParcelLayers.SelectedIndex = Me.cmbLandUseLayers.SelectedIndex Then
            'CREATE A POINT FEATURE CLASS IF LAND USE AND PARCEL LAYERS ARE DIFFERENT
            'POINT LAYER WILL BE INTERSECTED WITH LAND USE LAYER, THEN JOINED ON OBJECT ID TO CALC LAND USE ATTRIBUTES TO THE PARCEL LAYER
            Me.lblRunStatus.Text = "Creating Point Layer to intersect with Land Use"
            CreateLUCentroidLayer()

            If pGPSetup Is Nothing Then
                CreateSetupGP()
            End If
            If Not m_lyrLandUse Is Nothing And Not m_lyrLUCentroids Is Nothing Then
                Try
                    pDataset = CType(m_lyrLandUse.FeatureClass, IDataset)
                    strInFeatures = pDataset.Workspace.PathName & "\" & pDataset.BrowseName
                    sb.AppendFormat("{0};", strInFeatures)
                    pDataset = CType(m_lyrLUCentroids.FeatureClass, IDataset)
                    strInFeatures = pDataset.Workspace.PathName & "\" & pDataset.BrowseName
                    sb.AppendFormat("{0};", strInFeatures)

                    pIntersect = New ESRI.ArcGIS.AnalysisTools.Intersect
                    pIntersect.in_features = sb.ToString
                    pIntersect.join_attributes = "NO_FID"
                    pIntersect.cluster_tolerance = 1
                    pIntersect.out_feature_class = m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb" & "\PROJECT_INTERSECT_LU_CENTROIDS"
                    RunTool(pGPSetup, pIntersect)
                    pIntersect = Nothing

                    pWksFactory = New FileGDBWorkspaceFactory
                    pFeatWks = pWksFactory.OpenFromFile(m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb", 0)
                    pFeatClass = pFeatWks.OpenFeatureClass("PROJECT_INTERSECT_LU_CENTROIDS")
                    pTable = CType(pFeatClass, ITable)
                    m_lyrLUCentroids = Nothing
                    If pFeatClass.ShapeType = esriGeometryType.esriGeometryPoint Then
                        m_lyrLUCentroids = New FeatureLayer
                        m_lyrLUCentroids.FeatureClass = pFeatClass
                        m_lyrLUCentroids.Name = m_lyrGrid.Name & "<SETUP> " & "EX_LU CENTROID LAYER"
                    End If

                    pJoin = New ESRI.ArcGIS.DataManagementTools.JoinField
                    pJoin.in_data = m_lyrMainProcessingLayer
                    pJoin.in_field = "ET_TEMP_OBJ_ID2"
                    pJoin.join_field = "ET_TEMP_OBJ_ID2"
                    pJoin.join_table = m_lyrLUCentroids
                    pGPSetup.AddOutputsToMap = True
                    RunTool(pGPSetup, pJoin)
                    pGPSetup.AddOutputsToMap = False
                    pJoin = Nothing

                    'For intCount = 0 To m_lyrMainProcessingLayer.FeatureClass.Fields.FieldCount - 1
                    '    MessageBox.Show(m_lyrMainProcessingLayer.FeatureClass.Fields.Field(intCount).Name, intCount.ToString)
                    'Next

                    'pTable = CType(m_lyrMainProcessingLayer.FeatureClass, ITable)
                    'pCalc = New ESRI.ArcGIS.DataManagementTools.CalculateField
                    'pCalc.expression = "[EX_LU_1]"     '"""" & "[EX_LU_1]" & """"
                    ''pCalc.expression_type = "VB"
                    'pCalc.field = "EX_LU"
                    'pCalc.in_table = pTable
                    'RunTool(pGPSetup, pCalc)
                    'pCalc = Nothing

                    'pRemoveJoin = New ESRI.ArcGIS.DataManagementTools.RemoveJoin
                    'pRemoveJoin.in_layer_or_view = pTable
                    'RunTool(pGPSetup, pRemoveJoin)
                    'pRemoveJoin = Nothing

                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try
            End If

        End If



        GoTo CleanUp

CleanUp:
        m_strProcessing = m_strProcessing & vbNewLine & "End function CreateLandUseFeatClass: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
        pDataset = Nothing
        strParcelFilename = Nothing
        strLanduseFilename = Nothing
        pClip = Nothing
        pWksFactory = Nothing
        pFeatWks = Nothing
        pFeatClass = Nothing
        strLandUseField = Nothing
        pTable = Nothing
        intCount = Nothing
        pField = Nothing
        strFieldName = Nothing
        pFeatureCursor = Nothing
        pFeat = Nothing
        intFomLU = Nothing
        intToLU = Nothing
        strFomLU = Nothing
        strToLU = Nothing
        intLU = Nothing
        arrFieldNames = Nothing
        pUnion = Nothing
        strInFeatures = Nothing
        strUnionName = Nothing
        sb = Nothing
        'pDissolve = Nothing
        'pSummarize = Nothing
        pJoin = Nothing
        'pSummaryTable = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Function

    Private Sub CreateLUCentroidLayer()
        'CREATE A CENTROID POINT FEATURE LAYER FROM THE DEFINED EXISTING LAND USE LAYER
        m_strProcessing = m_strProcessing & vbNewLine & "Starting function CreateLUCentroidLayer: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
        Dim pWksFactory As IWorkspaceFactory = Nothing
        Dim pFeatWks As IFeatureWorkspace
        Dim pToFeatClass As IFeatureClass = Nothing
        Dim pTable As ITable
        Dim intToPntXFld As Integer = -1
        Dim intToPntYFld As Integer = -1
        Dim pToFeatureCursor As IFeatureCursor
        Dim intTotalCount As Integer = 0
        Dim pFeatSelection As IFeatureSelection
        Dim pToFeat As IFeature
        Dim pArea As IArea

        If pGPSetup Is Nothing Then
            CreateSetupGP()
        End If

        If Not m_lyrMainProcessingLayer Is Nothing Then

            Me.lblRunStatus.Text = "Creating Land Use centroid point layer"

            Try
                If m_lyrMainProcessingLayer Is Nothing Then
                    GoTo CleanUp
                Else
                    pToFeatClass = m_lyrMainProcessingLayer.FeatureClass
                    pTable = CType(pToFeatClass, Table)
                End If

                'ADD CENTROID POINT XY FIELDS IF THERE ARE SUB AREAS SELECTED
                If pToFeatClass.FindField("ET_CENTROID_X") <= -1 Then
                    AddEnvisionField(pTable, "ET_CENTROID_X", "DOUBLE", 16, 6)
                End If
                intToPntXFld = pToFeatClass.FindField("ET_CENTROID_X")
                If pToFeatClass.FindField("ET_CENTROID_Y") <= -1 Then
                    AddEnvisionField(pTable, "ET_CENTROID_Y", "DOUBLE", 16, 6)
                End If
                intToPntYFld = pToFeatClass.FindField("ET_CENTROID_Y")


                'CYCLE THROUGH EACH RECORD IN THE FINAL PARCEL LAYER
                pToFeatureCursor = pToFeatClass.Search(Nothing, False)
                intTotalCount = pToFeatClass.FeatureCount(Nothing)
                pFeatSelection = CType(m_lyrEnvisionParcel, IFeatureSelection)
                Me.lblRunStatus.Text = "Calc XY polygon centroids: " & intTotalCount.ToString
                Me.Refresh()
                m_strProcessing = m_strProcessing & vbNewLine & "Starting to process each output feature, total feature count: " & intTotalCount.ToString
                Me.lblRunStatus.Text = "Calculating XY Coordinates"
                intTotalCount = 1
                pToFeat = pToFeatureCursor.NextFeature
                Do While Not pToFeat Is Nothing
                    'POPULATE POLYGON CENTROID XY COORDINATES
                    Try
                        pArea = pToFeat.Shape
                        pToFeat.Value(intToPntXFld) = pArea.Centroid.X
                        pToFeat.Value(intToPntYFld) = pArea.Centroid.Y
                        pToFeat.Store()
                    Catch ex As Exception
                        m_strProcessing = m_strProcessing & vbNewLine & "Error writing centroid XY values.  Failed on feature count:  " & intTotalCount.ToString & vbNewLine & "Error: " & ex.Message
                    End Try
                    intTotalCount = intTotalCount + 1
                    Me.lblRunStatus.Text = "Calc XY polygon centroids: " & intTotalCount.ToString
                    Me.Refresh()
                    pToFeat = pToFeatureCursor.NextFeature
                Loop

                'CREATE A POINT FEATURE CLASS FROM THE CENTROID XY VALUES FOR USE IN SUBAREA MATCHING
                CreateLUCentroidFeatClass()
                CreateLUCentroidPointFeatures(pToFeatClass)

                GoTo CleanUp

            Catch ex As Exception
                m_strProcessing = m_strProcessing & vbNewLine & "Error in function CreateLUCentroidLayer.  Failed on feature count:  " & intTotalCount.ToString & vbNewLine & "Error: " & ex.Message
                GoTo CleanUp
            End Try
        End If

CleanUp:
       
        m_strProcessing = m_strProcessing & vbNewLine & "End function CreateLUCentroidLayer: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
        Me.lblRunStatus.Text = ""
        pWksFactory = Nothing
        pFeatWks = Nothing
        pToFeatClass = Nothing
        pTable = Nothing
        intToPntXFld = Nothing
        intToPntYFld = Nothing
        pToFeatureCursor = Nothing
        intTotalCount = Nothing
        pFeatSelection = Nothing
        pToFeat = Nothing
        pArea = Nothing
        GC.Collect()
        GC.WaitForPendingFinalizers()
    End Sub

    Private Function CreateLUCentroidFeatClass() As Boolean
        CreateLUCentroidFeatClass = True
        m_strProcessing = m_strProcessing & vbNewLine & "Starting function CreateLUCentroidFeatClass: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
        'CREATE THE PONT FEATURE CLASS, WHICH CONTAINS THE CENTROID POINT LOCATIONS OF ALL THE EXISTING LAND USE FEATURES
        Dim pCreateFeatClass As ESRI.ArcGIS.DataManagementTools.CreateFeatureclass
        Dim pWksFactory As IWorkspaceFactory = Nothing
        Dim pFeatWks As IFeatureWorkspace
        Dim pFeatClass As IFeatureClass
        Dim pTable As ITable

        Try
            If pGPSetup Is Nothing Then
                CreateSetupGP()
            End If

            pCreateFeatClass = New ESRI.ArcGIS.DataManagementTools.CreateFeatureclass
            pCreateFeatClass.out_path = m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb"
            pCreateFeatClass.out_name = "PROJECT_LU_CENTROIDS"
            pCreateFeatClass.geometry_type = "POINT"
            pCreateFeatClass.spatial_reference = m_pETSpatRefProject
            RunTool(pGPSetup, pCreateFeatClass)
            pCreateFeatClass = Nothing

            'DEFINE THE GRID LAYER
            pWksFactory = New FileGDBWorkspaceFactory
            pFeatWks = pWksFactory.OpenFromFile(m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb", 0)
            pFeatClass = pFeatWks.OpenFeatureClass("PROJECT_LU_CENTROIDS")
            pTable = CType(pFeatClass, ITable)
            'ADD FIELD FOR ENVISION IDS
            AddEnvisionField(pTable, "ET_TEMP_OBJ_ID2", "INTEGER", 16, 0)
            If pFeatClass.ShapeType = esriGeometryType.esriGeometryPoint Then
                m_lyrLUCentroids = New FeatureLayer
                m_lyrLUCentroids.FeatureClass = pFeatClass
                m_lyrLUCentroids.Name = m_lyrGrid.Name & "<SETUP> " & "EX_LU CENTROID LAYER"
            End If
        Catch ex As Exception
            m_strProcessing = m_strProcessing & vbNewLine & "Error in creating centroid feature class (PROJECT_LU_CENTROIDS): " & ex.Message
            CreateLUCentroidFeatClass = False
            GoTo CleanUp
        End Try
CleanUp:
        m_strProcessing = m_strProcessing & vbNewLine & "End function CreateLUCentroidFeatClass: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
        pCreateFeatClass = Nothing
        pWksFactory = Nothing
        pFeatWks = Nothing
        pFeatClass = Nothing
        GC.Collect()
        GC.WaitForPendingFinalizers()
    End Function

    Private Function CreateLUCentroidPointFeatures(ByVal pLUFeatClass As IFeatureClass) As Boolean
        'GENERATE THE CENTROIDS FOR EACH PARCEL POLYGON BASE UPON THE FEATURE'S AREA CENTROID
        m_strProcessing = m_strProcessing & vbNewLine & "Starting function CreateLUCentroidPointFeatures: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
        CreateLUCentroidPointFeatures = True

        Dim pCentroidFeatClass As IFeatureClass
        Dim intCount As Integer
        Dim pFeat As IFeature
        Dim pToFeatureCursor As IFeatureCursor
        Dim pArea As IArea
        Dim pInsertFeatureBuffer As IFeatureBuffer
        Dim pInsertFeatureCursor As IFeatureCursor = Nothing
        Dim pPointTemp As ESRI.ArcGIS.Geometry.Point
        Dim intIdToField As Integer

        Try
            'EXIT IF THERE IS NOT CENTROID FEATURECLASS
            If m_lyrLUCentroids Is Nothing Then
                CreateLUCentroidPointFeatures = False
                GoTo CleanUp
            Else
                pCentroidFeatClass = m_lyrLUCentroids.FeatureClass
            End If

            'RETRIEVE THE FROM AND TO ENVISION ID 2 FIELD NUMBERS
            intIdToField = pCentroidFeatClass.FindField("ET_TEMP_OBJ_ID2")
            If intIdToField <= -1 Then
                m_strProcessing = m_strProcessing & vbNewLine & "The id field, ET_TEMP_OBJ_ID2, was missing from the parcel or centroid feature classes. "
                GoTo CleanUp
            End If

            'CYCLE THROUGH THE PARCEL FEATURE CLASS TO BUILD A CENTROID FOR EACH PARCEL
            pToFeatureCursor = m_lyrMainProcessingLayer.Search(Nothing, False)
            pFeat = pToFeatureCursor.NextFeature
            Do While Not pFeat Is Nothing
                Try
                    pArea = pFeat.Shape
                    If Not pArea Is Nothing Then
                        pPointTemp = New ESRI.ArcGIS.Geometry.Point
                        pPointTemp.X = pArea.Centroid.X
                        pPointTemp.Y = pArea.Centroid.Y
                        'FEATURE BUFFER SETUP
                        pInsertFeatureCursor = pCentroidFeatClass.Insert(True)
                        pInsertFeatureBuffer = pCentroidFeatClass.CreateFeatureBuffer
                        pInsertFeatureBuffer.Shape = pPointTemp
                        pInsertFeatureBuffer.Value(intIdToField) = pFeat.Value(0)
                        pInsertFeatureCursor.InsertFeature(pInsertFeatureBuffer)
                        pPointTemp = Nothing
                        pArea = Nothing
                    End If
                Catch ex As Exception
                    m_strProcessing = m_strProcessing & vbNewLine & "Error writing centroid point.  Failed on feature count:  " & intCount.ToString & vbNewLine & "Error: " & ex.Message
                End Try
                intCount = intCount + 1
                Me.lblRunStatus.Text = "Creating Centroid: " & intCount.ToString
                Me.Refresh()
                pFeat = pToFeatureCursor.NextFeature
            Loop

            If Not pInsertFeatureCursor Is Nothing Then
                pInsertFeatureCursor.Flush()
            End If

        Catch ex As Exception
            m_strProcessing = m_strProcessing & vbNewLine & "Error in creating centroid point features.  " & vbNewLine & ex.Message
            CreateLUCentroidPointFeatures = False
            GoTo CleanUp
        End Try

CleanUp:
        m_strProcessing = m_strProcessing & vbNewLine & "End function CreateLUCentroidPointFeatures: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
        pCentroidFeatClass = Nothing
        intCount = Nothing
        pFeat = Nothing
        pToFeatureCursor = Nothing
        pArea = Nothing
        pInsertFeatureBuffer = Nothing
        pInsertFeatureCursor = Nothing
        pPointTemp = Nothing
        intIdToField = Nothing
        GC.Collect()
        GC.WaitForPendingFinalizers()
    End Function

    Private Function CreateConstraintsMashup() As Boolean
        'INTERSECT ALL THE LAYERS DEISGNATED AS CONSTRAINTS LAYER(S)
        m_strProcessing = m_strProcessing & vbNewLine & vbNewLine & "Starting function CreateFinalMashup: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
        CreateConstraintsMashup = True
        Dim intRow As Integer
        Dim arrConstrainedLanduses As ArrayList = New ArrayList
        Dim pTable As ITable
        Dim pCursor As ICursor
        Dim pCalc As ICalculator
        Dim pQFilter As IQueryFilter = New QueryFilter
        Dim pFeatSel As IFeatureSelection
        Dim intTotalRec As Integer
        Dim pDataset As IDataset = Nothing
        Dim sb As System.Text.StringBuilder = New System.Text.StringBuilder
        Dim strInput1 As String = ""
        Dim strInput2 As String = ""
        Dim pUnion As ESRI.ArcGIS.AnalysisTools.Union
        Dim strUnionName As String
        Dim pExplode As ESRI.ArcGIS.DataManagementTools.MultipartToSinglepart
        Dim pWksFactory As IWorkspaceFactory = Nothing
        Dim pFeatWks As IFeatureWorkspace
        Dim pFeatClass As IFeatureClass = Nothing
        Dim pMerge As ESRI.ArcGIS.DataManagementTools.Merge
        Dim pFeatLayer As IFeatureLayer
        Dim pSummarize As ESRI.ArcGIS.AnalysisTools.Statistics
        Dim pFeatClassParcelsWithConstraints As IFeatureClass = Nothing
        Dim pFeatClassMultiPart As IFeatureClass = Nothing
        Dim pFeatClassConstrainedParcels As IFeatureClass = Nothing

        Try
            If pGPSetup Is Nothing Then
                CreateSetupGP()
            End If

            'CHECK TO SEE IF THERE ARE ANY LAND USES LOADED AND IF ANY ARE CONSTRAINED.
            'CONSTRAINTS ARE ONLY APPLIED TO SELECTED TO LAND USES AND NO LAND USE ASSIGNED 
            If Me.dgvLandUseAttributes.Rows.Count > 0 Then
                For intRow = 0 To Me.dgvLandUseAttributes.Rows.Count - 1
                    If Me.dgvLandUseAttributes.Rows(intRow).Cells(8).Value = 1 Then
                        arrConstrainedLanduses.Add(Me.dgvLandUseAttributes.Rows(intRow).Cells(1).Value)
                    End If
                Next
            End If

            'ADD A TRACKING FIELD TO DESIGNATE THOSE FEATURE(S) TO WHICH CONSTRIANTS MAY BE APPLIED
            If m_lyrMainProcessingLayer Is Nothing Then
                GoTo CleanUp
            Else
                pTable = CType(m_lyrMainProcessingLayer.FeatureClass, ITable)
                intTotalRec = pTable.RowCount(Nothing)
                If pTable.FindField("APPLY_CONSTRAINTS") <= -1 Then
                    If Not AddEnvisionField(pTable, "APPLY_CONSTRAINTS", "INTEGER", 1, 0) Then
                        m_strProcessing = m_strProcessing & vbNewLine & "Unable to add required field, APPLY_CONSTRAINTS"
                        CreateConstraintsMashup = False
                        GoTo CleanUp
                    End If
                End If
                'BY DEFAULT MAKE EVERY FEATURE SET TO APPLY CONSTRAINTS
                Try
                    pCalc = New Calculator
                    pCursor = pTable.Update(Nothing, False)
                    With pCalc
                        .Cursor = pCursor
                        .Field = "APPLY_CONSTRAINTS"
                        .Expression = 1
                    End With
                    pCalc.Calculate()
                Catch ex As Exception
                    MessageBox.Show(ex.Message, "Calculate All Field Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
                End Try
                'QUERY EACH LAND USE WHERE CONSTRAINTS CAN NOT BE APPLIED AND CALC TO 0
                If Me.dgvLandUseAttributes.Rows.Count > 0 Then
                    For intRow = 0 To Me.dgvLandUseAttributes.Rows.Count - 1
                        If Me.dgvLandUseAttributes.Rows(intRow).Cells(8).Value = 0 Then
                            pQFilter = New QueryFilter
                            pQFilter.WhereClause = """" & "EX_LU" & """" & " = '" & Me.dgvLandUseAttributes.Rows(intRow).Cells(1).Value & "'"
                            Try
                                pCalc = New Calculator
                                pCursor = pTable.Update(pQFilter, False)
                                With pCalc
                                    .Cursor = pCursor
                                    .Field = "APPLY_CONSTRAINTS"
                                    .Expression = 0
                                End With
                                pCalc.Calculate()
                            Catch ex As Exception
                                MessageBox.Show(ex.Message, "Calculate Field Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
                            End Try
                        End If
                    Next
                End If
            End If

            pFeatSel = CType(m_lyrMainProcessingLayer, IFeatureSelection)
            pFeatSel.Clear()
            Try
                pQFilter = New QueryFilter
                pQFilter.WhereClause = """" & "APPLY_CONSTRAINTS" & """" & " = 1"
                pFeatSel.SelectFeatures(pQFilter, esriSelectionResultEnum.esriSelectionResultNew, False)
                If pFeatSel.SelectionSet.Count > 0 Then
                    If intTotalRec = pFeatSel.SelectionSet.Count Then
                        'ALL PARCELS CAN BE CONSTRAINTED, NO NEED TO SUBSET
                        m_lyrConstrainedParcels = m_lyrMainProcessingLayer
                    Else
                        'CREATE 2 LAYERS OF PARCELS CAN AND CAN NOT BE CONSTRAINED
                        If Not CreateConstrainedParcelFeatClasses() Then
                            CreateConstraintsMashup = False
                            GoTo CleanUp
                        End If
                    End If
                Else
                    m_strProcessing = m_strProcessing & vbNewLine & "No constraints layer could was found to union with parcel layer."
                    CreateConstraintsMashup = False
                    GoTo CleanUp
                End If
            Catch ex As Exception
                m_strProcessing = m_strProcessing & vbNewLine & vbNewLine & ex.Message
                GoTo CleanUp
            End Try

            'ADD CONSTRAINED PARCELS TO THE UNION STRING
            Try
                pDataset = CType(m_lyrConstrainedParcels.FeatureClass, IDataset)
                strInput2 = pDataset.Workspace.PathName & "\" & pDataset.BrowseName
                sb.AppendFormat("{0};", strInput2)
            Catch ex As Exception
                m_strProcessing = m_strProcessing & vbNewLine & "Unable to retrieve filename for the layer: m_lyrConstrainedParcels"
                CreateConstraintsMashup = False
                GoTo CleanUp
            End Try

            'ADD CONSTRAINTS TO THE UNION STRING
            Try
                pDataset = CType(m_lyrConstraints.FeatureClass, IDataset)
                strInput1 = pDataset.Workspace.PathName & "\" & pDataset.BrowseName
                If strInput1.Length <= 0 Then
                    m_strProcessing = m_strProcessing & vbNewLine & "No constraints layer could was found to union with parcel layer."
                    CreateConstraintsMashup = False
                    GoTo CleanUp
                Else
                    sb.AppendFormat("{0};", strInput1)
                    'ONLY CONTINUE IF THE INPUTS ARE FROM DIFFERENT FEATURE CLASSES
                    pUnion = New ESRI.ArcGIS.AnalysisTools.Union
                    pUnion.in_features = sb.ToString
                    pUnion.join_attributes = "ALL"
                    pUnion.gaps = "GAPS"
                    pUnion.cluster_tolerance = 1
                    pUnion.out_feature_class = m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb" & "\ENVISION_PARCELS_WITH_CONSTRAINTS"
                    RunTool(pGPSetup, pUnion)
                    pUnion = Nothing
                End If
            Catch ex As Exception
                m_strProcessing = m_strProcessing & vbNewLine & "Error in CreateFinalMashup in attempt to open project workspace:" & vbNewLine & ex.Message
                CreateConstraintsMashup = False
                GoTo CleanUp
            End Try

            'CHECK THAT THE MASHUP WAS CREATED
            'OPEN THE PROJECT WORKSPACE
            Try
                pWksFactory = New FileGDBWorkspaceFactory
                pFeatWks = pWksFactory.OpenFromFile(m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb", 0)
                'DELETE ALL FEATURES THAT DO NOT SHARE PARCEL AREA
                Try
                    pFeatClassParcelsWithConstraints = pFeatWks.OpenFeatureClass("ENVISION_PARCELS_WITH_CONSTRAINTS")
                    pQFilter = New QueryFilter
                    pQFilter.WhereClause = """" & "ET_TEMP_OBJ_ID2" & """" & " <= 0"
                    pTable = CType(pFeatClassParcelsWithConstraints, ITable)
                    pTable.DeleteSearchedRows(pQFilter)
                Catch ex As Exception
                    m_strProcessing = m_strProcessing & vbNewLine & "Error in opening the constrained parcel layer to delete excess polygon feature(s):" & vbNewLine & ex.Message
                    CreateConstraintsMashup = False
                    GoTo CleanUp
                End Try
            Catch ex As Exception
                m_strProcessing = m_strProcessing & vbNewLine & "Error in CreateFinalMashup in attempt to open project workspace:" & vbNewLine & ex.Message
                CreateConstraintsMashup = False
                GoTo CleanUp
            End Try

        Catch ex As Exception
            m_strProcessing = m_strProcessing & vbNewLine & "Error in CreateFinalMashup:" & vbNewLine & ex.Message
            CreateConstraintsMashup = False
            GoTo CleanUp
        End Try

        'EXPLODE THE MASHUP
        Try
            m_strProcessing = m_strProcessing & vbNewLine & "Multipart to Single part features from the Mash-up:  " & Date.Now.ToString("hh:mm:ss tt")
            pExplode = New ESRI.ArcGIS.DataManagementTools.MultipartToSinglepart
            If Not m_lyrUnConstrainedParcels Is Nothing Then
                pExplode.in_features = m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb" & "\ENVISION_PARCELS_WITH_CONSTRAINTS"
                pExplode.out_feature_class = m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb" & "\ENVISION_PARCELS_WITH_CONSTRAINTS_MULTIPART"
            Else
                pExplode.in_features = m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb" & "\ENVISION_PARCELS_WITH_CONSTRAINTS"
                pExplode.out_feature_class = m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb" & "\ENVISION_CONSTRAINED_PARCELS"
            End If
            RunTool(pGPSetup, pExplode)
            pExplode = Nothing
        Catch ex As Exception
            m_strProcessing = m_strProcessing & vbNewLine & "Creating Multipart feature class from constrained parcel layer failed to complete  "
        End Try


        'OPEN THE PROCESSING LAYERS TO DEFINE, WHICH LAYER WILL BE PROCESSED.  
        Try
            pWksFactory = New FileGDBWorkspaceFactory
            pFeatWks = pWksFactory.OpenFromFile(m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb", 0)
            'LAYER OF PARCELS UNIONED WITH CONSTRAINTS LAYER
            Try
                pFeatClassParcelsWithConstraints = pFeatWks.OpenFeatureClass("ENVISION_PARCELS_WITH_CONSTRAINTS")
            Catch ex As Exception
                m_strProcessing = m_strProcessing & vbNewLine & "ENVISION_PARCELS_WITH_CONSTRAINTS layer not found"
                pFeatClassParcelsWithConstraints = Nothing
            End Try
            'LAYER OF PARCELS UNIONED WITH CONSTRAINTS LAYER AND CONVERTED INTO MULTIPART POLYGON LAYER AND WILL BE MERGED WITH NON-CONSTRAINED PARCELS
            Try
                pFeatClassMultiPart = pFeatWks.OpenFeatureClass("ENVISION_PARCELS_WITH_CONSTRAINTS_MULTIPART")
            Catch ex As Exception
                m_strProcessing = m_strProcessing & vbNewLine & "ENVISION_PARCELS_WITH_CONSTRAINTS_MULTIPART layer not found"
                pFeatClassMultiPart = Nothing
            End Try
            'FINAL CONSTRAINED LAYER AS THERE WERE NO NON-CONSTRIANED PARCELS TO ADD BACK IN
            Try
                pFeatClassConstrainedParcels = pFeatWks.OpenFeatureClass("ENVISION_CONSTRAINED_PARCELS")
            Catch ex As Exception
                m_strProcessing = m_strProcessing & vbNewLine & "ENVISION_CONSTRAINED_PARCELS layer not found"
                pFeatClassConstrainedParcels = Nothing
            End Try
        Catch ex As Exception
            m_strProcessing = m_strProcessing & vbNewLine & "Error in CreateFinalMashup in attempt to open project workspace:" & vbNewLine & ex.Message
            CreateConstraintsMashup = False
            GoTo CleanUp
        End Try

        'SET WHICH FEATURE CLASS WILL BE PASSED ALONG FOR FURTHER PROCESSING
        sb = New System.Text.StringBuilder
        sb.AppendFormat("{0};", m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb" & "\PROJECT_UNCONSTRAINED_PARCELS")
        If pFeatClassParcelsWithConstraints Is Nothing Then
            sb.AppendFormat("{0};", m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb" & "\ENVISION_PARCELS_WITH_CONSTRAINTS")
            pFeatClass = pFeatClassParcelsWithConstraints
        End If
        If Not pFeatClassMultiPart Is Nothing Then
            sb.AppendFormat("{0};", m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb" & "\ENVISION_PARCELS_WITH_CONSTRAINTS_MULTIPART")
            pFeatClass = pFeatClassMultiPart
        End If
        If Not pFeatClassConstrainedParcels Is Nothing Then
            sb.AppendFormat("{0};", m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb" & "\ENVISION_CONSTRAINED_PARCELS")
            pFeatClass = pFeatClassConstrainedParcels
        End If
        pFeatClassConstrainedParcels = Nothing
        pFeatClassMultiPart = Nothing
        pFeatClassParcelsWithConstraints = Nothing
        If pFeatClass Is Nothing Then
            m_strProcessing = m_strProcessing & vbNewLine & "A feature class could not be defined for further processing:"
            CreateConstraintsMashup = False
            GoTo CleanUp
        End If


        'CHECK THAT THE MASHUP WAS CREATED
        'OPEN THE PROJECT WORKSPACE
        Try
            pWksFactory = New FileGDBWorkspaceFactory
            pFeatWks = pWksFactory.OpenFromFile(m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb", 0)
            'DELETE ALL FEATURES THAT DO NOT SHARE PARCEL AREA
            Try
                pQFilter = New QueryFilter
                pQFilter.WhereClause = """" & "ET_TEMP_OBJ_ID2" & """" & " <= 0"
                pTable = CType(pFeatClass, ITable)
                pTable.DeleteSearchedRows(pQFilter)
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        Catch ex As Exception
            m_strProcessing = m_strProcessing & vbNewLine & "Unable to open feature class and delete polygons that are not required" & vbNewLine & ex.Message
            CreateConstraintsMashup = False
            GoTo CleanUp
        End Try

        Try
            If Not m_lyrUnConstrainedParcels Is Nothing Then
                pMerge = New ESRI.ArcGIS.DataManagementTools.Merge
                pMerge.inputs = sb.ToString
                pMerge.field_mappings = "ALL"
                pMerge.output = m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb" & "\ENVISION_CONSTRAINED_PARCELS"
                RunTool(pGPSetup, pMerge)
                pUnion = Nothing
            End If
        Catch ex As Exception
            m_strProcessing = m_strProcessing & vbNewLine & "Error oin merged unconstrained parcels with constrained parcels:" & vbNewLine & ex.Message
            GoTo CleanUp
        End Try

        'OPEN AND DEFINE THE GLOBAL CONTRAINED PARCEL LAYER
        Try
            pWksFactory = New FileGDBWorkspaceFactory
            pFeatWks = pWksFactory.OpenFromFile(m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb", 0)
            pFeatClass = pFeatWks.OpenFeatureClass("ENVISION_CONSTRAINED_PARCELS")
            m_lyrConstrainedParcels = New FeatureLayer
            m_lyrConstrainedParcels.FeatureClass = pFeatClass
            m_lyrConstrainedParcels.Name = m_lyrConstrainedParcels.Name & "<SETUP> " & "PROJECT CONSTRAINED PARCELS"
        Catch ex As Exception
            m_strProcessing = m_strProcessing & vbNewLine & "Error opening the feature layer:  ENVISION_CONSTRAINED_PARCELS" & vbNewLine & ex.Message
            CreateConstraintsMashup = False
            GoTo CleanUp
        End Try

        'CREATE A SUMMARY TABLE THAT WILL BE ONLY THE CONSTRAINED AREA FOR EACH PARCEL.  THOSE PARCELS WITH ZERO CONSTRAINED AREA WILL NOT BE IN THE TABLE
        Try
            pSummarize = New ESRI.ArcGIS.AnalysisTools.Statistics
            pSummarize.case_field = "ET_TEMP_OBJ_ID2"
            pSummarize.statistics_fields = "ET_CONSTRAINED MAX;Shape_Area SUM"
            pFeatClass = pFeatWks.OpenFeatureClass("ENVISION_CONSTRAINED_PARCELS")
            pSummarize.in_table = m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb" & "\ENVISION_CONSTRAINED_PARCELS"
        Catch ex As Exception
            m_strProcessing = m_strProcessing & vbNewLine & "Error in creating summary table for constrained parcels." & vbNewLine & ex.Message
            CreateConstraintsMashup = False
            GoTo CleanUp
        End Try

        Try
            pSummarize.out_table = m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb" & "\ENVISION_CONSTRAINED_PARCEL_AREA_SUMMARIZED"
            RunTool(pGPSetup, pSummarize)
            pSummarize = Nothing
            pTable = pFeatWks.OpenTable("ENVISION_CONSTRAINED_PARCEL_AREA_SUMMARIZED")
            Try
                pQFilter = New QueryFilter
                pQFilter.WhereClause = """" & "MAX_ET_CONSTRAINED" & """" & " = 0"
                pTable.DeleteSearchedRows(pQFilter)
            Catch ex As Exception
                m_strProcessing = m_strProcessing & vbNewLine & "Error in removing constrained summary values from table, ENVISION_CONSTRAINED_PARCEL_AREA_SUMMARIZED: " & vbNewLine & ex.Message & vbNewLine & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
                CreateConstraintsMashup = False
                GoTo CleanUp
            End Try
        Catch ex As Exception
            m_strProcessing = m_strProcessing & vbNewLine & "Error in opening table, ENVISION_CONSTRAINED_PARCEL_AREA_SUMMARIZED: " & vbNewLine & ex.Message & vbNewLine & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
            CreateConstraintsMashup = False
            GoTo CleanUp
        End Try


CleanUp:
        m_strProcessing = m_strProcessing & vbNewLine & "End function CreateFinalMashup: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
        intRow = Nothing
        arrConstrainedLanduses = Nothing
        pTable = Nothing
        pCursor = Nothing
        pCalc = Nothing
        pQFilter = Nothing
        pFeatSel = Nothing
        intTotalRec = Nothing
        pDataset = Nothing
        sb = Nothing
        strInput1 = Nothing
        strInput2 = Nothing
        pUnion = Nothing
        strUnionName = Nothing
        pExplode = Nothing
        pWksFactory = Nothing
        pFeatWks = Nothing
        pFeatClass = Nothing
        pMerge = Nothing
        pFeatLayer = Nothing
        pFeatClassParcelsWithConstraints = Nothing
        pFeatClassMultiPart = Nothing
        pFeatClassConstrainedParcels = Nothing
        GC.Collect()
        GC.WaitForPendingFinalizers()

    End Function

    Private Function CreateConstrainedParcelFeatClasses() As Boolean
        'CREATE THE MINIMUM ACRE FEATURE LAYER
        'USED FOR HYBRID OPTION
        m_strProcessing = m_strProcessing & vbNewLine & "Starting function CreateConstrainedParcelFeatClasses: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
        CreateConstrainedParcelFeatClasses = True
        Dim strQuery As String
        Dim pSelect As ESRI.ArcGIS.AnalysisTools.Select
        Dim pWksFactory As IWorkspaceFactory = Nothing
        Dim pFeatWks As IFeatureWorkspace
        Dim pFeatClass As IFeatureClass
        Dim pDataset As IDataset

        Try
            'BUILD THE QUERY STRING TO SELECT OUT PARCELS TO WHICH CONSTRAINTS CAN BE APPLIED
            strQuery = """" & "APPLY_CONSTRAINTS" & """" & " = 1"
            If Not m_lyrConstrainedParcels Is Nothing Then
                pDataset = CType(m_lyrEnvisionParcel.FeatureClass, IDataset)
                pSelect = New ESRI.ArcGIS.AnalysisTools.Select
                pSelect.in_features = pDataset.Workspace.PathName & "\" & pDataset.BrowseName
                pSelect.where_clause = strQuery
                pSelect.out_feature_class = m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb" & "\PROJECT_POSSIBLE_CONSTRAINED_PARCELS"
                RunTool(pGPSetup, pSelect)
                pSelect = Nothing
            End If
            'DEFINE THE CONSTRAINTED LAYER
            pWksFactory = New FileGDBWorkspaceFactory
            pFeatWks = pWksFactory.OpenFromFile(m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb", 0)
            pFeatClass = pFeatWks.OpenFeatureClass("PROJECT_POSSIBLE_CONSTRAINED_PARCELS")
            If pFeatClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                m_lyrConstrainedParcels = New FeatureLayer
                m_lyrConstrainedParcels.FeatureClass = pFeatClass
                m_lyrConstrainedParcels.Name = m_lyrConstrainedParcels.Name & "<SETUP> " & "PROJECT POSSIBLE CONSTRAINED PARCELS"
            End If
        Catch ex As Exception
            m_strProcessing = m_strProcessing & vbNewLine & "An error occured attempting to select out parcels to which constraints can be applied." & vbNewLine & ex.Message
            CreateConstrainedParcelFeatClasses = False
            GoTo CleanUp
        End Try

        Try

            'BUILD THE QUERY STRING TO SELECT OUT PARCELS TO WHICH CONSTRAINTS CAN NOT BE APPLIED
            strQuery = """" & "APPLY_CONSTRAINTS" & """" & " = 0"
            If Not m_lyrConstrainedParcels Is Nothing Then
                pDataset = CType(m_lyrEnvisionParcel.FeatureClass, IDataset)
                pSelect = New ESRI.ArcGIS.AnalysisTools.Select
                pSelect.in_features = pDataset.Workspace.PathName & "\" & pDataset.BrowseName
                pSelect.where_clause = strQuery
                pSelect.out_feature_class = m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb" & "\PROJECT_UNCONSTRAINED_PARCELS"
                RunTool(pGPSetup, pSelect)
                pSelect = Nothing
            End If
            'DEFINE THE CONSTRAINTED LAYER
            pWksFactory = New FileGDBWorkspaceFactory
            pFeatWks = pWksFactory.OpenFromFile(m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb", 0)
            pFeatClass = pFeatWks.OpenFeatureClass("PROJECT_UNCONSTRAINED_PARCELS")
            If pFeatClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                m_lyrUnConstrainedParcels = New FeatureLayer
                m_lyrUnConstrainedParcels.FeatureClass = pFeatClass
                m_lyrUnConstrainedParcels.Name = m_lyrUnConstrainedParcels.Name & "<SETUP> " & "PROJECT UNCONSTRAINED PARCELS"
            End If
        Catch ex As Exception
            m_strProcessing = m_strProcessing & vbNewLine & "An error occured attempting to select out parcels to which constraints can NOT be applied." & vbNewLine & ex.Message
            CreateConstrainedParcelFeatClasses = False
            GoTo CleanUp
        End Try
CleanUp:
        m_strProcessing = m_strProcessing & vbNewLine & "End function CreateConstrainedParcelFeatClasses: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
        pSelect = Nothing
        pWksFactory = Nothing
        pFeatWks = Nothing
        pFeatClass = Nothing
        pDataset = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Function

    Private Function CopyEnvisionExcelfile() As Boolean
        CopyEnvisionExcelfile = True
        'COPY THE SELECTED OR DEFAULT ENVISION EXCEL FILE
        Dim strFile As String = ""
        Try
            strFile = Me.tbxExcel.Text.Split("\")(Me.tbxExcel.Text.Split("\").Length - 1)
            If Not File.Exists(m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & strFile) Then
                File.Copy(Me.tbxExcel.Text, m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & strFile)
            End If
        Catch ex As Exception
            CopyEnvisionExcelfile = False
        End Try
    End Function

    Private Function CreateSlopeFeatureClasses() As Boolean
        'CREATE THE SLOPE AND HILLSHADE GRID, AND THE CONSTRAINTS SLOPE FEATURE CLASS(ES)
        CreateSlopeFeatureClasses = True

        Dim pRLyr As IRasterLayer
        Dim pRaster2 As IRaster2 = Nothing
        Dim objExtent As Object
        Dim objOutputSpatref As Object
        Dim objWorkspace As Object
        Dim pTempRLyr As IRasterLayer
        Dim pCreateSlope As ESRI.ArcGIS.SpatialAnalystTools.Slope
        Dim pCreateHillshade As ESRI.ArcGIS.SpatialAnalystTools.HillShade
        Dim pCreate15Slope As ESRI.ArcGIS.SpatialAnalystTools.Reclassify
        Dim pCreate15Lyr As ESRI.ArcGIS.ConversionTools.RasterToPolygon
        Dim pCreate25Slope As ESRI.ArcGIS.SpatialAnalystTools.Reclassify
        Dim pCreate25Lyr As ESRI.ArcGIS.ConversionTools.RasterToPolygon

        Try
            If pGPSetup Is Nothing Then
                CreateSetupGP()
            End If

            If m_arrETSetup_RLayers.Count > 0 Then
                CreateSlopeFeatureClasses = False
            End If

            'RASTER LAYER VARIABLES
            pRLyr = CType(m_arrETSetup_RLayers.Item(Me.cmbGridLayers.SelectedIndex), IRasterLayer)
            pRaster2 = CType(pRLyr.Raster, IRaster2)

            'HILLSHADE
            Me.lblRunStatus.Text = "Creating Hillshade Grid"
            Me.Refresh()
            pCreateHillshade = New ESRI.ArcGIS.SpatialAnalystTools.HillShade
            pCreateHillshade.in_raster = pRaster2
            pCreateHillshade.altitude = "45"
            pCreateHillshade.azimuth = "315"
            pCreateHillshade.out_raster = (m_frmEnvisionProjectSetup.tbxWorkspace.Text + "\Hillshade")
            pCreateHillshade.z_factor = "3.2808"
            RunTool(pGPSetup, pCreateHillshade)
            pCreateHillshade = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()

            'Create Slope Grid from selected raster
            Me.lblRunStatus.Text = "Creating Slope Grid"
            Me.Refresh()
            pCreateSlope = New ESRI.ArcGIS.SpatialAnalystTools.Slope
            pCreateSlope.in_raster = pRaster2
            pCreateSlope.output_measurement = "DEGREE"
            pCreateSlope.out_raster = (m_frmEnvisionProjectSetup.tbxWorkspace.Text + "\Slope")
            pCreateSlope.z_factor = "3.2808"
            RunTool(pGPSetup, pCreateSlope)
            pCreateSlope = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()

            '15-25% SLOPES
            Me.lblRunStatus.Text = "Reclassing 15-25% Slope Grid"
            Me.Refresh()
            pCreate15Slope = New ESRI.ArcGIS.SpatialAnalystTools.Reclassify
            pCreate15Slope.in_raster = pRaster2
            pCreate15Slope.out_raster = (m_frmEnvisionProjectSetup.tbxWorkspace.Text + "\Slp15to25")
            pCreate15Slope.remap = "15 25 1"
            pCreate15Slope.missing_values = "NODATA"
            pCreate15Slope.reclass_field = "Value"
            RunTool(pGPSetup, pCreate15Slope)
            pCreate15Slope = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()

            'RASRTER TO POLYGONS 15-25% SLOPES
            Me.lblRunStatus.Text = "Please Wait, Coverting Raster to Polygons"
            Me.Refresh()
            pCreate15Lyr = New ESRI.ArcGIS.ConversionTools.RasterToPolygon
            pCreate15Lyr.in_raster = pRaster2
            pCreate15Lyr.out_polygon_features = (m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb" & "\SLOPE_15to25")
            pCreate15Lyr.simplify = "NO_SIMPLIFY"
            pCreate15Lyr.raster_field = "Value"
            RunTool(pGPSetup, pCreate15Lyr)
            pCreate15Lyr = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()

            'GREATER THAN 25% SLOPES
            Me.lblRunStatus.Text = "Reclassing 25+% Slope Grid"
            Me.Refresh()
            pCreate25Slope = New ESRI.ArcGIS.SpatialAnalystTools.Reclassify
            pRaster2 = CType(m_pSlopeLyr.Raster, IRaster2)
            pCreate25Slope.in_raster = pRaster2
            pCreate25Slope.out_raster = (m_frmEnvisionProjectSetup.tbxWorkspace.Text + "\" + "Slp25Plus")
            pCreate25Slope.remap = "25 100 1"
            pCreate25Slope.missing_values = "NODATA"
            pCreate25Slope.reclass_field = "Value"
            RunTool(pGPSetup, pCreate25Slope)
            pCreate25Slope = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()

            'RASRTER TO POLYGONS 15-25% SLOPES
            Me.lblRunStatus.Text = "Converting Raster to Polygons"
            Me.Refresh()
            pCreate25Lyr = New ESRI.ArcGIS.ConversionTools.RasterToPolygon
            pCreate25Lyr.in_raster = pRaster2
            pCreate25Lyr.out_polygon_features = (m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb" & "\SLOPE_25PLUS")
            pCreate25Lyr.simplify = "NO_SIMPLIFY"
            pCreate25Lyr.raster_field = "Value"
            RunTool(pGPSetup, pCreate25Lyr)
            pCreate25Lyr = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()

            GoTo CleanUp
        Catch ex As Exception
            CreateSlopeFeatureClasses = False
            MessageBox.Show(ex.Message.ToString, "Envision Slope/Hillshade Data Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo CleanUp
        End Try


CleanUp:
        Me.lblRunStatus.Text = ""
        pRLyr = Nothing
        pRaster2 = Nothing
        objExtent = Nothing
        objOutputSpatref = Nothing
        objWorkspace = Nothing
        pTempRLyr = Nothing
        pCreateSlope = Nothing
        pCreateHillshade = Nothing
        pCreate15Slope = Nothing
        pCreate15Lyr = Nothing
        pCreate25Slope = Nothing
        pCreate25Lyr = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()

    End Function

    Private Function CalAcres(ByVal pTable As ITable, ByVal strFieldName As String) As Boolean
        'ADD AN ACRES FIELD TO THE DISIGNATED TABLE GIVEN THE INPUT UNITS
        m_strProcessing = m_strProcessing & vbNewLine & "Starting function CalAcres: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
        CalAcres = True
        Dim pCursor As ICursor
        Dim pCalc As ICalculator
        Dim strExpression As String = ""
        Dim strUnits As String = ""

        If Me.lblProjectUnits.Text.Length <= 0 Then
            m_strProcessing = m_strProcessing & vbNewLine & "The projection units were not set."
            GoTo CleanUp
        Else
            strUnits = UCase(Me.lblProjectUnits.Text)
            If strUnits.Contains("FOOT") Then
                strExpression = "[Shape_Area]" & " * 0.0000229568411"
            ElseIf strUnits.Contains("YARD") Then
                strExpression = "[Shape_Area]" & " * 0.00020661157"
            ElseIf strUnits.Contains("MILE") Then
                strExpression = "[Shape_Area]" & " * 640"
            ElseIf strUnits = "METER" Then
                strExpression = "[Shape_Area]" & " * 0.000247105381"
            ElseIf strUnits = "KILOMETER" Then
                strExpression = "[Shape_Area]" & "* 247.105381"
            ElseIf strUnits = "CENTIMETER" Then
                strExpression = "[Shape_Area]" & " * 0.00000229568411"
            End If
        End If

        Try
            If pTable.FindField("Shape_Area") <= -1 Then
                m_strProcessing = m_strProcessing & vbNewLine & "The AREA field could not be found to calculate acre values."
            End If
            pCalc = New Calculator
            pCursor = pTable.Update(Nothing, False)
            With pCalc
                .Cursor = pCursor
                .Field = strFieldName
                .Expression = strExpression
            End With
            pCalc.Calculate()
            GoTo CleanUp
        Catch ex As Exception
            m_strProcessing = m_strProcessing & vbNewLine & "An error has occured calcualting acres: " & vbNewLine & ex.Message
            CalAcres = False
            GoTo CleanUp
        End Try

CleanUp:
        m_strProcessing = m_strProcessing & vbNewLine & "End function CalAcres: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
        pCalc = Nothing
        pCursor = Nothing
        pTable = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()

    End Function

    Private Function CalcTempObjId(ByVal pTable As ITable, ByVal strFieldName As String) As Boolean
        'CLAC THE DESIGNATED FIELD NAME TO BE THAT OF THE OBJECT ID FIELD, FIELD (0)
        m_strProcessing = m_strProcessing & vbNewLine & "Starting function CalcTempObjId: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
        CalcTempObjId = True
        Dim pCursor As ICursor
        Dim pCalc As ICalculator
        Dim strExpression As String = ""
        strExpression = "[" & pTable.Fields.Field(0).Name & "]"
        Try
            pCalc = New Calculator
            pCursor = pTable.Update(Nothing, False)
            With pCalc
                .Cursor = pCursor
                .Field = strFieldName
                .Expression = strExpression
            End With
            pCalc.Calculate()
            GoTo CleanUp
        Catch ex As Exception
            m_strProcessing = m_strProcessing & vbNewLine & "An error has occured calcualting temp object id: " & vbNewLine & ex.Message
            CalcTempObjId = False
            GoTo CleanUp
        End Try

CleanUp:
        m_strProcessing = m_strProcessing & vbNewLine & "End function CalcTempObjId: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
        pCalc = Nothing
        pCursor = Nothing
        pTable = Nothing
        GC.Collect()
        GC.WaitForPendingFinalizers()
    End Function

    Private Function CreateSetupGP() As Boolean
        'CREATE THE GEOPROCESSING OBJECT THAT WILL BE USED TO PROCESS ALL SETUP GEOPROCESSING FUNCTIONS
        CreateSetupGP = True
        Dim objExtent As Object
        Dim objOutputSpatref As Object
        Dim objWorkspace As Object

        Try
            If pGPSetup Is Nothing Then
                objExtent = m_pExtentEnv.XMax.ToString + " " + m_pExtentEnv.YMax.ToString + " " + m_pExtentEnv.XMin.ToString + " " + m_pExtentEnv.YMin.ToString
                objOutputSpatref = m_pETSpatRefProject.FactoryCode
                objWorkspace = m_frmEnvisionProjectSetup.tbxWorkspace.Text
                pGPSetup = New ESRI.ArcGIS.Geoprocessor.Geoprocessor
                pGPSetup.SetEnvironmentValue("cellSize", Me.tbxGridCellSize.Text)
                pGPSetup.SetEnvironmentValue("extent", objExtent)
                pGPSetup.SetEnvironmentValue("workspace", objWorkspace)
                Try
                    pGPSetup.SetEnvironmentValue("outputCoordinateSystem", objOutputSpatref)
                Catch ex As Exception
                    m_strProcessing = m_strProcessing & vbNewLine & "Error in function, CreateSetupGP, error message associated with the selected projection:" & vbNewLine & ex.Message
                End Try
                pGPSetup.OverwriteOutput = True
                pGPSetup.AddOutputsToMap = False
                pGPSetup.TemporaryMapLayers = False
            End If
        Catch ex As Exception
            CreateSetupGP = False
            m_strProcessing = m_strProcessing & vbNewLine & "Error in function, CreateSetupGP, error message:" & vbNewLine & ex.Message
            'MessageBox.Show(ex.Message, "")
        End Try

        GoTo CleanUp

CleanUp:
        objExtent = Nothing
        objOutputSpatref = Nothing
        objWorkspace = Nothing
        GC.Collect()
        GC.WaitForPendingFinalizers()
    End Function

    Private Sub CreateParcelCentroidLayer()
        'CREATE A CENTROID POINT FEATURE LAYER FROM THE DEFINED PARCEL LAYER
        m_strProcessing = m_strProcessing & vbNewLine & "Starting function CreateParcelCentroidLayer: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
        Dim pWksFactory As IWorkspaceFactory = Nothing
        Dim pFeatWks As IFeatureWorkspace
        Dim pToFeatClass As IFeatureClass = Nothing
        Dim pTable As ITable
        Dim intToPntXFld As Integer = -1
        Dim intToPntYFld As Integer = -1
        Dim pToFeatureCursor As IFeatureCursor
        Dim intTotalCount As Integer = 0
        Dim pFeatSelection As IFeatureSelection
        Dim pToFeat As IFeature
        Dim pArea As IArea

        If pGPSetup Is Nothing Then
            CreateSetupGP()
        End If

        If Not m_lyrEnvisionParcel Is Nothing Then

            Me.lblRunStatus.Text = "Acres Summary"

            Try
                If m_lyrEnvisionParcel Is Nothing Then
                    pWksFactory = New FileGDBWorkspaceFactory
                    pFeatWks = pWksFactory.OpenFromFile(m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb", 0)
                    pToFeatClass = pFeatWks.OpenFeatureClass("ENVISION_PARCELS_FINAL")
                    pTable = CType(pToFeatClass, Table)
                    m_lyrEnvisionParcel = Nothing
                    m_lyrEnvisionParcel = New FeatureLayer
                    m_lyrEnvisionParcel.FeatureClass = pToFeatClass
                    m_lyrEnvisionParcel.Name = "<SETUP> " & "PARCEL LAYER"
                Else
                    pToFeatClass = m_lyrEnvisionParcel.FeatureClass
                    pTable = CType(pToFeatClass, Table)
                End If

                'ADD CENTROID POINT XY FIELDS IF THERE ARE SUB AREAS SELECTED
                If pToFeatClass.FindField("ET_CENTROID_X") <= -1 Then
                    AddEnvisionField(pTable, "ET_CENTROID_X", "DOUBLE", 16, 6)
                End If
                intToPntXFld = pToFeatClass.FindField("ET_CENTROID_X")
                If pToFeatClass.FindField("ET_CENTROID_Y") <= -1 Then
                    AddEnvisionField(pTable, "ET_CENTROID_Y", "DOUBLE", 16, 6)
                End If
                intToPntYFld = pToFeatClass.FindField("ET_CENTROID_Y")


                'CYCLE THROUGH EACH RECORD IN THE FINAL PARCEL LAYER
                pToFeatureCursor = pToFeatClass.Search(Nothing, False)
                intTotalCount = pToFeatClass.FeatureCount(Nothing)
                pFeatSelection = CType(m_lyrEnvisionParcel, IFeatureSelection)
                Me.lblRunStatus.Text = "Calc XY polygon centroids: " & intTotalCount.ToString
                Me.Refresh()
                m_strProcessing = m_strProcessing & vbNewLine & "Starting to process each output feature, total feature count: " & intTotalCount.ToString
                intTotalCount = 1
                pToFeat = pToFeatureCursor.NextFeature
                Do While Not pToFeat Is Nothing
                    'POPULATE POLYGON CENTROID XY COORDINATES
                    Try
                        pArea = pToFeat.Shape
                        pToFeat.Value(intToPntXFld) = pArea.Centroid.X
                        pToFeat.Value(intToPntYFld) = pArea.Centroid.Y
                        pToFeat.Store()
                    Catch ex As Exception
                        m_strProcessing = m_strProcessing & vbNewLine & "Error writing centroid XY values.  Failed on feature count:  " & intTotalCount.ToString & vbNewLine & "Error: " & ex.Message
                    End Try
                    intTotalCount = intTotalCount + 1
                    Me.lblRunStatus.Text = "Calc XY polygon centroids: " & intTotalCount.ToString
                    Me.Refresh()
                    pToFeat = pToFeatureCursor.NextFeature
                Loop

                'CREATE A POINT FEATURE CLASS FROM THE CENTROID XY VALUES FOR USE IN SUBAREA MATCHING
                CreateParcelCentroidFeatClass()
                CreateParcelCentroidPointFeatures(pToFeatClass)

                GoTo CleanUp

            Catch ex As Exception
                m_strProcessing = m_strProcessing & vbNewLine & "Error in function CreateParcelCentroidLayer.  Failed on feature count:  " & intTotalCount.ToString & vbNewLine & "Error: " & ex.Message
                GoTo CleanUp
            End Try
        End If

CleanUp:
        Try
            If Not pToFeatClass Is Nothing Then
                m_lyrEnvisionParcel = New FeatureLayer
                m_lyrEnvisionParcel.FeatureClass = pToFeatClass
                m_lyrEnvisionParcel.Name = "<SETUP> " & "PARCEL LAYER"
            End If
        Catch ex As Exception

        End Try
        m_strProcessing = m_strProcessing & vbNewLine & "End function CreateParcelCentroidLayer: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
        Me.lblRunStatus.Text = ""
        pWksFactory = Nothing
        pFeatWks = Nothing
        pToFeatClass = Nothing
        pTable = Nothing
        intToPntXFld = Nothing
        intToPntYFld = Nothing
        pToFeatureCursor = Nothing
        intTotalCount = Nothing
        pFeatSelection = Nothing
        pToFeat = Nothing
        pArea = Nothing
        GC.Collect()
        GC.WaitForPendingFinalizers()
    End Sub

    Private Function CreateParcelCentroidFeatClass() As Boolean
        CreateParcelCentroidFeatClass = True
        m_strProcessing = m_strProcessing & vbNewLine & "Starting function CreateParcelCentroidFeatClass: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
        'CREATE THE PONT FEATURE CLASS, WHICH CONTAINS THE CENTROID POINT LOCATIONS OF ALL THE OUTPUT PARCEL FEATURES
        Dim pCreateFeatClass As ESRI.ArcGIS.DataManagementTools.CreateFeatureclass
        Dim pWksFactory As IWorkspaceFactory = Nothing
        Dim pFeatWks As IFeatureWorkspace
        Dim pFeatClass As IFeatureClass
        Dim pTable As ITable

        Try
            If pGPSetup Is Nothing Then
                CreateSetupGP()
            End If

            pCreateFeatClass = New ESRI.ArcGIS.DataManagementTools.CreateFeatureclass
            pCreateFeatClass.out_path = m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb"
            pCreateFeatClass.out_name = "PROJECT_PARCEL_CENTROIDS"
            pCreateFeatClass.geometry_type = "POINT"
            pCreateFeatClass.spatial_reference = m_pETSpatRefProject
            RunTool(pGPSetup, pCreateFeatClass)
            pCreateFeatClass = Nothing

            'DEFINE THE GRID LAYER
            pWksFactory = New FileGDBWorkspaceFactory
            pFeatWks = pWksFactory.OpenFromFile(m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb", 0)
            pFeatClass = pFeatWks.OpenFeatureClass("PROJECT_PARCEL_CENTROIDS")
            pTable = CType(pFeatClass, ITable)
            'ADD FIELD FOR ENVISION IDS
            AddEnvisionField(pTable, "ET_TEMP_OBJ_ID2", "INTEGER", 16, 0)
            If pFeatClass.ShapeType = esriGeometryType.esriGeometryPoint Then
                m_lyrParcelCentroids = New FeatureLayer
                m_lyrParcelCentroids.FeatureClass = pFeatClass
                m_lyrParcelCentroids.Name = m_lyrGrid.Name & "<SETUP> " & "PARCEL CENTROID LAYER"
            End If
        Catch ex As Exception
            m_strProcessing = m_strProcessing & vbNewLine & "Error in creating centroid feature class (PROJECT_PARCEL_CENTROIDS): " & ex.Message
            CreateParcelCentroidFeatClass = False
            GoTo CleanUp
        End Try
CleanUp:
        m_strProcessing = m_strProcessing & vbNewLine & "End function CreateParcelCentroidFeatClass: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
        pCreateFeatClass = Nothing
        pWksFactory = Nothing
        pFeatWks = Nothing
        pFeatClass = Nothing
        GC.Collect()
        GC.WaitForPendingFinalizers()
    End Function

    Private Function CreateParcelCentroidPointFeatures(ByVal pParcelFeatClass As IFeatureClass) As Boolean
        'GENERATE THE CENTROIDS FOR EACH PARCEL BASE UPON THE FEATURE'S AREA CENTROID
        m_strProcessing = m_strProcessing & vbNewLine & "Starting function CreateParcelCentroidPointFeatures: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
        CreateParcelCentroidPointFeatures = True

        Dim pCentroidFeatClass As IFeatureClass
        Dim intCount As Integer
        Dim pFeat As IFeature
        Dim pToFeatureCursor As IFeatureCursor
        Dim pArea As IArea
        Dim pInsertFeatureBuffer As IFeatureBuffer
        Dim pInsertFeatureCursor As IFeatureCursor = Nothing
        Dim pPointTemp As ESRI.ArcGIS.Geometry.Point
        Dim intIdToField As Integer

        Try
            'EXIT IF THERE IS NOT CENTROID FEATURECLASS
            If m_lyrParcelCentroids Is Nothing Then
                CreateParcelCentroidPointFeatures = False
                GoTo CleanUp
            Else
                pCentroidFeatClass = m_lyrParcelCentroids.FeatureClass
            End If

            'RETRIEVE THE FROM AND TO ENVISION ID 2 FIELD NUMBERS
            intIdToField = pCentroidFeatClass.FindField("ET_TEMP_OBJ_ID2")
            If intIdToField <= -1 Then
                m_strProcessing = m_strProcessing & vbNewLine & "The id field, ET_TEMP_OBJ_ID2, was missing from the parcel or centroid feature classes. "
                GoTo CleanUp
            End If

            'CYCLE THROUGH THE PARCEL FEATURE CLASS TO BUILD A CENTROID FOR EACH PARCEL
            pToFeatureCursor = pParcelFeatClass.Search(Nothing, False)
            pFeat = pToFeatureCursor.NextFeature
            Do While Not pFeat Is Nothing
                Try
                    pArea = pFeat.Shape
                    If Not pArea Is Nothing Then
                        pPointTemp = New ESRI.ArcGIS.Geometry.Point
                        pPointTemp.X = pArea.Centroid.X
                        pPointTemp.Y = pArea.Centroid.Y
                        'FEATURE BUFFER SETUP
                        pInsertFeatureCursor = pCentroidFeatClass.Insert(True)
                        pInsertFeatureBuffer = pCentroidFeatClass.CreateFeatureBuffer
                        pInsertFeatureBuffer.Shape = pPointTemp
                        pInsertFeatureBuffer.Value(intIdToField) = pFeat.Value(0)
                        pInsertFeatureCursor.InsertFeature(pInsertFeatureBuffer)
                        pPointTemp = Nothing
                        pArea = Nothing
                    End If
                Catch ex As Exception
                    m_strProcessing = m_strProcessing & vbNewLine & "Error writing centroid point.  Failed on feature count:  " & intCount.ToString & vbNewLine & "Error: " & ex.Message
                End Try
                intCount = intCount + 1
                Me.lblRunStatus.Text = "Creating Centroid: " & intCount.ToString
                Me.Refresh()
                pFeat = pToFeatureCursor.NextFeature
            Loop

            If Not pInsertFeatureCursor Is Nothing Then
                pInsertFeatureCursor.Flush()
            End If

        Catch ex As Exception
            m_strProcessing = m_strProcessing & vbNewLine & "Error in creating centroid point features.  " & vbNewLine & ex.Message
            CreateParcelCentroidPointFeatures = False
            GoTo CleanUp
        End Try

CleanUp:
        m_strProcessing = m_strProcessing & vbNewLine & "End function CreateParcelCentroidPointFeatures: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
        pCentroidFeatClass = Nothing
        intCount = Nothing
        pFeat = Nothing
        pToFeatureCursor = Nothing
        pArea = Nothing
        pInsertFeatureBuffer = Nothing
        pInsertFeatureCursor = Nothing
        pPointTemp = Nothing
        intIdToField = Nothing
        GC.Collect()
        GC.WaitForPendingFinalizers()
    End Function

    Public Sub ParcelLastStep()
        Dim pRename As ESRI.ArcGIS.DataManagementTools.Rename
        Dim pUnion As ESRI.ArcGIS.AnalysisTools.Union
        Dim pSelect As ESRI.ArcGIS.AnalysisTools.Select
        Dim strInFeatures As String
        Dim pDataset As IDataset = Nothing
        Dim sb As System.Text.StringBuilder = New System.Text.StringBuilder
        Dim pWksFactory As IWorkspaceFactory = Nothing
        Dim pFeatWks As IFeatureWorkspace
        Dim pFeatClass As IFeatureClass
        Dim pTable As ITable
        Dim arrKeepFields As ArrayList = New ArrayList
        Dim intField As Integer
        Dim strFieldName As String
        Dim pField As IField

        m_strProcessing = m_strProcessing & vbNewLine & "Start function ParcelLastStep: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
        If Not Me.btnConstraintedOption.Text = "Calculate Constrainted Acres Only" And Not m_lyrConstrainedParcels Is Nothing Then
            'REMOVE THE CONSTRAINTS FROM FROM FINAL PARCEL LAYER IF CHECKED
            Try
                pDataset = CType(m_lyrConstrainedParcels, IDataset)
                strInFeatures = pDataset.Workspace.PathName & "\" & pDataset.BrowseName
                m_strProcessing = m_strProcessing & vbNewLine & "Constraints Layer: " & strInFeatures
                sb.AppendFormat("{0};", strInFeatures)

                pSelect = New ESRI.ArcGIS.AnalysisTools.Select
                pSelect.in_features = strInFeatures 'm_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb" & "\ENVISION_PARCELS_WITH_CONSTRAINTS_MULTIPART"
                If Me.btnConstraintedOption.Text = "Retain Constrainted Area Mash-up" Then
                    pSelect.where_clause = """" & "ET_TEMP_OBJ_ID2" & """" & " >= 1"
                Else
                    pSelect.where_clause = """" & "ET_CONSTRAINED" & """" & " = 0 AND " & """" & "ET_TEMP_OBJ_ID2" & """" & " >= 1"
                End If
                pSelect.out_feature_class = m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb" & "\ENVISION_PARCELS_FINAL"
                RunTool(pGPSetup, pSelect)
                pSelect = Nothing
            Catch ex As Exception
                m_strProcessing = m_strProcessing & vbNewLine & "Failed to retrieve constraints feature layer information."
            End Try
        Else
            'IF NO OTHER WORK, THEN RENAME PARCEL LAYER TO FINAL OUTPUT LAYER NAME
            'REMOVE THE CONSTRAINTS FROM FROM FINAL PARCEL LAYER IF CHECKED
            Try
                pDataset = CType(m_lyrMainProcessingLayer, IDataset)
                strInFeatures = pDataset.Workspace.PathName & "\" & pDataset.BrowseName
                pRename = New ESRI.ArcGIS.DataManagementTools.Rename
                pRename.in_data = strInFeatures
                pRename.out_data = m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb" & "\ENVISION_PARCELS_FINAL"
                pGPSetup.AddOutputsToMap = True
                RunTool(pGPSetup, pRename)
                pGPSetup.AddOutputsToMap = False
            Catch ex As Exception
                m_strProcessing = m_strProcessing & vbNewLine & "Failed to rename main processing layer to final parcel layer name."
            End Try

        End If

        'OPEN THE PROJECT WORKSPACE
        Try
            pWksFactory = New FileGDBWorkspaceFactory
            pFeatWks = pWksFactory.OpenFromFile(m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb", 0)
            pFeatClass = pFeatWks.OpenFeatureClass("ENVISION_PARCELS_FINAL")
            pTable = CType(pFeatClass, ITable)
            m_lyrEnvisionParcel = Nothing
            m_lyrEnvisionParcel = New FeatureLayer
            m_lyrEnvisionParcel.FeatureClass = pFeatClass
            m_lyrEnvisionParcel.Name = "<SETUP> " & "FINAL PARCEL LAYER"
            pFeatClass = pFeatWks.OpenFeatureClass("ENVISION_PARCELS_FINAL")
        Catch ex As Exception
            m_strProcessing = m_strProcessing & vbNewLine & "Error in opening project workspace to retreive the final parcel layer feature class:" & vbNewLine & ex.Message
            GoTo CleanUp
        End Try

        'CALC ET_ACREs AGAIN ON THE FINAL LAYER
        Me.lblRunStatus.Text = "Calculating Acres for input parcel feature(s)"
        If Not pTable Is Nothing Then
            CalAcres(pTable, "ET_ACRES")
        End If


        'CLEAN UP THE FIELDS, MAKE SURE ONLY THE FIELDS FROM THE ORIGINAL INPUT LAYER OR REQUIRED FIELDS REMAIN
        If Not m_lyrOriginalParcels Is Nothing And (Me.rdbSourceHybrid.Checked Or Me.rdbSourceParcels.Checked) Then
            For intField = 0 To m_lyrOriginalParcels.FeatureClass.Fields.FieldCount - 1
                Try
                    strFieldName = m_lyrOriginalParcels.FeatureClass.Fields.Field(intField).Name
                    arrKeepFields.Add(strFieldName)
                Catch ex As Exception
                End Try
            Next
        End If

        'ADD REQUIRED FIELDS(S)
        arrKeepFields.Add("ET_TEMP_OBJ_ID1")
        arrKeepFields.Add("ET_TEMP_OBJ_ID2")
        arrKeepFields.Add("ET_ACRES")
        arrKeepFields.Add("EX_LU")
        arrKeepFields.Add("VAC_ACRE")
        arrKeepFields.Add("DEVD_ACRE")
        arrKeepFields.Add("CONSTRAINED_ACRE")
        arrKeepFields.Add("ET_CENTROID_X")
        arrKeepFields.Add("ET_CENTROID_Y")
        arrKeepFields.Add("APPLY_CONSTRAINTS")

        If arrKeepFields.Count > 0 And Not pFeatClass Is Nothing Then
            For intField = 1 To pFeatClass.Fields.FieldCount - 1
                Try
                    If arrKeepFields.IndexOf(pFeatClass.Fields.Field(intField).Name) <= -1 Then
                        pField = pFeatClass.Fields.Field(intField)
                        pFeatClass.DeleteField(pField)
                        pField = Nothing
                    End If
                Catch ex As Exception
                End Try
            Next
        End If


        GoTo CleanUp

CleanUp:
        m_strProcessing = m_strProcessing & vbNewLine & "End function ParcelLastStep: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
        pRename = Nothing
        pUnion = Nothing
        pSelect = Nothing
        strInFeatures = Nothing
        pDataset = Nothing
        sb = Nothing
        pWksFactory = Nothing
        pFeatWks = Nothing
        pFeatClass = Nothing
        pTable = Nothing
        arrKeepFields = Nothing
        intField = Nothing
        strFieldName = Nothing
        pField = Nothing
        GC.Collect()
        GC.WaitForPendingFinalizers()
    End Sub

    Private Sub FieldMap()
        'PROPORTION THE VALUES OF THE SELECTED FIELD(S) BASED ON AREA VALUES OF THE ORIGINAL AND FINAL FEATURES
        m_strProcessing = m_strProcessing & vbNewLine & "Starting sub FieldMap: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
        Dim pWksFactory As IWorkspaceFactory = Nothing
        Dim pFeatWks As IFeatureWorkspace
        Dim pFeatClass As IFeatureClass
        Dim pTable As ITable = Nothing
        Dim strInFeatures As String
        Dim pDataset As IDataset = Nothing
        Dim sb As System.Text.StringBuilder = New System.Text.StringBuilder
        Dim intRow As Integer
        Dim strInFieldName As String
        Dim strOutFieldname As String
        Dim pField As IField


        Try
            If pGPSetup Is Nothing Then
                CreateSetupGP()
            End If

            'CHECK THAT THE FINAL PARCEL LAYER WAS CREATED
                Try
                    'OPEN THE PROJECT WORKSPACE
                    Try
                        pWksFactory = New FileGDBWorkspaceFactory
                        pFeatWks = pWksFactory.OpenFromFile(m_frmEnvisionProjectSetup.tbxWorkspace.Text & "\" & m_frmEnvisionProjectSetup.tbxProjectName.Text & ".gdb", 0)
                    Catch ex As Exception
                    m_strProcessing = m_strProcessing & vbNewLine & "Error in FinalMap in attempt to open project workspace:" & vbNewLine & ex.Message
                        GoTo CleanUp
                    End Try
                pFeatClass = pFeatWks.OpenFeatureClass("ENVISION_PARCELS_FINAL")
                pTable = CType(pTable, IFeatureClass)
                Catch ex As Exception
                m_strProcessing = m_strProcessing & vbNewLine & "Error in FinalMap Opening Final Parcel feature class:" & vbNewLine & ex.Message
                    GoTo CleanUp
                End Try

            'CHECK FOR THE REQUIRED INPUT AND OUTPUT FIELD(S)
            If Me.dgvFieldMatch.RowCount > 0 Then
                For intRow = 0 To Me.dgvFieldMatch.RowCount - 1
                    If Me.dgvFieldMatch.Rows(Me.dgvFieldMatch.RowCount - 1).Cells(1).Value = 1 Then
                        strInFieldName = Me.dgvFieldMatch.Rows(Me.dgvFieldMatch.RowCount - 1).Cells(0).Value
                        strOutFieldname = Me.dgvFieldMatch.Rows(Me.dgvFieldMatch.RowCount - 1).Cells(1).Value
                        If strInFieldName.Length > 0 Then
                            pField = pFeatClass.Fields.Field(pFeatClass.FindField(strInFieldName))
                            If pFeatClass.FindField(strInFieldName) >= 0 Then
                                If strOutFieldname.Length > 0 Then
                                    If pFeatClass.FindField(strOutFieldname) <= -1 Then
                                        AddEnvisionField(pTable, strOutFieldname, pField.Type, pField.Length, pField.Scale)
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next
            End If
        Catch ex As Exception
            m_strProcessing = m_strProcessing & vbNewLine & "Error in sub FieldMap:" & vbNewLine & ex.Message
            GoTo CleanUp
        End Try

CleanUp:
        m_strProcessing = m_strProcessing & vbNewLine & "End sub FieldMap: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
        pWksFactory = Nothing
        pFeatWks = Nothing
        pFeatClass = Nothing
        strInFeatures = Nothing
        pDataset = Nothing
        sb = Nothing
        'pFeatLayer = Nothing
        ''intCount = Nothing
        'strUnionName = Nothing
        pTable = Nothing
        'strFieldName = Nothing
        GC.Collect()
        GC.WaitForPendingFinalizers()

    End Sub

    Private Sub btnCheckMapFields_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCheckMapFields.Click
        'CHECK ALL FIELDS 
        FieldCheckStatus(True)
    End Sub

    Private Sub btnUncheckMapFields_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUncheckMapFields.Click
        'UNCHECK ALL FIELDS
        FieldCheckStatus(False)
    End Sub

    Public Sub FieldCheckStatus(ByVal blnCheckStatus As Boolean)
        'SET CHECK STATUS TO ALL LAYERS AS DEFINED BY INUT VARIABLE
        Dim intRow As Integer
        Me.dgvFieldMatch.ClearSelection()
        For intRow = 0 To Me.dgvFieldMatch.Rows.Count - 1
            If blnCheckStatus Then
                Me.dgvFieldMatch.Rows(intRow).Cells(0).Value = "TRUE"
            Else
                Me.dgvFieldMatch.Rows(intRow).Cells(0).Value = "FALSE"
            End If
        Next
    End Sub

    Public Function ProjectDirCheck() As Boolean
        'ENSURE SELECTED ENVISION WORKSPACE EXISTS
        If Directory.Exists(m_frmEnvisionProjectSetup.tbxWorkspace.Text) Then
            ProjectDirCheck = True
        Else
            ProjectDirCheck = False
        End If
    End Function




    Public Function IsEnvisionExcel(ByVal strFile As String) As Boolean
        'INSURE THE SELECTED ENVISION EXCEL FILE CONTAINS THE rEQUIRED TABS
        IsEnvisionExcel = True
        Dim xlsApp As Microsoft.Office.Interop.Excel.Application = Nothing
        Dim xlsWB As Microsoft.Office.Interop.Excel.Workbook = Nothing
        Dim arrTabList As ArrayList = New ArrayList
        Dim strTab As String
        Try
            m_frmEnvisionProjectSetup.lblStatus1.Visible = True
            m_frmEnvisionProjectSetup.barStatus1.Visible = True
            m_frmEnvisionProjectSetup.barStatus1.Value = 10
            m_frmEnvisionProjectSetup.Refresh()

      xlsApp = New Microsoft.Office.Interop.Excel.Application
      xlsApp.DisplayAlerts = False
            xlsApp.WindowState = Interop.Excel.XlWindowState.xlMinimized
            xlsApp.Visible = True
            xlsWB = CType(xlsApp.Workbooks.Open(strFile), Microsoft.Office.Interop.Excel.Workbook)

        Catch ex As Exception
            IsEnvisionExcel = False
            GoTo CleanUp
        End Try

        'BUILD LIST OF REQUIRED TABS
        arrTabList.Add("Building Inputs")
        arrTabList.Add("Dev Type Attributes")
        arrTabList.Add("SCENARIO1")
        arrTabList.Add("SCENARIO2")
        arrTabList.Add("SCENARIO3")
        arrTabList.Add("SCENARIO4")
        arrTabList.Add("SCENARIO5")

        'CYCLE THROUGH TAB LIST 
        For Each strTab In arrTabList
            Try
                m_frmEnvisionProjectSetup.lblStatus1.Text = strTab & " Tab Check"
                m_frmEnvisionProjectSetup.barStatus1.Value = m_frmEnvisionProjectSetup.barStatus1.Value + 10
                m_frmEnvisionProjectSetup.Refresh()

                If Not TypeOf xlsWB.Sheets(strTab) Is Microsoft.Office.Interop.Excel.Worksheet Then
                    MessageBox.Show("The required tab, " & strTab & ", could not be found in the selected Excel file.  Please select another Envision Excel file.", "Invalid Envision Excel File", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
                    IsEnvisionExcel = False
                    Exit For
                End If
            Catch ex As Exception
                MessageBox.Show("The required tab, " & strTab & ", could not be found in the selected Excel file.  Please select another Envision Excel file.", "Invalid Envision Excel File", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
                IsEnvisionExcel = False
                Exit For
            End Try
        Next

        'CLOSE THE SELECTED EXCEL FILE
        Try
            If Not xlsApp Is Nothing Then
                xlsApp.Quit()
                Marshal.FinalReleaseComObject(xlsApp)
            End If
            xlsApp = Nothing

            If Not xlsWB Is Nothing Then
                xlsWB.Close()
                Marshal.FinalReleaseComObject(xlsWB)
            End If
            xlsWB = Nothing
            GoTo CleanUp
        Catch ex As Exception
            GoTo CleanUp
        End Try

CleanUp:
        m_frmEnvisionProjectSetup.barStatus1.Value = 0
        m_frmEnvisionProjectSetup.Refresh()
        xlsApp = Nothing
        xlsWB = Nothing
        arrTabList = Nothing
        strTab = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Function

    Public Function ParcelLyrCheck() As Boolean
        'CHANGE TEXT COLOR TO INDICATE IF THE TEXT MATCHES A LAYER IN THE CURRENT LIST
        If m_frmEnvisionProjectSetup.cmbParcelLayers.Items.Contains(m_frmEnvisionProjectSetup.cmbParcelLayers.Text) Then
            m_frmEnvisionProjectSetup.cmbParcelLayers.ForeColor = Color.Black
            ParcelLyrCheck = True
        Else
            m_frmEnvisionProjectSetup.cmbParcelLayers.ForeColor = Color.Red
            ParcelLyrCheck = False
        End If

        'LOAD THE NUMERIC FIELDS FROM THE SELECTED PARCEL LAYER INTO THE ACRES FIELD CONTROL

    End Function

    Public Function GridCellParametersReview() As Boolean
        'REVIEW THE GRID CELL SIZE VALUE
        GridCellParametersReview = True
        m_frmEnvisionProjectSetup.tbxGridCellSize.ForeColor = Color.Black
        If m_frmEnvisionProjectSetup.tbxGridCellSize.Text.Length <= 0 Then
            GridCellParametersReview = False
        Else
            If Not IsNumeric(m_frmEnvisionProjectSetup.tbxGridCellSize.Text) Then
                GridCellParametersReview = False
                m_frmEnvisionProjectSetup.tbxGridCellSize.ForeColor = Color.Red
            End If
        End If
    End Function

    Public Function ParcelSizeParametersReview() As Boolean
        ParcelSizeParametersReview = True
        m_frmEnvisionProjectSetup.tbxMaxParcelSize.ForeColor = Color.Black
        If m_frmEnvisionProjectSetup.tbxMaxParcelSize.Text.Length <= 0 Then
            ParcelSizeParametersReview = False
        Else
            If Not IsNumeric(m_frmEnvisionProjectSetup.tbxMaxParcelSize.Text) Then
                ParcelSizeParametersReview = False
                m_frmEnvisionProjectSetup.tbxMaxParcelSize.ForeColor = Color.Red
            End If
        End If

        'm_frmEnvisionProjectSetup.cmbAcresField.ForeColor = Color.Black
        'If m_frmEnvisionProjectSetup.cmbAcresField.Text.Length <= 0 Then
        '    ParcelSizeParametersReview = False
        'Else
        '    If Not m_frmEnvisionProjectSetup.cmbAcresField.Items.Contains(m_frmEnvisionProjectSetup.cmbAcresField.Text) Then
        '        ParcelSizeParametersReview = False
        '        m_frmEnvisionProjectSetup.cmbAcresField.ForeColor = Color.Red
        '    End If
        'End If
    End Function

    Public Sub LoadEnvisionExcelFile()
        'OPEN SELECTED ENVISION EXCEL FILE
        Dim xlPaintApp As Microsoft.Office.Interop.Excel.Application = Nothing
        Dim xlPaintWB1 As Microsoft.Office.Interop.Excel.Workbook = Nothing
        Dim shtExistingDevArea As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        Dim shtExistingLandUse As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        Dim strFldCellValue As String = ""
        Dim intRow As Integer
        Dim intFieldRow As Integer = -1
        Dim intStartRow As Integer = -1
        Dim intCol As Integer = -1
        Dim intCount As Integer = 0
        Dim intTotalCount As Integer = 0
        Dim intExLUCol As Integer = -1
        Dim dgvcFieldMap As DataGridViewComboBoxColumn
        Dim dgvcLandUse As DataGridViewComboBoxColumn

        Try
            'OPEN THE SELECTE ENVISION EXECEL FILE
            m_appEnvision.StatusBar.Message(0) = "Opening selected excel file"
            If m_frmEnvisionProjectSetup.tbxExcel.Text.Length <= 0 Then
                GoTo CleanUp
            End If
      xlPaintApp = New Microsoft.Office.Interop.Excel.Application
      xlPaintApp.DisplayAlerts = False
            xlPaintApp.WindowState = Interop.Excel.XlWindowState.xlMinimized
            xlPaintApp.Visible = True
            xlPaintWB1 = CType(xlPaintApp.Workbooks.Open(m_frmEnvisionProjectSetup.tbxExcel.Text), Microsoft.Office.Interop.Excel.Workbook)

            '------------------------------------------------------------------------------------------
            'RETRIEVE THE EXISTING DEVELOPED AREA TAB TO RETRIEVE VALUES FROM
            Try
                m_appEnvision.StatusBar.Message(0) = "Retrieving 'Existing Developed Area' tab"
                shtExistingDevArea = xlPaintWB1.Sheets("Existing Developed Area")

                'FIND THE STARTING POINT
                For intRow = 1 To 10
                    strFldCellValue = CStr(CType(shtExistingDevArea.Cells(intRow, 1), Microsoft.Office.Interop.Excel.Range).Value)
                    If Not strFldCellValue Is Nothing Then
                        If UCase(strFldCellValue) = "EXISTING" Then
                            intFieldRow = intRow
                            intStartRow = intRow + 1
                            Exit For
                        End If
                    End If
                Next

                'RETRIEVE EXISITNG LAND USE ABBREVIATION COLUMN
                m_appEnvision.StatusBar.Message(0) = "'Existing Developed Area' tab - building Existing Land Use list"
                intCount = 0
                For intCol = 1 To 25
                    strFldCellValue = CStr(CType(shtExistingDevArea.Cells(intFieldRow, intCol), Microsoft.Office.Interop.Excel.Range).Value)
                    If Not strFldCellValue Is Nothing Then
                        'If pFeatClass.FindField(strFldCellValue) >= 0 Then
                        If UCase(strFldCellValue) = "EX_LU" Then
                            intExLUCol = intCol
                            Exit For
                        End If
                    End If
                Next
                If intExLUCol <= -1 Then
                    GoTo CloseExcel
                End If

                'CLEAR EXISTING LAND USE ABBREVIATIONS
                dgvcLandUse = m_frmEnvisionProjectSetup.dgvLandUseAttributes.Columns(1)
                dgvcLandUse.Items.Clear()

                'BUILD LIST OF EXISTING LAND USE ABBREVIATIONS
                For intRow = intStartRow To (intStartRow + 20)
                    strFldCellValue = CStr(CType(shtExistingDevArea.Cells(intRow, intExLUCol), Microsoft.Office.Interop.Excel.Range).Value)
                    If Not strFldCellValue Is Nothing Then
                        dgvcLandUse.Items.Add(strFldCellValue)
                    End If
                Next

                'CLEAR ANY EXISTING SELECTION IF NOT FOUND
                For intRow = 0 To m_frmEnvisionProjectSetup.dgvLandUseAttributes.Rows.Count - 1
                    strFldCellValue = CStr(m_frmEnvisionProjectSetup.dgvLandUseAttributes.Rows(intRow).Cells(1).Value)
                    If Not strFldCellValue Is Nothing Then
                        If dgvcLandUse.Items.IndexOf(strFldCellValue) <= -1 Then
                            m_frmEnvisionProjectSetup.dgvLandUseAttributes.Rows(intRow).Cells(1).Value = ""
                        End If
                    End If
                Next

                '------------------------------------------------------------------------------------------
                'RETRIEVE THE EXISTING LAND USE TAB TO RETRIEVE VALUES FROM
                Try
                    m_appEnvision.StatusBar.Message(0) = "Retrieving 'Existing Land Use' tab"
                    shtExistingLandUse = xlPaintWB1.Sheets("Existing Land Use")

                    'FIND THE STARTING POINT
                    For intRow = 1 To 10
                        strFldCellValue = CStr(CType(shtExistingLandUse.Cells(intRow, 1), Microsoft.Office.Interop.Excel.Range).Value)
                        If Not strFldCellValue Is Nothing Then
                            If UCase(strFldCellValue) = "EXISTING" Then
                                intStartRow = intRow
                                Exit For
                            End If
                        End If
                    Next

                    'CLEAR EXISTING LAND USE ABBREVIATIONS
                    dgvcFieldMap = m_frmEnvisionProjectSetup.dgvFieldMatch.Columns(2)
                    dgvcFieldMap.Items.Clear()

                    'RETRIEVE EXISITNG LAND USE ABBREVIATION COLUMN
                    m_appEnvision.StatusBar.Message(0) = "'Existing Land Use' tab - building existing conditions field list"
                    intCount = 0
                    For intCol = 2 To 50
                        strFldCellValue = CStr(CType(shtExistingLandUse.Cells(intStartRow, intCol), Microsoft.Office.Interop.Excel.Range).Value)
                        If Not strFldCellValue Is Nothing Then
                            dgvcFieldMap.Items.Add(strFldCellValue)
                        End If
                    Next

                    'CLEAR ANY EXISTING SELECTION IF NOT FOUND
                    For intRow = 0 To m_frmEnvisionProjectSetup.dgvFieldMatch.Rows.Count - 1
                        strFldCellValue = CStr(m_frmEnvisionProjectSetup.dgvFieldMatch.Rows(intRow).Cells(2).Value)
                        If Not strFldCellValue Is Nothing Then
                            If dgvcFieldMap.Items.IndexOf(strFldCellValue) <= -1 Then
                                m_frmEnvisionProjectSetup.dgvFieldMatch.Rows(intRow).Cells(2).Value = ""
                            End If
                        End If
                    Next
                Catch ex As Exception

                End Try
                '---------------------------------------------------------------------------------------

                'CLOSE EXCEL
                GoTo CloseExcel

            Catch ex As Exception
                GoTo CloseExcel
            End Try


        Catch ex As Exception
            GoTo CloseExcel
        End Try

        'CLOSE THE ENVISION EXCEL APPLICATION
CloseExcel:
        Try
            shtExistingDevArea = Nothing
            shtExistingLandUse = Nothing
            If Not xlPaintWB1 Is Nothing Then
                xlPaintWB1.Close(True)
                Marshal.FinalReleaseComObject(xlPaintWB1)
            End If
            If Not xlPaintApp Is Nothing Then
                xlPaintApp.Quit()
                Marshal.FinalReleaseComObject(xlPaintApp)
            End If
            xlPaintApp = Nothing
            If Not xlPaintWB1 Is Nothing Then
                xlPaintWB1.Close(False)
                Marshal.FinalReleaseComObject(xlPaintWB1)
            End If
            xlPaintWB1 = Nothing
            GoTo CleanUp
        Catch ex As Exception
            GoTo CleanUp
        End Try

CleanUp:
        m_appEnvision.StatusBar.Message(0) = ""
        xlPaintApp = Nothing
        xlPaintWB1 = Nothing
        shtExistingDevArea = Nothing
        shtExistingLandUse = Nothing
        strFldCellValue = Nothing
        intRow = Nothing
        intFieldRow = Nothing
        intStartRow = Nothing
        intCol = Nothing
        intCount = Nothing
        intTotalCount = Nothing
        intExLUCol = Nothing
        dgvcFieldMap = Nothing
        dgvcLandUse = Nothing
        GC.Collect()
        GC.WaitForPendingFinalizers()
    End Sub

    Private Sub cmbLandUseField_Click(sender As Object, e As EventArgs) Handles cmbLandUseField.Click

    End Sub
End Class