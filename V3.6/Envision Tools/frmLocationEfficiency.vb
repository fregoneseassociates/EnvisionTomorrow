Option Explicit On
Option Strict On

Imports System
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Runtime.InteropServices
Imports System.Collections
Imports System.IO
Imports System.Drawing.Printing
Imports System.Data
Imports System.Math

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
Imports ESRI.ArcGIS.CartoUI
Imports ESRI.ArcGIS.Catalog
Imports ESRI.ArcGIS.CatalogUI
Imports ESRI.ArcGIS.DataSourcesGDB
Imports ESRI.ArcGIS.DataSourcesFile
Imports ESRI.ArcGIS.Display
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.GeoDatabaseUI
Imports ESRI.ArcGIS.SystemUI
Imports ESRI.ArcGIS.SpatialAnalyst
Imports ESRI.ArcGIS.SpatialAnalystUI
Imports ESRI.ArcGIS.Geoprocessing
Imports ESRI.ArcGIS.Editor

Public Class frmLocationEfficiency

  Public strMessage As String = ""
  Public arrLayers As ArrayList = New ArrayList
  Public arrZoneLayers As ArrayList = New ArrayList
  Public strProcessing As String = ""
  Public m_cmbLayers As ComboBox = Nothing
  Public m_cmbValueFlds As ComboBox = Nothing
  Public m_tbxInfluence As TextBox = Nothing
  Public m_btnAddFactor As System.Windows.Forms.Button = Nothing
  Public m_dgvWeights As DataGridView = Nothing
  Public m_tbxTotalInfluence As TextBox = Nothing
  Public pGP_SLTool As ESRI.ArcGIS.Geoprocessor.Geoprocessor
  Public m_pSpatRefSLTool As ISpatialReference = Nothing
  Public m_pExtentSLToolEnv As IEnvelope = Nothing
  Public m_lyrSLToolAOI As IFeatureLayer
  Public m_extentUpperX As Double
  Public m_extentLowerX As Double
  Public m_extentUpperY As Double
  Public m_extentLowerY As Double
  Private m_toolTip As ToolTip = New ToolTip()
  Private m_SpatialAnalystExtEnabled As Boolean

  Private Sub frmStationLocation_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
    Try
      If m_SpatialAnalystExtEnabled Then DisableSpatialAnalystExtension()
      m_frmLocationEfficiency = Nothing
      strMessage = Nothing
      arrLayers = Nothing
      strProcessing = Nothing
      m_cmbLayers = Nothing
      m_cmbValueFlds = Nothing
      m_tbxInfluence = Nothing
      m_btnAddFactor = Nothing
      m_dgvWeights = Nothing
      m_tbxTotalInfluence = Nothing
      pGP_SLTool = Nothing
      m_pSpatRefSLTool = Nothing
      m_pExtentSLToolEnv = Nothing
      m_lyrSLToolAOI = Nothing
      GC.Collect()
      GC.WaitForPendingFinalizers()
      GoTo CleanUp
    Catch ex As Exception
      'MessageBox.Show(me,ex.Message, "Closing Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
      GoTo CleanUp
    End Try

CleanUp:
    m_frmLocationEfficiency = Nothing
    GC.Collect()
    GC.WaitForPendingFinalizers()
  End Sub

  Private Sub frmStationLocation_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Dim mxApplication As IMxApplication = Nothing
    Dim pMxDocument As IMxDocument = Nothing
    Dim mapCurrent As Map
    Dim uid As ESRI.ArcGIS.esriSystem.IUID = New ESRI.ArcGIS.esriSystem.UIDClass
    Dim enumLayer As ESRI.ArcGIS.Carto.IEnumLayer
    Dim pActiveView As IActiveView = Nothing
    Dim pLyr As ILayer
    Dim pFeatLyr As IFeatureLayer

    ' Check for and enable the spatial analyst extension
    If Not IsSpatialAnalystExtensionEnabled() Then
      If Not EnableSpatialAnalystExtension() Then
        MessageBox.Show(Me, "The Spatial Analyst extension is required for this tool and is currently not available.", "Spatial Analyst Extension", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1)
        Exit Sub
      Else
        m_SpatialAnalystExtEnabled = True
      End If
    End If

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

    'SELECT THE EXTENT OPTION FOR THE CURENT VIEW
    m_frmLocationEfficiency.rdbExtentView.Checked = True

    'SET THE CURRENT TAB TO THE WEIGHTS TAB
    Me.InfoStationLocationSettings.SelectTab(1)
    Me.InfoStationLocationSettings.SelectTab(0)

    'RETRIEVE THE FEATURE AND RASTER LAYERS TO POPULATE LAYER OPTIONS
    Try
      m_appEnvision.StatusBar.Message(0) = "Building list of layers"
      arrLayers.Clear()
      arrZoneLayers.Clear()
      Me.cmbLayers.Items.Clear()
      Me.cmbExtentLayers.Items.Clear()
      cboZoneLayer.Items.Clear()
      Me.cmbLayers.Text = ""
      Me.cmbExtentLayers.Text = ""
      cboZoneLayer.Text = ""
      uid.Value = "{6CA416B1-E160-11D2-9F4E-00C04F6BC78E}" '= IDataLayer
      enumLayer = mapCurrent.Layers((CType(uid, ESRI.ArcGIS.esriSystem.UID)), True) ' Explicit Cast
      enumLayer.Reset()
      pLyr = enumLayer.Next
      Do While pLyr IsNot Nothing
        Dim isValidLayer As Boolean = False
        Dim isRaster As Boolean = False
        Dim rasterHasTable As Boolean = False
        If TryCast(pLyr, IRasterLayer) IsNot Nothing Then
          isRaster = True
          If pLyr.Name.ToLower.Contains("location") AndAlso pLyr.Name.ToLower.Contains("efficiency") Then
            isValidLayer = False
          Else
            isValidLayer = True

            ' Verify there is an attribute table
            Dim raster As IRaster = DirectCast(pLyr, IRasterLayer).Raster
            Dim hasTable As Boolean
            DirectCast(raster, IRasterBandCollection).Item(0).HasTable(hasTable)
            If hasTable Then rasterHasTable = True
          End If
        ElseIf TryCast(pLyr, IFeatureLayer) IsNot Nothing Then
          pFeatLyr = DirectCast(pLyr, IFeatureLayer)
          If pFeatLyr.FeatureClass.ShapeType = esriGeometryType.esriGeometryPoint OrElse pFeatLyr.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline OrElse pFeatLyr.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
            isValidLayer = True
          End If
        End If
        If isValidLayer Then
          arrLayers.Add(pLyr)
          Me.cmbLayers.Items.Add(pLyr.Name)
          Me.cmbExtentLayers.Items.Add(pLyr.Name)
          If Not isRaster OrElse rasterHasTable Then
            cboZoneLayer.Items.Add(pLyr.Name)
            arrZoneLayers.Add(pLyr)
          End If
        End If
        pLyr = enumLayer.Next()
      Loop
      If arrLayers.Count = 0 Then
        MessageBox.Show(Me, "No feature or raster layers could be found in the current view document. Please add point, polyline or polygon feature layers, or raster layers to utilize the Location Efficiency tool.", "Feature or Raster Layers Not Found", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1)
        GoTo CloseForm
      End If
    Catch ex As System.Exception
      arrLayers = Nothing
      GoTo CloseForm
    End Try

    GoTo CleanUp

CloseForm:
    Me.Close()
    GoTo CleanUp
CleanUp:
    mxApplication = Nothing
    pMxDocument = Nothing
    mapCurrent = Nothing
    uid = Nothing
    enumLayer = Nothing
    pActiveView = Nothing
    pLyr = Nothing
    pFeatLyr = Nothing
    GC.Collect()
    GC.WaitForPendingFinalizers()

  End Sub

  Private Sub SetActiveGroupControls() Handles InfoStationLocationSettings.SelectedIndexChanged
    'DETERMINE WHICH TAB IS ACTIVE, IF A GROUP TAB, THEN SET ACTIVE CONTROLS
    Dim intCurrentTab As Integer = 0
    intCurrentTab = Me.InfoStationLocationSettings.SelectedIndex

    m_cmbLayers = Nothing
    m_cmbValueFlds = Nothing
    m_tbxInfluence = Nothing
    m_btnAddFactor = Nothing
    m_dgvWeights = Nothing
    m_tbxTotalInfluence = Nothing
    If intCurrentTab = 0 Then
      m_cmbLayers = CType(Me.cmbLayers, ComboBox)
      m_cmbValueFlds = CType(Me.cmbValuesFld, ComboBox)
      m_tbxInfluence = CType(Me.tbxInfluence, TextBox)
      m_btnAddFactor = CType(Me.btnAddFactor, System.Windows.Forms.Button)
      m_dgvWeights = CType(Me.dgvWeights, DataGridView)
      m_tbxTotalInfluence = CType(Me.tbxTotalInfluence, TextBox)
    End If

    intCurrentTab = Nothing
    GC.Collect()
    GC.WaitForPendingFinalizers()
  End Sub

  Private Sub cmbLayers_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbLayers.SelectedIndexChanged
    'LOAD A LIST OF STRING | INTEGER | DOUBLE FIELDS INTO FIELD COMBO BOX
    Dim pField As IField
    Dim intFldCount As Integer
    Dim pFeatLayer As IFeatureLayer
    Dim pLayer As ILayer

    pLayer = DirectCast(arrLayers.Item(m_cmbLayers.SelectedIndex), ILayer)
    If pLayer Is Nothing Then GoTo CleanUp

    'CLEAR OUT PREVIOUS 
    m_cmbValueFlds.Items.Clear()
    m_cmbValueFlds.Text = ""
    txtAlias.Text = ""
    m_tbxInfluence.Text = ""

    ' Add fields to the dropdown list
    If TryCast(arrLayers.Item(m_cmbLayers.SelectedIndex), IFeatureLayer) IsNot Nothing Then
      'ADD THE OPTION OF NONE FOR POINT AND LINE FEATURE LAYERS
      pFeatLayer = DirectCast(arrLayers.Item(m_cmbLayers.SelectedIndex), IFeatureLayer)
      If pFeatLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPoint Or pFeatLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline Then
        m_cmbValueFlds.Items.Add("<None>")
      End If

      ' Add numeric fields for point layers, and numeric and string fields for polygon layers.
      For intFldCount = 1 To pFeatLayer.FeatureClass.Fields.FieldCount - 1
        pField = pFeatLayer.FeatureClass.Fields.Field(intFldCount)
        If pFeatLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPoint OrElse pFeatLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
          If pField.Type = esriFieldType.esriFieldTypeInteger OrElse pField.Type = esriFieldType.esriFieldTypeDouble OrElse pField.Type = esriFieldType.esriFieldTypeSingle OrElse pField.Type = esriFieldType.esriFieldTypeSmallInteger Then
            m_cmbValueFlds.Items.Add(pField.Name)
          End If
        End If
        If pFeatLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
          If pField.Type = esriFieldType.esriFieldTypeString Then
            m_cmbValueFlds.Items.Add(pField.Name)
          End If
        End If
      Next
    Else
      ' Add None for raster layers
      m_cmbValueFlds.Items.Add("<None>")
    End If

    ' If the only field is "<None>" then select it
    If m_cmbValueFlds.Items.Count = 1 AndAlso m_cmbValueFlds.Items(0).ToString = "<None>" Then
      m_cmbValueFlds.SelectedItem = m_cmbValueFlds.Items(0)
    End If

    AddFactorButton_Status()
    GoTo CleanUp

CleanUp:
    pField = Nothing
    intFldCount = Nothing
    pFeatLayer = Nothing
    GC.WaitForPendingFinalizers()
    GC.Collect()
  End Sub

  Private Sub cmbValuesFld_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbValuesFld.SelectedIndexChanged, tbxInfluence.TextChanged
    ' Set the default alias
    'TRIGGER THE VALIDATION OF THE ADD FACTOR BUTTON
    AddFactorButton_Status()

    ' Validate the alias field
    txtAlias.Text = getValidAlias(txtAlias.Text, cmbLayers.Text, cmbValuesFld.Text)

    'TRIGGER THE VALIDATION OF THE ADD FACTOR BUTTON
    AddFactorButton_Status()
  End Sub

  Private Sub AddFactorButton_Status()
    'SETS THE ENABLED STATAUS OF THE ADD FACTOR BUTTON DEPENDING UPON THE INPUTS PROVIDED
    Dim blnLayerFound As Boolean = False
    Dim blnFieldFound As Boolean = False
    Dim blnInfluence As Boolean = False
    Dim blnAliasFound As Boolean
    Dim dblInfluence As Double = 0

    'CHECK FOR THE SELECTED LAYER
    If m_cmbLayers.FindStringExact(m_cmbLayers.Text) >= 0 Then
      blnLayerFound = True
    End If
    'CHECK FOR THE SELECTED VALUE FIELD
    If m_cmbValueFlds.FindStringExact(m_cmbValueFlds.Text) >= 0 Then
      blnFieldFound = True
    End If

    ' Check the alias
    If txtAlias.Text <> String.Empty Then blnAliasFound = True

    'REVIEW THE Influence PERCENTAGE VALUE, MUST BE NUMERIC AND GREATER THAN ZERO
    If IsNumeric(m_tbxInfluence.Text) Then
      dblInfluence = CDbl(m_tbxInfluence.Text)
      If dblInfluence > 0 Then
        m_tbxInfluence.ForeColor = Color.Black
        blnInfluence = True
      Else
        m_tbxInfluence.ForeColor = Color.Red
      End If
    Else
      m_tbxInfluence.ForeColor = Color.Red
    End If

    'SET THE 'ADD FACTOR' BUTTON ENABLED STATUS
    If blnLayerFound AndAlso blnFieldFound AndAlso blnInfluence AndAlso blnAliasFound Then
      m_btnAddFactor.Enabled = True
    Else
      m_btnAddFactor.Enabled = False
    End If

    'TotalInfluencePercentageGroupX()
    GoTo CleanUp

CleanUp:
    blnLayerFound = Nothing
    blnFieldFound = Nothing
    blnInfluence = Nothing
    dblInfluence = Nothing
    GC.Collect()
    GC.WaitForPendingFinalizers()
  End Sub

  Private Sub btnGroupXAddFactor_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddFactor.Click
    'ADD A ROW FOR THE SELECTED FEATURE LAYER, VALUE FIELD AND Influence PERCENTAGE TO BE UTILIZED IN PROCESSING
    Dim strLayer As String = m_cmbLayers.Text
    Dim strField As String = m_cmbValueFlds.Text
    Dim strInfluence As String = m_tbxInfluence.Text
    Dim intLayerIndex As Integer = m_cmbLayers.SelectedIndex
    Dim aliasName As String = txtAlias.Text
    Dim blnFound As Boolean = False
    Dim intRow As Integer
    Dim intCount As Integer

    Try
      'ENSURE LAYER AND FIELD COMBINATION DOES NOT ALREADY EXIST
      For intRow = 0 To m_dgvWeights.Rows.Count - 1
        If strLayer = CStr(m_dgvWeights.Rows(intRow).Cells(1).Value) And strField = CStr(m_dgvWeights.Rows(intRow).Cells(3).Value) And intLayerIndex = CType(m_dgvWeights.Rows(intRow).Cells(6).Value, Int32) Then
          MessageBox.Show(Me, "The selected feature layer and value field have already been designated as a Demographics factor.", "Demographic Factor Found", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
          blnFound = True
          GoTo CleanUp
        End If
      Next

      ' Ask the user if they want to use kernel density or natural neighbor interpolation
      Dim useKernelDensity As Boolean = False
      Dim disableKernelDensity As Boolean = True
      If TryCast(arrLayers.Item(m_cmbLayers.SelectedIndex), IFeatureLayer) IsNot Nothing Then
        Dim pFeatLayer As IFeatureLayer = CType(arrLayers.Item(m_cmbLayers.SelectedIndex), IFeatureLayer)
        If pFeatLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPoint OrElse pFeatLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline Then
          If strField <> "<None>" Then
            useKernelDensity = MessageBox.Show(Me, "This layer can be processed as a density overlay or a sampling overlay.  If you would like to display the density of your data, select yes (Kernel Density).  Otherwise, select no if your data includes sample points or polylines that will be used to create an interpolated surface (Natural Neighbor).", "Layer Processing Method", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) = Windows.Forms.DialogResult.Yes
            disableKernelDensity = False
          Else
            useKernelDensity = True
          End If
        End If
      End If

      'REVIEW CURRENT ROWS FOR A MATCH OF SELECTED SUBAREA INPUTS
      m_dgvWeights.Rows.Add()
      m_dgvWeights.Rows(m_dgvWeights.RowCount - 1).Cells(1).Value = strLayer
      m_dgvWeights.Rows(m_dgvWeights.RowCount - 1).Cells(2).Value = aliasName
      m_dgvWeights.Rows(m_dgvWeights.RowCount - 1).Cells(3).Value = strField
      m_dgvWeights.Rows(m_dgvWeights.RowCount - 1).Cells(4).Value = strInfluence
      m_dgvWeights.Rows(m_dgvWeights.RowCount - 1).Cells(5).Value = useKernelDensity
      If disableKernelDensity Then
        With m_dgvWeights.Rows(m_dgvWeights.RowCount - 1).Cells(5)
          .ReadOnly = True
          .Style.BackColor = Color.LightGray
          .Style.ForeColor = Color.DarkGray
        End With
      End If
      m_dgvWeights.Rows(m_dgvWeights.RowCount - 1).Cells(6).Value = intLayerIndex
      m_dgvWeights.Refresh()

      TotalInfluencePercentage()

      'CLEAR INPUT CONTROLS
      m_cmbLayers.Text = ""
      m_cmbValueFlds.Text = ""
      m_tbxInfluence.Text = ""
      m_btnAddFactor.Enabled = False
      txtAlias.Text = ""

      GoTo CleanUp
    Catch ex As Exception
      GoTo CleanUp
    End Try

CleanUp:
    strLayer = Nothing
    strField = Nothing
    strInfluence = Nothing
    intLayerIndex = Nothing
    blnFound = Nothing
    intRow = Nothing
    intCount = Nothing
    GC.WaitForPendingFinalizers()
    GC.Collect()
  End Sub

  Private Sub dgvWeights_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dgvWeights.MouseDown
    'REMOVE THE SELECTED LINE FROM THE DATA GRID
    Dim hit As DataGridView.HitTestInfo = m_dgvWeights.HitTest(e.X, e.Y)
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
        m_dgvWeights.Rows.RemoveAt(intDeleteRow)
        TotalInfluencePercentage()
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

  Private Sub dgvWeights_CellValueChanged(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvWeights.CellValueChanged

    If m_dgvWeights Is Nothing Then Exit Sub

    TotalInfluencePercentage()

    ' Validate the alias field
    m_dgvWeights.Rows(e.RowIndex).Cells(2).Value = getValidAlias(CStr(m_dgvWeights.Rows(e.RowIndex).Cells(2).Value), CStr(m_dgvWeights.Rows(e.RowIndex).Cells(1).Value), CStr(m_dgvWeights.Rows(e.RowIndex).Cells(3).Value))
  End Sub

  ''' <summary>
  ''' Gets a valid alias name for the input
  ''' </summary>
  ''' <param name="currentAlias"></param>
  ''' <param name="layerName"></param>
  ''' <param name="fieldName"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function getValidAlias(currentAlias As String, layerName As String, fieldName As String) As String
    Dim aliasName As String
    If currentAlias <> String.Empty Then
      aliasName = ReplaceSpecialChar(currentAlias)
      If aliasName.Length > 20 Then
        aliasName = aliasName.Substring(0, 20)
      End If
    Else
      If fieldName <> "<None>" Then
        aliasName = ReplaceSpecialChar(fieldName)
      Else
        aliasName = ReplaceSpecialChar(layerName)
      End If
    End If
    Return aliasName
  End Function

  Private Sub TotalInfluencePercentage()
    'SUM THE TOTAL PERCENTAGE FOR DEMOGRAPHIC FACTORS
    Dim strValue As String
    Dim dblTotalInfluence As Double = 0
    Dim intRow As Integer
    Dim intCount As Integer

    Try
      For intRow = 0 To m_dgvWeights.Rows.Count - 1
        strValue = CStr(m_dgvWeights.Rows(intRow).Cells(4).Value)
        If IsNumeric(strValue) Then
          dblTotalInfluence = dblTotalInfluence + CDbl(m_dgvWeights.Rows(intRow).Cells(4).Value)
        End If
      Next

      m_tbxTotalInfluence.Text = dblTotalInfluence.ToString
      If m_dgvWeights.RowCount = 0 Then
        m_tbxTotalInfluence.ForeColor = Color.Black
        GoTo CleanUp
      End If

      If dblTotalInfluence <> 100 Then
        m_tbxTotalInfluence.ForeColor = Color.Red
      Else
        m_tbxTotalInfluence.ForeColor = Color.Black
      End If

      GoTo CleanUp
    Catch ex As Exception
      GoTo CleanUp
    End Try

CleanUp:
    strValue = Nothing
    dblTotalInfluence = Nothing
    intRow = Nothing
    intCount = Nothing
    GC.WaitForPendingFinalizers()
    GC.Collect()
  End Sub

  Private Sub btnWorkspace_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnWorkspace.Click
    'PROMPT THE USER TO SELECT THE DIRECTORY WHERE PROCESSING OF THE LOCATION EFFICIENCY INPUTS WILL RESULT
    'WORKSPACE DIRECTORY
    Dim MyDialog As New FolderBrowserDialog
    Dim intCount As Integer
    Dim strFGeo As String = ""
    MyDialog.Description = "Select the Location Efficiency Workspace Directory:"
    Me.tbxProjectName.Enabled = False
    If (MyDialog.ShowDialog() = Windows.Forms.DialogResult.OK) Then
      Me.tbxWorkspace.Text = MyDialog.SelectedPath.ToString
      Me.tbxProjectName.Enabled = True
    End If

    'CYCLE THROUGH 200 FOR DEFAULT FILE GEODATABASE 
    For intCount = 1 To 200
      strFGeo = "LOCATION_EFFICIENCY_" & CStr(intCount)
      If Not Directory.Exists(Me.tbxWorkspace.Text & "\" & strFGeo & ".gdb") Then
        Me.tbxProjectName.Text = strFGeo
        Exit For
      End If
    Next

    GoTo CleanUp
CleanUp:
    MyDialog = Nothing
    intCount = Nothing
    strFGeo = Nothing
    GC.WaitForPendingFinalizers()
    GC.Collect()
  End Sub

  Private Sub tbxProjectName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbxProjectName.TextChanged
    'CHECK TO MAKE SURE ANY INPUT FILE GEODATABASE NAME DOES NOT EXIST
    If Directory.Exists(Me.tbxWorkspace.Text & "\" & tbxProjectName.Text & ".gdb") Then
      m_toolTip.SetToolTip(tbxProjectName, "The name provided already exists.  Please enter a new name.")
      Me.tbxProjectName.ForeColor = Color.Red
    Else
      Me.tbxProjectName.ForeColor = Color.Black
      m_toolTip.SetToolTip(tbxProjectName, "")
    End If
  End Sub

  Private Sub btnNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNext.Click
    'DETERMINE THE CURRENT TAB AND MOVE TO NEXT
    Me.InfoStationLocationSettings.SelectTab(Me.InfoStationLocationSettings.SelectedIndex + 1)
  End Sub

  Private Sub btnPrevious_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrevious.Click
    'DETERMINE THE CURRENT TAB AND MOVE TO PREVIOUS
    Me.InfoStationLocationSettings.SelectTab(Me.InfoStationLocationSettings.SelectedIndex - 1)
  End Sub

  Private Sub cmbExtentLayers_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbExtentLayers.SelectedIndexChanged, rdbExtentLayer.CheckedChanged
    'SET THE RASTER OUTPUT EXTENT TO THE EXTENT OF THE SELECTED LAYER
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
      mxApplication = CType(m_appEnvision, IMxApplication)
      mxDoc = CType(m_appEnvision.Document, IMxDocument)
      mapCurrent = CType(mxDoc.FocusMap, Map)

      'SET ENABLED STATUS TO COORDINATE INPUT CONTROLS
      Me.cmbExtentLayers.Enabled = Me.rdbExtentLayer.Checked

      'CLEAR THE EXTENT COORDINATES
      m_extentLowerX = 0
      m_extentLowerY = 0
      m_extentUpperX = 0
      m_extentUpperY = 0

      'RETRIEVE EXTENT ENVELOPE FROM SELECTED FEATURE LAYER 
      If Me.cmbExtentLayers.SelectedIndex >= 0 Then
        pLayer = CType(arrLayers.Item(Me.cmbExtentLayers.SelectedIndex), ILayer)
        pGeodataset = DirectCast(pLayer, IGeoDataset)
        pExtentEnv = pLayer.AreaOfInterest
        pSR = pGeodataset.SpatialReference
      Else
        GoTo CleanUp
      End If

      'CHECK THE SPATIAL REFERENCES TO MAKE SURE THEY MATCH, REPROJECT IF THEY DO NOT MATCH
      Try
        If Not m_pSpatRefSLTool Is Nothing And Not pSR Is Nothing Then
          pClone1 = DirectCast(m_pSpatRefSLTool, IClone)
          pClone2 = DirectCast(pSR, IClone)
          If Not pClone1.IsEqual(pClone2) And Not pExtentEnv Is Nothing Then
            pGeometry = DirectCast(pExtentEnv, IGeometry5)
            pGeometry.Project(m_pSpatRefSLTool)
            pExtentEnv = DirectCast(pGeometry, IEnvelope)
          End If
        End If
      Catch ex As Exception
        MessageBox.Show(Me, ex.Message, "Extent Retrieval Error - Layer Select", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
        GoTo CleanUp
      End Try

      'WRITE EXTENT COORDINATES TO THE REPRESENTATIVE CONTROLS
      If Not pExtentEnv Is Nothing Then
        m_extentLowerX = pExtentEnv.XMin
        m_extentLowerY = pExtentEnv.YMin
        m_extentUpperX = pExtentEnv.XMax
        m_extentUpperY = pExtentEnv.YMax
      End If
      m_pExtentSLToolEnv = pExtentEnv

    Catch ex As Exception
      MessageBox.Show(Me, ex.Message, "Extent Retrieval Error - Layer Select", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
      m_pExtentSLToolEnv = Nothing
      GoTo CleanUp
    End Try

    GoTo CleanUp
CleanUp:
    mxApplication = Nothing
    mxDoc = Nothing
    mapCurrent = Nothing
    pActiveView = Nothing
    pLayer = Nothing
    pExtentEnv = Nothing
    pGeodataset = Nothing
    pClone1 = Nothing
    pClone2 = Nothing
    pSR = Nothing
    pGeometry = Nothing
    GC.Collect()
    GC.WaitForPendingFinalizers()
  End Sub

  Private Sub rdbExtentView_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdbExtentView.CheckedChanged
    'RETRIEVE THE EXTENT FOR THE SELECTED PROCESSING EXTENT OPTION
    Dim mxApplication As IMxApplication
    Dim mxDoc As IMxDocument
    Dim mapCurrent As Map
    Dim pActiveView As IActiveView
    Dim pExtentEnv As IEnvelope = Nothing

    'RETRIEVE THE EXTENT FROM THE SELECTED OPTION
    Try
      mxApplication = CType(m_appEnvision, IMxApplication)
      mxDoc = CType(m_appEnvision.Document, IMxDocument)
      mapCurrent = CType(mxDoc.FocusMap, Map)

      'SET ENABLED STATUS TO COORDINATE INPUT CONTROLS
      'Me.gpbCoordinates.Enabled = Me.rdbExtentCustom.Checked

      'CLEAR THE EXTENT COORDINATES
      If Not Me.rdbExtentView.Checked Then
        m_extentLowerX = 0
        m_extentLowerY = 0
        m_extentUpperX = 0
        m_extentUpperY = 0
        GoTo CleanUp
      End If

      'RETRIEVE VIEW EXTENT ENVELOPE
      pActiveView = CType(mxDoc.FocusMap, IActiveView)
      pExtentEnv = pActiveView.Extent

      'WRITE EXTENT COORDINATES TO THE REPRESENTATIVE CONTROLS
      If Not pExtentEnv Is Nothing Then
        m_extentLowerX = pExtentEnv.XMin
        m_extentLowerY = pExtentEnv.YMin
        m_extentUpperX = pExtentEnv.XMax
        m_extentUpperY = pExtentEnv.YMax
      End If
      m_pExtentSLToolEnv = pExtentEnv

    Catch ex As Exception
      MessageBox.Show(Me, ex.Message, "Extent Retrieval Error - View Extent", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
      m_pExtentSLToolEnv = Nothing
      GoTo CleanUp
    End Try

    GoTo CleanUp
CleanUp:
    mxApplication = Nothing
    mxDoc = Nothing
    mapCurrent = Nothing
    pActiveView = Nothing
    pExtentEnv = Nothing
    GC.Collect()
    GC.WaitForPendingFinalizers()
  End Sub

  Private Sub btnRun_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRun.Click
    Dim blnFactors As Boolean = False
    Dim mxApplication As IMxApplication = Nothing
    Dim pMxDocument As IMxDocument = Nothing
    Dim mapCurrent As Map
    Dim pSpatRef As ISpatialReference
    Dim pPCS As IProjectedCoordinateSystem
    Dim intCellSize As Integer
    Dim intSearchRadius As Integer
    Dim pCreateFileGDB As ESRI.ArcGIS.DataManagementTools.CreateFileGDB
    Dim intRow As Integer
    Dim strField As String = ""
    Dim strAlias As String
    Dim dblInfluence As Double = 0
    Dim intLayerIndex As Integer
    Dim useKernelDensity As Boolean
    Dim pFeatLayer As IFeatureLayer
    Dim pRasterLayer As IRasterLayer
    Dim pDataset As IDataset
    Dim strInFeatures As String
    Dim objExtent As Object
    Dim CreateInputRaster As ESRI.ArcGIS.ConversionTools.PolygonToRaster
    Dim CreateTimesRaster As ESRI.ArcGIS.SpatialAnalystTools.Times
    Dim CreateKernelRaster As ESRI.ArcGIS.SpatialAnalystTools.KernelDensity
    Dim CreateNaturalNghbrRaster As ESRI.ArcGIS.SpatialAnalystTools.NaturalNeighbor
    Dim CreateSliceRaster As ESRI.ArcGIS.SpatialAnalystTools.Slice
    Dim CreateRasterCalc As ESRI.ArcGIS.SpatialAnalystTools.RasterCalculator
    Dim CopyRaster As ESRI.ArcGIS.DataManagementTools.CopyRaster
    Dim strExpression As String = ""
    Dim strGroupExpression As String = ""
    Dim intCount As Integer
    Dim strFGeo As String
    Dim dblWeight As Double
    Dim blnAddWeights As Boolean = True
    Dim oFile As System.IO.File = Nothing
    Dim oWrite As System.IO.StreamWriter
    Dim wsRaster As IRasterWorkspace2
    Dim zoneLayer As ILayer = Nothing
    Dim zoneLayerIsOverlappingPolys As Boolean = False
    Dim zoneFeatureCursor As IFeatureCursor = Nothing
    Dim zoneFeatureLayer As IFeatureLayer = Nothing
    Dim zonalStatsInsertCursor As ICursor = Nothing

    '--------------------------------------------------------------------------------------------
    'PROMPT THE USER TO CREATE PROCEED WITH THE ENVISION PROJECT SETUP
    'If MessageBox.Show(me,"Would you like to proceed with processing?", "Location Efficiency Tool", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2, False) = Windows.Forms.DialogResult.No Then
    '  GoTo CleanUp
    'Else
    strProcessing = "LOCATION EFFICIENCY PROCESSING START TIME: " & Date.Now.ToString("M/d/yyyy hh:mm:ss tt") & vbNewLine & "---------------------------------------------------------------------------------------------" & vbNewLine
    m_appEnvision.StatusBar.Message(0) = "Starting Location Efficiency Processes: Review Inputs"
    m_frmLocationEfficiency.barStatusRun.Visible = True
    m_frmLocationEfficiency.lblRunStatus.Visible = True
    m_frmLocationEfficiency.barStatusRun.Value = 0
    m_frmLocationEfficiency.lblRunStatus.Text = "Starting Location Efficiency Processes: Review Inputs"
    m_frmLocationEfficiency.InfoSettings.Refresh()
    'End If

    ' Make sure the cell size is properly set
    If txtCellSize.Text = String.Empty OrElse Not Int32.TryParse(txtCellSize.Text, intCellSize) OrElse Int32.Parse(txtCellSize.Text) <= 0 Then
      MessageBox.Show(Me, "Please enter a cell size in feet for the output raster.", "Invalid Cell Size", MessageBoxButtons.OK, MessageBoxIcon.Information)
      GoTo CleanUp
    End If
    Dim cellSize As Int32 = Int32.Parse(txtCellSize.Text)

    ' Make sure the search radius is properly set
    If txtSearchRadius.Text = String.Empty OrElse Not Int32.TryParse(txtSearchRadius.Text, intSearchRadius) OrElse Int32.Parse(txtSearchRadius.Text) <= 0 Then
      MessageBox.Show(Me, "Please enter a search radius in feet.  This value will be used for the kernel density operation on point and line inputs.", "Invalid Search Radius", MessageBoxButtons.OK, MessageBoxIcon.Information)
      GoTo CleanUp
    End If
    Dim searchRadius As Int32 = Int32.Parse(txtSearchRadius.Text)

    ' Make sure all alias names are unique
    Dim aliasList As List(Of String) = New List(Of String)
    For intRow = 0 To dgvWeights.Rows.Count - 1
      Dim aliasStr As String = CStr(dgvWeights.Rows(intRow).Cells(2).Value)
      If aliasList.Contains(aliasStr) Then
        Me.InfoStationLocationSettings.SelectTab(0)
        MessageBox.Show(Me, "Please ensure all alias values are unique", "Invalid Inputs", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1)
        GoTo CleanUp
      Else
        aliasList.Add(aliasStr)
      End If
    Next

    '--------------------------------------------------------------------------------------------
    'REVIEW ALL FACTOR THE INPUTS - ARE THERE DESIGNATED FACTORS AND IS THE TOTAL INFLUENCE 100% - ASSUMPTION IS ALL THE INFLUENCE TOTALS ARE ACCURATE
    m_frmLocationEfficiency.barStatusRun.Value = 2
    m_frmLocationEfficiency.lblRunStatus.Text = "Starting Location Efficiency Processes: Review Inputs"
    If Me.dgvWeights.RowCount > 0 Then
      If Me.tbxTotalInfluence.Text = "100" Then
        blnFactors = True
      Else
        Me.InfoStationLocationSettings.SelectTab(0)
        MessageBox.Show(Me, "Please review the influence percentages as the total does not equal 100%.", "Invalid Inputs", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1)
        GoTo CleanUp
      End If
    Else
      Me.InfoStationLocationSettings.SelectTab(0)
      MessageBox.Show(Me, "Please enter layer weights.", "Invalid Inputs", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1)
      GoTo CleanUp
    End If

    ' Verify the extent is specified
    If rdbExtentLayer.Checked AndAlso cmbExtentLayers.Text = String.Empty Then
      MessageBox.Show(Me, "Please specify the extent of the raster output.", "Invalid Extent", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1)
      GoTo CleanUp
    End If

    ' Verify there is a zone field selected if there is a zone data layer selected
    If cboZoneLayer.Text <> String.Empty AndAlso cboZoneField.Text = String.Empty Then
      MessageBox.Show(Me, "Please select a zone field to calculate zonal statistics.", "Missing Zone Field", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1)
      GoTo CleanUp
    End If

    'RETRIEVE THE VIEW DOCUMENT COORDINATE SYSTEM AND UNITS OF MEASUERSURE
    'BY DEFUALT RETRIEVE AND LOAD THE VIEW DOCUMENT SPATIAL REFERENCE PROJECTION
    Try
      If Not TypeOf m_appEnvision Is IApplication Then
        GoTo CleanUp
      Else
        pMxDocument = DirectCast(m_appEnvision.Document, IMxDocument)
        mapCurrent = CType(pMxDocument.FocusMap, Map)
        'pActiveView = CType(pMxDocument.FocusMap, IActiveView)
      End If
      pSpatRef = pMxDocument.FocusMap.SpatialReference
      pPCS = CType(pSpatRef, IProjectedCoordinateSystem)
    Catch ex As Exception
      MessageBox.Show(Me, "Please set a Projected Coordinate System for the data frame.", "Projection Not Defined", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1)
      'MessageBox.Show(me,"Error in retrieving the view document projection information." & vbNewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1)
      GoTo CleanUp
    End Try

    'RECALCULATE THE GRID CELL SIZE AND SEARCH RADIUS VALUES
    If UCase(pPCS.CoordinateUnit.Name) = "YARDS" Then
      intCellSize = CType(cellSize * 0.33333, Int32)
      intSearchRadius = CType(searchRadius * 0.33333, Int32)
    ElseIf UCase(pPCS.CoordinateUnit.Name) = "MILES" Then
      intCellSize = CType(cellSize * 0.00018939, Int32)
      intSearchRadius = CType(searchRadius * 0.00018939, Int32)
    ElseIf UCase(pPCS.CoordinateUnit.Name) = "INCHES" Then
      intCellSize = cellSize * 12
      intSearchRadius = searchRadius * 12
    ElseIf UCase(pPCS.CoordinateUnit.Name) = "CENTIMETERS" Then
      intCellSize = CType(cellSize * 30.48, Int32)
      intSearchRadius = CType(searchRadius * 30.48, Int32)
    ElseIf UCase(pPCS.CoordinateUnit.Name) = "METERS" Then
      intCellSize = CType(cellSize * 0.3048, Int32)
      intSearchRadius = CType(searchRadius * 0.3048, Int32)
    ElseIf UCase(pPCS.CoordinateUnit.Name) = "KILOMETERS" Then
      intCellSize = CType(cellSize * 0.0003048, Int32)
      intSearchRadius = CType(searchRadius * 0.0003048, Int32)
    End If

    'CREATE THE GEOPROCESSOR TO BE UTILIZED BY THE LOCATION EFFICIENCY TOOL
    If Not CreateSLToolGP(intCellSize.ToString, pSpatRef, Me.tbxWorkspace.Text) Then
      GoTo CleanUp
    End If
    strProcessing = strProcessing & vbNewLine & vbNewLine & "PROJECTION INFORMATION"
    strProcessing = strProcessing & vbNewLine & "Spatial Reference Name : " & pSpatRef.Name
    strProcessing = strProcessing & vbNewLine & "Coordinate Units : " & pPCS.CoordinateUnit.Name
    strProcessing = strProcessing & vbNewLine & "Cell Size : " & intCellSize.ToString
    strProcessing = strProcessing & vbNewLine & "Search Radius : " & intSearchRadius.ToString & vbNewLine

    strProcessing = strProcessing & vbNewLine & "Area of Interest (extent): "
    If m_frmLocationEfficiency.rdbExtentView.Checked Then
      strProcessing = strProcessing & "View Extent"
    Else
      strProcessing = strProcessing & "Layer Extent - " & m_frmLocationEfficiency.cmbExtentLayers.Text
    End If
    strProcessing = strProcessing & vbNewLine & "Extent Upper X : " & m_extentUpperX.ToString
    strProcessing = strProcessing & vbNewLine & "Extent Upper Y : " & m_extentUpperY.ToString
    strProcessing = strProcessing & vbNewLine & "Extent Lower X : " & m_extentLowerX.ToString
    strProcessing = strProcessing & vbNewLine & "Extent Lower Y : " & m_extentLowerY.ToString & vbNewLine

    'CHECK THAT THE SELECTED WORKSPACE DIRECTORY HAS BEEN SELECTED AND IS VALID
    If Not Me.tbxWorkspace.Text.Length > 0 Then
      MessageBox.Show(Me, "Please select a workspace directory.", "Workspace Directory Not Defined", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1)
      GoTo CleanUp
    Else
      If Not Directory.Exists(Me.tbxWorkspace.Text) Then
        MessageBox.Show(Me, "Please select a valid workspace directory.", "Workspace Directory Not Defined", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1)
        GoTo CleanUp
      End If
    End If

    'REVIEW THE FILE GEODATABASE NAME TO ENSURE IT IS NOT EXIST
    If Me.lblNewFileGDBName.Text.Length > 0 Then
      If Directory.Exists(Me.tbxWorkspace.Text & "\" & Me.lblNewFileGDBName.Text & ".gdb") Then
        MessageBox.Show(Me, "Please enter a new file geodatabase name.", "File Geodatabase Exists", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1)
        GoTo CleanUp
      End If
    Else
      MessageBox.Show(Me, "Please enter a new file geodatabase name.", "File Geodatabase Not Defined", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1)
      GoTo CleanUp
    End If
    strProcessing = strProcessing & vbNewLine & "WORKSPACE DIRECTORY AND PROCESSING FILE GEODATABASE"
    strProcessing = strProcessing & vbNewLine & "Workspace : " & Me.tbxWorkspace.Text
    strProcessing = strProcessing & vbNewLine & "File Geodatabase Name : " & Me.tbxProjectName.Text & vbNewLine

    '--------------------------------------------------------------------------------------------
    'CREATE A FILE GEODATABASE FROM THE USER INPUTS, WHICH WILL HOLD ALL THE PROCESSING FILES
    m_appEnvision.StatusBar.Message(0) = "Creating Location Efficiency File Geodatabase"
    m_frmLocationEfficiency.barStatusRun.Value = 5
    m_frmLocationEfficiency.lblRunStatus.Text = "Creating Location Efficiency File Geodatabase"
    Try
      pCreateFileGDB = New ESRI.ArcGIS.DataManagementTools.CreateFileGDB
      pCreateFileGDB.out_folder_path = Me.tbxWorkspace.Text
      pCreateFileGDB.out_name = Me.tbxProjectName.Text
      RunTool(pGP_SLTool, pCreateFileGDB)
      pCreateFileGDB = Nothing
    Catch ex As Exception
      MessageBox.Show(Me, "Error in creating the Location Efficiency Raster Processing File geodatabase, " & Me.tbxProjectName.Text & "." & vbNewLine & ex.Message, "Processing Error", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1)
      GoTo CleanUp
    End Try

    'CREATE AREA OF INTEREST (AOI) FEAGTURE CLASS
    m_appEnvision.StatusBar.Message(0) = "Creating Area of Interest (AOI) Layer"
    If CreateAOIFeatClass() Then
      If Not CreateAOIFeature() Then
        GoTo CleanUp
      End If
    Else
      GoTo CleanUp
    End If

    Me.barStatusRun.Visible = True
    strProcessing = strProcessing & vbNewLine & "CREATE INPUT RASTERS, and INFLUENCE PERCENTAGE RASTERS"
    m_appEnvision.StatusBar.Message(0) = "Processing factors"
    m_frmLocationEfficiency.lblRunStatus.Text = m_appEnvision.StatusBar.Message(0)
    For intRow = 0 To dgvWeights.Rows.Count - 1
      Try
        m_frmLocationEfficiency.barStatusRun.Value = 5 + CType(70 * (intRow + 1 / dgvWeights.Rows.Count), Int32)
      Catch ex As Exception
      End Try

      Try
        strAlias = CStr(dgvWeights.Rows(intRow).Cells(2).Value)
        strField = CStr(dgvWeights.Rows(intRow).Cells(3).Value)
        intLayerIndex = CType(dgvWeights.Rows(intRow).Cells(6).Value, Int32)
        useKernelDensity = CType(dgvWeights.Rows(intRow).Cells(5).Value, Boolean)
        dblInfluence = CDbl(dgvWeights.Rows(intRow).Cells(4).Value) / 100
        If TryCast(arrLayers.Item(intLayerIndex), IFeatureLayer) IsNot Nothing Then
          pDataset = DirectCast(DirectCast(arrLayers.Item(intLayerIndex), IFeatureLayer).FeatureClass, IDataset)
        Else
          pDataset = DirectCast(DirectCast(arrLayers.Item(intLayerIndex), IRasterLayer), IDataset)
        End If
        If pDataset.Category.ToString.Contains("Shapefile") Then
          strInFeatures = pDataset.Workspace.PathName & "\" & pDataset.BrowseName & ".shp"
        Else
          strInFeatures = pDataset.Workspace.PathName & "\" & pDataset.BrowseName
        End If
        strProcessing = strProcessing & vbNewLine & "START TIME: " & Date.Now.ToString("M/d/yyyy hh:mm:ss tt")
        strProcessing = strProcessing & vbNewLine & "Factor Layer " & intRow.ToString & " (INPUT LAYER): " & strInFeatures
        strProcessing = strProcessing & vbNewLine & "Factor Layer " & intRow.ToString & " (ALIAS): " & strAlias
        strProcessing = strProcessing & vbNewLine & "Factor Layer " & intRow.ToString & " (VALUE FIELD): " & strField
        strProcessing = strProcessing & vbNewLine & "Factor Layer " & intRow.ToString & " (PERCENT INFLUENCE): " & dgvWeights.Rows(intRow).Cells(4).Value.ToString

        'CREATE A GRID FROM THE INPUTS PROVIDED
        If TryCast(arrLayers.Item(intLayerIndex), IFeatureLayer) IsNot Nothing Then
          pFeatLayer = DirectCast(arrLayers.Item(intLayerIndex), IFeatureLayer)
          If pFeatLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
            CreateInputRaster = New ESRI.ArcGIS.ConversionTools.PolygonToRaster
            CreateInputRaster.in_features = pFeatLayer
            CreateInputRaster.value_field = strField
            CreateInputRaster.out_rasterdataset = (m_frmLocationEfficiency.tbxWorkspace.Text + "\" & m_frmLocationEfficiency.tbxProjectName.Text & ".gdb\INPUT_" & strAlias)
            strProcessing = strProcessing & vbNewLine & "Factor Layer " & intRow.ToString & " (GRID NAME - INPUT): INPUT_" & strAlias
            If Not RunTool(pGP_SLTool, CreateInputRaster) Then
              GoTo CleanUp
            End If
          ElseIf pFeatLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPoint Or pFeatLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline Then
            ' Allow either kernel density of natural neighbor
            If strField = "<None>" OrElse useKernelDensity Then
              objExtent = ""
              pGP_SLTool.SetEnvironmentValue("extent", objExtent)
              CreateKernelRaster = New ESRI.ArcGIS.SpatialAnalystTools.KernelDensity
              CreateKernelRaster.in_features = pFeatLayer
              If strField <> "<None>" Then CreateKernelRaster.population_field = strField
              CreateKernelRaster.search_radius = intSearchRadius
              CreateKernelRaster.area_unit_scale_factor = "SQUARE_MILES"
              CreateKernelRaster.out_raster = (m_frmLocationEfficiency.tbxWorkspace.Text + "\" & m_frmLocationEfficiency.tbxProjectName.Text & ".gdb\INPUT_" & strAlias)
              strProcessing = strProcessing & vbNewLine & "Factor Layer " & intRow.ToString & " (GRID NAME - INPUT): INPUT_" & strAlias
              If Not RunTool(pGP_SLTool, CreateKernelRaster) Then
                GoTo CleanUp
              End If
              objExtent = m_extentUpperX.ToString + " " + m_extentUpperY.ToString + " " + m_extentLowerX.ToString + " " + m_extentLowerY.ToString
              pGP_SLTool.SetEnvironmentValue("extent", objExtent)

            Else
              CreateNaturalNghbrRaster = New ESRI.ArcGIS.SpatialAnalystTools.NaturalNeighbor
              CreateNaturalNghbrRaster.in_point_features = pFeatLayer
              CreateNaturalNghbrRaster.z_field = strField
              CreateNaturalNghbrRaster.out_raster = (m_frmLocationEfficiency.tbxWorkspace.Text + "\" & m_frmLocationEfficiency.tbxProjectName.Text & ".gdb\INPUT_" & strAlias)
              strProcessing = strProcessing & vbNewLine & "Factor Layer " & intRow.ToString & " (GRID NAME - INPUT): INPUT_" & strAlias
              If Not RunTool(pGP_SLTool, CreateNaturalNghbrRaster) Then
                GoTo CleanUp
              End If
            End If
          End If
        Else
          ' Copy the raster as it is
          pRasterLayer = DirectCast(arrLayers.Item(intLayerIndex), IRasterLayer)
          CopyRaster = New ESRI.ArcGIS.DataManagementTools.CopyRaster
          CopyRaster.in_raster = pRasterLayer
          CopyRaster.out_rasterdataset = (m_frmLocationEfficiency.tbxWorkspace.Text + "\" & m_frmLocationEfficiency.tbxProjectName.Text & ".gdb\INPUT_" & strAlias)
          strProcessing = strProcessing & vbNewLine & "Factor Layer " & intRow.ToString & " (GRID NAME - INPUT): INPUT_" & strAlias
          If Not RunTool(pGP_SLTool, CopyRaster) Then
            GoTo CleanUp
          End If
        End If

        CreateSliceRaster = New ESRI.ArcGIS.SpatialAnalystTools.Slice
        CreateSliceRaster.in_raster = (m_frmLocationEfficiency.tbxWorkspace.Text + "\" & m_frmLocationEfficiency.tbxProjectName.Text & ".gdb\INPUT_" & strAlias)
        CreateSliceRaster.base_output_zone = 1
        CreateSliceRaster.number_zones = 10
        CreateSliceRaster.slice_type = "EQUAL_AREA"
        Dim sliceName As String = String.Empty
        If dgvWeights.RowCount = 1 Then
          sliceName = "LOCATION_EFFICIENCY"
          strProcessing = strProcessing & vbNewLine & vbNewLine & "GRID NAME - FINAL: LOCATION_EFFICIENCY"
        Else
          sliceName = "SLICE_" & strAlias
          strProcessing = strProcessing & vbNewLine & "Factor Layer " & intRow.ToString & " (GRID NAME - SLICE): SLICE_" & strAlias
        End If
        CreateSliceRaster.out_raster = (m_frmLocationEfficiency.tbxWorkspace.Text + "\" & m_frmLocationEfficiency.tbxProjectName.Text & ".gdb\" & sliceName)
        If sliceName = "LOCATION_EFFICIENCY" Then pGP_SLTool.AddOutputsToMap = True
        If Not RunSLTool(pGP_SLTool, CreateSliceRaster) Then
          GoTo CleanUp
        Else
          'REPLACE ALL NO DATA VALUES FOR A VALUE OF 0
          m_appEnvision.StatusBar.Message(0) = "Processing the grid,SLICE_" & strAlias & ", to set all NoData values to 0."
          ClearNoData(sliceName)
        End If
        pGP_SLTool.AddOutputsToMap = False

        If dgvWeights.RowCount <> 1 Then
          CreateTimesRaster = New ESRI.ArcGIS.SpatialAnalystTools.Times
          CreateTimesRaster.in_raster_or_constant1 = (m_frmLocationEfficiency.tbxWorkspace.Text + "\" & m_frmLocationEfficiency.tbxProjectName.Text & ".gdb\SLICE_" & strAlias)
          CreateTimesRaster.in_raster_or_constant2 = dblInfluence.ToString
          CreateTimesRaster.out_raster = (m_frmLocationEfficiency.tbxWorkspace.Text + "\" & m_frmLocationEfficiency.tbxProjectName.Text & ".gdb\PERCENTAGE_" & strAlias)
          strProcessing = strProcessing & vbNewLine & "Factor Layer " & intRow.ToString & " (GRID NAME - PERCENT INFLUENCE): PERCENTAGE_" & strAlias
          If Not RunSLTool(pGP_SLTool, CreateTimesRaster) Then
            GoTo CleanUp
          End If
        End If

        CreateInputRaster = Nothing
        CreateTimesRaster = Nothing
        CreateSliceRaster = Nothing
        CreateKernelRaster = Nothing
        CreateNaturalNghbrRaster = Nothing
        GC.Collect()
        GC.WaitForPendingFinalizers()

      Catch ex As Exception
        MessageBox.Show(ex.Message, "Grid Processing Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
        GoTo CleanUp
      End Try
      strProcessing = strProcessing & vbNewLine & "END TIME: " & Date.Now.ToString("M/d/yyyy hh:mm:ss tt")
    Next

    'USE THE RASTER CALCULATOR TO SUM ALL THE INFLUENCE GRIDS INTO A SINGLE GRID
    If dgvWeights.RowCount <> 1 Then
      'CREATE THE RASTER CALCULATOR EXPRESSION
      For intRow = 0 To dgvWeights.Rows.Count - 1
        strAlias = CStr(dgvWeights.Rows(intRow).Cells(2).Value)
        'If intRow = 0 Then
        '  strExpression = "'PERCENTAGE_" & strAlias & "'"
        'Else
        '  strExpression = strExpression & " + 'PERCENTAGE_" & strAlias & "'"
        'End If
        Dim pctRasterPath As String = "'" & m_frmLocationEfficiency.tbxWorkspace.Text + "\" & m_frmLocationEfficiency.tbxProjectName.Text & ".gdb\PERCENTAGE_" & strAlias & "'"
        If intRow = 0 Then
          strExpression = pctRasterPath
        Else
          strExpression = strExpression & " + " & pctRasterPath
        End If
      Next
      CreateRasterCalc = New ESRI.ArcGIS.SpatialAnalystTools.RasterCalculator
      CreateRasterCalc.expression = strExpression
      CreateRasterCalc.output_raster = (m_frmLocationEfficiency.tbxWorkspace.Text + "\" & m_frmLocationEfficiency.tbxProjectName.Text & ".gdb\LOCATION_EFFICIENCY")
      strProcessing = strProcessing & vbNewLine & vbNewLine & "GRID NAME - FINAL: LOCATION_EFFICIENCY"

      pGP_SLTool.AddOutputsToMap = True
      If Not RunSLTool(pGP_SLTool, CreateRasterCalc) Then
        GoTo CleanUp
      End If
      strProcessing = strProcessing & vbNewLine & "END TIME: " & Date.Now.ToString("M/d/yyyy hh:mm:ss tt")
      pGP_SLTool.AddOutputsToMap = False
    End If

    ' Zonal statistics processing
    If cboZoneLayer.Text <> String.Empty Then
      ' Make a list of all the value layers to processing zonal statistics
      Dim valueLayerList As List(Of String) = New List(Of String)
      If dgvWeights.RowCount <> 1 Then
        For n As Int32 = 0 To dgvWeights.Rows.Count - 1
          strAlias = CStr(dgvWeights.Rows(n).Cells(2).Value)
          valueLayerList.Add("SLICE_" & strAlias)
        Next
      End If
      valueLayerList.Add("LOCATION_EFFICIENCY")

      ' Get the zone layer and determine if it is a polygon featurelayer
      For Each lyr As ILayer In arrLayers
        If lyr.Name = cboZoneLayer.Text Then
          zoneLayer = lyr
          If ChkDataHasOverlappingPolys.Checked AndAlso TryCast(zoneLayer, IFeatureLayer) IsNot Nothing AndAlso DirectCast(zoneLayer, IFeatureLayer).FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
            zoneLayerIsOverlappingPolys = True
            zoneFeatureLayer = DirectCast(zoneLayer, IFeatureLayer)
          End If
          Exit For
        End If
      Next

      ' Loop through the value layers and calculate zonal stats
      strProcessing = strProcessing & vbNewLine & vbNewLine & "CALCULATE ZONAL STATISTICS TABLES"
      For Each valueLayer As String In valueLayerList
        ' Set the output table name
        Dim fileGDBPath As String = m_frmLocationEfficiency.tbxWorkspace.Text + "\" & m_frmLocationEfficiency.tbxProjectName.Text & ".gdb"
        Dim zonalStatsTableName As String = valueLayer & "_ZONALSTATS"
        Dim tablePath As String = fileGDBPath & "\" & zonalStatsTableName

        ' Process each overlapping polygon separately
        If zoneLayerIsOverlappingPolys Then
          ' Get the layer definition and use this when searching for records
          Dim zoneLayerFeatureLayerDef As IFeatureLayerDefinition2 = DirectCast(zoneLayer, IFeatureLayerDefinition2)
          Dim queryFilter As IQueryFilter = New QueryFilter
          queryFilter.WhereClause = zoneLayerFeatureLayerDef.DefinitionExpression

          ' Get a featurecursor and the first feature
          zoneFeatureCursor = zoneFeatureLayer.FeatureClass.Search(queryFilter, True)
          Dim feature As IFeature = zoneFeatureCursor.NextFeature
          If feature Is Nothing Then Throw New Exception("There are no features in the feature zone data.")

          ' Call zonal statistics as table for the first feature
          Dim zonalStatsTable As ITable = CalculateZonalStatsAsTableSingleFeature(valueLayer, zoneLayer, feature, zoneFeatureLayer, tablePath, fileGDBPath, zonalStatsTableName)
          If zonalStatsTable Is Nothing Then GoTo cleanup
          zonalStatsInsertCursor = zonalStatsTable.Insert(True)
          Dim rowBuffer As IRowBuffer = zonalStatsTable.CreateRowBuffer

          ' Loop through the remaining features
          feature = zoneFeatureCursor.NextFeature
          Do While feature IsNot Nothing
            ' Call zonal statistics as table for this feature
            zonalStatsTableName = valueLayer & "_ZONALSTATS_" & feature.OID.ToString
            Dim table As ITable = CalculateZonalStatsAsTableSingleFeature(valueLayer, zoneLayer, feature, zoneFeatureLayer, tablePath & "_" & feature.OID.ToString, fileGDBPath, zonalStatsTableName)
            If table Is Nothing Then GoTo cleanup

            ' Copy the values to the first table
            Dim row As IRow = table.GetRow(1)
            For n As Int32 = 0 To zonalStatsTable.Fields.FieldCount - 1
              Dim fld As IField = zonalStatsTable.Fields.Field(n)
              If fld.Name <> zonalStatsTable.OIDFieldName Then rowBuffer.Value(n) = row.Value(row.Fields.FindField(fld.Name))
            Next
            zonalStatsInsertCursor.InsertRow(rowBuffer)

            ' Delete the table
            Dim dataset As IDataset = DirectCast(table, IDataset)
            If dataset.CanDelete Then dataset.Delete()

            ' Get the next feature
            feature = zoneFeatureCursor.NextFeature
          Loop
          ' Flush the insert cursor
          zonalStatsInsertCursor.Flush()
        Else
          ' Process the entire layer
          If CalculateZonalStatsAsTable(valueLayer, zoneLayer, tablePath, fileGDBPath, zonalStatsTableName) Is Nothing Then GoTo cleanup
        End If
      Next
    End If

    strProcessing = strProcessing & vbNewLine & "'----------------------------------------------------------------------" & vbNewLine

    'REMOVE ALL EXCEPT THE LOCATION EFFICIENCY GRID FROM THE VIEW
    If LocationEfficiencyCleanUpTOC() Then
      MessageBox.Show(Me, "Location Efficiency processing has completed.  Please review the outputs.  If zonal statistics was selected, summary tables can be found in the workspace file geodatabase.", "Processing Completed", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
    Else
      MessageBox.Show(Me, "Location Efficiency processing has completed.  Please review the outputs.  If zonal statistics was selected, summary tables can be found in the workspace file geodatabase." & vbNewLine & "A new classification legend failed to update.", "Processing Completed", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
    End If

    GoTo CleanUp
CleanUp:

    'WRITE THE PROCESSING TEXT TO FILE FOR REVIEW 
    m_appEnvision.StatusBar.Message(0) = "Writing processing information to text file"
    m_frmLocationEfficiency.barStatusRun.Value = 95
    m_frmLocationEfficiency.lblRunStatus.Text = m_appEnvision.StatusBar.Message(0)
    If Me.tbxWorkspace.Text <> String.Empty AndAlso Directory.Exists(Me.tbxWorkspace.Text) Then
      oWrite = oFile.CreateText(Me.tbxWorkspace.Text & "\ET_LOCATION_EFFICIENCY_Tool_Processing_" & Me.tbxProjectName.Text & ".txt")
      oWrite.Write(strProcessing)
      oWrite.WriteLine("Processing End Time: " & Date.Now.ToString("M/d/yyyy hh:mm:ss tt"))
      oWrite.Close()
    End If
    Try
      'SET THE PROJECT NAME THE NEXT VALID PROJECT NAME NOT IN USE
      'CYCLE THROUGH 200 FOR DEFAULT FILE GEODATABASE 
      For intCount = 1 To 200
        strFGeo = "LOCATION_EFFICIENCY_" & CStr(intCount)
        If Not Directory.Exists(m_frmLocationEfficiency.tbxWorkspace.Text & "\" & strFGeo & ".gdb") Then
          Me.tbxProjectName.Text = strFGeo
          Exit For
        End If
      Next

    Catch ex As Exception

    End Try

    m_frmLocationEfficiency.barStatusRun.Value = 0
    m_frmLocationEfficiency.lblRunStatus.Text = ""
    Me.barStatusRun.Visible = False
    mxApplication = Nothing
    pMxDocument = Nothing
    mapCurrent = Nothing
    pSpatRef = Nothing
    pPCS = Nothing
    intCellSize = Nothing
    intSearchRadius = Nothing
    pCreateFileGDB = Nothing
    intRow = Nothing
    strField = Nothing
    dblInfluence = Nothing
    intLayerIndex = Nothing
    pFeatLayer = Nothing
    pRasterLayer = Nothing
    pDataset = Nothing
    strInFeatures = Nothing
    objExtent = Nothing
    CreateInputRaster = Nothing
    CreateTimesRaster = Nothing
    CreateKernelRaster = Nothing
    CreateNaturalNghbrRaster = Nothing
    CreateSliceRaster = Nothing
    CreateRasterCalc = Nothing
    strExpression = Nothing
    intCount = Nothing
    strFGeo = Nothing
    dblWeight = Nothing
    blnAddWeights = Nothing
    oFile = Nothing
    oWrite = Nothing
    wsRaster = Nothing
    If zoneFeatureCursor IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(zoneFeatureCursor)
    If zonalStatsInsertCursor IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(zonalStatsInsertCursor)
    GC.Collect()
    GC.WaitForPendingFinalizers()
  End Sub

  ''' <summary>
  ''' Calculate zonal statistics as table for a single feature
  ''' </summary>
  ''' <param name="valueLayer"></param>
  ''' <param name="zoneLayer"></param>
  ''' <param name="feature"></param>
  ''' <param name="zoneFeatureLayer"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function CalculateZonalStatsAsTableSingleFeature(valueLayer As String, zoneLayer As ILayer, feature As IFeature, zoneFeatureLayer As IFeatureLayer, _
                                                           outputTableName As String, fileGDBPath As String, zonalStatsTableName As String) As ITable
    'Dim currentExtent As String = String.Empty
    Dim zoneLayerFeatureLayerDef As IFeatureLayerDefinition2 = DirectCast(zoneLayer, IFeatureLayerDefinition2)
    Dim existingZoneLayerDef As String = zoneLayerFeatureLayerDef.DefinitionExpression

    Try
      ' Modify the layer definition
      If existingZoneLayerDef = String.Empty Then
        zoneLayerFeatureLayerDef.DefinitionExpression = zoneFeatureLayer.FeatureClass.OIDFieldName & " = " & feature.OID
      Else
        zoneLayerFeatureLayerDef.DefinitionExpression = "(" & existingZoneLayerDef & ") AND " & zoneFeatureLayer.FeatureClass.OIDFieldName & " = " & feature.OID
      End If

      ' Temporarily change the extent environment setting to be the extent of the feature
      'currentExtent = pGP_SLTool.GetEnvironmentValue("extent").ToString
      'Dim env As IEnvelope = feature.ShapeCopy.Envelope
      'env.Expand(1.1, 1.1, True)
      'env.Project(DirectCast(m_appEnvision.Document, IMxDocument).FocusMap.SpatialReference)
      'Dim featureExtent As String = env.XMax.ToString + " " + env.YMax.ToString + " " + env.XMin.ToString + " " + env.YMin.ToString
      'pGP_SLTool.SetEnvironmentValue("extent", featureExtent)

      ' Calculate the zonal statistics table
      Return CalculateZonalStatsAsTable(valueLayer, zoneLayer, outputTableName, fileGDBPath, zonalStatsTableName)
    Catch ex As Exception
      MessageBox.Show(Me, "Error in calculating zonal statistics for a single feature: " & vbNewLine & ex.Message, "Processing Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
      Return Nothing
    Finally
      'pGP_SLTool.SetEnvironmentValue("extent", currentExtent)

      ' Restore the original layer definition
      zoneLayerFeatureLayerDef.DefinitionExpression = existingZoneLayerDef
    End Try
  End Function

  ''' <summary>
  ''' Calculate zonal statistics as table
  ''' </summary>
  ''' <param name="valueLayer"></param>
  ''' <param name="zoneLayer"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function CalculateZonalStatsAsTable(valueLayer As String, zoneLayer As ILayer, outputTableName As String, fileGDBPath As String, zonalStatsTableName As String) As ITable
    Try
      ' Call the zonalstatsastable tool
      Dim createZonalStatsAsTable As ESRI.ArcGIS.SpatialAnalystTools.ZonalStatisticsAsTable
      createZonalStatsAsTable = New ESRI.ArcGIS.SpatialAnalystTools.ZonalStatisticsAsTable
      createZonalStatsAsTable.in_value_raster = m_frmLocationEfficiency.tbxWorkspace.Text + "\" & m_frmLocationEfficiency.tbxProjectName.Text & ".gdb\" & valueLayer
      createZonalStatsAsTable.in_zone_data = zoneLayer
      createZonalStatsAsTable.out_table = outputTableName
      createZonalStatsAsTable.zone_field = cboZoneField.Text
      strProcessing = strProcessing & vbNewLine & "Zonal Statistics Table Name: " & zonalStatsTableName
      If Not RunSLTool(pGP_SLTool, createZonalStatsAsTable) Then
        Return Nothing
      End If

      ' Return the table
      Dim workspaceFactory As IWorkspaceFactory = New FileGDBWorkspaceFactory
      Dim featureWorkspace As IFeatureWorkspace = DirectCast(workspaceFactory.OpenFromFile(fileGDBPath, 0), IFeatureWorkspace)
      Return featureWorkspace.OpenTable(zonalStatsTableName)
    Catch ex As Exception
      MessageBox.Show(Me, "Error in calculating zonal statistics: " & vbNewLine & ex.Message, "Processing Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
      Return Nothing
    End Try
  End Function

  Private Function CreateAOIFeatClass() As Boolean
    CreateAOIFeatClass = True
    'CREATE THE AOI FEATURE CLASS, WHICH CONTAIN THE AOI RECTANGLE FEATURE
    'WITH JUST ONE FEATURE REPRESENTING THE PROCESSING EXTENT FOR THE LOCATION EFFICIENCY TOOL
    Dim pCreateFeatClass As ESRI.ArcGIS.DataManagementTools.CreateFeatureclass
    Dim pWksFactory As IWorkspaceFactory = Nothing
    Dim pFeatWks As IFeatureWorkspace
    Dim pFeatClass As IFeatureClass
    Try

      pCreateFeatClass = New ESRI.ArcGIS.DataManagementTools.CreateFeatureclass
      pCreateFeatClass.out_path = m_frmLocationEfficiency.tbxWorkspace.Text & "\" & m_frmLocationEfficiency.tbxProjectName.Text & ".gdb"
      pCreateFeatClass.out_name = "PROJECT_AOI"
      pCreateFeatClass.spatial_reference = m_pSpatRefSLTool
      RunTool(pGP_SLTool, pCreateFeatClass)
      pCreateFeatClass = Nothing

      'DEFINE THE AOI LAYER
      pWksFactory = New FileGDBWorkspaceFactory
      pFeatWks = DirectCast(pWksFactory.OpenFromFile(m_frmLocationEfficiency.tbxWorkspace.Text & "\" & m_frmLocationEfficiency.tbxProjectName.Text & ".gdb", 0), IFeatureWorkspace)
      pFeatClass = pFeatWks.OpenFeatureClass("PROJECT_AOI")
      If pFeatClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
        m_lyrSLToolAOI = New FeatureLayer
        m_lyrSLToolAOI.FeatureClass = pFeatClass
        m_lyrSLToolAOI.Name = m_lyrSLToolAOI.Name & "<Location Efficiency Tool> " & "PROJECT_AOI"
      End If
    Catch ex As Exception
      MessageBox.Show(Me, "Error in creating AOI featureclass (PROJECT_AOI): " & vbNewLine & ex.Message, "Processing Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
      CreateAOIFeatClass = False
      GoTo CleanUp
    End Try

CleanUp:
    pCreateFeatClass = Nothing
    pWksFactory = Nothing
    pFeatWks = Nothing
    pFeatClass = Nothing
    GC.WaitForPendingFinalizers()
    GC.Collect()
  End Function

  Private Function CreateAOIFeature() As Boolean
    'GENERATE THE GRID CELL POLYGONS
    CreateAOIFeature = True

    Dim pntLowerLeft As ESRI.ArcGIS.Geometry.IPoint = New ESRI.ArcGIS.Geometry.Point
    Dim pntUpperRight As ESRI.ArcGIS.Geometry.IPoint = New ESRI.ArcGIS.Geometry.Point
    Dim intCellSize As Integer = 0
    Dim pInsertFeatureBuffer As IFeatureBuffer
    Dim pInsertFeatureCursor As IFeatureCursor
    Dim intNewFeatureCount As Integer = 0
    Dim pFClass As IFeatureClass
    Dim pPolyTemp As IPolygon
    Dim pPointCollection As IPointCollection
    Dim pPointTemp As ESRI.ArcGIS.Geometry.Point
    Dim pNewFeat As IFeature
    Dim pArea As IArea
    Dim intRecCount As Integer = 0

    Try
      'RETRIEVE GRID CELL LAYER FEATURE CLASS
      If m_lyrSLToolAOI Is Nothing Then
        CreateAOIFeature = False
        GoTo CleanUp
      Else
        pFClass = m_lyrSLToolAOI.FeatureClass
      End If

      'CHECK THE SELECTED EXTENT
      If m_pExtentSLToolEnv Is Nothing Then
        CreateAOIFeature = False
        GoTo CleanUp
      End If

      'FEATURE BUFFER SETUP
      pInsertFeatureCursor = pFClass.Insert(True)
      pInsertFeatureBuffer = pFClass.CreateFeatureBuffer

      Me.barStatusRun.Visible = False

      pPolyTemp = New Polygon
      pPointCollection = CType(pPolyTemp, IPointCollection)

      pPointTemp = New ESRI.ArcGIS.Geometry.Point
      pPointTemp.X = m_pExtentSLToolEnv.XMin
      pPointTemp.Y = m_pExtentSLToolEnv.YMax
      pPointCollection.AddPoint(pPointTemp)

      pPointTemp = New ESRI.ArcGIS.Geometry.Point
      pPointTemp.X = m_pExtentSLToolEnv.XMax
      pPointTemp.Y = m_pExtentSLToolEnv.YMax
      pPointCollection.AddPoint(pPointTemp)

      pPointTemp = New ESRI.ArcGIS.Geometry.Point
      pPointTemp.X = m_pExtentSLToolEnv.XMax
      pPointTemp.Y = m_pExtentSLToolEnv.YMin
      pPointCollection.AddPoint(pPointTemp)

      pPointTemp = New ESRI.ArcGIS.Geometry.Point
      pPointTemp.X = m_pExtentSLToolEnv.XMin
      pPointTemp.Y = m_pExtentSLToolEnv.YMin
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
      MessageBox.Show(Me, "Error in creating the AOI feature: " & vbNewLine & ex.Message, "Processing Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
      CreateAOIFeature = False
      GoTo CleanUp
    End Try

CleanUp:
    pntLowerLeft = Nothing
    pntUpperRight = Nothing
    intCellSize = Nothing
    pInsertFeatureBuffer = Nothing
    pInsertFeatureCursor = Nothing
    intNewFeatureCount = Nothing
    pFClass = Nothing
    pPolyTemp = Nothing
    pPointCollection = Nothing
    pPointTemp = Nothing
    pNewFeat = Nothing
    pArea = Nothing
    intRecCount = Nothing
    GC.WaitForPendingFinalizers()
    GC.Collect()
  End Function

  Public Function CreateSLToolGP(ByVal strCellSize As String, ByVal pSpatRef As ISpatialReference, ByVal strWorkspace As String) As Boolean
    'CREATE THE GEOPROCESSOR FOR THE STATION LOCATION TOOL
    CreateSLToolGP = True
    Dim objExtent As Object
    Dim objOutputSpatref As Object = pSpatRef.FactoryCode
    Dim objWorkspace As Object = strWorkspace

    Try

      pGP_SLTool = Nothing
      GC.Collect()
      GC.WaitForPendingFinalizers()
      objExtent = m_extentUpperX.ToString + " " + m_extentUpperY.ToString + " " + m_extentLowerX.ToString + " " + m_extentLowerY.ToString
      objWorkspace = Me.tbxWorkspace.Text
      pGP_SLTool = New ESRI.ArcGIS.Geoprocessor.Geoprocessor
      pGP_SLTool.SetEnvironmentValue("cellSize", strCellSize)
      pGP_SLTool.SetEnvironmentValue("extent", objExtent)
      pGP_SLTool.SetEnvironmentValue("workspace", objWorkspace)
      pGP_SLTool.SetEnvironmentValue("outputCoordinateSystem", objOutputSpatref)
      pGP_SLTool.OverwriteOutput = True
      pGP_SLTool.AddOutputsToMap = False
      pGP_SLTool.TemporaryMapLayers = False
    Catch ex As Exception
      CreateSLToolGP = False
      pGP_SLTool = Nothing
      MessageBox.Show(Me, "Error in creating a geoprocessor" & vbNewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
      GoTo CleanUp
    End Try

    GoTo CleanUp

CleanUp:
    objExtent = Nothing
    objOutputSpatref = Nothing
    objWorkspace = Nothing
    GC.WaitForPendingFinalizers()
    GC.Collect()
  End Function

  Public Function RunSLTool(ByVal geoProcessor As ESRI.ArcGIS.Geoprocessor.Geoprocessor, ByVal process As IGPProcess) As Boolean
    '*******************************************************
    ' Set the overwrite output option to true
    '*******************************************************
    Try
      geoProcessor.OverwriteOutput = True
      geoProcessor.Execute(process, Nothing)
      RunSLTool = ReturnMessages(geoProcessor)
      GoTo CleanUp
    Catch err As Exception
      RunSLTool = ReturnMessages(geoProcessor)
      GoTo CleanUp
    End Try
CleanUp:
    process = Nothing
    geoProcessor = Nothing
    GC.WaitForFullGCComplete()
    GC.WaitForPendingFinalizers()
    GC.Collect()
  End Function

  Private Function ReturnMessages(ByVal GP As ESRI.ArcGIS.Geoprocessor.Geoprocessor) As Boolean
    '*******************************************************
    ' Function for returning the tool messages.
    ' Print out the messages from tool executions
    '*******************************************************
    ReturnMessages = True
    Dim strMessage As String = ""
    Dim Count As Integer
    Try
      If GP.MessageCount > 0 Then
        For Count = 0 To GP.MessageCount - 1
          Try
            strMessage = strMessage & vbNewLine & GP.GetMessage(Count)
          Catch ex As Exception
            GoTo CleanUp
          End Try
        Next
        If strMessage.Contains("ERROR") = True Or strMessage.Contains("ERR:") = True Then
          MessageBox.Show(Me, strMessage, "Geoprocessing Error")
          ReturnMessages = False
        End If
        GoTo CleanUp
      End If
    Catch ex As Exception
      GoTo CleanUp
    End Try

CleanUp:
    'm_strProcessing = m_strProcessing & strMessage
    GP = Nothing
    GC.WaitForFullGCComplete()
    GC.WaitForPendingFinalizers()
    GC.Collect()
    Return ReturnMessages
  End Function

  Private Sub ClearNoData(ByVal inputDatasetName As String)
    'CYCLE THROGUH A SLICE GRID TO REPLACE ALL NO DATA CELL VALUES WITH A VALUE OF 0
    Dim inputWorkspace As String = m_frmLocationEfficiency.tbxWorkspace.Text & "\" & m_frmLocationEfficiency.tbxProjectName.Text & ".gdb"
    Dim workspaceFactory As ESRI.ArcGIS.Geodatabase.IWorkspaceFactory = New FileGDBWorkspaceFactory()
    Dim rasterWorkspace As IRasterWorkspaceEx = CType(workspaceFactory.OpenFromFile(inputWorkspace, 0), IRasterWorkspaceEx)
    Dim rasterDataset As IRasterDataset3 = DirectCast(rasterWorkspace.OpenRasterDataset(inputDatasetName), IRasterDataset3)
    Dim geoDs As IGeoDataset = DirectCast(rasterDataset, IGeoDataset)
    Dim rasterLayer As ESRI.ArcGIS.Carto.IRasterLayer = New RasterLayerClass()
    Dim bands As IRasterBandCollection
    Dim band As IRasterBand
    Dim rasterProps As IRasterProps
    Dim pixelBlock As IPixelBlock3
    Dim blockSize As IPnt
    Dim pRaster As IRaster
    Dim pRaster2 As IRaster2
    Dim col As Integer = 0
    Dim row As Integer = 0
    Dim pixelValue As Double = 0
    Dim pixels As System.Array
    Dim i, j As Integer
    Dim upperLeft As IPnt
    Dim rasterEdit As IRasterEdit

    Try
      rasterLayer.CreateFromDataset(DirectCast(geoDs, IRasterDataset))
      bands = DirectCast(rasterDataset, IRasterBandCollection)
      band = bands.Item(0)
      rasterProps = DirectCast(band, IRasterProps)
      rasterProps.NoDataValue = 99

      'Create a pixel block.
      pRaster = rasterDataset.CreateFullRaster
      pRaster2 = DirectCast(rasterLayer.Raster, IRaster2)
      blockSize = New Pnt
      blockSize.SetCoords(rasterProps.Width, rasterProps.Height)
      pixelBlock = DirectCast(pRaster.CreatePixelBlock(blockSize), IPixelBlock3)

      'Populate some pixel values to the pixel block.
      pixels = CType(pixelBlock.PixelData(0), System.Array)
      For i = 0 To rasterProps.Width - 1
        For j = 0 To rasterProps.Height - 1
          pixelValue = CDbl(pRaster2.GetPixelValue(0, i, j))
          If pixelValue <= 0 Then
            'IF NoData, THEN CHANGE TO NEW VALUE
            pixels.SetValue(Convert.ToByte(0), i, j)
          Else
            'IF VALUE IS FOUND, THEN KEEP VALUE
            pixels.SetValue(Convert.ToByte(pixelValue), i, j)
          End If
        Next j
      Next i
      pixelBlock.PixelData(0) = CType(pixels, System.Array)

      'Define the location that the upper left corner of the pixel block is to write.
      upperLeft = New Pnt
      upperLeft.SetCoords(0, 0)

      'Write the pixel block.
      rasterEdit = DirectCast(pRaster, IRasterEdit)
      rasterEdit.Write(upperLeft, DirectCast(pixelBlock, IPixelBlock))

      'Release rasterEdit explicitly.
      System.Runtime.InteropServices.Marshal.ReleaseComObject(rasterEdit)

      GoTo CleanUp
    Catch ex As Exception
      GoTo CleanUp
    End Try

CleanUp:
    inputWorkspace = Nothing
    workspaceFactory = Nothing
    rasterWorkspace = Nothing
    rasterDataset = Nothing
    geoDs = Nothing
    rasterLayer = Nothing
    bands = Nothing
    band = Nothing
    rasterProps = Nothing
    pixelBlock = Nothing
    blockSize = Nothing
    pRaster = Nothing
    pRaster2 = Nothing
    col = Nothing
    row = Nothing
    pixelValue = Nothing
    pixels = Nothing
    i = Nothing
    j = Nothing
    upperLeft = Nothing
    rasterEdit = Nothing
    GC.Collect()
    GC.WaitForPendingFinalizers()
  End Sub

  Private Function ReplaceSpecialChar(ByVal strString As String) As String
    ReplaceSpecialChar = strString
    ReplaceSpecialChar = Replace(ReplaceSpecialChar, " ", "_")
    ReplaceSpecialChar = Replace(ReplaceSpecialChar, "~", "_")
    ReplaceSpecialChar = Replace(ReplaceSpecialChar, "!", "_")
    ReplaceSpecialChar = Replace(ReplaceSpecialChar, "@", "_")
    ReplaceSpecialChar = Replace(ReplaceSpecialChar, "#", "_")
    ReplaceSpecialChar = Replace(ReplaceSpecialChar, "$", "_")
    ReplaceSpecialChar = Replace(ReplaceSpecialChar, "%", "_")
    ReplaceSpecialChar = Replace(ReplaceSpecialChar, "^", "_")
    ReplaceSpecialChar = Replace(ReplaceSpecialChar, "&", "_")
    ReplaceSpecialChar = Replace(ReplaceSpecialChar, "*", "_")
    ReplaceSpecialChar = Replace(ReplaceSpecialChar, "(", "_")
    ReplaceSpecialChar = Replace(ReplaceSpecialChar, ")", "_")
    ReplaceSpecialChar = Replace(ReplaceSpecialChar, "-", "_")
    ReplaceSpecialChar = Replace(ReplaceSpecialChar, "<", "_")
    ReplaceSpecialChar = Replace(ReplaceSpecialChar, ">", "_")
    ReplaceSpecialChar = Replace(ReplaceSpecialChar, "'", "_")
    ReplaceSpecialChar = Replace(ReplaceSpecialChar, ":", "_")
    ReplaceSpecialChar = Replace(ReplaceSpecialChar, ":", "_")
    ReplaceSpecialChar = Replace(ReplaceSpecialChar, "/", "_")
    ReplaceSpecialChar = Replace(ReplaceSpecialChar, "\", "_")
  End Function

  Public Function LocationEfficiencyCleanUpTOC() As Boolean
    LocationEfficiencyCleanUpTOC = True
    Dim pMxDocument As IMxDocument
    Dim mapCurrent As Map
    Dim pRasterLayer As IRasterLayer = Nothing
    Dim intLayer As Integer
    Dim strPath As String
    Dim arrRemove As ArrayList = New ArrayList
    Dim strMatchPath As String = m_frmLocationEfficiency.tbxWorkspace.Text + "\" & m_frmLocationEfficiency.tbxProjectName.Text & ".gdb"
    Dim pClassRen As IRasterClassifyColorRampRenderer
    Dim pLayerAffects As ILayerEffects
    pMxDocument = DirectCast(m_appEnvision.Document, IMxDocument)
    mapCurrent = CType(pMxDocument.FocusMap, Map)

    If mapCurrent.LayerCount > 0 Then
      Try
        'CYCLE LIST OF LAYERS AND REVIEW RASTER LAYERS.  RASTER LAYERS CONTAINING THE SAME PATH AS THE 
        'EXISTING PROJECT FILE GEODATABASE ARE ADDED TO ARRAY TO BE REMOVED FROM TOC
        For intLayer = 0 To mapCurrent.LayerCount - 1
          If TypeOf mapCurrent.Layer(intLayer) Is IRasterLayer Then
            pRasterLayer = CType(mapCurrent.Layer(intLayer), IRasterLayer)
            strPath = pRasterLayer.FilePath
            If strPath.Contains(strMatchPath) Then
              If Not strPath = m_frmLocationEfficiency.tbxWorkspace.Text + "\" & m_frmLocationEfficiency.tbxProjectName.Text & ".gdb\LOCATION_EFFICIENCY" Then
                arrRemove.Add(pRasterLayer)
              Else
                Try
                  pClassRen = StandardDevGridRenderer(pRasterLayer.Raster)
                  pRasterLayer.Renderer = DirectCast(pClassRen, IRasterRenderer)
                  pLayerAffects = DirectCast(pRasterLayer, ILayerEffects)
                  pLayerAffects.Transparency = 35
                  pMxDocument.ActiveView.Refresh()
                  pMxDocument.UpdateContents()
                Catch ex As Exception
                  LocationEfficiencyCleanUpTOC = False
                End Try
              End If
            End If
          End If
        Next

        If arrRemove.Count > 0 Then
          For intLayer = 0 To arrRemove.Count - 1
            pRasterLayer = DirectCast(arrRemove.Item(intLayer), IRasterLayer)
            mapCurrent.DeleteLayer(pRasterLayer)
          Next
        End If
      Catch ex As Exception
        GoTo CleanUp

      End Try

    End If
    GoTo CleanUp

CleanUp:
    pMxDocument = Nothing
    mapCurrent = Nothing
    pRasterLayer = Nothing
    intLayer = Nothing
    strPath = Nothing
    arrRemove = Nothing
    pClassRen = Nothing
    pLayerAffects = Nothing
    GC.WaitForPendingFinalizers()
    GC.Collect()
  End Function

  Private Function StandardDevGridRenderer(ByVal pRaster As IRaster) As RasterClassifyColorRampRendererClass

    Dim NumOfClass As Integer = 10
    Dim pRLayer As IRasterLayer
    Dim pClassRen As IRasterClassifyColorRampRenderer = New RasterClassifyColorRampRendererClass()
    Dim pRasRen As IRasterRenderer = DirectCast(pClassRen, IRasterRenderer)
    pRasRen.Raster = pRaster
    Dim pBandCol As IRasterBandCollection = DirectCast(pRaster, IRasterBandCollection)
    Dim pRasBand As IRasterBand = pBandCol.Item(0)
    pRasBand.ComputeStatsAndHist()
    Dim pRasterHistogram As IRasterHistogram = pRasBand.Histogram
    Dim temp As Double() = TryCast(pRasterHistogram.Counts, Double())
    Dim size As Integer = temp.Length
    Dim bound0 As Double = CDbl(pRasBand.Statistics.Minimum)
    Dim bound1 As Double = CDbl(pRasBand.Statistics.Maximum)
    Dim range As Double = bound1 - bound0
    Dim pVal As Double() = New Double(size - 1) {}
    Dim currentcount As Integer = -1
    Dim incrementSize As Double = range / 255.0
    Dim j As Double = 0
    Dim pFreq As Integer()
    Dim sd As StandardDeviation = New StandardDeviationClass()
    Dim pClassifyGen As IClassifyGEN
    Dim pDeviationInterval As IDeviationInterval
    Dim classbreak As Double()
    Dim pUID As UID
    Dim pClassProp As IRasterClassifyUIProperties = Nothing
    Dim fromColor As IRgbColor = New RgbColorClass()
    Dim toColor As IRgbColor = New RgbColorClass()
    Dim colorRamp As IAlgorithmicColorRamp = New AlgorithmicColorRampClass()
    Dim createColorRamp As Boolean
    Dim pFSymbol As IFillSymbol = New SimpleFillSymbol()

    Try

      While j <= range
        If pRasterHistogram.Bin(pRasBand.Statistics.Minimum + j) > currentcount Then
          currentcount += 1
          pVal(currentcount) = pRasBand.Statistics.Minimum + j
        End If
        j = j + incrementSize
      End While
      pFreq = New Integer(size - 1) {}
      For k As Integer = 0 To size - 1
        pFreq(k) = Convert.ToInt32(temp(k))
      Next
      pClassifyGen = TryCast(sd, IClassifyGEN)
      pDeviationInterval = TryCast(sd, IDeviationInterval)
      pDeviationInterval.Mean = pRasBand.Statistics.Mean
      pDeviationInterval.StandardDev = pRasBand.Statistics.StandardDeviation
      pDeviationInterval.DeviationInterval = 0.5
      ' Note the code below fails to classify (pClassifyGen.ClassBreaks is nothing) when there are less than 10 non zero frequency values
      ' It seems there are always only 4 unique values despite the values in the selected field
      Try
        pClassifyGen.Classify(pVal, pFreq, NumOfClass)
      Catch ex As Exception
        MessageBox.Show(Me, ex.Message)
      End Try
      classbreak = TryCast(pClassifyGen.ClassBreaks, [Double]())
      'FAILS HERE FOR GRID - null
      pUID = pClassifyGen.ClassID
      pClassProp = DirectCast(pClassRen, IRasterClassifyUIProperties)
      pClassProp.ClassificationMethod = pUID
      ' Set raster classify renderer
      pClassRen.ClassCount = NumOfClass
      'pClassRen.ClassField = sField;
      For i As Integer = 0 To NumOfClass
        pClassRen.Break(i) = classbreak(i)
      Next

      'Define the from and to colors for the color ramp.
      fromColor.Red = 255
      fromColor.Green = 255
      fromColor.Blue = 128
      toColor.Red = 12
      toColor.Green = 16
      toColor.Blue = 120

      'Create the color ramp.
      colorRamp.Size = NumOfClass
      colorRamp.FromColor = fromColor
      colorRamp.ToColor = toColor
      colorRamp.CreateRamp(createColorRamp)

      ' Create symbol for the classes
      ' Loop through the classes and apply the color and label
      For i As Integer = 0 To pClassRen.ClassCount - 1
        pFSymbol.Color = colorRamp.Color(i)
        pClassRen.Symbol(i) = DirectCast(pFSymbol, ISymbol)
        pClassRen.Label(i) = classbreak(i).ToString()
      Next
      pRasRen.Update()
      StandardDevGridRenderer = DirectCast(pClassRen, RasterClassifyColorRampRendererClass)

    Catch ex As Exception
      Return Nothing
    End Try
    GoTo CleanUp

CleanUp:
    NumOfClass = Nothing
    pRLayer = Nothing
    pClassRen = Nothing
    pRasRen = Nothing
    pBandCol = Nothing
    pRasBand = Nothing
    pRasterHistogram = Nothing
    temp = Nothing
    size = Nothing
    bound0 = Nothing
    bound1 = Nothing
    range = Nothing
    pVal = Nothing
    currentcount = Nothing
    incrementSize = Nothing
    j = Nothing
    pFreq = Nothing
    sd = Nothing
    pClassifyGen = Nothing
    pDeviationInterval = Nothing
    classbreak = Nothing
    pUID = Nothing
    pClassProp = Nothing
    fromColor = Nothing
    toColor = Nothing
    colorRamp = Nothing
    createColorRamp = Nothing
    pFSymbol = Nothing

    GC.WaitForPendingFinalizers()
    GC.Collect()
  End Function

  Public Function ClassifyRenderer(ByVal pRaster As IRaster) As IRasterRenderer
    Try
      'Create the classify renderer.
      Dim pClassifyRenderer As IRasterClassifyColorRampRenderer = New RasterClassifyColorRampRendererClass()
      Dim rasterRenderer As IRasterRenderer = CType(pClassifyRenderer, IRasterRenderer)

      'Set up the renderer properties.
      'Dim pRaster As IRaster = RasterDataset.CreateDefaultRaster()
      rasterRenderer.Raster = pRaster
      pClassifyRenderer.ClassCount = 10
      rasterRenderer.Update()

      'Define the from and to colors for the color ramp.
      Dim fromColor As IRgbColor = New RgbColorClass()
      fromColor.Red = 190
      fromColor.Green = 210
      fromColor.Blue = 255
      Dim toColor As IRgbColor = New RgbColorClass()
      toColor.Red = 0
      toColor.Green = 77
      toColor.Blue = 168

      'Create the color ramp.
      Dim colorRamp As IAlgorithmicColorRamp = New AlgorithmicColorRampClass()
      colorRamp.Size = 11
      colorRamp.FromColor = fromColor
      colorRamp.ToColor = toColor
      Dim createColorRamp As Boolean
      colorRamp.CreateRamp(createColorRamp)

      'Create the symbol for the classes.
      Dim fillSymbol As IFillSymbol = New SimpleFillSymbolClass()
      Dim i As Integer
      For i = 0 To pClassifyRenderer.ClassCount - 1
        fillSymbol.Color = colorRamp.Color(i)
        pClassifyRenderer.Symbol(i) = CType(fillSymbol, ISymbol)
        pClassifyRenderer.Label(i) = Convert.ToString(i)
      Next
      Return rasterRenderer
    Catch ex As Exception
      System.Diagnostics.Debug.WriteLine(ex.Message)
      Return Nothing
    End Try
  End Function


  Private Sub txtAlias_LostFocus(sender As Object, e As EventArgs) Handles txtAlias.LostFocus
    ' Validate the alias field
    txtAlias.Text = getValidAlias(txtAlias.Text, cmbLayers.Text, cmbValuesFld.Text)

    'TRIGGER THE VALIDATION OF THE ADD FACTOR BUTTON
    AddFactorButton_Status()
  End Sub

  Private Sub cboZoneLayer_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboZoneLayer.SelectedIndexChanged
    ' Add fields to the dropdown list
    Dim layer As ILayer = DirectCast(arrZoneLayers.Item(cboZoneLayer.SelectedIndex), ILayer)
    If layer Is Nothing Then Exit Sub
    cboZoneField.Items.Clear()
    cboZoneField.Text = ""
    If TryCast(arrZoneLayers.Item(cboZoneLayer.SelectedIndex), IFeatureLayer) IsNot Nothing Then
      Dim featureLayer As IFeatureLayer = DirectCast(arrZoneLayers.Item(cboZoneLayer.SelectedIndex), IFeatureLayer)
      For n As Int32 = 1 To featureLayer.FeatureClass.Fields.FieldCount - 1
        Dim fld As IField = featureLayer.FeatureClass.Fields.Field(n)
        If featureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPoint OrElse featureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon _
          OrElse featureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryLine Then
          If fld.Type = esriFieldType.esriFieldTypeInteger OrElse fld.Type = esriFieldType.esriFieldTypeSmallInteger OrElse fld.Type = esriFieldType.esriFieldTypeString Then
            cboZoneField.Items.Add(fld.Name)
          End If
        End If
      Next
    Else
      ' Get fields from the raster layer
      Dim tableFields As ITableFields = DirectCast(layer, ITableFields)
      For n As Int32 = 0 To tableFields.FieldCount - 1
        Dim fld As IField = tableFields.Field(n)
        If fld.Type = esriFieldType.esriFieldTypeInteger OrElse fld.Type = esriFieldType.esriFieldTypeSmallInteger OrElse fld.Type = esriFieldType.esriFieldTypeString Then
          If fld.Name <> "Count" Then cboZoneField.Items.Add(fld.Name)
        End If
      Next
    End If

    ' If there is only one field then select it
    If cboZoneField.Items.Count = 1 Then
      cboZoneField.SelectedItem = cboZoneField.Items(0)
    End If
  End Sub
End Class