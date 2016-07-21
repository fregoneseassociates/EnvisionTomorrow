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


Public Class frmLocalJobsHousingBalance
  Private m_PolygonLyrs As ArrayList
  Private m_TractLayer As IFeatureLayer
  Dim blnOpenForm As Boolean = True

  Private Sub frmLocalJobsHousingBalance_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
    Dim pUID As New UID
    Try
      m_frmLocalJHBalance = Nothing
      m_PolygonLyrs = Nothing
      m_TractLayer = Nothing
      blnOpenForm = False
      GC.WaitForPendingFinalizers()
      GC.Collect()

      Me.WindowState = FormWindowState.Minimized

      pUID.Value = "esriCore.SelectTool"
      If Not m_appEnvision Is Nothing Then
        m_appEnvision.CurrentTool = m_appEnvision.Document.CommandBars.Find(pUID)
      End If
      GoTo CleanUp
    Catch ex As Exception
      MessageBox.Show(Me, ex.Message, "Closing Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
      GoTo CleanUp
    End Try

CleanUp:
    pUID = Nothing
    m_frmEnvisionProjectSetup = Nothing
    GC.WaitForPendingFinalizers()
    GC.Collect()
  End Sub

  Private Sub frmLocalJobsHousingBalance_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Dim mxApplication As IMxApplication = Nothing
    Dim pMxDocument As IMxDocument = Nothing
    Dim mapCurrent As Map
    Dim intCount As Integer
    Dim pLyr As ILayer
    Dim pFeatLyr As IFeatureLayer
    Dim pDataset As IDataset
    Dim intLayer As Integer
    Dim intFeatCount As Integer = 0
    Dim pActiveView As IActiveView = Nothing

    Try
      '********************************************************************
      'Populate the combo boxes with layer information
      '********************************************************************
      If Not TypeOf m_appEnvision Is IApplication Then
        GoTo CleanUp
      End If

      pMxDocument = DirectCast(m_appEnvision.Document, IMxDocument)
      mapCurrent = CType(pMxDocument.FocusMap, Map)
      pActiveView = CType(pMxDocument.FocusMap, IActiveView)
      If mapCurrent.LayerCount = 0 Then
        MessageBox.Show(Me, "No polygon feature layers could be found in the current view document.  Please added polygon layers to view before utilizing this tool.", "No Polygon Layers Found", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1)
        GoTo CleanUp
      End If

      'BUILD LIST OF AVAILABLE FEATURE CLASSES
      m_PolygonLyrs = New ArrayList
      m_PolygonLyrs.Clear()
      Me.cmbParcelLayers.Items.Clear()
      Me.cmbEmployedRes.Items.Clear()
      Me.cmbEmployees.Items.Clear()
      Me.cmbAvgResIncome.Items.Clear()
      Me.cmbAvgWorkersIncome.Items.Clear()
      Me.cmbParcelLayers.Text = ""
      Me.cmbEmployedRes.Text = ""
      Me.cmbEmployees.Text = ""
      Me.cmbAvgResIncome.Text = ""
      Me.cmbAvgWorkersIncome.Text = ""

      For intLayer = 0 To mapCurrent.LayerCount - 1
        pLyr = CType(mapCurrent.Layer(intLayer), ILayer)
        If TypeOf mapCurrent.Layer(intLayer) Is IFeatureLayer Then
          pFeatLyr = CType(mapCurrent.Layer(intLayer), IFeatureLayer)
          pDataset = CType(pFeatLyr.FeatureClass, IDataset)
          If pFeatLyr.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon And Not (pDataset.Category.ToString.Contains("Shapefile")) Then
            Me.cmbParcelLayers.Items.Add(pFeatLyr.Name)
            m_PolygonLyrs.Add(pFeatLyr)
            intFeatCount = intFeatCount + 1
          End If
          pFeatLyr = Nothing
        End If
      Next
      If intFeatCount <= 0 Then
        MessageBox.Show(Me, "No file geodatabase polygon feature layers could be found in the current view document.  Please added polygon layers to view before utilizing this tool.", "No Polygon Layers Found", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1)
        GoTo CleanUp
      End If

      GoTo CleanUp
    Catch ex As Exception
      MessageBox.Show(Me, ex.Message, "Redevelopment Timing Calculator Form Opening Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
      GoTo CleanUp
    End Try
    GoTo CleanUp

CleanUp:
    mxApplication = Nothing
    pMxDocument = Nothing
    mapCurrent = Nothing
    intCount = Nothing
    pLyr = Nothing
    pFeatLyr = Nothing
    intLayer = Nothing
    intFeatCount = Nothing
    pActiveView = Nothing
    GC.WaitForPendingFinalizers()
    GC.Collect()
  End Sub

  Private Sub cmbParcelLayers_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbParcelLayers.SelectedIndexChanged
    Dim intCount As Integer
    Dim blnLayerFound As Boolean = False
    Dim pFeatureClass As IFeatureClass
    Dim pField As IField
    Dim intFld As Integer

    If Not blnOpenForm Then
      GoTo CleanUp
    End If
    'LOAD THE FIELDS FROM THE SELECTED TRACT LAYER
    If Me.cmbParcelLayers.Text.Length <= 0 Then
      GoTo CleanUp
    Else
      'MAKE SURE THE SELECTED LAYER IS VALID
      If Me.cmbParcelLayers.Items.Contains(Me.cmbParcelLayers.Text) Then
        blnLayerFound = True
      End If
    End If

    If m_PolygonLyrs.Count > 0 And blnLayerFound Then
      Try
        m_TractLayer = CType(m_PolygonLyrs.Item(Me.cmbParcelLayers.SelectedIndex), IFeatureLayer)
        pFeatureClass = m_TractLayer.FeatureClass
      Catch ex As Exception
        MessageBox.Show(Me, ex.Message, "Return Parcel Layer Fields Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
        GoTo CleanUp
      End Try

      Me.cmbEmployedRes.Items.Clear()
      Me.cmbEmployees.Items.Clear()
      Me.cmbAvgResIncome.Items.Clear()
      Me.cmbAvgWorkersIncome.Items.Clear()
      Me.cmbEmployedRes.Text = ""
      Me.cmbEmployees.Text = ""
      Me.cmbAvgResIncome.Text = ""
      Me.cmbAvgWorkersIncome.Text = ""
      For intFld = 0 To pFeatureClass.Fields.FieldCount - 1
        pField = pFeatureClass.Fields.Field(intFld)
        If pField.Type = esriFieldType.esriFieldTypeInteger Or pField.Type = esriFieldType.esriFieldTypeDouble Or pField.Type = esriFieldType.esriFieldTypeSmallInteger Or pField.Type = esriFieldType.esriFieldTypeSingle Then
          If Not UCase(pField.Name) = "OBJECTID" And Not UCase(pField.Name) = "SHAPE_LENGTH" And Not UCase(pField.Name) = "SHAPE_AREA" Then
            Me.cmbEmployedRes.Items.Add(pField.Name)
            Me.cmbEmployees.Items.Add(pField.Name)
            Me.cmbAvgResIncome.Items.Add(pField.Name)
            Me.cmbAvgWorkersIncome.Items.Add(pField.Name)
          End If
        End If
      Next
    End If

CleanUp:
    intCount = Nothing
    blnLayerFound = Nothing
    pFeatureClass = Nothing
    pField = Nothing
    intFld = Nothing
    GC.WaitForPendingFinalizers()
    GC.Collect()
  End Sub

  Private Sub btnRun_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRun.Click
    Dim pSFeatClass As IFeatureClass
    Dim pSTable As ITable
    Dim dblBuffDist As Double = 0
    Dim blnFailed As Boolean = False
    Dim intEmployedResFld As Integer = -1
    Dim intEmployeeFld As Integer = -1
    Dim intResSalaryFld As Integer = -1
    Dim intWorkerSalaryFld As Integer = -1
    Dim intJobWorkerBalanceFld As Integer = -1
    Dim intIncomeBalanceFld As Integer = -1
    Dim intRecCount As Integer = 0
    Dim intCount As Integer = 0

    Dim dblEmployedRes As Double = 0
    Dim dblEmployedResSum As Double = 0
    Dim dblWorker As Double = 0
    Dim dblWorkerSum As Double = 0

    Dim dblResSalary As Double = 0
    Dim dblResSalarySum As Double = 0
    Dim dblWorkerSalary As Double = 0
    Dim dblWorkerSalarySum As Double = 0

    Dim dblTotalWage As Double = 0
    Dim dblTotalIncome As Double = 0
    Dim dblAverageWageWeighted As Double = 0
    Dim dblAverageIncomeWeighted As Double = 0

    Dim dblJobWorkerBalance As Double = 0
    Dim dblIncomeBalance As Double = 0
    Dim pDef As IFeatureLayerDefinition2
    Dim pFSel As IFeatureSelection = Nothing
    Dim pFeatureCursor As IFeatureCursor = Nothing
    Dim pFeatureCursorSel As IFeatureCursor
    Dim pFeat As IFeature
    Dim strObjFld As String
    Dim pQFilter As IQueryFilter
    Dim strQString As String = Nothing
    Dim pQBLayer As IQueryByLayer
    Dim pCursor As ICursor = Nothing
    Dim pFeatSel As IFeature
    Dim dblValue As Double = 0

    'CHECK FOR TRACTING LAYER
    If m_TractLayer Is Nothing Then
      MessageBox.Show(Me, "Please select a Tracking layer.", "No Tracking Layer Selected", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1)
      GoTo CleanUp
    Else
      pSFeatClass = m_TractLayer.FeatureClass
      pSTable = CType(pSFeatClass, Table)
      pDef = m_TractLayer
      pFSel = m_TractLayer
      intRecCount = pSFeatClass.FeatureCount(Nothing)
      barStatusRun.Visible = True
      barStatusRun.Value = 0
    End If

    'CHECK BUFFER DISTANCE VALUE, IS IT A NUMBER
    If Me.tbxBufferDistance.Text.Length > 0 Then
      If Not IsNumeric(Me.tbxBufferDistance.Text) Then
        MessageBox.Show(Me, "Please provide a valid buffer distance.", "Invalid Buffer Distance", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1)
        blnFailed = True
        GoTo CleanUp
      End If
    Else
      MessageBox.Show(Me, "Please provide a valid buffer distance.", "Invalid Buffer Distance", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1)
      blnFailed = True
      GoTo CleanUp
    End If
    dblBuffDist = CDbl(Me.tbxBufferDistance.Text)

    'REVIEW INPUT FIELDS 
    If Me.cmbEmployedRes.Text.Length <= 0 Then
      MessageBox.Show(Me, "Please select a 'Population' field.", "Field Required", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1)
      blnFailed = True
      GoTo CleanUp
    Else
      intEmployedResFld = pSFeatClass.FindField(Me.cmbEmployedRes.Text)
    End If
    If Me.cmbEmployees.Text.Length <= 0 Then
      MessageBox.Show(Me, "Please select a 'Workers' field.", "Field Required", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1)
      blnFailed = True
      GoTo CleanUp
    Else
      intEmployeeFld = pSFeatClass.FindField(Me.cmbEmployees.Text)
    End If
    If Me.cmbAvgResIncome.Text.Length <= 0 Then
      MessageBox.Show(Me, "Please select a 'Avg. Income of Residents' field.", "Field Required", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1)
      blnFailed = True
      GoTo CleanUp
    Else
      intResSalaryFld = pSFeatClass.FindField(Me.cmbAvgResIncome.Text)
    End If
    If Me.cmbAvgWorkersIncome.Text.Length <= 0 Then
      MessageBox.Show(Me, "Please select a 'Avg. Income of Workers' field.", "Field Required", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1)
      blnFailed = True
      GoTo CleanUp
    Else
      intWorkerSalaryFld = pSFeatClass.FindField(Me.cmbAvgWorkersIncome.Text)
    End If

    'REVIEW OUTPUT FIELDS, ADD IF MISSING, CALCULATE ALL RECORDS TO ZERO
    Try
      If pSFeatClass.FindField("Job_Worker_Balance") <= -1 Then
        If Not AddEnvisionField(pSTable, "Job_Worker_Balance", "DOUBLE", 16, 6) Then
          GoTo CleanUp
        End If
      End If
      Me.barStatusRun.Value = 25
      Me.Refresh()
      CalcFldValues(pSTable, "Job_Worker_Balance", "0", "DOUBLE")
      Me.barStatusRun.Value = 50
      Me.Refresh()
      If pSFeatClass.FindField("Income_Balance") <= -1 Then
        If Not AddEnvisionField(pSTable, "Income_Balance", "DOUBLE", 16, 6) Then
          GoTo CleanUp
        End If
      End If
      Me.barStatusRun.Value = 75
      Me.Refresh()
      CalcFldValues(pSTable, "Income_Balance", "0", "DOUBLE")
      Me.barStatusRun.Value = 100
      Me.Refresh()
      strObjFld = pSFeatClass.Fields.Field(0).Name
    Catch ex As Exception
      MessageBox.Show(Me, "Error:" & vbNewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1)
      GoTo CleanUp
    End Try


    'RETRIEVE CURRENT RECORD COUNT
    barStatusRun.Value = 0
    intCount = 0
    Try
      pFeatureCursor = pSFeatClass.Search(Nothing, False)
      pFeat = pFeatureCursor.NextFeature
      intJobWorkerBalanceFld = pSFeatClass.FindField("Job_Worker_Balance")
      intIncomeBalanceFld = pSFeatClass.FindField("Income_Balance")
    Catch ex As Exception
      MessageBox.Show(Me, "Error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1)
      GoTo CleanUp
    End Try
    Do While Not pFeat Is Nothing
      intCount = intCount + 1
      'CLEAR SUMMARY INPUTS AND OUTPUTS
      dblEmployedResSum = 0
      dblEmployedRes = 0
      dblWorker = 0
      dblWorkerSum = 0
      dblResSalarySum = 0
      dblResSalary = 0
      dblWorkerSalarySum = 0
      dblWorkerSalary = 0
      dblTotalWage = 0
      dblTotalIncome = 0
      dblAverageWageWeighted = 0
      dblAverageIncomeWeighted = 0
      dblJobWorkerBalance = 0
      dblIncomeBalance = 0

      'QUERY THE CURRENT FEATURE AND THOSE WITH BUFFER DISTANCE
      pQFilter = New QueryFilter
      strQString = strObjFld & " = " & intCount.ToString
      pQFilter.WhereClause = strQString
      pFSel.Clear()
      pFSel.SelectFeatures(pQFilter, esriSelectionResultEnum.esriSelectionResultNew, True)
      pQBLayer = New QueryByLayer
      With pQBLayer
        .ByLayer = m_TractLayer
        .FromLayer = pFSel
        .BufferDistance = dblBuffDist
        .BufferUnits = esriUnits.esriMiles
        .ResultType = esriSelectionResultEnum.esriSelectionResultNew
        .LayerSelectionMethod = esriLayerSelectionMethod.esriLayerSelectHaveTheirCenterIn
        .UseSelectedFeatures = True
        pFSel.SelectionSet = .Select
      End With

      'CYCLE THROUGH THE SELECTED FEATURE(S) AND SUM REQUIRED VALUES
      pFSel.SelectionSet.Search(Nothing, False, pCursor)
      pFeatureCursorSel = DirectCast(pCursor, IFeatureCursor)
      If pFSel.SelectionSet.Count > 0 Then
        pFeatSel = pFeatureCursorSel.NextFeature
        Do While Not pFeatSel Is Nothing
          'EMPLOYED RESIDENTS
          Try
            dblEmployedRes = pFeatSel.Value(intEmployedResFld)
          Catch ex As Exception
            dblEmployedRes = 0
          End Try
          dblEmployedResSum = dblEmployedResSum + dblEmployedRes
          'WORKERS
          Try
            dblWorker = pFeatSel.Value(intEmployeeFld)
          Catch ex As Exception
            dblWorker = 0
          End Try
          dblWorkerSum = dblWorkerSum + dblWorker
          'AVERAGE WAGE RES
          Try
            dblResSalary = pFeatSel.Value(intResSalaryFld)
          Catch ex As Exception
            dblResSalary = 0
          End Try
          dblResSalarySum = dblResSalarySum + dblResSalary
          'AVERAGE WAGE WORKERS
          Try
            dblWorkerSalary = pFeatSel.Value(intWorkerSalaryFld)
          Catch ex As Exception
            dblWorkerSalary = 0
          End Try
          dblWorkerSalarySum = dblWorkerSalarySum + dblWorkerSalary

          dblTotalWage = dblTotalWage + (dblWorker * dblWorkerSalary)
          dblTotalIncome = dblTotalIncome + (dblEmployedRes * dblResSalary)

          'WRITE VALUES
          pFeatSel = pFeatureCursorSel.NextFeature
        Loop
        'CALCUALTE EQUATIONS
        dblJobWorkerBalance = 0
        dblIncomeBalance = 0
        dblAverageWageWeighted = dblTotalWage / dblWorkerSum
        dblAverageIncomeWeighted = dblTotalIncome / dblEmployedResSum
        Try
          dblJobWorkerBalance = (dblEmployedResSum - dblWorkerSum) / (dblEmployedResSum + dblWorkerSum)
        Catch ex As Exception
        End Try
        dblIncomeBalance = 0
        Try
          'dblIncomeBalance = ((dblResSalarySum / dblEmployedResSum) - (dblWorkerSalarySum / dblEmployedResSum)) / ((dblResSalarySum / dblEmployedResSum) + (dblWorkerSalarySum / dblEmployedResSum))
          dblIncomeBalance = (dblAverageIncomeWeighted - dblAverageWageWeighted) / (dblAverageIncomeWeighted + dblAverageWageWeighted)
        Catch ex As Exception
        End Try


        pFeat.Value(intJobWorkerBalanceFld) = dblJobWorkerBalance
        pFeat.Value(intIncomeBalanceFld) = dblIncomeBalance
        pFeat.Store()
        pFeat = pFeatureCursor.NextFeature
        Me.barStatusRun.Value = (intCount / intRecCount) * 100
        Me.Refresh()
      End If
    Loop
    pFSel.Clear()
    'APPLY THE SELECTED LEGEND
    If Me.rdbJobWorkerLegend.Checked Then
      JobWorkerLegend()
    Else
      IncomeBalanceLegend()
    End If
    MessageBox.Show(Me, "Jobs-Housing Balance processing has completed", "Processing Complete", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
    GoTo CleanUp

CleanUp:
    pSFeatClass = Nothing
    pSTable = Nothing
    dblBuffDist = Nothing
    blnFailed = Nothing
    intEmployedResFld = Nothing
    intEmployeeFld = Nothing
    intResSalaryFld = Nothing
    intWorkerSalaryFld = Nothing
    intJobWorkerBalanceFld = Nothing
    intIncomeBalanceFld = Nothing
    intRecCount = Nothing
    intCount = Nothing
    dblEmployedResSum = Nothing
    dblResSalarySum = Nothing
    dblWorkerSalarySum = Nothing
    dblJobWorkerBalance = Nothing
    dblIncomeBalance = Nothing
    pDef = Nothing
    pFSel = Nothing
    pFeatureCursor = Nothing
    pFeatureCursorSel = Nothing
    pFeat = Nothing
    strObjFld = Nothing
    pQFilter = Nothing
    strQString = Nothing
    pQBLayer = Nothing
    pCursor = Nothing
    pFeatSel = Nothing
    dblValue = Nothing
    Me.Close()
    GC.WaitForPendingFinalizers()
    GC.Collect()
  End Sub

  Private Sub tbxPlanningHorizon_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbxBufferDistance.TextChanged
    'MAKE SURE WHAT USER ENTER IS A NUMERIC VALUE
    If IsNumeric(Me.tbxBufferDistance.Text) Then
      Me.tbxBufferDistance.ForeColor = Color.Black
    Else
      Me.tbxBufferDistance.ForeColor = Color.Red
    End If
  End Sub

  Private Sub TrackingFieldCheck() Handles cmbEmployedRes.TextChanged, cmbEmployees.TextChanged, cmbAvgResIncome.TextChanged, cmbAvgWorkersIncome.TextChanged
    'REVIEW FIELD INPUTS FOR ACCURACY
    Dim pFeatClass As IFeatureClass

    'CHECK FOR TRACTING LAYER
    If m_TractLayer Is Nothing Then
      GoTo CleanUp
    Else
      pFeatClass = m_TractLayer.FeatureClass
    End If

    If Me.cmbEmployedRes.Text.Length > 0 Then
      Me.cmbEmployedRes.ForeColor = Color.Black
      If pFeatClass.FindField(Me.cmbEmployedRes.Text) <= -1 Then
        Me.cmbEmployedRes.ForeColor = Color.Red
      End If
    End If
    If Me.cmbEmployees.Text.Length > 0 Then
      Me.cmbEmployees.ForeColor = Color.Black
      If pFeatClass.FindField(Me.cmbEmployees.Text) <= -1 Then
        Me.cmbEmployees.ForeColor = Color.Red
      End If
    End If
    If Me.cmbAvgResIncome.Text.Length > 0 Then
      Me.cmbAvgResIncome.ForeColor = Color.Black
      If pFeatClass.FindField(Me.cmbAvgResIncome.Text) <= -1 Then
        Me.cmbAvgResIncome.ForeColor = Color.Red
      End If
    End If
    If Me.cmbAvgResIncome.Text.Length > 0 Then
      Me.cmbAvgResIncome.ForeColor = Color.Black
      If pFeatClass.FindField(Me.cmbAvgResIncome.Text) <= -1 Then
        Me.cmbAvgResIncome.ForeColor = Color.Red
      End If
    End If

CleanUp:
    pFeatClass = Nothing
    GC.WaitForPendingFinalizers()
    GC.Collect()
  End Sub

  '    Private Sub BufferSummary(ByVal strAttribFld As String, ByVal strSummaryFld As String, ByVal dblMiles As Double, ByVal blnInteger As Boolean)

  '        Dim pFeatClass As IFeatureClass
  '        Dim pTable As ITable
  '        Dim strFieldName As String
  '        Dim intMatrixRow As Integer
  '        Dim blnFound As Boolean = True
  '        Dim pDef As IFeatureLayerDefinition2
  '        Dim pQFilter As IQueryFilter
  '        Dim strDefExpression As String = ""
  '        Dim strQString As String = Nothing
  '        Dim pFeatureCursor As IFeatureCursor = Nothing
  '        Dim pFeatureCursorSel As IFeatureCursor
  '        Dim intTotalCount As Integer
  '        Dim pQBLayer As IQueryByLayer
  '        Dim pFSel As IFeatureSelection = Nothing
  '        Dim pFeat As IFeature
  '        Dim intObjId As Integer = 0
  '        Dim intObjFld As Integer
  '        Dim intDevTypeFld As Integer = 0
  '        Dim intAttribFld As Integer = 0
  '        Dim intSummaryFld As Integer = 0
  '        Dim strDevType As String = ""
  '        Dim dblCurrentValue As Double = 0
  '        Dim intCurrentValue As Double = 0
  '        Dim dblValue As Double = 0
  '        Dim intValue As Double = 0
  '        Dim dblSummary As Double = 0
  '        Dim dblSummaryTotal As Double = 0
  '        Dim intSummaryTotal As Double = 0
  '        Dim pCursor As ICursor = Nothing
  '        Dim pFeatSel As IFeature
  '        Dim intCount As Integer = 0
  '        Dim intFailed As Integer = 0
  '        Dim mxApplication As IMxApplication
  '        Dim pMxDoc As IMxDocument
  '        Dim pFeatLyr As IFeatureLayer = Nothing

  '        Try
  '            'CHECK AND RETRIEVE THE PROCESSING OPTION
  '            If Me.rdbNeighborhoodLayer.Checked Then
  '                pFeatLyr = m_arrPolyFeatureLayers.Item(Me.cmbLayers.FindString(Me.cmbLayers.Text))
  '                pFeatClass = pFeatLyr.FeatureClass
  '                pDef = pFeatLyr
  '                pFSel = pFeatLyr
  '            Else
  '                pFeatClass = m_pEditFeatureLyr.FeatureClass
  '                pDef = m_pEditFeatureLyr
  '                pFSel = m_pEditFeatureLyr
  '            End If
  '            strDefExpression = pDef.DefinitionExpression
  '            pQFilter = New QueryFilter
  '            If Me.rdbPartial.Checked Then
  '                strQString = "NOT DEV_TYPE = ''"
  '                If strDefExpression.Length > 0 Then
  '                    pQFilter.WhereClause = "(" & strQString & ") AND (" & pDef.DefinitionExpression & ")"
  '                    pFeatureCursor = pFeatClass.Search(pQFilter, False)
  '                Else
  '                    pQFilter.WhereClause = strQString
  '                    pFeatureCursor = pFeatClass.Search(pQFilter, False)
  '                End If
  '                intTotalCount = pFeatClass.FeatureCount(pQFilter)
  '            ElseIf Me.rdbFull.Checked Then
  '                If strDefExpression.Length > 0 Then
  '                    pQFilter.WhereClause = pDef.DefinitionExpression
  '                    pFeatureCursor = pFeatClass.Search(pQFilter, False)
  '                    intTotalCount = pFeatClass.FeatureCount(pQFilter)
  '                Else
  '                    pFeatureCursor = pFeatClass.Search(Nothing, False)
  '                    intTotalCount = pFeatClass.FeatureCount(Nothing)
  '                End If
  '            ElseIf Me.rdbSelected.Checked Then
  '                pFSel.SelectionSet.Search(Nothing, False, pCursor)
  '                pFeatureCursor = DirectCast(pCursor, IFeatureCursor)
  '                intTotalCount = pFSel.SelectionSet.Count
  '            End If

  '            pQBLayer = New QueryByLayer
  '            intObjFld = pFeatClass.FindField("OBJECTID")
  '            'intDevTypeFld = pFeatClass.FindField("DEV_TYPE")
  '            intAttribFld = pFeatClass.FindField(strAttribFld)
  '            intSummaryFld = pFeatClass.FindField(strSummaryFld)
  '            pFeat = pFeatureCursor.NextFeature
  '            Do While Not pFeat Is Nothing
  '                intCount = intCount + 1
  '                Me.barStatusRun.Value = (intCount / intTotalCount) * 100
  '                Me.Refresh()
  '                'strDevType = ""
  '                'Try
  '                '    strDevType = pFeat.Value(intDevTypeFld)
  '                'Catch ex As Exception
  '                'End Try
  '                dblCurrentValue = 0
  '                intCurrentValue = 0
  '                If blnInteger Then
  '                    Try
  '                        intCurrentValue = pFeat.Value(intAttribFld)
  '                    Catch ex As Exception
  '                    End Try
  '                Else
  '                    Try
  '                        dblCurrentValue = pFeat.Value(intAttribFld)
  '                    Catch ex As Exception
  '                    End Try
  '                End If

  '                Try
  '                    'If strDevType.Length > 0 Then
  '                    intObjId = pFeat.Value(intObjFld)
  '                    pQFilter = New QueryFilter
  '                    strQString = "OBJECTID = " & intObjId.ToString
  '                    pQFilter.WhereClause = strQString
  '                    pFSel.Clear()
  '                    pFSel.SelectFeatures(pQFilter, esriSelectionResultEnum.esriSelectionResultNew, True)
  '                    pQBLayer = New QueryByLayer
  '                    With pQBLayer
  '                        If Me.rdbNeighborhoodLayer.Checked Then
  '                            .ByLayer = pFeatLyr
  '                        Else
  '                            .ByLayer = m_pEditFeatureLyr
  '                        End If
  '                        .FromLayer = pFSel
  '                        .BufferDistance = dblMiles
  '                        .BufferUnits = esriUnits.esriMiles
  '                        .ResultType = esriSelectionResultEnum.esriSelectionResultNew
  '                        .LayerSelectionMethod = esriLayerSelectionMethod.esriLayerSelectIntersect
  '                        .UseSelectedFeatures = True
  '                        pFSel.SelectionSet = .Select
  '                    End With

  '                    pFSel.SelectionSet.Search(Nothing, False, pCursor)
  '                    pFeatureCursorSel = DirectCast(pCursor, IFeatureCursor)
  '                    dblSummaryTotal = 0
  '                    intSummaryTotal = 0
  '                    If pFSel.SelectionSet.Count > 0 Then
  '                        pFeatSel = pFeatureCursorSel.NextFeature
  '                        Do While Not pFeatSel Is Nothing
  '                            If blnInteger Then
  '                                Try
  '                                    intValue = pFeatSel.Value(intAttribFld)
  '                                Catch ex As Exception
  '                                    intValue = 0
  '                                End Try
  '                                intSummaryTotal = intSummaryTotal + intValue
  '                            Else
  '                                Try
  '                                    dblValue = pFeatSel.Value(intAttribFld)
  '                                Catch ex As Exception
  '                                    dblValue = 0
  '                                End Try
  '                                dblSummaryTotal = dblSummaryTotal + dblValue
  '                            End If
  '                            pFeatSel = pFeatureCursorSel.NextFeature
  '                        Loop
  '                    End If

  '                    If blnInteger Then
  '                        pFeat.Value(intSummaryFld) = intSummaryTotal
  '                    Else
  '                        pFeat.Value(intSummaryFld) = dblSummaryTotal
  '                    End If
  '                    pFeat.Store()
  '                    'End If
  '                Catch ex As Exception
  '                    'MessageBox.Show(me,"Error: " & ex.Message & vbNewLine & "Count: " & intCount.ToString & vbNewLine & "OBJECTID = " & intObjId.ToString, "Buffer Summary Process Error", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
  '                    intFailed = intFailed + 1
  '                End Try

  '                pFeat = pFeatureCursor.NextFeature
  '            Loop

  '            'CLEAR ANY SELECTION ON THE ENVISION EDIT LAYER AND REFRESH THE VIEW DOCUMENT
  '            pFSel.Clear()
  '            mxApplication = CType(m_appEnvision, IMxApplication)
  '            pMxDoc = CType(m_appEnvision.Document, IMxDocument)
  '            pMxDoc.ActivatedView.Refresh()

  '        Catch ex As Exception
  '            MessageBox.Show(me,"Error: " & ex.Message & vbNewLine & "Count: " & intCount.ToString & vbNewLine & "OBJECTID = " & intObjId.ToString, "Buffer Summary Process Error", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
  '            GoTo CleanUp
  '        End Try
  'CleanUp:
  '        If intFailed > 0 Then
  '            MessageBox.Show(me,intFailed.ToString & " feature(s) failed to processing for the buffering.", "Incomplete Processing", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1)
  '        End If
  '        pFeatClass = Nothing
  '        pTable = Nothing
  '        strFieldName = Nothing
  '        intMatrixRow = Nothing
  '        blnFound = Nothing
  '        pDef = Nothing
  '        pQFilter = Nothing
  '        strDefExpression = Nothing
  '        strQString = Nothing
  '        pFeatureCursor = Nothing
  '        pFeatureCursorSel = Nothing
  '        intTotalCount = Nothing
  '        pQBLayer = Nothing
  '        pFSel = Nothing
  '        pFeat = Nothing
  '        intObjId = Nothing
  '        intObjFld = Nothing
  '        intDevTypeFld = Nothing
  '        intAttribFld = Nothing
  '        intSummaryFld = Nothing
  '        strDevType = Nothing
  '        dblCurrentValue = Nothing
  '        intCurrentValue = Nothing
  '        dblValue = Nothing
  '        intValue = Nothing
  '        dblSummary = Nothing
  '        dblSummaryTotal = Nothing
  '        intSummaryTotal = Nothing
  '        pCursor = Nothing
  '        pFeatSel = Nothing
  '        intCount = Nothing
  '        intFailed = Nothing
  '        mxApplication = Nothing
  '        pMxDoc = Nothing
  '        pFeatLyr = Nothing
  '        GC.WaitForFullGCComplete()
  '        GC.Collect()
  '    End Sub

  Public Sub JobWorkerLegend()

    If m_TractLayer Is Nothing Then
      Exit Sub
    End If

    Dim mxApplication As IMxApplication
    Dim pDoc As IMxDocument
    Dim pMap As IMap
    Dim strFieldName As String = ""
    Dim pRender As IClassBreaksRenderer
    Dim symd As ISimpleFillSymbol
    Dim intCount As Integer = 0
    Dim blnDevTypFldIsString As Boolean = True
    Dim strDevType As String = ""
    Dim intDevType As Integer = 0
    Dim intRed As Integer = 0
    Dim intGreen As Integer = 0
    Dim intBlue As Integer = 0
    Dim pNewColor As IRgbColor
    Dim symx As ISimpleFillSymbol
    Dim pLyr As IGeoFeatureLayer
    Dim pOutlineSymbol As ILineSymbol
    Dim pLayerAffects As ILayerEffects
    Dim pTable As ITable

    'SET THE LEGEND FIELD
    strFieldName = "Job_Worker_Balance"

    'LOOK FOR THE DEV_TYPE FIELD FIRST...ADD IF MISSING
    If m_TractLayer.FeatureClass.FindField("Job_Worker_Balance") <= -1 Then
      pTable = CType(m_TractLayer.FeatureClass, ITable)
      If Not AddEnvisionField(pTable, "Job_Worker_Balance", "DOUBLE", 16, 6) Then
        GoTo CleanUp
      End If
    End If

    'EXIT IF FIELD NOT FOUND
    If m_TractLayer.FeatureClass.FindField(strFieldName) <= -1 Then
      GoTo CleanUp
    End If

    'DEFINE SYMBOL OUTLINE
    pOutlineSymbol = New SimpleLineSymbol
    pOutlineSymbol.Width = 0.1
    pNewColor = New RgbColor
    pNewColor.Red = 0
    pNewColor.Blue = 0
    pNewColor.Green = 0
    pOutlineSymbol.Color = pNewColor

    Try
      m_strProcessing = m_strProcessing & "Setting Legend objects" & vbNewLine
      mxApplication = CType(m_appEnvision, IMxApplication)
      pDoc = CType(m_appEnvision.Document, IMxDocument)
      pMap = pDoc.FocusMap

      m_strProcessing = m_strProcessing & "Erase Color" & vbNewLine
      pNewColor = New RgbColor
      pNewColor.Red = 225
      pNewColor.Blue = 225
      pNewColor.Green = 225
      pOutlineSymbol.Color = pNewColor
      pOutlineSymbol.Width = 0.0
      'BACKGROUND HOLLOW SYMBOL
      symd = New SimpleFillSymbol
      symd.Style = esriSimpleFillStyle.esriSFSHollow
      symd.Outline = pOutlineSymbol

      '** Make the renderer
      '** These properties should be set prior to adding values
      m_strProcessing = m_strProcessing & "Construct Renderer" & vbNewLine
      pRender = New ClassBreaksRenderer
      pRender.BreakCount = 10
      pRender.Field = strFieldName


      'pRender.BackgroundSymbol = symd
      'DEFINE MINIMUM BREAK VALUES
      pRender.Break(0) = -0.8
      pRender.Break(1) = -0.6
      pRender.Break(2) = -0.4
      pRender.Break(3) = -0.2
      pRender.Break(4) = 0.0
      pRender.Break(5) = 0.2
      pRender.Break(6) = 0.4
      pRender.Break(7) = 0.6
      pRender.Break(8) = 0.8
      pRender.Break(9) = 1.0
      'DEFINE LABELS
      pRender.Label(0) = "-1.0 to -0.8"
      pRender.Label(1) = "-0.8 to -0.6"
      pRender.Label(2) = "-0.6 to -0.4"
      pRender.Label(3) = "-0.4 to -0.2"
      pRender.Label(4) = "-0.2 to 0.0"
      pRender.Label(5) = "0.0 to 0.2"
      pRender.Label(6) = "0.2 to 0.4"
      pRender.Label(7) = "0.4 to 0.6"
      pRender.Label(8) = "0.6 to 0.8"
      pRender.Label(9) = "0.8 to 1.0"
      'DEFINE SYMBOLS
      '------------------------------------------------------------------------------------------
      'SYMBOL FOR BREAK 1
      pNewColor = New RgbColor
      pNewColor.Red = 40
      pNewColor.Green = 146
      pNewColor.Blue = 199

      symx = New SimpleFillSymbol
      symx.Style = esriSimpleFillStyle.esriSFSSolid
      symx.Outline = pOutlineSymbol
      symx.Color = pNewColor
      pRender.Symbol(0) = symx
      '------------------------------------------------------------------------------------------
      'SYMBOL FOR BREAK 2 
      pNewColor = New RgbColor
      pNewColor.Red = 104
      pNewColor.Green = 166
      pNewColor.Blue = 179

      symx = New SimpleFillSymbol
      symx.Style = esriSimpleFillStyle.esriSFSSolid
      symx.Outline = pOutlineSymbol
      symx.Color = pNewColor
      pRender.Symbol(1) = symx
      '------------------------------------------------------------------------------------------
      'SYMBOL FOR BREAK 3 
      pNewColor = New RgbColor
      pNewColor.Red = 149
      pNewColor.Green = 189
      pNewColor.Blue = 159

      symx = New SimpleFillSymbol
      symx.Style = esriSimpleFillStyle.esriSFSSolid
      symx.Outline = pOutlineSymbol
      symx.Color = pNewColor
      pRender.Symbol(2) = symx
      '------------------------------------------------------------------------------------------
      'SYMBOL FOR BREAK 4 
      pNewColor = New RgbColor
      pNewColor.Red = 191
      pNewColor.Green = 212
      pNewColor.Blue = 138

      symx = New SimpleFillSymbol
      symx.Style = esriSimpleFillStyle.esriSFSSolid
      symx.Outline = pOutlineSymbol
      symx.Color = pNewColor
      pRender.Symbol(3) = symx
      '------------------------------------------------------------------------------------------
      'SYMBOL FOR BREAK 5 
      pNewColor = New RgbColor
      pNewColor.Red = 231
      pNewColor.Green = 237
      pNewColor.Blue = 114

      symx = New SimpleFillSymbol
      symx.Style = esriSimpleFillStyle.esriSFSSolid
      symx.Outline = pOutlineSymbol
      symx.Color = pNewColor
      pRender.Symbol(4) = symx
      '------------------------------------------------------------------------------------------
      'SYMBOL FOR BREAK 6
      pNewColor = New RgbColor
      pNewColor.Red = 252
      pNewColor.Green = 228
      pNewColor.Blue = 91

      symx = New SimpleFillSymbol
      symx.Style = esriSimpleFillStyle.esriSFSSolid
      symx.Outline = pOutlineSymbol
      symx.Color = pNewColor
      pRender.Symbol(5) = symx
      '------------------------------------------------------------------------------------------
      'SYMBOL FOR BREAK 7 
      pNewColor = New RgbColor
      pNewColor.Red = 252
      pNewColor.Green = 179
      pNewColor.Blue = 68

      symx = New SimpleFillSymbol
      symx.Style = esriSimpleFillStyle.esriSFSSolid
      symx.Outline = pOutlineSymbol
      symx.Color = pNewColor
      pRender.Symbol(6) = symx
      '------------------------------------------------------------------------------------------
      'SYMBOL FOR BREAK 8 
      pNewColor = New RgbColor
      pNewColor.Red = 250
      pNewColor.Green = 133
      pNewColor.Blue = 50

      symx = New SimpleFillSymbol
      symx.Style = esriSimpleFillStyle.esriSFSSolid
      symx.Outline = pOutlineSymbol
      symx.Color = pNewColor
      pRender.Symbol(7) = symx
      '------------------------------------------------------------------------------------------
      'SYMBOL FOR BREAK 9 
      pNewColor = New RgbColor
      pNewColor.Red = 242
      pNewColor.Green = 86
      pNewColor.Blue = 34

      symx = New SimpleFillSymbol
      symx.Style = esriSimpleFillStyle.esriSFSSolid
      symx.Outline = pOutlineSymbol
      symx.Color = pNewColor
      pRender.Symbol(8) = symx
      '------------------------------------------------------------------------------------------
      'SYMBOL FOR BREAK 10 
      pNewColor = New RgbColor
      pNewColor.Red = 232
      pNewColor.Green = 16
      pNewColor.Blue = 20

      symx = New SimpleFillSymbol
      symx.Style = esriSimpleFillStyle.esriSFSSolid
      symx.Outline = pOutlineSymbol
      symx.Color = pNewColor
      pRender.Symbol(9) = symx
      '------------------------------------------------------------------------------------------

      pLyr = m_TractLayer
      pLyr.Renderer = pRender
      pLyr.DisplayField = strFieldName
      '** Set Layer Transparency
      pLayerAffects = pLyr
      pLayerAffects.Transparency = 35
      '** Refresh the TOC
      pDoc.ActiveView.ContentsChanged()
      pDoc.UpdateContents()
      '** Draw the map
      pDoc.ActiveView.Refresh()

      GoTo CleanUp

    Catch ex As Exception
      MessageBox.Show(ex.Message, "JobWorkerLegend Sub Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
      GoTo CleanUp
    End Try


CleanUp:
    mxApplication = Nothing
    pDoc = Nothing
    pMap = Nothing
    strFieldName = Nothing
    pRender = Nothing
    symd = Nothing
    intCount = Nothing
    blnDevTypFldIsString = Nothing
    strDevType = Nothing
    intDevType = Nothing
    intRed = Nothing
    intGreen = Nothing
    intBlue = Nothing
    pNewColor = Nothing
    symx = Nothing
    pLyr = Nothing
    pOutlineSymbol = Nothing
    pLayerAffects = Nothing
    GC.WaitForPendingFinalizers()
    GC.Collect()
  End Sub

  Public Sub IncomeBalanceLegend()

    If m_TractLayer Is Nothing Then
      Exit Sub
    End If

    Dim mxApplication As IMxApplication
    Dim pDoc As IMxDocument
    Dim pMap As IMap
    Dim strFieldName As String = ""
    Dim pRender As IClassBreaksRenderer
    Dim symd As ISimpleFillSymbol
    Dim intCount As Integer = 0
    Dim blnDevTypFldIsString As Boolean = True
    Dim strDevType As String = ""
    Dim intDevType As Integer = 0
    Dim intRed As Integer = 0
    Dim intGreen As Integer = 0
    Dim intBlue As Integer = 0
    Dim pNewColor As IRgbColor
    Dim symx As ISimpleFillSymbol
    Dim pLyr As IGeoFeatureLayer
    Dim pOutlineSymbol As ILineSymbol
    Dim pLayerAffects As ILayerEffects

    'SET THE LEGEND FIELD
    strFieldName = "Income_Balance"

    'EXIT IF FIELD NOT FOUND
    If m_TractLayer.FeatureClass.FindField(strFieldName) <= -1 Then
      GoTo CleanUp
    End If

    'DEFINE SYMBOL OUTLINE
    pOutlineSymbol = New SimpleLineSymbol
    pOutlineSymbol.Width = 0.1
    pNewColor = New RgbColor
    pNewColor.Red = 0
    pNewColor.Blue = 0
    pNewColor.Green = 0
    pOutlineSymbol.Color = pNewColor

    Try
      m_strProcessing = m_strProcessing & "Setting Legend objects" & vbNewLine
      mxApplication = CType(m_appEnvision, IMxApplication)
      pDoc = CType(m_appEnvision.Document, IMxDocument)
      pMap = pDoc.FocusMap

      m_strProcessing = m_strProcessing & "Erase Color" & vbNewLine
      pNewColor = New RgbColor
      pNewColor.Red = 225
      pNewColor.Blue = 225
      pNewColor.Green = 225
      pOutlineSymbol.Color = pNewColor
      pOutlineSymbol.Width = 0.0
      'BACKGROUND HOLLOW SYMBOL
      symd = New SimpleFillSymbol
      symd.Style = esriSimpleFillStyle.esriSFSHollow
      symd.Outline = pOutlineSymbol

      '** Make the renderer
      '** These properties should be set prior to adding values
      m_strProcessing = m_strProcessing & "Construct Renderer" & vbNewLine
      pRender = New ClassBreaksRenderer
      pRender.BreakCount = 10
      pRender.Field = strFieldName
      'pRender.BackgroundSymbol = symd
      'DEFINE MINIMUM BREAK VALUES
      pRender.Break(0) = -0.2
      pRender.Break(1) = -0.15
      pRender.Break(2) = -0.1
      pRender.Break(3) = -0.05
      pRender.Break(4) = 0
      pRender.Break(5) = 0.05
      pRender.Break(6) = 0.1
      pRender.Break(7) = 0.15
      pRender.Break(8) = 0.2
      pRender.Break(9) = 1
      'DEFINE LABELS
      pRender.Label(0) = "-1.00 to -0.20"
      pRender.Label(1) = "-0.20 to -0.15"
      pRender.Label(2) = "-0.15 to -0.10"
      pRender.Label(3) = "-0.10 to -0.05"
      pRender.Label(4) = "-0.05 to 0.00"
      pRender.Label(5) = "0.0 to 0.05"
      pRender.Label(6) = "0.05 to 0.10"
      pRender.Label(7) = "0.10 to 0.15"
      pRender.Label(8) = "0.15 to 0.20"
      pRender.Label(9) = "0.20 to 1.00"
      'DEFINE SYMBOLS
      '------------------------------------------------------------------------------------------
      'SYMBOL FOR BREAK 1
      pNewColor = New RgbColor
      pNewColor.Red = 40
      pNewColor.Green = 146
      pNewColor.Blue = 199

      symx = New SimpleFillSymbol
      symx.Style = esriSimpleFillStyle.esriSFSSolid
      symx.Outline = pOutlineSymbol
      symx.Color = pNewColor
      pRender.Symbol(0) = symx
      '------------------------------------------------------------------------------------------
      'SYMBOL FOR BREAK 2 
      pNewColor = New RgbColor
      pNewColor.Red = 104
      pNewColor.Green = 166
      pNewColor.Blue = 179

      symx = New SimpleFillSymbol
      symx.Style = esriSimpleFillStyle.esriSFSSolid
      symx.Outline = pOutlineSymbol
      symx.Color = pNewColor
      pRender.Symbol(1) = symx
      '------------------------------------------------------------------------------------------
      'SYMBOL FOR BREAK 3 
      pNewColor = New RgbColor
      pNewColor.Red = 149
      pNewColor.Green = 189
      pNewColor.Blue = 159

      symx = New SimpleFillSymbol
      symx.Style = esriSimpleFillStyle.esriSFSSolid
      symx.Outline = pOutlineSymbol
      symx.Color = pNewColor
      pRender.Symbol(2) = symx
      '------------------------------------------------------------------------------------------
      'SYMBOL FOR BREAK 4 
      pNewColor = New RgbColor
      pNewColor.Red = 191
      pNewColor.Green = 212
      pNewColor.Blue = 138

      symx = New SimpleFillSymbol
      symx.Style = esriSimpleFillStyle.esriSFSSolid
      symx.Outline = pOutlineSymbol
      symx.Color = pNewColor
      pRender.Symbol(3) = symx
      '------------------------------------------------------------------------------------------
      'SYMBOL FOR BREAK 5 
      pNewColor = New RgbColor
      pNewColor.Red = 231
      pNewColor.Green = 237
      pNewColor.Blue = 114

      symx = New SimpleFillSymbol
      symx.Style = esriSimpleFillStyle.esriSFSSolid
      symx.Outline = pOutlineSymbol
      symx.Color = pNewColor
      pRender.Symbol(4) = symx
      '------------------------------------------------------------------------------------------
      'SYMBOL FOR BREAK 6
      pNewColor = New RgbColor
      pNewColor.Red = 252
      pNewColor.Green = 228
      pNewColor.Blue = 91

      symx = New SimpleFillSymbol
      symx.Style = esriSimpleFillStyle.esriSFSSolid
      symx.Outline = pOutlineSymbol
      symx.Color = pNewColor
      pRender.Symbol(5) = symx
      '------------------------------------------------------------------------------------------
      'SYMBOL FOR BREAK 7 
      pNewColor = New RgbColor
      pNewColor.Red = 252
      pNewColor.Green = 179
      pNewColor.Blue = 68

      symx = New SimpleFillSymbol
      symx.Style = esriSimpleFillStyle.esriSFSSolid
      symx.Outline = pOutlineSymbol
      symx.Color = pNewColor
      pRender.Symbol(6) = symx
      '------------------------------------------------------------------------------------------
      'SYMBOL FOR BREAK 8 
      pNewColor = New RgbColor
      pNewColor.Red = 250
      pNewColor.Green = 133
      pNewColor.Blue = 50

      symx = New SimpleFillSymbol
      symx.Style = esriSimpleFillStyle.esriSFSSolid
      symx.Outline = pOutlineSymbol
      symx.Color = pNewColor
      pRender.Symbol(7) = symx
      '------------------------------------------------------------------------------------------
      'SYMBOL FOR BREAK 9 
      pNewColor = New RgbColor
      pNewColor.Red = 242
      pNewColor.Green = 86
      pNewColor.Blue = 34

      symx = New SimpleFillSymbol
      symx.Style = esriSimpleFillStyle.esriSFSSolid
      symx.Outline = pOutlineSymbol
      symx.Color = pNewColor
      pRender.Symbol(8) = symx
      '------------------------------------------------------------------------------------------
      'SYMBOL FOR BREAK 10 
      pNewColor = New RgbColor
      pNewColor.Red = 232
      pNewColor.Green = 16
      pNewColor.Blue = 20

      symx = New SimpleFillSymbol
      symx.Style = esriSimpleFillStyle.esriSFSSolid
      symx.Outline = pOutlineSymbol
      symx.Color = pNewColor
      pRender.Symbol(9) = symx
      '------------------------------------------------------------------------------------------

      pLyr = m_TractLayer
      pLyr.Renderer = pRender
      pLyr.DisplayField = strFieldName
      '** Set Layer Transparency
      pLayerAffects = pLyr
      pLayerAffects.Transparency = 35
      '** Refresh the TOC
      pDoc.ActiveView.ContentsChanged()
      pDoc.UpdateContents()
      '** Draw the map
      pDoc.ActiveView.Refresh()

      GoTo CleanUp

    Catch ex As Exception
      MessageBox.Show(Me, ex.Message, "JobWorkerLegend Sub Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
      GoTo CleanUp
    End Try


CleanUp:
    mxApplication = Nothing
    pDoc = Nothing
    pMap = Nothing
    strFieldName = Nothing
    pRender = Nothing
    symd = Nothing
    intCount = Nothing
    blnDevTypFldIsString = Nothing
    strDevType = Nothing
    intDevType = Nothing
    intRed = Nothing
    intGreen = Nothing
    intBlue = Nothing
    pNewColor = Nothing
    symx = Nothing
    pLyr = Nothing
    pOutlineSymbol = Nothing
    pLayerAffects = Nothing
    GC.WaitForPendingFinalizers()
    GC.Collect()
  End Sub

  Private Sub btnApplyLegend_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnApplyLegend.Click
    'APPLY A DEFINED LEGEND TO THE SELECTED LAYER
    If Me.rdbJobWorkerLegend.Checked Then
      JobWorkerLegend()
    Else
      IncomeBalanceLegend()
    End If
  End Sub

  Private Sub lblParcelLyr_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblParcelLyr.Click

  End Sub

  Private Sub rdbIncomeBalance_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdbIncomeBalance.CheckedChanged

  End Sub
End Class