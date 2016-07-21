Imports System
Imports System.Windows.Forms
Imports System.IO
Imports System.Drawing
Imports Microsoft.Office

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
Imports ESRI.ArcGIS.AnalysisTools

Public Class frmProximity
    Dim arrFeatLyr As ArrayList = New ArrayList
    Dim intFieldRow As Integer = 0
    Dim intStartRow As Integer = 0
    Dim dgvcBufferOption As DataGridViewComboBoxColumn
    Dim blnOpening As Boolean = True
    Dim strExEmpFld As String
    Dim strExHUFld As String
    Dim strEmpFld As String
    Dim strHUFld As String
    Dim shtProximty As Microsoft.Office.Interop.Excel.Worksheet = Nothing
    Dim m_arrFeatureLayers As ArrayList = New ArrayList


    Private Sub frmProximity_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        m_frmProximtySummary = Nothing
        arrFeatLyr = Nothing
        intFieldRow = Nothing
        intStartRow = Nothing
        dgvcBufferOption = Nothing
        blnOpening = Nothing
        strExEmpFld = Nothing
        strExHUFld = Nothing
        strEmpFld = Nothing
        strHUFld = Nothing
        shtProximty = Nothing
        m_arrFeatureLayers = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub

    Private Sub frmProximity_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        blnOpening = True
        Dim mxApplication As IMxApplication = Nothing
        Dim pMxDocument As IMxDocument = Nothing
        Dim mapCurrent As Map
        Dim uid As ESRI.ArcGIS.esriSystem.IUID = New ESRI.ArcGIS.esriSystem.UIDClass
        Dim enumLayer As ESRI.ArcGIS.Carto.IEnumLayer
        Dim pLyr As ILayer
        Dim pFeatLyr As IFeatureLayer
        Dim pActiveView As IActiveView = Nothing
        Dim dgvcLayers As DataGridViewComboBoxColumn
        Dim dgvcUnits As DataGridViewComboBoxColumn
        Dim intCount As Integer
        Dim shtProximty As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        Dim intRow As Integer
        Dim strCellValue As String

        Try

            'EXIT IF NO ENVISION LAYER
            If m_pEditFeatureLyr Is Nothing Then
                MessageBox.Show("The Envision Layer has not been defined.", "Envision Layer Not Defined", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
                GoTo CloseForm
            End If

            'EXIT IF NO ENVISION EXCEL FILE
            If m_xlPaintWB1 Is Nothing Or m_xlPaintApp Is Nothing Then
                MessageBox.Show("The Envision Excel file has not been defined.", "Envision Excel File Not Defined", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
                GoTo CloseForm
            Else
                Try
                    shtProximty = m_xlPaintWB1.Sheets("Proximity")
                Catch ex As Exception
                    MessageBox.Show("The selected Envision Excel file does not contain the worksheet, Proximity.", "Proximity Worksheet Not Found", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
                    GoTo CloseForm
                End Try
            End If

            'CLEAR EXISTING LAND USE ABBREVIATIONS
            dgvcLayers = m_frmProximtySummary.dgvProximity.Columns(1)
            dgvcLayers.Items.Clear()
            dgvcUnits = m_frmProximtySummary.dgvProximity.Columns(4)
            dgvcBufferOption = m_frmProximtySummary.dgvProximity.Columns(2)

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
            m_appEnvision.StatusBar.Message(0) = "Building list of feature layers"
            m_arrFeatureLayers = New ArrayList
            uid.Value = "{E156D7E5-22AF-11D3-9F99-00C04F6BC78E}" '= IGeoFeatureLayer
            enumLayer = mapCurrent.Layers((CType(uid, ESRI.ArcGIS.esriSystem.UID)), True) ' Explicit Cast
            enumLayer.Reset()
            pLyr = enumLayer.Next
            Do While Not (pLyr Is Nothing)
                pFeatLyr = CType(pLyr, IFeatureLayer)
                If Not pFeatLyr.Name = m_pEditFeatureLyr.Name Then
                    arrFeatLyr.Add(pFeatLyr)
                    dgvcLayers.Items.Add(pFeatLyr.Name)
                End If
                pLyr = enumLayer.Next()
            Loop
            'EXIT IF NO FEATURE LAYERS FOUND
            If arrFeatLyr.Count <= 0 Then
                MessageBox.Show("No feature layer(s) were found in the current map document.", "No Feature Layer(s) Found", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
                GoTo CloseForm
            End If

            'ADD THE 10 INPUT ROWS
            For intCount = 1 To 20
                m_frmProximtySummary.dgvProximity.Rows.Add()
                Me.dgvProximity.Rows(Me.dgvProximity.RowCount - 1).Cells(4).Value = "MILES"
                Me.dgvProximity.Rows(Me.dgvProximity.RowCount - 1).Cells(2).Value = "Custom"
            Next

            'FIND THE STARTING POINT
            For intRow = 1 To 20
                strCellValue = CStr(CType(shtProximty.Cells(intRow, 1), Microsoft.Office.Interop.Excel.Range).Value)
                If Not strCellValue Is Nothing Then
                    If strCellValue = "GIS Layer" Then
                        intFieldRow = intRow
                        intStartRow = intRow + 1
                        Exit For
                    End If
                End If
            Next

            'CYCLE THROUGH NEXT 20 ROWS FOR INPUTS
            For intRow = 1 To 20
                'CHECK FOR AN AMENITY LAYER
                strCellValue = CStr(CType(shtProximty.Cells(intRow + intFieldRow, 1), Microsoft.Office.Interop.Excel.Range).Value)
                If Not strCellValue Is Nothing Then
                    If dgvcLayers.Items.IndexOf(strCellValue) >= 0 Then
                        Me.dgvProximity.Rows(intRow - 1).Cells(1).Value = strCellValue
                    End If
                End If
                'CHECK FOR AN BUFFER OPTION
                strCellValue = CStr(CType(shtProximty.Cells(intRow + intFieldRow, 2), Microsoft.Office.Interop.Excel.Range).Value)
                If Not strCellValue Is Nothing Then
                    If strCellValue.Contains(" ") Then
                        'APPLY DISTANCE VALUE INPUT
                        If IsNumeric(strCellValue.Split(" ")(0)) Then
                            Me.dgvProximity.Rows(intRow - 1).Cells(3).Value = strCellValue.Split(" ")(0)
                        End If
                        'APPLY UNIT OF MEASURE INPUT
                        If dgvcLayers.Items.Contains(strCellValue.Split(" ")(1)) Then
                            Me.dgvProximity.Rows(intRow - 1).Cells(4).Value = strCellValue.Split(" ")(1)
                        End If
                        'SELECT CORRESPONDNIG BUFFER OPTION
                        Me.dgvProximity.Rows(intRow - 1).Cells(2).Value = "Custom"
                        If strCellValue = "0.25 MILES" Then
                            Me.dgvProximity.Rows(intRow - 1).Cells(2).Value = "1/4 MILE"
                        ElseIf strCellValue = "0.50 MILES" Then
                            Me.dgvProximity.Rows(intRow - 1).Cells(2).Value = "1/2 MILE"
                        ElseIf strCellValue = "1.00 MILES" Then
                            Me.dgvProximity.Rows(intRow - 1).Cells(2).Value = "1 MILE"
                        End If
                    End If
                End If
            Next

            'RETREIVE THE SUMMARY FIELD NAMES
            strExEmpFld = CStr(CType(shtProximty.Cells(intFieldRow, 3), Microsoft.Office.Interop.Excel.Range).Value)
            strExHUFld = CStr(CType(shtProximty.Cells(intFieldRow, 4), Microsoft.Office.Interop.Excel.Range).Value)
            If m_intEditScenario = 1 Then
                strEmpFld = CStr(CType(shtProximty.Cells(intFieldRow, 5), Microsoft.Office.Interop.Excel.Range).Value)
                strHUFld = CStr(CType(shtProximty.Cells(intFieldRow, 6), Microsoft.Office.Interop.Excel.Range).Value)
            ElseIf m_intEditScenario = 2 Then
                strEmpFld = CStr(CType(shtProximty.Cells(intFieldRow, 7), Microsoft.Office.Interop.Excel.Range).Value)
                strHUFld = CStr(CType(shtProximty.Cells(intFieldRow, 8), Microsoft.Office.Interop.Excel.Range).Value)
            ElseIf m_intEditScenario = 3 Then
                strEmpFld = CStr(CType(shtProximty.Cells(intFieldRow, 9), Microsoft.Office.Interop.Excel.Range).Value)
                strHUFld = CStr(CType(shtProximty.Cells(intFieldRow, 10), Microsoft.Office.Interop.Excel.Range).Value)
            ElseIf m_intEditScenario = 4 Then
                strEmpFld = CStr(CType(shtProximty.Cells(intFieldRow, 11), Microsoft.Office.Interop.Excel.Range).Value)
                strHUFld = CStr(CType(shtProximty.Cells(intFieldRow, 12), Microsoft.Office.Interop.Excel.Range).Value)
            ElseIf m_intEditScenario = 5 Then
                strEmpFld = CStr(CType(shtProximty.Cells(intFieldRow, 13), Microsoft.Office.Interop.Excel.Range).Value)
                strHUFld = CStr(CType(shtProximty.Cells(intFieldRow, 14), Microsoft.Office.Interop.Excel.Range).Value)
            End If

        Catch ex As System.Exception
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
        pLyr = Nothing
        pFeatLyr = Nothing
        pActiveView = Nothing
        blnOpening = False
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub

    Private Sub btnCheckAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCheckAll.Click
        'CHECK ALL LAYERS
        CheckStatus(True)
    End Sub

    Private Sub btnUncheckAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUncheckAll.Click
        'UNCHECK ALL LAYERS
        CheckStatus(False)
    End Sub

    Public Sub CheckStatus(ByVal blnCheckStatus As Boolean)
        'SET CHECK STATUS TO ALL LAYERS AS DEFINED BY INUT VARIABLE
        Dim intRow As Integer
        Me.dgvProximity.ClearSelection()
        For intRow = 0 To Me.dgvProximity.Rows.Count - 1
            If blnCheckStatus Then
                Me.dgvProximity.Rows(intRow).Cells(0).Value = 1
            Else
                Me.dgvProximity.Rows(intRow).Cells(0).Value = 0
            End If
        Next
    End Sub

    Private Sub btnRun_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRun.Click
        Dim pFeatLyr As IFeatureLayer2
        Dim dgvcLayers As DataGridViewComboBoxColumn
        Dim intRow As Integer
        Dim strLayerName As String
        Dim strBuffOption As String
        Dim strDistance As String
        Dim strUnits As String
        Dim intExcelFormulaCalc As Integer = 0

        'MAKE SURE THE ENVISION PARCEL LAYER IS AVAILABLE
        If m_pEditFeatureLyr Is Nothing Then
            MessageBox.Show("The Envision Layer has not been defined.", "Envision Layer Not Defined", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo CleanUp
        End If

        Me.barStatus.Value = 0
        Me.barStatus.Visible = True

        'MAKE SURE THE SUMMARY FIELDS EXIST
        If Me.btnSumExisting.Text = "Sum Existing Conditions - YES" Then
            If m_pEditFeatureLyr.FeatureClass.FindField(strExEmpFld) <= -1 Then
                MessageBox.Show("The existing conditions field, " & strExEmpFld & ", is missing.  Please add this field before using this option.", "Missing Required Field", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
                GoTo CleanUp
            End If
            If m_pEditFeatureLyr.FeatureClass.FindField(strExHUFld) <= -1 Then
                MessageBox.Show("The existing conditions field, " & strExHUFld & ", is missing.  Please add this field before using this option.", "Missing Required Field", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
                GoTo CleanUp
            End If
        Else
            If m_pEditFeatureLyr.FeatureClass.FindField(strEmpFld) <= -1 Then
                MessageBox.Show("The field, " & strEmpFld & ", is missing.  Please add this field before using this option.", "Missing Required Field", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
                GoTo CleanUp
            End If
            If m_pEditFeatureLyr.FeatureClass.FindField(strHUFld) <= -1 Then
                MessageBox.Show("The field, " & strHUFld & ", is missing.  Please add this field before using this option.", "Missing Required Field", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
                GoTo CleanUp
            End If
        End If

        'RETRIEVE THE LAYER COLUMNS
        dgvcLayers = m_frmProximtySummary.dgvProximity.Columns(1)
        Me.dgvProximity.EndEdit()

        'RETREIVE THE PROXIMITY WORKSHEET
        Try
            shtProximty = m_xlPaintWB1.Sheets("Proximity")
        Catch ex As Exception
            MessageBox.Show("The selected Envision Excel file does not contain the worksheet, Proximity.", "Proximity Worksheet Not Found", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo CleanUp
        End Try

        'FIRST REVIEW EACH INPUT ROW TO ENSURE THE INPUTS ARE VALID
        Try
            For intRow = 1 To 20
                If Me.dgvProximity.Rows(intRow - 1).Cells(0).Value = 1 Then
                    strLayerName = Me.dgvProximity.Rows(intRow - 1).Cells(1).Value
                    If strLayerName Is Nothing Then
                        strLayerName = ""
                    End If
                    strBuffOption = Me.dgvProximity.Rows(intRow - 1).Cells(2).Value
                    strDistance = Me.dgvProximity.Rows(intRow - 1).Cells(3).Value
                    strUnits = Me.dgvProximity.Rows(intRow - 1).Cells(4).Value

                    If Not IsNumeric(strDistance) Then
                        MessageBox.Show("The buffer distance for row " & intRow.ToString & " is not valid.", "Numerical Buffer Distance Required", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
                        GoTo CleanUp
                    End If

                    If strLayerName.Length <= 0 Then
                        MessageBox.Show("The amenity layer for row " & intRow.ToString & " has not been defined.", "Amenity Layer Required", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
                        GoTo CleanUp
                    End If
                End If
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            GoTo CleanUp
        End Try

        'DETERMINE THE CURRENT FORMULA CALC SETTING TO RESET AFTER FUNCTION EXECUTES
        If Not m_xlPaintApp Is Nothing Then
            If m_xlPaintApp.Calculation = Interop.Excel.XlCalculation.xlCalculationAutomatic Then
                intExcelFormulaCalc = 1
            ElseIf m_xlPaintApp.Calculation = Interop.Excel.XlCalculation.xlCalculationManual Then
                intExcelFormulaCalc = 2
            ElseIf m_xlPaintApp.Calculation = Interop.Excel.XlCalculation.xlCalculationSemiautomatic Then
                intExcelFormulaCalc = 3
            End If
            'SET EXCEL FORMULA CALC TO MANUAL
            m_xlPaintApp.Calculation = Interop.Excel.XlCalculation.xlCalculationManual
        End If

        Try

            'RETRIEVE SUMMARY FIELD NAMES
            If Not Me.btnSumExisting.Text = "Sum Existing Conditions - YES" Then
                strExEmpFld = ""
                strExHUFld = ""
            Else
                strExHUFld = CType(shtProximty.Cells(intFieldRow, 3), Microsoft.Office.Interop.Excel.Range).Value
                If m_pEditFeatureLyr.FeatureClass.FindField(strExHUFld) <= -1 Then
                    strExHUFld = ""
                End If
                strExEmpFld = CType(shtProximty.Cells(intFieldRow, 4), Microsoft.Office.Interop.Excel.Range).Value
                If m_pEditFeatureLyr.FeatureClass.FindField(strExEmpFld) <= -1 Then
                    strExEmpFld = ""
                End If
            End If
            If m_intEditScenario = 1 Then
                strHUFld = CType(shtProximty.Cells(intFieldRow, 5), Microsoft.Office.Interop.Excel.Range).Value
                If m_pEditFeatureLyr.FeatureClass.FindField(strHUFld) <= -1 Then
                    strHUFld = ""
                End If
            ElseIf m_intEditScenario = 2 Then
                strHUFld = CType(shtProximty.Cells(intFieldRow, 7), Microsoft.Office.Interop.Excel.Range).Value
                If m_pEditFeatureLyr.FeatureClass.FindField(strHUFld) <= -1 Then
                    strHUFld = ""
                End If
            ElseIf m_intEditScenario = 3 Then
                strHUFld = CType(shtProximty.Cells(intFieldRow, 9), Microsoft.Office.Interop.Excel.Range).Value
                If m_pEditFeatureLyr.FeatureClass.FindField(strHUFld) <= -1 Then
                    strHUFld = ""
                End If
            ElseIf m_intEditScenario = 4 Then
                strHUFld = CType(shtProximty.Cells(intFieldRow, 11), Microsoft.Office.Interop.Excel.Range).Value
                If m_pEditFeatureLyr.FeatureClass.FindField(strHUFld) <= -1 Then
                    strHUFld = ""
                End If
            ElseIf m_intEditScenario = 5 Then
                strHUFld = CType(shtProximty.Cells(intFieldRow, 13), Microsoft.Office.Interop.Excel.Range).Value
                If m_pEditFeatureLyr.FeatureClass.FindField(strHUFld) <= -1 Then
                    strHUFld = ""
                End If
            End If
            If m_intEditScenario = 1 Then
                strEmpFld = CType(shtProximty.Cells(intFieldRow, 6), Microsoft.Office.Interop.Excel.Range).Value
                If m_pEditFeatureLyr.FeatureClass.FindField(strEmpFld) <= -1 Then
                    strEmpFld = ""
                End If
            ElseIf m_intEditScenario = 2 Then
                strEmpFld = CType(shtProximty.Cells(intFieldRow, 8), Microsoft.Office.Interop.Excel.Range).Value
                If m_pEditFeatureLyr.FeatureClass.FindField(strEmpFld) <= -1 Then
                    strEmpFld = ""
                End If
            ElseIf m_intEditScenario = 3 Then
                strEmpFld = CType(shtProximty.Cells(intFieldRow, 10), Microsoft.Office.Interop.Excel.Range).Value
                If m_pEditFeatureLyr.FeatureClass.FindField(strEmpFld) <= -1 Then
                    strEmpFld = ""
                End If
            ElseIf m_intEditScenario = 4 Then
                strEmpFld = CType(shtProximty.Cells(intFieldRow, 12), Microsoft.Office.Interop.Excel.Range).Value
                If m_pEditFeatureLyr.FeatureClass.FindField(strEmpFld) <= -1 Then
                    strEmpFld = ""
                End If
            ElseIf m_intEditScenario = 5 Then
                strEmpFld = CType(shtProximty.Cells(intFieldRow, 14), Microsoft.Office.Interop.Excel.Range).Value
                If m_pEditFeatureLyr.FeatureClass.FindField(strEmpFld) <= -1 Then
                    strEmpFld = ""
                End If
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
            GoTo CleanUp
        End Try

        'CYCLE THROUGH EACH LAYER TO DETERMINE HOW MANY EMP/HU ARE WITHIN EACH FEATURE LAYER
        Try
            Me.barStatus.Value = 1
            For intRow = 1 To 20
                If Me.dgvProximity.Rows(intRow - 1).Cells(0).Value = 1 Then
                    strLayerName = Me.dgvProximity.Rows(intRow - 1).Cells(1).Value
                    If strLayerName.Length <= 0 Then
                        Continue For
                    End If
                    strBuffOption = Me.dgvProximity.Rows(intRow - 1).Cells(2).Value
                    strDistance = Me.dgvProximity.Rows(intRow - 1).Cells(3).Value
                    If Not IsNumeric(strDistance) Then
                        Continue For
                    End If
                    strUnits = Me.dgvProximity.Rows(intRow - 1).Cells(4).Value

                    If dgvcLayers.Items.IndexOf(strLayerName) >= 0 Then
                        pFeatLyr = arrFeatLyr.Item(dgvcLayers.Items.IndexOf(strLayerName))
                        ProximitySelection(intRow, pFeatLyr, m_pEditFeatureLyr, strDistance, strUnits)
                    Else
                        Continue For
                    End If
                End If
                Me.barStatus.Value = (intRow / 20) * 100
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            GoTo CleanUp
        End Try

        MessageBox.Show("Proximity calculations have completed.", "Processing complete")
        GoTo CleanUp

CleanUp:
        Me.barStatus.Value = 0
        Me.barStatus.Visible = False
        'RESET FORMULA CALC SETTING
        If Not m_xlPaintApp Is Nothing Then
            If intExcelFormulaCalc = 1 Then
                m_xlPaintApp.Calculation = Interop.Excel.XlCalculation.xlCalculationAutomatic
            ElseIf intExcelFormulaCalc = 2 Then
                m_xlPaintApp.Calculation = Interop.Excel.XlCalculation.xlCalculationManual
            ElseIf intExcelFormulaCalc = 3 Then
                m_xlPaintApp.Calculation = Interop.Excel.XlCalculation.xlCalculationSemiautomatic
            End If
            m_xlPaintWB1.Save()
        End If

        Me.Close()
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub

    Private Function ProximitySelection(ByVal intRow As Integer, ByVal pByFeatLayer As IFeatureLayer2, ByVal pSelLayer As IFeatureLayer2, ByVal strBuffDist As String, ByVal strBuffUnits As String) As Boolean

        Dim pQBLayer As ESRI.ArcGIS.Carto.IQueryByLayer
        Dim pFSelOther As IFeatureSelection = Nothing
        Dim intFeatCount As Integer
        Dim pDataStats As IDataStatistics
        Dim dblSumValue As Double = 0
        Dim dblBuffDist As Double = 0
        Dim pCursor As ICursor = Nothing
        Dim pLayer As ILayer

        'SELECT THE CORRESPONDING BUFFER FOR THE SELECTED NEIGHBORHOOD FEATURE
        Try
            dblBuffDist = CDbl(strBuffDist)
            pFSelOther = CType(pSelLayer, IFeatureSelection)
            intFeatCount = pFSelOther.SelectionSet.Count
            pQBLayer = New QueryByLayer
            With pQBLayer
                .ByLayer = pByFeatLayer
                .FromLayer = pSelLayer
                .ResultType = esriSelectionResultEnum.esriSelectionResultNew
                .BufferDistance = dblBuffDist
                .BufferUnits = ESRI.ArcGIS.esriSystem.esriUnits.esriMiles
                If strBuffUnits = "FEET" Then
                    .BufferUnits = ESRI.ArcGIS.esriSystem.esriUnits.esriFeet
                ElseIf strBuffUnits = "KILOMETERS" Then
                    .BufferUnits = ESRI.ArcGIS.esriSystem.esriUnits.esriKilometers
                ElseIf strBuffUnits = "METERS" Then
                    .BufferUnits = ESRI.ArcGIS.esriSystem.esriUnits.esriMeters
                ElseIf strBuffUnits = "YARDS" Then
                    .BufferUnits = ESRI.ArcGIS.esriSystem.esriUnits.esriYards
                End If
                If Me.itmSelectOption.Text = "INTERSECT - ON" Then
                    .LayerSelectionMethod = esriLayerSelectionMethod.esriLayerSelectIntersect
                Else
                    .LayerSelectionMethod = esriLayerSelectionMethod.esriLayerSelectHaveTheirCenterIn
                End If
                .UseSelectedFeatures = False
                pFSelOther.SelectionSet = .Select
            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message, pSelLayer.Name.ToString)
            GoTo CleanUp
        End Try

        Try
            'WRITE THE COUNT TO THE INPUT FEATURE
            intFeatCount = pFSelOther.SelectionSet.Count

            Try
                'WRITE INPUTS TO EXCEL FILE
                pLayer = pByFeatLayer
                CType(shtProximty.Cells(intRow + intFieldRow, 1), Microsoft.Office.Interop.Excel.Range).Value = pLayer.Name
                pLayer = Nothing
                CType(shtProximty.Cells(intRow + intFieldRow, 2), Microsoft.Office.Interop.Excel.Range).Value = strBuffDist & " " & strBuffUnits.ToString
                If intFeatCount > 0 Then

                Else
                    GoTo CleanUp
                End If
                'EXISTING HU SUMMARY
                If strExEmpFld.Length > 0 Then
                    If pSelLayer.FeatureClass.FindField(strExEmpFld) >= 0 Then
                        pFSelOther.SelectionSet.Search(Nothing, False, pCursor)
                        pDataStats = New DataStatistics
                        pDataStats.Cursor = pCursor
                        pDataStats.Field = strExHUFld
                        dblSumValue = pDataStats.Statistics.Sum
                        CType(shtProximty.Cells(intRow + intFieldRow, 3), Microsoft.Office.Interop.Excel.Range).Value = dblSumValue
                    End If
                End If
                'EXISTING EMP SUMMARY
                If strExEmpFld.Length > 0 Then
                    If pSelLayer.FeatureClass.FindField(strExEmpFld) >= 0 Then
                        pFSelOther.SelectionSet.Search(Nothing, False, pCursor)
                        pDataStats = New DataStatistics
                        pDataStats.Cursor = pCursor
                        pDataStats.Field = strExEmpFld
                        dblSumValue = pDataStats.Statistics.Sum
                        CType(shtProximty.Cells(intRow + intFieldRow, 4), Microsoft.Office.Interop.Excel.Range).Value = dblSumValue
                    End If
                End If
                'HU SUMMARY
                If strHUFld.Length > 0 Then
                    If pSelLayer.FeatureClass.FindField(strHUFld) >= 0 Then
                        pFSelOther.SelectionSet.Search(Nothing, False, pCursor)
                        pDataStats = New DataStatistics
                        pDataStats.Cursor = pCursor
                        pDataStats.Field = strHUFld
                        dblSumValue = pDataStats.Statistics.Sum
                        If m_intEditScenario = 1 Then
                            CType(shtProximty.Cells(intRow + intFieldRow, 5), Microsoft.Office.Interop.Excel.Range).Value = dblSumValue
                        ElseIf m_intEditScenario = 2 Then
                            CType(shtProximty.Cells(intRow + intFieldRow, 7), Microsoft.Office.Interop.Excel.Range).Value = dblSumValue
                        ElseIf m_intEditScenario = 3 Then
                            CType(shtProximty.Cells(intRow + intFieldRow, 9), Microsoft.Office.Interop.Excel.Range).Value = dblSumValue
                        ElseIf m_intEditScenario = 4 Then
                            CType(shtProximty.Cells(intRow + intFieldRow, 11), Microsoft.Office.Interop.Excel.Range).Value = dblSumValue
                        ElseIf m_intEditScenario = 5 Then
                            CType(shtProximty.Cells(intRow + intFieldRow, 13), Microsoft.Office.Interop.Excel.Range).Value = dblSumValue
                        End If
                    End If
                End If
                ' EMP SUMMARY
                If strEmpFld.Length > 0 Then
                    If pSelLayer.FeatureClass.FindField(strEmpFld) >= 0 Then
                        pFSelOther.SelectionSet.Search(Nothing, False, pCursor)
                        pDataStats = New DataStatistics
                        pDataStats.Cursor = pCursor
                        pDataStats.Field = strEmpFld
                        dblSumValue = pDataStats.Statistics.Sum
                        If m_intEditScenario = 1 Then
                            CType(shtProximty.Cells(intRow + intFieldRow, 6), Microsoft.Office.Interop.Excel.Range).Value = dblSumValue
                        ElseIf m_intEditScenario = 2 Then
                            CType(shtProximty.Cells(intRow + intFieldRow, 8), Microsoft.Office.Interop.Excel.Range).Value = dblSumValue
                        ElseIf m_intEditScenario = 3 Then
                            CType(shtProximty.Cells(intRow + intFieldRow, 10), Microsoft.Office.Interop.Excel.Range).Value = dblSumValue
                        ElseIf m_intEditScenario = 4 Then
                            CType(shtProximty.Cells(intRow + intFieldRow, 12), Microsoft.Office.Interop.Excel.Range).Value = dblSumValue
                        ElseIf m_intEditScenario = 5 Then
                            CType(shtProximty.Cells(intRow + intFieldRow, 14), Microsoft.Office.Interop.Excel.Range).Value = dblSumValue
                        End If
                    End If
                End If


            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try

            'CLEAR ANY SELECTION MADE
            pFSelOther.Clear()

        Catch ex As Exception
            GoTo CleanUp
        End Try

CleanUp:
        pByFeatLayer = Nothing
        pQBLayer = Nothing
        pFSelOther = Nothing
        pFSelOther = Nothing
        pSelLayer = Nothing
        intFeatCount = Nothing
        pDataStats = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()

    End Function


    Private Sub btnSumExisting_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSumExisting.Click
        If Me.btnSumExisting.Text = "Sum Existing Conditions - YES" Then
            Me.btnSumExisting.Text = "Sum Existing Conditions - NO"
        Else
            Me.btnSumExisting.Text = "Sum Existing Conditions - YES"
        End If
    End Sub

    Private Sub UseIntersectsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UseIntersectsToolStripMenuItem.Click
        Me.itmSelectOption.Text = "INTERSECT - ON"
    End Sub

    Private Sub UseCentroidWithinToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UseCentroidWithinToolStripMenuItem.Click
        Me.itmSelectOption.Text = "CENTROID WITHIN - ON"
    End Sub
End Class