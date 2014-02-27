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

Public Class frm7DModelCalc

    Dim arrFeatLyr As ArrayList = New ArrayList
    Dim intStartRow As Integer = -1
    Dim intScenarioRow As Integer = -1
    Dim arrConstants As ArrayList = New ArrayList
    Dim arrWriteFields As ArrayList = New ArrayList
    Dim arrGISFields As ArrayList = New ArrayList
    Dim arrVMT_Values As ArrayList = New ArrayList
    Dim arrAutoTrips_Values As ArrayList = New ArrayList
    Dim arrWalkTrips_Values As ArrayList = New ArrayList
    Dim arrBike_Values As ArrayList = New ArrayList
    Dim arrTransit_Values As ArrayList = New ArrayList
    Dim arrVMT_Prob_Values As ArrayList = New ArrayList
    Dim arrWalkTrips_Prob_Values As ArrayList = New ArrayList
    Dim arrBike_Prob_Values As ArrayList = New ArrayList
    Dim arrTransit_Prob_Values As ArrayList = New ArrayList
    Dim m_shtTravelModel As Microsoft.Office.Interop.Excel.Worksheet = Nothing
    Dim m_xlApp As Microsoft.Office.Interop.Excel.Application = Nothing
    Dim m_xlWB As Microsoft.Office.Interop.Excel.Workbook = Nothing


    Private Property strPrefix As String

    Private Sub frm7DModelCalc_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'CLOSE THE ENVISION EXCEL APPLICATION
        If Me.rdbExisting.Checked Then
            If MessageBox.Show("Would you like to close the current HH 7Ds Transportation Excel file?", "Close Excel File", MessageBoxButtons.YesNo, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False) Then
                Try
                    If Not m_xlApp Is Nothing Then
                        m_xlApp.Quit()
                        Marshal.FinalReleaseComObject(m_xlApp)
                    End If
                    m_xlApp = Nothing

                    If Not m_xlWB Is Nothing Then
                        m_xlPaintWB1.Close()
                        Marshal.FinalReleaseComObject(m_xlPaintWB1)
                    End If
                    m_xlWB = Nothing
                    GoTo CleanUp
                Catch ex As Exception
                    'MessageBox.Show("Error in closing the Envision Excel file.  You may need to terminate the application using the Task Manager." & vbNewLine & ex.Message, "Envision Excel Closing Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
                    GoTo CleanUp
                End Try
            Else
                Try
                    If Not m_xlApp Is Nothing Then
                        Marshal.FinalReleaseComObject(m_xlApp)
                    End If
                    m_xlApp = Nothing
                    If Not m_xlWB Is Nothing Then
                        Marshal.FinalReleaseComObject(m_xlPaintWB1)
                    End If
                    m_xlWB = Nothing
                    GoTo CleanUp
                Catch ex As Exception
                    'MessageBox.Show("Error in closing the Envision Excel file.  You may need to terminate the application using the Task Manager." & vbNewLine & ex.Message, "Envision Excel Closing Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
                    GoTo CleanUp
                End Try
            End If
        Else
            m_xlApp = Nothing
            m_xlWB = Nothing
            GoTo CleanUp
        End If

        
CleanUp:
        m_frm7DModelCalc = Nothing
        arrFeatLyr = Nothing
        intStartRow = Nothing
        intScenarioRow = Nothing
        arrConstants = Nothing
        arrWriteFields = Nothing
        arrGISFields = Nothing
        arrVMT_Values = Nothing
        arrAutoTrips_Values = Nothing
        arrWalkTrips_Values = Nothing
        arrBike_Values = Nothing
        arrTransit_Values = Nothing
        arrVMT_Prob_Values = Nothing
        arrWalkTrips_Prob_Values = Nothing
        arrBike_Prob_Values = Nothing
        arrTransit_Prob_Values = Nothing
        m_shtTravelModel = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub

    Private Sub frm7DModel_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'LOAD IN POLYGON LAYERS FOR USER TO SELECT FROM
        If Not Form_LoadData(sender, e) Then
            m_frm7DModelCalc.Close()
        End If
    End Sub

    Public Function Form_LoadData(ByVal sender As Object, ByVal e As System.EventArgs) As Boolean
        Form_LoadData = True
        Dim mxApplication As IMxApplication = Nothing
        Dim pMxDocument As IMxDocument = Nothing
        Dim mapCurrent As Map
        Dim intCount As Integer
        Dim pLyr As ILayer
        Dim pFeatLyr As IFeatureLayer
        Dim intLayer As Integer
        Dim intFeatCount As Integer = 0
        Dim pActiveView As IActiveView = Nothing

        '********************************************************************
        'Populate the combo boxes with layer information
        '********************************************************************
        If Not TypeOf m_appEnvision Is IApplication Then
            GoTo CloseForm
        End If

        pMxDocument = DirectCast(m_appEnvision.Document, IMxDocument)
        mapCurrent = CType(pMxDocument.FocusMap, Map)
        pActiveView = CType(pMxDocument.FocusMap, IActiveView)
        If mapCurrent.LayerCount = 0 Then
            GoTo CloseForm
        End If

        'BUILD LIST OF AVAILABLE FEATURE CLASSES
        m_frm7DModelCalc.cmbLayers.Items.Clear()
        arrFeatLyr = New ArrayList
        If mapCurrent.LayerCount > 0 Then
            For intLayer = 0 To mapCurrent.LayerCount - 1
                pLyr = CType(mapCurrent.Layer(intLayer), ILayer)
                If TypeOf mapCurrent.Layer(intLayer) Is IFeatureLayer Then
                    pFeatLyr = CType(mapCurrent.Layer(intLayer), IFeatureLayer)
                    If pFeatLyr.FeatureClass.ShapeType = ESRI.ArcGIS.Geometry.esriGeometryType.esriGeometryPolygon Then
                        arrFeatLyr.Add(pFeatLyr)
                        Me.cmbLayers.Items.Add(pFeatLyr.Name)
                        intFeatCount = intFeatCount + 1
                    End If
                    pFeatLyr = Nothing
                End If
            Next
        Else
            MessageBox.Show("Please add an input point or polygon layer in the current view document to use this tool.", "No Layer(s) Found", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo CloseForm
        End If

        'RETRIEVE OR OPEN THE SELECTED ENVISON EXCEL FILE
        Me.rdbEnvision.Enabled = True
        If m_strEnvisionExcelFile = "" Or (m_xlPaintWB1 Is Nothing Or m_xlPaintApp Is Nothing) Then
            Me.rdbEnvision.Enabled = False
            Me.rdbExisting.Checked = True
        Else
            Try
                If Not TypeOf m_xlPaintWB1.Sheets("HH Travel Model") Is Microsoft.Office.Interop.Excel.Worksheet Then
                End If
            Catch ex As Exception
                Me.rdbEnvision.Enabled = False
                Me.rdbExisting.Checked = True
            End Try
        End If

        GoTo CleanUp

CloseForm:
        Form_LoadData = False
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
    End Function

    Private Sub cmbLayers_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbLayers.SelectedIndexChanged
        Dim pFeatSelection As IFeatureSelection
        Dim pCursor As ICursor = Nothing
        Dim pFeatLyr As IFeatureLayer
        Dim pFeatureClass As IFeatureClass = Nothing
        Dim pTable As ITable
        Dim intFld As Integer = 0
        Dim strFld As String
        Dim intMissingfld As Integer
        Dim strMissingFld As String = "The following input value fields were missing and added to the selected neighborhood layer:"
        Dim arrVariables As ArrayList = New ArrayList

        'NEIGHBORHOOD LAYER SELECTED
        Try
            pFeatLyr = arrFeatLyr.Item(Me.cmbLayers.SelectedIndex)
            pFeatureClass = pFeatLyr.FeatureClass
            pTable = CType(pFeatureClass, ITable)
            pFeatSelection = CType(pFeatLyr, IFeatureSelection)
            pFeatSelection.SelectionSet.Search(Nothing, False, pCursor)
            If pFeatSelection.SelectionSet.Count > 0 Then
                Me.chkUseSelected.Enabled = True
                Me.chkUseSelected.Text = "Selected Features (" & pFeatSelection.SelectionSet.Count.ToString & ")"
            Else
                Me.chkUseSelected.Enabled = False
                Me.chkUseSelected.Text = "Selected Features"
            End If

            GoTo CleanUp
        Catch ex As Exception
            MessageBox.Show("Error occured in retrieving neighborhood layer. " & vbNewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo CleanUp
        End Try


CleanUp:
        pFeatSelection = Nothing
        pCursor = Nothing
        pFeatLyr = Nothing
        pFeatureClass = Nothing
        intFld = Nothing
        strFld = Nothing
        intMissingfld = Nothing
        strMissingFld = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub

    Private Function ReviewTravelModelInputs() As Boolean
        ReviewTravelModelInputs = True
        'REVIEW THE SELECTED ENVISION FILE TO ENSURE ALL INPUTS ARE AVAILABLE
        Dim intRow As Integer
        Dim intColumn As Integer
        Dim strCellValue As String = ""
        Dim intScenarioRow As Integer = 3
        Dim intStartRow As Integer = 9

        'RETRIEVE TRAVEL MODEL WORKSHEET
        If m_shtTravelModel Is Nothing Then
            ReviewTravelModelInputs = False
            GoTo CleanUp
        Else
            'FIND THE STARTING POINT
            For intRow = 1 To 10
                strCellValue = CStr(CType(m_shtTravelModel.Cells(intRow, 2), Microsoft.Office.Interop.Excel.Range).Value)
                If Not strCellValue Is Nothing Then
                    If Me.rdbExisting.Checked Then
                        If UCase(strCellValue) = "EXISTING" Then
                            intScenarioRow = intRow
                        End If
                    Else
                        If UCase(strCellValue) = "SCENARIO " & m_intEditScenario.ToString Then
                            intScenarioRow = intRow
                        End If
                    End If
                    If UCase(strCellValue) = "CONSTANT" Then
                        intStartRow = intRow
                        Exit For
                    End If
                End If
            Next

            'RETRIEVE AND BUILD LIST OF CONSTANT VALUES AND WRITE FIELD NAMES
            For intColumn = 4 To 12
                strCellValue = CStr(CType(m_shtTravelModel.Cells((intStartRow), intColumn), Microsoft.Office.Interop.Excel.Range).Value)
                If Not strCellValue Is Nothing Then
                    arrConstants.Add(strCellValue)
                Else
                    arrConstants.Add("0")
                End If
                strCellValue = CStr(CType(m_shtTravelModel.Cells(2, intColumn), Microsoft.Office.Interop.Excel.Range).Value)
                If Not strCellValue Is Nothing Then
                    arrWriteFields.Add(strCellValue)
                Else
                    arrWriteFields.Add(" ")
                End If
            Next


            'BUILD INPUT LISTS
            For intRow = 1 To 30
                'GIS INPUT FIELD NAME
                strCellValue = CStr(CType(m_shtTravelModel.Cells((intStartRow + intRow), 3), Microsoft.Office.Interop.Excel.Range).Value)
                If Not strCellValue Is Nothing Then
                    If strCellValue.Length > 0 Then
                        arrGISFields.Add(Trim(strCellValue))
                        'VMT PROBABILITY INPUTS
                        strCellValue = CStr(CType(m_shtTravelModel.Cells((intStartRow + intRow), 4), Microsoft.Office.Interop.Excel.Range).Value)
                        If Not strCellValue Is Nothing Then
                            arrVMT_Prob_Values.Add(strCellValue)
                        Else
                            arrVMT_Prob_Values.Add("0")
                        End If
                        'VMT INPUTS
                        strCellValue = CStr(CType(m_shtTravelModel.Cells((intStartRow + intRow), 5), Microsoft.Office.Interop.Excel.Range).Value)
                        If Not strCellValue Is Nothing Then
                            arrVMT_Values.Add(strCellValue)
                        Else
                            arrVMT_Values.Add("0")
                        End If
                        'AUTO INPUTS
                        strCellValue = CStr(CType(m_shtTravelModel.Cells((intStartRow + intRow), 6), Microsoft.Office.Interop.Excel.Range).Value)
                        If Not strCellValue Is Nothing Then
                            arrAutoTrips_Values.Add(strCellValue)
                        Else
                            arrAutoTrips_Values.Add("0")
                        End If
                        'WALK PROBABILITY INPUTS
                        strCellValue = CStr(CType(m_shtTravelModel.Cells((intStartRow + intRow), 7), Microsoft.Office.Interop.Excel.Range).Value)
                        If Not strCellValue Is Nothing Then
                            arrWalkTrips_Prob_Values.Add(strCellValue)
                        Else
                            arrWalkTrips_Prob_Values.Add("0")
                        End If
                        'WALK INPUTS
                        strCellValue = CStr(CType(m_shtTravelModel.Cells((intStartRow + intRow), 8), Microsoft.Office.Interop.Excel.Range).Value)
                        If Not strCellValue Is Nothing Then
                            arrWalkTrips_Values.Add(strCellValue)
                        Else
                            arrWalkTrips_Values.Add("0")
                        End If
                        'BIKE PROBABILITY INPUTS
                        strCellValue = CStr(CType(m_shtTravelModel.Cells((intStartRow + intRow), 9), Microsoft.Office.Interop.Excel.Range).Value)
                        If Not strCellValue Is Nothing Then
                            arrBike_Prob_Values.Add(strCellValue)
                        Else
                            arrBike_Prob_Values.Add("0")
                        End If
                        'BIKE INPUTS
                        strCellValue = CStr(CType(m_shtTravelModel.Cells((intStartRow + intRow), 10), Microsoft.Office.Interop.Excel.Range).Value)
                        If Not strCellValue Is Nothing Then
                            arrBike_Values.Add(strCellValue)
                        Else
                            arrBike_Values.Add("0")
                        End If
                        'TRANSIT PROBABILITY INPUTS
                        strCellValue = CStr(CType(m_shtTravelModel.Cells((intStartRow + intRow), 11), Microsoft.Office.Interop.Excel.Range).Value)
                        If Not strCellValue Is Nothing Then
                            arrTransit_Prob_Values.Add(strCellValue)
                        Else
                            arrTransit_Prob_Values.Add("0")
                        End If
                        'TRANSIT INPUTS
                        strCellValue = CStr(CType(m_shtTravelModel.Cells((intStartRow + intRow), 12), Microsoft.Office.Interop.Excel.Range).Value)
                        If Not strCellValue Is Nothing Then
                            arrTransit_Values.Add(strCellValue)
                        Else
                            arrTransit_Values.Add("0")
                        End If
                    Else
                        Continue For
                    End If
                Else
                    Continue For
                End If
            Next

            If Not arrGISFields.Count = arrVMT_Values.Count Or Not arrGISFields.Count = arrAutoTrips_Values.Count Or Not arrGISFields.Count = arrWalkTrips_Values.Count Or Not arrGISFields.Count = arrBike_Values.Count Or Not arrGISFields.Count = arrTransit_Values.Count Then
                MessageBox.Show("The input variables for the 7D calculations do not match.", "INVALID INPUTS", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
                ReviewTravelModelInputs = False
                GoTo CleanUp
            End If

            'EXIT IF THE SCENARIO ROW IS NOT FOUND
            If intScenarioRow = -1 Then
                MessageBox.Show("The edit scenario row (SCENARIO #) could not be fould in column B. Please select another Envision Excel file.", "Value not Found", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
                ReviewTravelModelInputs = False
                GoTo CleanUp
            End If

            GoTo CleanUp
        End If

CleanUp:
        intRow = Nothing
        intColumn = Nothing
        strCellValue = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Function

    Private Sub btnRun_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRun.Click
        'EXECUTE THE CALCULATION AND SUMMARY OF THE 7D VARIABLES
        Dim intRow As Integer = 0
        Dim intScenarioRow As Integer = 3
        Dim intColumn As Integer = 0
        Dim strCellValue As String = ""
        Dim pFeatLyr As IFeatureLayer
        Dim pFeatureClass As IFeatureClass = Nothing
        Dim pFeatSelection As IFeatureSelection
        Dim pCursor As ICursor = Nothing
        Dim pFeatureCursor As IFeatureCursor
        Dim pFeat As IFeature
        Dim intTotalCount As Integer
        Dim intRecCount As Integer = 0
        Dim strPrefix1 As String = "TOT_AVG_HH_"
        Dim strPrefix2 As String = "TOT_"
        Dim dblSumVMT As Double = 0
        Dim dblSumVehicle As Double = 0
        Dim dblSumWalk As Double = 0
        Dim dblSumBike As Double = 0
        Dim dblSumTransit As Double = 0
        Dim dblVMT As Double = 0
        Dim dblVehicle As Double = 0
        Dim dblWalk As Double = 0
        Dim dblBike As Double = 0
        Dim dblTransit As Double = 0
        Dim dblProbVMT As Double = 0
        Dim dblProbWalk As Double = 0
        Dim dblProbBike As Double = 0
        Dim dblProbTransit As Double = 0
        Dim dblHU_VMT As Double = 0
        Dim dblHU_Vehicle As Double = 0
        Dim dblHU_Walk As Double = 0
        Dim dblHU_Bike As Double = 0
        Dim dblHU_Transit As Double = 0
        Dim strFld As String = ""
        Dim arrFldValues As ArrayList = New ArrayList
        Dim strField As String
        Dim dblValue As Double
        Dim intCount As Integer
        Dim strValue As String
        Dim dblCoeff As Double
        Dim intExcelFormulaCalc As Integer
        Dim pTable As ITable
        Dim dblInputValue As Double
        Dim intHUFld As Integer = -1
        Dim dblHU As Double = 0
        Dim strReport As String = ""
        Dim datWriteLC As StreamWriter

        'DETERMINE THE SCENARIO TO BE EVALUATED
        intScenarioRow = 0
        If Me.rdbExisting.Checked Then
            intScenarioRow = 3
        Else
            If Me.cmbScenarios.Text = "Scenario 1" Then
                intScenarioRow = 4
            ElseIf Me.cmbScenarios.Text = "Scenario 2" Then
                intScenarioRow = 5
            ElseIf Me.cmbScenarios.Text = "Scenario 3" Then
                intScenarioRow = 6
            ElseIf Me.cmbScenarios.Text = "Scenario 4" Then
                intScenarioRow = 7
            ElseIf Me.cmbScenarios.Text = "Scenario 5" Then
                intScenarioRow = 8
            End If
        End If
        If intScenarioRow = 0 Then
            MessageBox.Show("Please select a scenario.", "Scenario Not Defined", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo CleanUp
        End If

        'CHECK FOR THE SELECTED NEIGHBORHOOD LAYER
        If Me.cmbLayers.Text = "" Then
            MessageBox.Show("Please select a neighbor layer.", "Layer Not Selected", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo CleanUp
        Else
            If Me.cmbLayers.Items.IndexOf(Me.cmbLayers.Text) <= -1 Then
                MessageBox.Show("Please select a neighbor layer.", "Layer Not Selected", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
                GoTo CleanUp
            End If
        End If

        'RETRIEVE OR OPEN THE SELECTED EXCEL FILE
        If m_shtTravelModel Is Nothing Then
            If Me.rdbEnvision.Checked Then
                If (m_xlPaintWB1 Is Nothing Or m_xlPaintApp Is Nothing) Then
                    MessageBox.Show("Please reload the Envision excel file or select the HH Travel Model excel file option.", "Envision Excel File Could Not be Retrieved", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
                    Me.rdbEnvision.Enabled = False
                    Me.rdbExisting.Checked = True
                    GoTo CleanUp
                Else
                    m_xlApp = m_xlPaintApp
                    m_xlWB = m_xlPaintWB1
                End If
            Else
                If Not File.Exists(Me.tbxExcel.Text) Then
                    MessageBox.Show("Please select an HH Travel Model excel file.", "Excel File Not Defined", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
                    GoTo CleanUp
                Else
                    Try
                        m_xlApp = New Microsoft.Office.Interop.Excel.Application
                        m_xlApp.WindowState = Interop.Excel.XlWindowState.xlMinimized
                        m_xlApp.Visible = True
                        m_xlWB = CType(m_xlApp.Workbooks.Open(Me.tbxExcel.Text), Microsoft.Office.Interop.Excel.Workbook)
                    Catch ex As Exception
                        MessageBox.Show(ex.Message, "Opening HH Travel Model Excel File Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
                        Me.rdbEnvision.Enabled = False
                        Me.rdbExisting.Checked = True
                    End Try
                End If
            End If

            'RETRIEVE THE HH TRAVEL MODEL WORKSHEET
            Try
                m_shtTravelModel = m_xlWB.Sheets("HH Travel Model")
            Catch ex As Exception
                MessageBox.Show("Unable to open the sheet, " & "HH Travel Model." & vbNewLine & "Error: " & ex.Message, "Excel Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
                GoTo CleanUp
            End Try

            If Not ReviewTravelModelInputs() Then
                'CLOSE THE FORM IF ALL THE ENVISION EXCEL FILE ARE NOT FOUND
                GoTo CleanUp
            End If
        End If


        'DETERMINE THE CURRENT FORMULA CALC SETTING TO RESET AFTER FUNCTION EXECUTES
        If m_xlApp.Calculation = Interop.Excel.XlCalculation.xlCalculationAutomatic Then
            intExcelFormulaCalc = 1
        ElseIf m_xlApp.Calculation = Interop.Excel.XlCalculation.xlCalculationManual Then
            intExcelFormulaCalc = 2
        ElseIf m_xlApp.Calculation = Interop.Excel.XlCalculation.xlCalculationSemiautomatic Then
            intExcelFormulaCalc = 3
        End If
        'SET EXCEL FORMULA CALC TO MANUAL
        m_xlApp.Calculation = Interop.Excel.XlCalculation.xlCalculationManual


        'SET THE GIS FIELD PREFIX
        If Me.rdbExisting.Checked Then
            strPrefix1 = "EX_PRE_AVG_HH_"
            strPrefix2 = "EX_PRE_"
        Else
            strPrefix1 = "TOT_AVG_HH_"
            strPrefix2 = "TOT_"
        End If

        'RESET THE SCENARIO OUTPUTS TO ZERO
        CType(m_shtTravelModel.Cells(intScenarioRow, 5), Microsoft.Office.Interop.Excel.Range).Value = 0
        CType(m_shtTravelModel.Cells(intScenarioRow, 6), Microsoft.Office.Interop.Excel.Range).Value = 0
        CType(m_shtTravelModel.Cells(intScenarioRow, 8), Microsoft.Office.Interop.Excel.Range).Value = 0
        CType(m_shtTravelModel.Cells(intScenarioRow, 10), Microsoft.Office.Interop.Excel.Range).Value = 0
        CType(m_shtTravelModel.Cells(intScenarioRow, 12), Microsoft.Office.Interop.Excel.Range).Value = 0
        dblSumVMT = 0
        dblSumVehicle = 0
        dblSumWalk = 0
        dblSumBike = 0
        dblSumTransit = 0

        'CHECK FOR REQUIRED OUTPUT FIELDS
        pFeatLyr = arrFeatLyr.Item(Me.cmbLayers.SelectedIndex)
        pFeatureClass = pFeatLyr.FeatureClass
        pTable = pFeatureClass
        'For Each strFld In arrWriteFields
        '    If Not strFld.Contains("_Prob") Then
        '        If pFeatureClass.FindField(strPrefix1 & strFld) = -1 Then
        '            AddEnvisionField(pTable, strPrefix1 & strFld, "DOUBLE", 16, 6)
        '        End If
        '    End If
        '    If pFeatureClass.FindField(strPrefix2 & strFld) = -1 Then
        '        AddEnvisionField(pTable, strPrefix2 & strFld, "DOUBLE", 16, 6)
        '    End If
        'Next

        If (pFeatureClass.FindField(strPrefix1 & arrWriteFields(0))) = -1 Then
            AddEnvisionField(pTable, (strPrefix1 & arrWriteFields(0)), "DOUBLE", 16, 6)
        End If
        If (pFeatureClass.FindField(strPrefix1 & arrWriteFields(1))) = -1 Then
            AddEnvisionField(pTable, (strPrefix1 & arrWriteFields(1)), "DOUBLE", 16, 6)
        End If
        If (pFeatureClass.FindField(strPrefix2 & arrWriteFields(1))) = -1 Then
            AddEnvisionField(pTable, (strPrefix2 & arrWriteFields(1)), "DOUBLE", 16, 6)
        End If
        If (pFeatureClass.FindField(strPrefix1 & arrWriteFields(2))) = -1 Then
            AddEnvisionField(pTable, (strPrefix1 & arrWriteFields(2)), "DOUBLE", 16, 6)
        End If
        If (pFeatureClass.FindField(strPrefix2 & arrWriteFields(2))) = -1 Then
            AddEnvisionField(pTable, (strPrefix2 & arrWriteFields(2)), "DOUBLE", 16, 6)
        End If
        If (pFeatureClass.FindField(strPrefix1 & arrWriteFields(3))) = -1 Then
            AddEnvisionField(pTable, (strPrefix1 & arrWriteFields(3)), "DOUBLE", 16, 6)
        End If
        If (pFeatureClass.FindField(strPrefix1 & arrWriteFields(4))) = -1 Then
            AddEnvisionField(pTable, (strPrefix1 & arrWriteFields(4)), "DOUBLE", 16, 6)
        End If
        If (pFeatureClass.FindField(strPrefix2 & arrWriteFields(4))) = -1 Then
            AddEnvisionField(pTable, (strPrefix2 & arrWriteFields(4)), "DOUBLE", 16, 6)
        End If
        If (pFeatureClass.FindField(strPrefix1 & arrWriteFields(5))) = -1 Then
            AddEnvisionField(pTable, (strPrefix1 & arrWriteFields(5)), "DOUBLE", 16, 6)
        End If
        If (pFeatureClass.FindField(strPrefix1 & arrWriteFields(6))) = -1 Then
            AddEnvisionField(pTable, (strPrefix1 & arrWriteFields(6)), "DOUBLE", 16, 6)
        End If
        If (pFeatureClass.FindField(strPrefix2 & arrWriteFields(6))) = -1 Then
            AddEnvisionField(pTable, (strPrefix2 & arrWriteFields(6)), "DOUBLE", 16, 6)
        End If
        If (pFeatureClass.FindField(strPrefix1 & arrWriteFields(7))) = -1 Then
            AddEnvisionField(pTable, (strPrefix1 & arrWriteFields(7)), "DOUBLE", 16, 6)
        End If
        If (pFeatureClass.FindField(strPrefix1 & arrWriteFields(8))) = -1 Then
            AddEnvisionField(pTable, (strPrefix1 & arrWriteFields(8)), "DOUBLE", 16, 6)
        End If
        If (pFeatureClass.FindField(strPrefix2 & arrWriteFields(8))) = -1 Then
            AddEnvisionField(pTable, (strPrefix2 & arrWriteFields(8)), "DOUBLE", 16, 6)
        End If


        'SET THE HU FIELD NUMBER
        If Me.rdbExisting.Checked Then
            intHUFld = pFeatureClass.FindField("EX_PRE_HU")
        Else
            intHUFld = pFeatureClass.FindField("TOT_HU")
        End If

        'RETRIEVE THE SELECTED LAYER AND DETERMINE IF THE FULL OR SELECTION SET WILL BE PROCESSED
        pFeatureClass = pFeatLyr.FeatureClass
        pFeatSelection = CType(pFeatLyr, IFeatureSelection)
        If Me.chkUseSelected.Checked Then
            pFeatSelection.SelectionSet.Search(Nothing, False, pCursor)
            pFeatureCursor = DirectCast(pCursor, IFeatureCursor)
            intTotalCount = pFeatSelection.SelectionSet.Count
        Else
            pFeatureCursor = pFeatureClass.Search(Nothing, False)
            intTotalCount = pFeatureClass.FeatureCount(Nothing)
        End If

        'SET THE SUMMARY VALUES
        'CYCLE THROUGH FEATURES AND CALCULATE 7D TRAVEL VARIABLES
        pFeat = pFeatureCursor.NextFeature
        Do While Not pFeat Is Nothing
            intRecCount = intRecCount + 1

            dblVMT = 0
            dblVehicle = 0
            dblWalk = 0
            dblBike = 0
            dblTransit = 0
            dblProbVMT = 0
            dblProbWalk = 0
            dblProbBike = 0
            dblProbTransit = 0
            'HU TOTALS
            dblHU = 0
            dblHU_VMT = 0
            dblHU_Vehicle = 0
            dblHU_Walk = 0
            dblHU_Bike = 0
            dblHU_Transit = 0

            arrFldValues.Clear()
            arrFldValues = New ArrayList
            'strReport = strReport & "Input Field Values: " & vbNewLine
            For Each strField In arrGISFields
                'strReport = strReport & "Field: " & (strPrefix2 & strField)
                If pFeatureClass.FindField(strPrefix2 & strField) >= 0 Then
                    Try
                        dblValue = CDbl(pFeat.Value(pFeatureClass.FindField(strPrefix2 & strField)))
                        'strReport = strReport & "Field: " & (strPrefix2 & strField) & " - Value: " & dblValue.ToString & vbNewLine
                        arrFldValues.Add(dblValue)
                    Catch ex As Exception
                        'strReport = strReport & "Field: " & (strPrefix2 & strField) & " - Value: 0" & vbNewLine
                        arrFldValues.Add(0)
                    End Try
                Else
                    'strReport = strReport & "Field: " & (strPrefix2 & strField) & " - Value: 0" & vbNewLine
                    arrFldValues.Add(0)
                End If
            Next
            'strReport = strReport & vbNewLine

            If Not arrFldValues.Count = arrGISFields.Count Then
                Try
                    'IF FIELD INPUT VALUES COUNT DOESN'T MATCH INPUT FIELD NAME LIST COUNT, THEN WRITE A ZERO TO THE FEATURE FIELDS
                    Try
                        pFeat.Value(pFeatureClass.FindField(strPrefix1 & arrWriteFields(0))) = 0
                        pFeat.Value(pFeatureClass.FindField(strPrefix1 & arrWriteFields(1))) = 0
                        pFeat.Value(pFeatureClass.FindField(strPrefix2 & arrWriteFields(1))) = 0
                    Catch ex As Exception
                        ' MessageBox.Show(ex.Message & vbNewLine & dblVMT.ToString, arrWriteFields(0))
                    End Try
                    Try
                        pFeat.Value(pFeatureClass.FindField(strPrefix1 & arrWriteFields(2))) = 0
                        pFeat.Value(pFeatureClass.FindField(strPrefix2 & arrWriteFields(2))) = 0
                    Catch ex As Exception
                        'MessageBox.Show(ex.Message & vbNewLine & dblVehicle.ToString, arrWriteFields(1))
                    End Try
                    Try
                        pFeat.Value(pFeatureClass.FindField(strPrefix1 & arrWriteFields(3))) = 0
                        pFeat.Value(pFeatureClass.FindField(strPrefix1 & arrWriteFields(4))) = 0
                        pFeat.Value(pFeatureClass.FindField(strPrefix2 & arrWriteFields(4))) = 0
                    Catch ex As Exception
                        'MessageBox.Show(ex.Message & vbNewLine & dblWalk.ToString, arrWriteFields(2))
                    End Try
                    Try
                        pFeat.Value(pFeatureClass.FindField(strPrefix1 & arrWriteFields(5))) = 0
                        pFeat.Value(pFeatureClass.FindField(strPrefix1 & arrWriteFields(6))) = 0
                        pFeat.Value(pFeatureClass.FindField(strPrefix2 & arrWriteFields(6))) = 0
                    Catch ex As Exception
                        'MessageBox.Show(ex.Message & vbNewLine & dblBike.ToString, arrWriteFields(3))
                    End Try
                    Try
                        pFeat.Value(pFeatureClass.FindField(strPrefix1 & arrWriteFields(7))) = 0
                        pFeat.Value(pFeatureClass.FindField(strPrefix1 & arrWriteFields(8))) = 0
                        pFeat.Value(pFeatureClass.FindField(strPrefix2 & arrWriteFields(8))) = 0
                    Catch ex As Exception
                        'MessageBox.Show(ex.Message & vbNewLine & dblTransit.ToString, arrWriteFields(4))
                    End Try
                    pFeat.Store()
                Catch ex As Exception
                End Try
            Else
                dblVMT = 0
                dblVehicle = 0
                dblWalk = 0
                dblBike = 0
                dblTransit = 0
                dblProbVMT = 0
                dblProbWalk = 0
                dblProbBike = 0
                dblProbTransit = 0
                dblHU = 0
                dblHU_VMT = 0
                dblHU_Vehicle = 0
                dblHU_Walk = 0
                dblHU_Bike = 0
                dblHU_Transit = 0

                'FIRST ADD IN THE CONSTANT VALUE
                'strReport = strReport & "Constant Values: " & vbNewLine

                'strReport = strReport & "Prob VMT: " & dblProbVMT.ToString & vbNewLine
                'strReport = strReport & "VMT: " & dblVMT.ToString & vbNewLine
                'strReport = strReport & "Vehicle: " & dblVehicle.ToString & vbNewLine
                'strReport = strReport & "Prob Walk: " & dblProbWalk.ToString & vbNewLine
                'strReport = strReport & "Walk: " & dblWalk.ToString & vbNewLine
                'strReport = strReport & "Prob Bike: " & dblProbBike.ToString & vbNewLine
                'strReport = strReport & "Bike: " & dblBike.ToString & vbNewLine
                'strReport = strReport & "Prob Transit: " & dblProbTransit.ToString & vbNewLine
                'strReport = strReport & "Transit: " & dblTransit.ToString & vbNewLine

                'strReport = strReport & vbNewLine


                If arrFldValues.Count > 3 Then
                    If arrFldValues.Item(0) > 0 Then 'And arrFldValues.Item(1) <= 0 And arrFldValues.Item(2) <= 0 Then
                        dblProbVMT = CDbl(arrConstants(0))
                        dblVMT = CDbl(arrConstants(1))
                        dblVehicle = CDbl(arrConstants(2))
                        dblProbWalk = CDbl(arrConstants(3))
                        dblWalk = CDbl(arrConstants(4))
                        dblProbBike = CDbl(arrConstants(5))
                        dblBike = CDbl(arrConstants(6))
                        dblProbTransit = CDbl(arrConstants(7))
                        dblTransit = CDbl(arrConstants(8))
                        For intCount = 0 To arrFldValues.Count - 1
                            'RETRIEVE THE INPUT PROCESSING VALUE
                            dblInputValue = arrFldValues(intCount)
                            'strReport = strReport & vbNewLine & vbNewLine
                            'strReport = strReport & "Input Value: " & dblInputValue.ToString & vbNewLine


                            '-------------------------------------------------------------------------------------
                            'VMT PROBABILITY
                            Try
                                strValue = "0"
                                strValue = CStr(arrVMT_Prob_Values(intCount))
                                If strValue.Contains("*") Then
                                    strValue = strValue.Replace("*", "0")
                                    If strValue.Contains("+") And dblInputValue > 0 Then
                                        strValue = strValue.Replace("+", "0")
                                        dblCoeff = CDbl(strValue)
                                        dblProbVMT = dblProbVMT + ((dblCoeff * Math.Log(dblInputValue)) / 1000)
                                        'strReport = strReport & "VMT PROBABILITY with Log and /1000: " & ((dblCoeff * Math.Log(dblInputValue)) / 1000).ToString & "      Coeff: " & dblCoeff.ToString & vbNewLine
                                    Else
                                        dblCoeff = CDbl(strValue)
                                        dblProbVMT = dblProbVMT + (dblCoeff * Math.Log(dblInputValue))
                                        'strReport = strReport & "VMT PROBABILITY with Log only: " & (dblCoeff * Math.Log(dblInputValue)).ToString & "      Coeff: " & dblCoeff.ToString & vbNewLine
                                    End If
                                Else
                                    If strValue.Contains("+") Then
                                        strValue = strValue.Replace("+", "0")
                                        dblCoeff = CDbl(strValue)
                                        dblProbVMT = dblProbVMT + ((dblCoeff * dblInputValue) / 1000)
                                        'strReport = strReport & "VMT PROBABILITY with /1000 only: " & ((dblCoeff * dblInputValue) / 1000).ToString & "      Coeff: " & dblCoeff.ToString & vbNewLine
                                    Else
                                        dblCoeff = CDbl(strValue)
                                        dblProbVMT = dblProbVMT + (dblCoeff * dblInputValue)
                                        'strReport = strReport & "VMT PROBABILITY 'NO' log or /1000: " & (dblCoeff * dblInputValue).ToString & "      Coeff: " & dblCoeff.ToString & vbNewLine
                                    End If
                                End If
                            Catch ex As Exception
                                MessageBox.Show(ex.Message, "VMT Prob")
                            End Try
                            '-------------------------------------------------------------------------------------
                            'VMT 
                            Try
                                strValue = "0"
                                strValue = CStr(arrVMT_Values(intCount))
                                If strValue.Contains("*") Then
                                    strValue = strValue.Replace("*", "0")
                                    If strValue.Contains("+") And dblInputValue > 0 Then
                                        strValue = strValue.Replace("+", "0")
                                        dblCoeff = CDbl(strValue)
                                        dblVMT = dblVMT + ((dblCoeff * Math.Log(dblInputValue)) / 1000)
                                        'strReport = strReport & "VMT with Log and /1000: " & ((dblCoeff * Math.Log(dblInputValue)) / 1000).ToString & "      Coeff: " & dblCoeff.ToString & vbNewLine
                                    Else
                                        dblCoeff = CDbl(strValue)
                                        dblVMT = dblVMT + (dblCoeff * Math.Log(dblInputValue))
                                        'strReport = strReport & "VMT with Log only: " & (dblCoeff * Math.Log(dblInputValue)).ToString & "      Coeff: " & dblCoeff.ToString & vbNewLine
                                    End If
                                Else
                                    If strValue.Contains("+") Then
                                        strValue = strValue.Replace("+", "0")
                                        dblCoeff = CDbl(strValue)
                                        dblVMT = dblVMT + ((dblCoeff * dblInputValue) / 1000)
                                        'strReport = strReport & "VMT with /1000 only: " & ((dblCoeff * dblInputValue) / 1000).ToString & "      Coeff: " & dblCoeff.ToString & vbNewLine
                                    Else
                                        dblCoeff = CDbl(strValue)
                                        dblVMT = dblVMT + (dblCoeff * dblInputValue)
                                        'strReport = strReport & "VMT 'NO' log or /1000: " & (dblCoeff * dblInputValue).ToString & "      Coeff: " & dblCoeff.ToString & vbNewLine
                                    End If
                                End If
                            Catch ex As Exception
                                MessageBox.Show(ex.Message, "VMT")
                            End Try
                            '-------------------------------------------------------------------------------------
                            'VEHICLE 
                            Try
                                strValue = "0"
                                strValue = CStr(arrAutoTrips_Values(intCount))
                                If strValue.Contains("*") Then
                                    strValue = strValue.Replace("*", "0")
                                    If strValue.Contains("+") And dblInputValue > 0 Then
                                        strValue = strValue.Replace("+", "0")
                                        dblCoeff = CDbl(strValue)
                                        dblVehicle = dblVehicle + ((dblCoeff * Math.Log(dblInputValue)) / 1000)
                                        'strReport = strReport & "VEHICLE with Log and /1000: " & ((dblCoeff * Math.Log(dblInputValue)) / 1000).ToString & "      Coeff: " & dblCoeff.ToString & vbNewLine
                                    Else
                                        dblCoeff = CDbl(strValue)
                                        dblVehicle = dblVehicle + (dblCoeff * Math.Log(dblInputValue))
                                        'strReport = strReport & "VEHICLE with Log only: " & (dblCoeff * Math.Log(dblInputValue)).ToString & "      Coeff: " & dblCoeff.ToString & vbNewLine
                                    End If
                                Else
                                    If strValue.Contains("+") Then
                                        strValue = strValue.Replace("+", "0")
                                        dblCoeff = CDbl(strValue)
                                        dblVehicle = dblVehicle + ((dblCoeff * dblInputValue) / 1000)
                                        'strReport = strReport & "VEHICLE with /1000 only: " & ((dblCoeff * dblInputValue) / 1000).ToString & "      Coeff: " & dblCoeff.ToString & vbNewLine
                                    Else
                                        dblCoeff = CDbl(strValue)
                                        dblVehicle = dblVehicle + (dblCoeff * dblInputValue)
                                        'strReport = strReport & "VEHICLE 'NO' log or /1000: " & (dblCoeff * dblInputValue).ToString & "      Coeff: " & dblCoeff.ToString & vbNewLine
                                    End If
                                End If
                            Catch ex As Exception
                                MessageBox.Show(ex.Message, "Vehicle")
                            End Try
                            '-------------------------------------------------------------------------------------
                            'WALK PROBABILITY
                            Try
                                strValue = "0"
                                strValue = CStr(arrWalkTrips_Prob_Values(intCount))
                                If strValue.Contains("*") Then
                                    strValue = strValue.Replace("*", "0")
                                    If strValue.Contains("+") And dblInputValue > 0 Then
                                        strValue = strValue.Replace("+", "0")
                                        dblCoeff = CDbl(strValue)
                                        dblProbWalk = dblProbWalk + ((dblCoeff * Math.Log(dblInputValue)) / 1000)
                                        'strReport = strReport & "WALK PROBABILITY with Log and /1000: " & ((dblCoeff * Math.Log(dblInputValue)) / 1000).ToString & "      Coeff: " & dblCoeff.ToString & vbNewLine
                                    Else
                                        dblCoeff = CDbl(strValue)
                                        dblProbWalk = dblProbWalk + (dblCoeff * Math.Log(dblInputValue))
                                        'strReport = strReport & "WALK PROBABILITY with Log only: " & (dblCoeff * Math.Log(dblInputValue)).ToString & "      Coeff: " & dblCoeff.ToString & vbNewLine
                                    End If
                                Else
                                    If strValue.Contains("+") Then
                                        strValue = strValue.Replace("+", "0")
                                        dblCoeff = CDbl(strValue)
                                        dblProbWalk = dblProbWalk + ((dblCoeff * dblInputValue) / 1000)
                                        'strReport = strReport & "WALK PROBABILITY with /1000 only: " & ((dblCoeff * dblInputValue) / 1000).ToString & "      Coeff: " & dblCoeff.ToString & vbNewLine
                                    Else
                                        dblCoeff = CDbl(strValue)
                                        dblProbWalk = dblProbWalk + (dblCoeff * dblInputValue)
                                        ' strReport = strReport & "WALK PROBABILITY 'NO' log or /1000: " & (dblCoeff * dblInputValue).ToString & "      Coeff: " & dblCoeff.ToString & vbNewLine
                                    End If
                                End If
                            Catch ex As Exception
                                MessageBox.Show(ex.Message, "Walk Prob")
                            End Try
                            '-------------------------------------------------------------------------------------
                            'WALK 
                            Try
                                strValue = "0"
                                strValue = CStr(arrWalkTrips_Values(intCount))
                                If strValue.Contains("*") Then
                                    strValue = strValue.Replace("*", "0")
                                    If strValue.Contains("+") And dblInputValue > 0 Then
                                        strValue = strValue.Replace("+", "0")
                                        dblCoeff = CDbl(strValue)
                                        dblWalk = dblWalk + ((dblCoeff * Math.Log(dblInputValue)) / 1000)
                                        'strReport = strReport & "WALK with Log and /1000: " & ((dblCoeff * Math.Log(dblInputValue)) / 1000).ToString & "      Coeff: " & dblCoeff.ToString & vbNewLine
                                    Else
                                        dblCoeff = CDbl(strValue)
                                        dblWalk = dblWalk + (dblCoeff * Math.Log(dblInputValue))
                                        'strReport = strReport & "WALK with Log only: " & (dblCoeff * Math.Log(dblInputValue)).ToString & "      Coeff: " & dblCoeff.ToString & vbNewLine
                                    End If
                                Else
                                    If strValue.Contains("+") Then
                                        strValue = strValue.Replace("+", "0")
                                        dblCoeff = CDbl(strValue)
                                        dblWalk = dblWalk + ((dblCoeff * dblInputValue) / 1000)
                                        'strReport = strReport & "WALK with /1000 only: " & ((dblCoeff * dblInputValue) / 1000).ToString & "      Coeff: " & dblCoeff.ToString & vbNewLine
                                    Else
                                        dblCoeff = CDbl(strValue)
                                        dblWalk = dblWalk + (dblCoeff * dblInputValue)
                                        'strReport = strReport & "WALK 'NO' log or /1000: " & (dblCoeff * dblInputValue).ToString & "      Coeff: " & dblCoeff.ToString & vbNewLine
                                    End If
                                End If
                            Catch ex As Exception
                                MessageBox.Show(ex.Message, "Walk")
                            End Try
                            '-------------------------------------------------------------------------------------
                            'BIKE PROBABILITY
                            Try
                                strValue = "0"
                                strValue = CStr(arrBike_Prob_Values(intCount))
                                If strValue.Contains("*") Then
                                    strValue = strValue.Replace("*", "0")
                                    If strValue.Contains("+") And dblInputValue > 0 Then
                                        strValue = strValue.Replace("+", "0")
                                        dblCoeff = CDbl(strValue)
                                        dblProbBike = dblProbBike + ((dblCoeff * Math.Log(dblInputValue)) / 1000)
                                        'strReport = strReport & "BIKE PROBABILITY with Log and /1000: " & ((dblCoeff * Math.Log(dblInputValue)) / 1000).ToString & "      Coeff: " & dblCoeff.ToString & vbNewLine
                                    Else
                                        dblCoeff = CDbl(strValue)
                                        dblProbBike = dblProbBike + (dblCoeff * Math.Log(dblInputValue))
                                        'strReport = strReport & "BIKE PROBABILITY with Log only: " & (dblCoeff * Math.Log(dblInputValue)).ToString & "      Coeff: " & dblCoeff.ToString & vbNewLine
                                    End If
                                Else
                                    If strValue.Contains("+") Then
                                        strValue = strValue.Replace("+", "0")
                                        dblCoeff = CDbl(strValue)
                                        dblProbBike = dblProbBike + ((dblCoeff * dblInputValue) / 1000)
                                        'strReport = strReport & "BIKE PROBABILITY with /1000 only: " & ((dblCoeff * dblInputValue) / 1000).ToString & "      Coeff: " & dblCoeff.ToString & vbNewLine
                                    Else
                                        dblCoeff = CDbl(strValue)
                                        dblProbBike = dblProbBike + (dblCoeff * dblInputValue)
                                        'strReport = strReport & "BIKE PROBABILITY 'NO' log or /1000: " & (dblCoeff * dblInputValue).ToString & "      Coeff: " & dblCoeff.ToString & vbNewLine
                                    End If
                                End If
                            Catch ex As Exception
                                MessageBox.Show(ex.Message, "Bike Prob")
                            End Try
                            '-------------------------------------------------------------------------------------
                            'BIKE
                            Try
                                strValue = "0"
                                strValue = CStr(arrBike_Values(intCount))
                                If strValue.Contains("*") Then
                                    strValue = strValue.Replace("*", "0")
                                    If strValue.Contains("+") And dblInputValue > 0 Then
                                        strValue = strValue.Replace("+", "0")
                                        dblCoeff = CDbl(strValue)
                                        dblBike = dblBike + ((dblCoeff * Math.Log(dblInputValue)) / 1000)
                                        'strReport = strReport & "BIKE with Log and /1000: " & ((dblCoeff * Math.Log(dblInputValue)) / 1000).ToString & "      Coeff: " & dblCoeff.ToString & vbNewLine
                                    Else
                                        dblCoeff = CDbl(strValue)
                                        dblBike = dblBike + (dblCoeff * Math.Log(dblInputValue))
                                        'strReport = strReport & "BIKE with Log only: " & (dblCoeff * Math.Log(dblInputValue)).ToString & "      Coeff: " & dblCoeff.ToString & vbNewLine
                                    End If
                                Else
                                    If strValue.Contains("+") Then
                                        strValue = strValue.Replace("+", "0")
                                        dblCoeff = CDbl(strValue)
                                        dblBike = dblBike + ((dblCoeff * dblInputValue) / 1000)
                                        'strReport = strReport & "BIKE with /1000 only: " & ((dblCoeff * dblInputValue) / 1000).ToString & "      Coeff: " & dblCoeff.ToString & vbNewLine
                                    Else
                                        dblCoeff = CDbl(strValue)
                                        dblBike = dblBike + (dblCoeff * dblInputValue)
                                        'strReport = strReport & "BIKE 'NO' log or /1000: " & (dblCoeff * dblInputValue).ToString & "      Coeff: " & dblCoeff.ToString & vbNewLine
                                    End If
                                End If
                            Catch ex As Exception
                                MessageBox.Show(ex.Message, "Bike")
                            End Try
                            '-------------------------------------------------------------------------------------
                            'TRANSIT PROBABILITY
                            Try
                                strValue = "0"
                                strValue = CStr(arrTransit_Prob_Values(intCount))
                                If strValue.Contains("*") Then
                                    strValue = strValue.Replace("*", "0")
                                    If strValue.Contains("+") And dblInputValue > 0 Then
                                        strValue = strValue.Replace("+", "0")
                                        dblCoeff = CDbl(strValue)
                                        dblProbTransit = dblProbTransit + ((dblCoeff * Math.Log(dblInputValue)) / 1000)
                                        'strReport = strReport & "TRANSIT PROBABILITY with Log and /1000: " & ((dblCoeff * Math.Log(dblInputValue)) / 1000).ToString & "      Coeff: " & dblCoeff.ToString & vbNewLine
                                    Else
                                        dblCoeff = CDbl(strValue)
                                        dblProbTransit = dblProbTransit + (dblCoeff * Math.Log(dblInputValue))
                                        'strReport = strReport & "TRANSIT PROBABILITY with Log only: " & (dblCoeff * Math.Log(dblInputValue)).ToString & "      Coeff: " & dblCoeff.ToString & vbNewLine
                                    End If
                                Else
                                    If strValue.Contains("+") Then
                                        strValue = strValue.Replace("+", "0")
                                        dblCoeff = CDbl(strValue)
                                        dblProbTransit = dblProbTransit + ((dblCoeff * dblInputValue) / 1000)
                                        'strReport = strReport & "TRANSIT PROBABILITY with /1000 only: " & ((dblCoeff * dblInputValue) / 1000).ToString & "      Coeff: " & dblCoeff.ToString & vbNewLine
                                    Else
                                        dblCoeff = CDbl(strValue)
                                        dblProbTransit = dblProbTransit + (dblCoeff * dblInputValue)
                                        'strReport = strReport & "TRANSIT PROBABILITY 'NO' log or /1000: " & (dblCoeff * dblInputValue).ToString & "      Coeff: " & dblCoeff.ToString & vbNewLine
                                    End If
                                End If
                            Catch ex As Exception
                                MessageBox.Show(ex.Message, "Transit Prob")
                            End Try
                            '-------------------------------------------------------------------------------------
                            'TRANSIT
                            Try
                                strValue = "0"
                                strValue = CStr(arrTransit_Values(intCount))
                                If strValue.Contains("*") Then
                                    strValue = strValue.Replace("*", "0")
                                    If strValue.Contains("+") And dblInputValue > 0 Then
                                        strValue = strValue.Replace("+", "0")
                                        dblCoeff = CDbl(strValue)
                                        dblTransit = dblTransit + ((dblCoeff * Math.Log(dblInputValue)) / 1000)
                                        'strReport = strReport & "TRANSIT with Log and /1000: " & ((dblCoeff * Math.Log(dblInputValue)) / 1000).ToString & "      Coeff: " & dblCoeff.ToString & vbNewLine
                                    Else
                                        dblCoeff = CDbl(strValue)
                                        dblTransit = dblTransit + (dblCoeff * Math.Log(dblInputValue))
                                        'strReport = strReport & "TRANSIT with Log only: " & (dblCoeff * Math.Log(dblInputValue)).ToString & "      Coeff: " & dblCoeff.ToString & vbNewLine
                                    End If
                                Else
                                    If strValue.Contains("+") Then
                                        strValue = strValue.Replace("+", "0")
                                        dblCoeff = CDbl(strValue)
                                        dblTransit = dblTransit + ((dblCoeff * dblInputValue) / 1000)
                                        'strReport = strReport & "TRANSIT with /1000 only: " & ((dblCoeff * dblInputValue) / 1000).ToString & "      Coeff: " & dblCoeff.ToString & vbNewLine
                                    Else
                                        dblCoeff = CDbl(strValue)
                                        dblTransit = dblTransit + (dblCoeff * dblInputValue)
                                        'strReport = strReport & "TRANSIT 'NO' log or /1000: " & (dblCoeff * dblInputValue).ToString & "      Coeff: " & dblCoeff.ToString & vbNewLine
                                    End If
                                End If
                            Catch ex As Exception
                                MessageBox.Show(ex.Message, "Transit")
                            End Try

                        Next
                        strReport = strReport & vbNewLine

                        'EXPONENTIATE THE RESULTS
                        'strReport = strReport & "EXPONENTIATE THE RESULTS" & vbNewLine
                        dblVMT = Exp(dblVMT)
                        'strReport = strReport & "VMT Exp:" & dblVMT.ToString & vbNewLine
                        dblProbVMT = Exp(dblProbVMT)
                        'strReport = strReport & "VMT Prob Exp:" & dblProbVMT.ToString & vbNewLine
                        dblProbVMT = dblProbVMT / (1 + dblProbVMT)
                        'strReport = strReport & "VMT Prob Divided:" & dblProbVMT.ToString & vbNewLine
                        dblVMT = dblVMT * dblProbVMT
                        'strReport = strReport & "VMT:" & dblVMT.ToString & vbNewLine

                        dblVehicle = Exp(dblVehicle)
                        strReport = strReport & "Vehicle Exp:" & dblVehicle.ToString & vbNewLine

                        dblWalk = Exp(dblWalk)
                        ' strReport = strReport & "Walk Exp:" & dblWalk.ToString & vbNewLine
                        dblProbWalk = Exp(dblProbWalk)
                        ' strReport = strReport & "Walk Prob Exp:" & dblProbWalk.ToString & vbNewLine
                        dblProbWalk = dblProbWalk / (1 + dblProbWalk)
                        ' strReport = strReport & "Walk Prob Divided:" & dblProbWalk.ToString & vbNewLine
                        dblWalk = dblWalk * dblProbWalk
                        ' strReport = strReport & "Walk:" & dblWalk.ToString & vbNewLine

                        dblBike = Exp(dblBike)
                        ' strReport = strReport & "Bike Exp:" & dblBike.ToString & vbNewLine
                        dblProbBike = Exp(dblProbBike)
                        'strReport = strReport & "Bike Prob Exp:" & dblProbBike.ToString & vbNewLine
                        dblProbBike = dblProbBike / (1 + dblProbBike)
                        'strReport = strReport & "Bike Prob Divided:" & dblProbBike.ToString & vbNewLine
                        dblBike = dblBike * dblProbBike
                        'strReport = strReport & "Bike:" & dblBike.ToString & vbNewLine

                        dblTransit = Exp(dblTransit)
                        'strReport = strReport & "Transit Exp:" & dblTransit.ToString & vbNewLine
                        dblProbTransit = Exp(dblProbTransit)
                        'strReport = strReport & "Transit Prob Exp:" & dblProbTransit.ToString & vbNewLine
                        dblProbTransit = dblProbTransit / (1 + dblProbTransit)
                        'strReport = strReport & "Transit Prob Divided:" & dblProbTransit.ToString & vbNewLine
                        dblTransit = dblTransit * dblProbTransit
                        'strReport = strReport & "Transit:" & dblTransit.ToString & vbNewLine
                        'strReport = strReport & vbNewLine

                        'HU CALCS
                        If intHUFld >= 0 Then
                            dblHU = pFeat.Value(intHUFld)
                            If dblHU > 0 Then
                                dblHU_VMT = dblVMT * dblHU
                                dblHU_Vehicle = dblVehicle * dblHU
                                dblHU_Walk = dblWalk * dblHU
                                dblHU_Bike = dblBike * dblHU
                                dblHU_Transit = dblTransit * dblHU
                                'strReport = strReport & "HU Calc Value:" & dblHU.ToString & vbNewLine
                                'strReport = strReport & "VMT HU Calc Value:" & dblHU_VMT.ToString & vbNewLine
                                'strReport = strReport & "Vehicle Calc Value:" & dblHU_Vehicle.ToString & vbNewLine
                                'strReport = strReport & "Walk Calc Value:" & dblHU_Walk.ToString & vbNewLine
                                'strReport = strReport & "Bike Calc Value:" & dblHU_Bike.ToString & vbNewLine
                                'strReport = strReport & "Trnasit Calc Value:" & dblHU_Transit.ToString & vbNewLine
                                'strReport = strReport & vbNewLine
                            End If
                        End If

                    End If

                End If

                'ADD INDIVIDUAL TRANSIT MODEL VALUES TO SUMMARY VALUES
                dblSumVMT = dblSumVMT + dblHU_VMT
                dblSumVehicle = dblSumVehicle + dblHU_Vehicle
                dblSumWalk = dblSumWalk + dblHU_Walk
                dblSumBike = dblSumBike + dblHU_Bike
                dblSumTransit = dblSumTransit + dblHU_Transit

                'WRITE TRANSIT MODEL VALUES TO FEATURE 
                Try
                    pFeat.Value(pFeatureClass.FindField(strPrefix1 & arrWriteFields(0))) = dblProbVMT
                    pFeat.Value(pFeatureClass.FindField(strPrefix1 & arrWriteFields(1))) = dblVMT
                    pFeat.Value(pFeatureClass.FindField(strPrefix2 & arrWriteFields(1))) = dblHU_VMT
                Catch ex As Exception
                    'MessageBox.Show(ex.Message & vbNewLine & dblVMT.ToString, arrWriteFields(0))
                End Try
                Try
                    pFeat.Value(pFeatureClass.FindField(strPrefix1 & arrWriteFields(2))) = dblVehicle
                    pFeat.Value(pFeatureClass.FindField(strPrefix2 & arrWriteFields(2))) = dblHU_Vehicle
                Catch ex As Exception
                    ' MessageBox.Show(ex.Message & vbNewLine & dblVehicle.ToString, arrWriteFields(1))
                End Try
                Try
                    pFeat.Value(pFeatureClass.FindField(strPrefix1 & arrWriteFields(3))) = dblProbWalk
                    pFeat.Value(pFeatureClass.FindField(strPrefix1 & arrWriteFields(4))) = dblWalk
                    pFeat.Value(pFeatureClass.FindField(strPrefix2 & arrWriteFields(4))) = dblHU_Walk
                Catch ex As Exception
                    'MessageBox.Show(ex.Message & vbNewLine & dblWalk.ToString, arrWriteFields(2))
                End Try
                Try
                    pFeat.Value(pFeatureClass.FindField(strPrefix1 & arrWriteFields(5))) = dblProbBike
                    pFeat.Value(pFeatureClass.FindField(strPrefix1 & arrWriteFields(6))) = dblBike
                    pFeat.Value(pFeatureClass.FindField(strPrefix2 & arrWriteFields(6))) = dblHU_Bike
                Catch ex As Exception
                    'MessageBox.Show(ex.Message & vbNewLine & dblBike.ToString, arrWriteFields(3))
                End Try
                Try
                    pFeat.Value(pFeatureClass.FindField(strPrefix1 & arrWriteFields(7))) = dblProbTransit
                    pFeat.Value(pFeatureClass.FindField(strPrefix1 & arrWriteFields(8))) = dblTransit
                    pFeat.Value(pFeatureClass.FindField(strPrefix2 & arrWriteFields(8))) = dblHU_Transit
                Catch ex As Exception
                    'MessageBox.Show(ex.Message & vbNewLine & dblTransit.ToString, arrWriteFields(4))
                End Try
                pFeat.Store()
            End If


            'datWriteLC = New StreamWriter(("C:\Temp\SevenDs_Review.txt"))
            'datWriteLC.Write(strReport)
            'datWriteLC.Close()

            GC.WaitForPendingFinalizers()
            GC.Collect()
            pFeat = pFeatureCursor.NextFeature
            barStatus.Value = (intRecCount / intTotalCount) * 100
        Loop

        'WRITE THE SCENARIO OUTPUTS
        CType(m_shtTravelModel.Cells(intScenarioRow, 5), Microsoft.Office.Interop.Excel.Range).Value = dblSumVMT
        CType(m_shtTravelModel.Cells(intScenarioRow, 6), Microsoft.Office.Interop.Excel.Range).Value = dblSumVehicle
        CType(m_shtTravelModel.Cells(intScenarioRow, 8), Microsoft.Office.Interop.Excel.Range).Value = dblSumWalk
        CType(m_shtTravelModel.Cells(intScenarioRow, 10), Microsoft.Office.Interop.Excel.Range).Value = dblSumBike
        CType(m_shtTravelModel.Cells(intScenarioRow, 12), Microsoft.Office.Interop.Excel.Range).Value = dblSumTransit

        'RESET FORMULA CALC SETTING
        If intExcelFormulaCalc = 1 Then
            m_xlApp.Calculation = Interop.Excel.XlCalculation.xlCalculationAutomatic
        ElseIf intExcelFormulaCalc = 2 Then
            m_xlApp.Calculation = Interop.Excel.XlCalculation.xlCalculationManual
        ElseIf intExcelFormulaCalc = 3 Then
            m_xlApp.Calculation = Interop.Excel.XlCalculation.xlCalculationSemiautomatic
        End If

        m_xlWB.Save()

        'DISABLE THE EXCLE ENTRY CONTROLS IN CASE THE USER WANTS TO RUN ANOTHER SCENARIO
        Me.pnlExcel.Enabled = False
        Me.btnExcelFile.Enabled = False
        Me.tbxExcel.Enabled = False

        MessageBox.Show("The 7Ds calcualtions have completed.", "Processing Completed", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
        GoTo CleanUp

CleanUp:
        intRow = Nothing
        intScenarioRow = Nothing
        intColumn = Nothing
        strCellValue = Nothing
        pFeatLyr = Nothing
        pFeatureClass = Nothing
        pFeatSelection = Nothing
        pCursor = Nothing
        pFeatureCursor = Nothing
        pFeat = Nothing
        intTotalCount = Nothing
        intRecCount = Nothing
        strPrefix1 = Nothing
        strPrefix2 = Nothing
        dblSumVMT = Nothing
        dblSumVehicle = Nothing
        dblSumWalk = Nothing
        dblSumBike = Nothing
        dblSumTransit = Nothing
        dblVMT = Nothing
        dblVehicle = Nothing
        dblWalk = Nothing
        dblBike = Nothing
        dblTransit = Nothing
        dblProbVMT = Nothing
        dblProbWalk = Nothing
        dblProbBike = Nothing
        dblProbTransit = Nothing
        dblHU_VMT = Nothing
        dblHU_Vehicle = Nothing
        dblHU_Walk = Nothing
        dblHU_Bike = Nothing
        dblHU_Transit = Nothing
        strFld = Nothing
        arrFldValues = Nothing
        strField = Nothing
        dblValue = Nothing
        intCount = Nothing
        strValue = Nothing
        dblCoeff = Nothing
        intExcelFormulaCalc = Nothing
        pTable = Nothing
        dblInputValue = Nothing
        intHUFld = Nothing
        dblHU = Nothing
        strReport = Nothing
        datWriteLC = Nothing
        barStatus.Value = 0
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub

    Private Function intExcelFormulaCalc() As Integer
        Throw New NotImplementedException
    End Function

    Private Sub btnExcelFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExcelFile.Click
        '****************************************************************
        'Provide the user with a Open File dialog to select the Envision Excel file
        '****************************************************************
        Dim FileDialog1 As New OpenFileDialog

        Try
            FileDialog1.Title = "Select an Envision Excel File"
            FileDialog1.Filter = "Excel Files (*.xls,*.xlsx,*.xlsm)|*.xls;*.xlsx;*.xlsm"
            If FileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                Me.tbxExcel.Text = FileDialog1.FileName.ToString
                'Me.WindowState = FormWindowState.Minimized
                'LoadEnvisionExcelFile()
                'Me.WindowState = FormWindowState.Normal
            End If
            GoTo CleanUp
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Excel File Error")
            GoTo CleanUp
        End Try

CleanUp:
        FileDialog1 = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub

    'Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
    '    Dim current As Process = Process.GetCurrentProcess()
    '    Dim processes As Process() = Process.GetProcesses

    '    Dim ThisProcess As Process
    '    For Each ThisProcess In processes
    '        '-- Ignore the current process 
    '        If ThisProcess.Id <> current.Id Then
    '            '-- Only list processes that have a Main Window Title 
    '            If ThisProcess.MainWindowTitle <> "" Then
    '                MessageBox.Show(ThisProcess.ProcessName)
    '            End If
    '        End If
    '    Next
    'End Sub
End Class