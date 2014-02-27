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
'Imports ESRI.ArcGIS.Utility.CATIDs
'Imports ESRI.ArcGIS.SpatialAnalyst
'Imports ESRI.ArcGIS.SpatialAnalystUI
Imports ESRI.ArcGIS.Geoprocessing

Module modEnvisionSetup

    'ENVISION PROJECT SETUP
    Public m_frmEnvisionProjectSetup As frmEnvisionProjectSetup
    Public m_strWorkspacePath As String = ""
    Public m_pSpatRefProject As ISpatialReference
    Public m_blnSAFound As Boolean = True
    Public m_blnEnvisonSetupIsOpening As Boolean = True
    Public m_arrFeatureLayers As New ArrayList
    Public m_arrPolyFeatureLayers As New ArrayList
    Public m_arrRasterLayers As New ArrayList
    Public m_strLandUseLyrName As String
    Public m_strLandUseFldName As String
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
                m_arrFeatureLayers = New ArrayList
                m_arrPolyFeatureLayers = New ArrayList
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
                    m_arrFeatureLayers.Add(pFeatLyr)
                    pLyr = enumLayer.Next()
                Loop
                If m_frmEnvisionProjectSetup.cmbExtentLayers.Items.Count > 0 Then
                    m_frmEnvisionProjectSetup.cmbExtentLayers.Text = m_frmEnvisionProjectSetup.cmbExtentLayers.Items.Item(0)
                End If
            Catch ex As System.Exception
                m_arrFeatureLayers = New ArrayList
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
            If m_arrFeatureLayers.Count = 0 Then
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
                    m_arrRasterLayers.Add(pRasterLyr)
                    m_appEnvision.StatusBar.Message(0) = "Building list of raster layers : " & pRasterLyr.Name
                    m_frmEnvisionProjectSetup.cmbGridLayers.Items.Add(pRasterLyr.Name)
                    pLyr = enumLayer.Next()
                Loop
            Catch ex As System.Exception
                m_arrRasterLayers = New ArrayList
            End Try

            'DISABLE THE CONTROLS REQUIRING A RASTER LAYER
            If m_arrRasterLayers.Count = 0 Then
                'DISABLE SLOPE CONSTRAINTS CONTROLS
                m_frmEnvisionProjectSetup.pnlSlope.Enabled = False
            End If

            'BY DEFUALT RETRIEVE AND LOAD THE VIEW DOCUMENT SPATIAL REFERENCE PROJECTION
            pSpatRef = pMxDocument.FocusMap.SpatialReference
            Try
                pPCS = pSpatRef
                m_frmEnvisionProjectSetup.tbxProjection.Text = pSpatRef.Name
                m_pSpatRefProject = pSpatRef
                m_frmEnvisionProjectSetup.lblProjectUnits.Text = pPCS.CoordinateUnit.Name
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

    Public Function ProjectDirCheck() As Boolean
        'ENSURE SELECTED ENVISION WORKSPACE EXISTS
        If Directory.Exists(m_frmEnvisionProjectSetup.tbxWorkspace.Text) Then
            ProjectDirCheck = True
        Else
            ProjectDirCheck = False
        End If
    End Function

    Public Function EvisionFGBNameCheck() As Boolean
        'CHECK IF THE FILE GEODATABASE PROJECT ALREADY EXISTS
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
End Module
