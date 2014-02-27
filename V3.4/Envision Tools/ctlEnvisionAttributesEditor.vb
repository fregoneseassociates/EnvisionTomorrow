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

Public Class ctlEnvisionAttributesEditor
    Public blnShowColor As Boolean = False
    Public blnShowDevTypes As Boolean = False
    Public blnTempEdit As Boolean = False
    Public blnLoading As Boolean = False
    Public lbLoadingSubaraeas As Boolean = False

    Private Sub itmOpenEnvisionFGDB_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles itmOpenEnvisionFGDB.Click
        blnLoading = True
        m_strProcessing = "DEFINE ENVISION EDIT WORKSPACE AND SCEANRIO" & vbNewLine
        m_strProcessing = m_strProcessing & vbNewLine & "PROCESSING START TIME: " & Date.Now.ToString("M/d/yyyy hh:mm:ss tt") & vbNewLine & "---------------------------------------------------------------------------------------------" & vbNewLine

        If Me.btnEditing.Text = "END EDIT" Then
            MessageBox.Show("Please exit the current edit session before selecting a new Envision edit layer.", "OPEN EDIT SESSION", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo CleanUp
        End If

        'SELECT THE FILE GEODATABASE TO EDIT
        Dim pGxDialog As ESRI.ArcGIS.CatalogUI.IGxDialog
        Dim pGxObject As ESRI.ArcGIS.Catalog.IGxObject = Nothing
        Dim pLyrFilter As ESRI.ArcGIS.Catalog.IGxObjectFilter
        Dim pFGDBFilterColl As ESRI.ArcGIS.Catalog.IGxObjectFilterCollection
        Dim pGxEnumObject As ESRI.ArcGIS.Catalog.IEnumGxObject = Nothing
        Dim intRowCount As Integer
        Dim intFld As Integer
        Dim strPath As String
        Dim strName As String

        Try
            'CHECK TO SEE IF THE EDIT OPTION HAS BEEN ENABLED
            pGxDialog = New GxDialog
            pFGDBFilterColl = CreateObject("esriCatalogUI.GxObjectFiltercollection")
            pLyrFilter = New GxFilterFileGeodatabases
            pFGDBFilterColl.AddFilter(pLyrFilter, True)

            With pGxDialog
                .ObjectFilter = pFGDBFilterColl
                .Title = "Select an Envision File Geodatabase"
                .AllowMultiSelect = False
                .RememberLocation = True
                If Not .DoModalOpen(0, pGxEnumObject) Then
                    GoTo CleanUp
                Else
                    Try
                        'DISABLE THE SCENARIO MENU 
                        Me.itmEditScenarios.Enabled = False

                        pGxObject = pGxEnumObject.Next
                        strPath = pGxObject.Parent.FullName
                        strName = pGxObject.Name
                        pGxObject = Nothing
                        pGxEnumObject = Nothing
                        pGxEnumObject = Nothing
                        GC.WaitForPendingFinalizers()
                        GC.Collect()

                        'CLOSE THE DEV TYPE FORM, CLEAR THE GRAPHS FORM, CLEAR THE DEV TYPES, DELETE GRAPHS
                        Me.dgvDevelopmentTypes.Rows.Clear()
                        m_tblDevelopmentTypes = Nothing
                        m_tblTravel = Nothing
                        m_tblScSummary = Nothing
                        m_tblSc1 = Nothing
                        m_tblSc2 = Nothing
                        m_tblSc3 = Nothing
                        m_tblSc4 = Nothing
                        m_tblSc5 = Nothing
                        m_pEditFeatureLyr = Nothing

                        'RETRIEVE THE LOOKUP TABLES FOR THE SELECTED FILE GEODATABASE
                        'm_appEnvision.StatusBar.Message(0) = "Reviewing lookup tables"
                        If LookUpTablesEnvisionCheck(strPath & "\" & strName) Then
                        Me.itmOpenEnvisionFGDB.ToolTipText = strPath & "\" & strName
                        m_strFeaturePath = strPath & "\" & strName
                        If m_strFeaturePath.Contains("\\") Then
                            m_strFeaturePath = m_strFeaturePath.Replace("\\", "\")
                        End If
                        SetScenarioCntrlLabels()
                        Else
                        GoTo CleanUp
                        End If
                        'BUILD LISTS OF TRACKING FIELDS
                        'm_appEnvision.StatusBar.Message(0) = "Defining field tracking variables"
                        EnvisionFieldTrack()
                        EnvisionDevTypeFieldTrack()
                        'PROMPT USER TO SELECT AN EXCEL FILE TO RETRIEVE DEV TYPES
                        If blnLoadDevTypes Then
                            blnLoadDevTypes = False
                            itmSelectEnvisionFile_Click(sender, e)
                            LoadExcelDevTypes()
                        End If

                    Catch ex As Exception
            m_strProcessing = m_strProcessing & "Error 1 in sub itmOpenEnvisionLyr_Click: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
            m_strProcessing = m_strProcessing & "Error Message:  " & vbNewLine & ex.Message & vbNewLine
            MessageBox.Show(m_strProcessing, "Error")
            GoTo CleanUp
        End Try
                End If
                'LOAD THE SELECTED SCENARIO LAYER
                blnLoading = False
                If ChangeEnvisionLayer() Then
                    'ENABLE THE SCENARIO MENU 
                    Me.itmEditScenarios.Enabled = True
                End If
                GoTo CleanUp
            End With
        Catch ex As Exception
            m_strProcessing = m_strProcessing & "Error 2 in sub itmOpenEnvisionLyr_Click: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
            m_strProcessing = m_strProcessing & "Error Message:  " & vbNewLine & ex.Message & vbNewLine
            MessageBox.Show(ex.Message, "Open Envision File Geodatabase Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo CleanUp
        End Try

CleanUp:
        blnLoading = False
        m_strProcessing = m_strProcessing & "PROCESSING END TIME: " & Date.Now.ToString("M/d/yyyy hh:mm:ss tt") & vbNewLine & "---------------------------------------------------------------------------------------------" & vbNewLine
        WriteToProcessingFile("DefiningEnvisionScenioeditLayer.txt")
        pGxDialog = Nothing
        pGxObject = Nothing
        pLyrFilter = Nothing
        pFGDBFilterColl = Nothing
        pGxEnumObject = Nothing
        intRowCount = Nothing
        intFld = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub

    Public Function ChangeEnvisionLayer() As Boolean
        'CHECK FOR AND RETRIEVE THE SELECTED ENVISON LAYER
        m_strProcessing = m_strProcessing & "Starting Function ChangeEnvisionLayer: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine & vbNewLine
        ChangeEnvisionLayer = True

        Dim rowTemp As IRow = Nothing
        Dim strLayerName As String = ""
        'Dim strNextLayerName As String = ""
        Dim arrLyrList As ArrayList = New ArrayList
        Dim intCount As Integer
        Dim intLayerName As Integer = 1
        Dim blnFound As Boolean = False
        Dim pDataset As IDataset
        Dim strCurrentName As String = ""

        Dim pMxDocument As IMxDocument
        Dim mapCurrent As Map
        Dim blnLayerFound As Boolean = False
        Dim pFeatLayer As IFeatureLayer
        Dim pWksFactory As IWorkspaceFactory = Nothing
        Dim pFeatWks As IFeatureWorkspace
        Dim pFeatClass As IFeatureClass
        Dim intLayer As Integer
        Dim pTable As ITable = Nothing
        Dim strLyrPath As String = ""
        Dim strName As String = ""
        Dim pLayerAffects As ILayerEffects
        Dim blnValid As Boolean = True


        pMxDocument = DirectCast(m_appEnvision.Document, IMxDocument)
        mapCurrent = CType(pMxDocument.FocusMap, Map)

        'END ANY EDIT SESSIONS
        If Not Me.btnEditing.Text = "Start Edit" Then
            m_strProcessing = m_strProcessing & "An edit session is currently started and needs to be stopped before proceeding: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine & vbNewLine
            EditSession()
        End If

        'BUILD LIST OF ENVISION EDIT LAYER NAMES
        m_strProcessing = m_strProcessing & "Starting to build list of feature class names for defined Scenarios: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine & vbNewLine
        If Not m_tblScSummary Is Nothing Then
            Try
                For intCount = 1 To 4
                    rowTemp = m_tblScSummary.GetRow(intCount)
                    strLayerName = rowTemp.Value(m_tblScSummary.FindField("LAYER_NAME"))
                    arrLyrList.Add(strLayerName)
                Next
            Catch ex As Exception
                m_strProcessing = m_strProcessing & "Error in building list of feature class names for defined Scenarios: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
                m_strProcessing = m_strProcessing & "Error Message:  " & vbNewLine & ex.Message & vbNewLine
                ChangeEnvisionLayer = False
                GoTo CleanUp
            End Try
        Else
            m_strProcessing = m_strProcessing & "Envision Summary Table object not found. Exiting function: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
            ChangeEnvisionLayer = False
            GoTo CleanUp
        End If
        m_strProcessing = m_strProcessing & "Completed list of feature class names for defined Scenarios: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine

        'RETRIEVE THE EDIT LAYER NAME
        m_strProcessing = m_strProcessing & "Retrieve the associated layer name to the selected scenario : " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
        If m_tblScSummary.FindField("LAYER_NAME") <= -1 Then
            m_strProcessing = m_strProcessing & "The field name, LAYER_NAME, could not be found in the current Envision Scenario Summary Table: " & vbNewLine
            m_strProcessing = m_strProcessing & "Exiting function ChangeEnvisionLayer : " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
            ChangeEnvisionLayer = False
            GoTo CleanUp
        Else
            rowTemp = m_tblScSummary.GetRow(m_intEditScenario)
            strLayerName = rowTemp.Value(m_tblScSummary.FindField("LAYER_NAME"))
        End If
        m_strProcessing = m_strProcessing & "Associated layer name is " & strLayerName & strLayerName & ": " & vbNewLine

        'RETRIEVE THE CURRENT EDIT LAYER NAME
        If Not m_pEditFeatureLyr Is Nothing Then
            pDataset = CType(m_pEditFeatureLyr.FeatureClass, IDataset)
            strCurrentName = pDataset.Name
            pDataset = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()
            m_strProcessing = m_strProcessing & "The previous edit layer name is " & strCurrentName & ": " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
        End If

        'CHECK TO MAKE SURE ALL THE REQUIRED LOOKUP TABLES ARE AVAILABLE 
        m_strProcessing = m_strProcessing & "Execute function LookUpTablesEnvisionCheck to ensure all required Envision tables are in the selected Envision file geodatabase " & strCurrentName & ": " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
        If Not LookUpTablesEnvisionCheck(m_strFeaturePath) Then
            m_strProcessing = m_strProcessing & "Exiting function ChangeEnvisionLayer : " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
            ChangeEnvisionLayer = False
            GoTo CleanUp
        End If

        'CHECK TO SEE IF LAYER IS VALID
        Try
            pWksFactory = New FileGDBWorkspaceFactory
            pFeatWks = pWksFactory.OpenFromFile(m_strFeaturePath, 0)
            pFeatClass = pFeatWks.OpenFeatureClass(strLayerName)
            pTable = CType(pFeatClass, ITable)
        Catch ex As Exception
            blnValid = False
        End Try
        pFeatWks = Nothing
        pFeatClass = Nothing

        'IF NO LAYER NAME, THEN EXECUTE FUNCTIONAL TO PROMPT USER TO SELECT ONE
        If strLayerName.Length <= 0 Or Not blnValid Then
            'If Not SetEnvisionLayer(m_strFeaturePath, rowTemp, arrLyrList, strLayerName) Then
            '    m_strProcessing = m_strProcessing & "Exiting function ChangeEnvisionLayer: " & strLayerName & ": " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
            '    ChangeEnvisionLayer = False
            '    GoTo CleanUp
            'End If
            Me.mnuDefineLayers.Enabled = True
            MessageBox.Show("Please use the options under pull down menu (FILE | Define Scenario Layers) to designate scenario layer.", "Define Scenario Layer", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            ChangeEnvisionLayer = False
            'GoTo CleanUp
        Else
            'SET THE SELECTED LAYER
            m_pEditFeatureLyr = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()
            m_strFeatureName = strLayerName
        End If

        pDataset = Nothing
        m_pEditFeatureLyr = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()

        'DISABLE THE CONTROLS WHICH ARE DEPENDANT UPON THE ENVISION LAYER FILE
        m_strProcessing = m_strProcessing & "Set common controls enabled status to FALSE:  " & vbNewLine
        m_dockEnvisionWinForm.btnEditing.Enabled = False
        m_dockEnvisionWinForm.itmEnvisionExcel.Enabled = False
        m_dockEnvisionWinForm.itmSaveToExcelFile.Enabled = False
        m_dockEnvisionWinForm.itmSynchronize.Enabled = False
        m_dockEnvisionWinForm.itmFieldManager.Enabled = False
        m_dockEnvisionWinForm.itmEditScenarios.Enabled = False
        m_dockEnvisionWinForm.mnuDefineLayers.Enabled = False
        m_dockEnvisionWinForm.tlsPaintOptions.Enabled = False
        m_dockEnvisionWinForm.itmAccessFunctions.Enabled = False

        'REVIEW CURRENT LAYERS FOR ENVISION LAYER
        m_strProcessing = m_strProcessing & "Review current document for previous Envision edit layers and the current edit layer:  " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
        If mapCurrent.LayerCount > 0 Then
            For intLayer = 0 To mapCurrent.LayerCount - 1
                Try
                    If TypeOf mapCurrent.Layer(intLayer) Is IFeatureLayer Then
                        pFeatLayer = CType(mapCurrent.Layer(intLayer), IFeatureLayer)
                        pFeatClass = pFeatLayer.FeatureClass
                        pDataset = CType(pFeatClass, IDataset)
                        strLyrPath = pDataset.Workspace.PathName
                        strName = pDataset.BrowseName
                        If m_strFeaturePath = strLyrPath And (strName = strLayerName Or strName = strLayerName & ".shp") And pFeatClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                            blnLayerFound = True
                            m_pEditFeatureLyr = pFeatLayer
                            pFeatLayer.Name = pDataset.Name & " <ACTIVE>"
                            pFeatLayer.Visible = True
                        Else
                            If pFeatLayer.Name.Contains("<ACTIVE>") Then
                                pFeatLayer.Name = pDataset.Name
                                pFeatLayer.Visible = False
                            End If
                        End If
                        GC.WaitForPendingFinalizers()
                        GC.Collect()
                    End If
                Catch ex As Exception

                End Try
            Next
        End If
        m_strProcessing = m_strProcessing & "Completed document review:  " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine

        Try
            If Not blnLayerFound And ChangeEnvisionLayer Then
                pWksFactory = New FileGDBWorkspaceFactory
                m_strProcessing = m_strProcessing & "Open Workspace:  " & m_strFeaturePath & vbNewLine
                m_strProcessing = m_strProcessing & "Open Feature Class:  " & m_strFeatureName & vbNewLine
                pFeatWks = pWksFactory.OpenFromFile(m_strFeaturePath, 0)
                pFeatClass = pFeatWks.OpenFeatureClass(m_strFeatureName)
                pDataset = CType(pFeatClass, IDataset)
                If pFeatClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                    m_strProcessing = m_strProcessing & "Load layer into current document" & vbNewLine
                    m_pEditFeatureLyr = New FeatureLayer
                    m_pEditFeatureLyr.FeatureClass = pFeatClass
                    m_pEditFeatureLyr.Name = pDataset.Name & " <ACTIVE>"
                    pLayerAffects = m_pEditFeatureLyr
                    pLayerAffects.Transparency = 35
                    m_dockEnvisionWinForm.itmOpenEnvisionFGDB.ToolTipText = m_strFeaturePath & "\" & m_strFeatureName
                    mapCurrent.AddLayer(m_pEditFeatureLyr)
                    pTable = CType(pFeatClass, Table)
                End If
                GC.WaitForPendingFinalizers()
                GC.Collect()
            End If
        Catch ex As Exception
            m_strProcessing = m_strProcessing & "Error in Opening feature class: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
            m_strProcessing = m_strProcessing & "Error Message:  " & vbNewLine & ex.Message & vbNewLine
            m_strProcessing = m_strProcessing & "Please review the available feature classes to ensure this layer exists." & vbNewLine
            'GoTo CleanUp
        End Try

        If ChangeEnvisionLayer Then
            'RELOAD THE SCENARIO LABELS
            SetScenarioCntrlLabels()

            'LOAD THE DEVELOPMENT TYPES
            RetrieveDevTypeData()

            If m_dockEnvisionWinForm.dgvDevelopmentTypes.Rows.Count > 0 Then
                'UPDATE THE LEGEND
                UpdateEnvisionLyrLegend()
            End If

            'CHECK FOR REQUIRED FIELDS
            If Not pTable Is Nothing Then
                EnvisionLyrRequiredFieldsCheck(pTable)
            End If
        End If

        'POPULATE THE ENVISION FGDB COMBO BOXES WITH POLY LAYER NAMES
        blnLoading = True
        LoadPolyLayers()
        blnLoading = False

        'ENABLE THE CONTROLS WHICH ARE DEPENDANT UPON THE ENVISION LAYER FILE
        m_strProcessing = m_strProcessing & "Set common controls enabled status to TRUE:  " & vbNewLine
        m_dockEnvisionWinForm.btnEditing.Enabled = ChangeEnvisionLayer
        m_dockEnvisionWinForm.itmEnvisionExcel.Enabled = ChangeEnvisionLayer
        m_dockEnvisionWinForm.itmSaveToExcelFile.Enabled = ChangeEnvisionLayer
        m_dockEnvisionWinForm.itmSynchronize.Enabled = ChangeEnvisionLayer
        m_dockEnvisionWinForm.itmFieldManager.Enabled = ChangeEnvisionLayer
        m_dockEnvisionWinForm.itmEditScenarios.Enabled = True
        m_dockEnvisionWinForm.mnuDefineLayers.Enabled = True
        m_dockEnvisionWinForm.tlsPaintOptions.Enabled = ChangeEnvisionLayer
        m_dockEnvisionWinForm.itmAccessFunctions.Enabled = ChangeEnvisionLayer
        If Not ChangeEnvisionLayer Then
            GoTo CleanUp
        End If

        'UPDATE THE INDICATOR TABLE SELECTION
        Try
            UpdateSummaryTableSelection()
        Catch ex As Exception
            m_strProcessing = m_strProcessing & "Error in UpdateSummaryTableSelection: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
            m_strProcessing = m_strProcessing & "Error Message:  " & vbNewLine & ex.Message & vbNewLine
            MessageBox.Show(ex.Message, "Open Layer Error: UpdateSummaryTableSelection Sub Section", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo CleanUp
        End Try

        'LOAD THE SELECTED ENVISION EDIT LAYERS FIELDS
        Try
            RetrieveEvisionFields()
        Catch ex As Exception
            m_strProcessing = m_strProcessing & "Error in RetrieveEvisionFields: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
            m_strProcessing = m_strProcessing & "Error Message:  " & vbNewLine & ex.Message & vbNewLine
            MessageBox.Show(ex.Message, "Open Layer Error: RetrieveEvisionFields Sub Section", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo CleanUp
        End Try

        'EXECUTE FUNCTIONS TO LOAD THE SELECTED ENVISION EDIT LAYER
        Try
            'LOAD THE DEVELOPMENT TYPES
            RetrieveDevTypeData()
            Me.itmEnvisionExcel.Enabled = True
        Catch ex As Exception
            m_strProcessing = m_strProcessing & "Error in RetrieveIndicatorData or RetrieveDevTypeData: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
            m_strProcessing = m_strProcessing & "Error Message:  " & vbNewLine & ex.Message & vbNewLine
            ChangeEnvisionLayer = False
            GoTo CleanUp
        End Try

        'EXECUTE SYNCHRONIZE SCRIPTS
        'm_appEnvision.StatusBar.Message(0) = "Execute synchronize scripts"
        DeleteEnvisionTempSummaryTables(True)
        'EnvisionSyncEX_LU()
        'EnvisionSyncEX_LU_CreateSummaryTables()
        'SynchronizeData("QUICK")

CleanUp:
        m_strProcessing = m_strProcessing & "Ending Function ChangeEnvisionLayer: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
        rowTemp = Nothing
        strLayerName = Nothing
        arrLyrList = Nothing
        intCount = Nothing
        intLayerName = Nothing
        blnFound = Nothing
        pDataset = Nothing
        pMxDocument = Nothing
        mapCurrent = Nothing
        blnLayerFound = Nothing
        pFeatLayer = Nothing
        pWksFactory = Nothing
        pFeatWks = Nothing
        pFeatClass = Nothing
        pDataset = Nothing
        intLayer = Nothing
        pTable = Nothing
        pLayerAffects = Nothing

        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Function

    Public Function SetEnvisionLayer(ByVal strFileGDB As String, ByVal rowTemp As IRow, ByVal arrList As ArrayList, ByVal strNextName As String) As Boolean
        SetEnvisionLayer = True
        'PROMPT USER TO SELECT A FEATURE CLASS FOR THE CURRENT SCENARIO
        'THE SELECTED FEATURE CLASS MUST BE IN THE LOACTED IN THE SELECTED FILE GEODATABASE
        If blnLoading Then
            GoTo CleanUp
        End If

        m_strProcessing = m_strProcessing & "Starting function SetEnvisionLayer: " & Date.Now.ToString("hh:mm:ss tt")
        Dim arrSelectedLayer As ArrayList = New ArrayList
        Dim strLayerPath As String = ""
        Dim strLayerName As String = ""
        Dim blnLyrSelected As Boolean = False
        Dim intCount As Integer
        Dim blnFound As Boolean = False
        Dim rowName As IRow

        Try
            Do While blnLyrSelected = False
                arrSelectedLayer = New ArrayList
                arrSelectedLayer = SelectEnvisionLayer()
                If Not arrSelectedLayer.Count = 2 Then
                    SetEnvisionLayer = False
                    GoTo CleanUp
                Else
                    strLayerPath = arrSelectedLayer.Item(0).ToString
                    strLayerName = arrSelectedLayer.Item(1).ToString
                    'REVIEW THE SELECTED FILE TO ENSURE IT DOESN'T ALREADY EXIST
                    For intCount = 1 To 4
                        rowName = m_tblScSummary.GetRow(intCount)
                        If rowName.Value(m_tblScSummary.FindField("LAYER_NAME")) = strLayerName Then
                            blnFound = True
                            Exit For
                        End If
                    Next
                    rowName = Nothing
                    If blnFound Then
                        If (Not strLayerPath = strFileGDB) Or strLayerName = m_strFeatureName Then
                            strNextName = MakeNewName(m_strFeatureName)
                        Else
                            strNextName = MakeNewName(strLayerName)
                        End If
                        If CopyEnvisionLyr((strLayerPath & "\" & strLayerName), (m_strFeaturePath & "\" & strNextName)) Then
                            'WRITE TO THE LOOKUP TABLE
                            m_strProcessing = m_strProcessing & "Saving Edit feature layer name to scenario summary table:  " & strNextName & vbNewLine
                            strLayerName = strNextName
                            m_strFeatureName = strLayerName
                            rowTemp.Value(m_tblScSummary.FindField("LAYER_NAME")) = strLayerName
                            rowTemp.Store()
                            GoTo CleanUp
                        End If
                    Else
                        If Not strLayerPath = strFileGDB Or strLayerName = m_strFeatureName Then
                            strNextName = MakeNewName(m_strFeatureName)
                            If CopyEnvisionLyr((strLayerPath & "\" & strLayerName), (m_strFeaturePath & "\" & strNextName)) Then
                                'WRITE TO THE LOOKUP TABLE
                                m_strProcessing = m_strProcessing & "Saving Edit feature layer name to scenario summary table:  " & strNextName & vbNewLine
                                strLayerName = strNextName
                                m_strFeatureName = strLayerName
                                rowTemp.Value(m_tblScSummary.FindField("LAYER_NAME")) = strLayerName
                                rowTemp.Store()
                                GoTo CleanUp
                            End If
                        Else
                            m_strFeatureName = strLayerName
                            m_strProcessing = m_strProcessing & "Saving Edit feature layer name to scenario summary table:  " & strLayerName & vbNewLine
                            rowTemp.Value(m_tblScSummary.FindField("LAYER_NAME")) = strLayerName
                            rowTemp.Store()
                            GoTo CleanUp
                        End If
                    End If
                    End If
            Loop
            blnLoading = True
            LoadPolyLayers()
            blnLoading = False

        Catch ex As Exception
            m_strProcessing = m_strProcessing & "Error in SetEnvisionLayer: " & Date.Now.ToString("hh:mm:ss tt")
            m_strProcessing = m_strProcessing & "Error Message:  " & vbNewLine & ex.Message
            SetEnvisionLayer = False
            MessageBox.Show(ex.Message, "Open Layer Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo CleanUp
        End Try

CleanUp:
        arrList = Nothing
        rowTemp = Nothing
        arrSelectedLayer = Nothing
        strLayerPath = Nothing
        strLayerName = Nothing
        blnLyrSelected = Nothing
        intCount = Nothing
        blnFound = Nothing
        rowName = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Function

    Private Function MakeNewName(ByVal strName As String) As String
        'APPEND THE SCENARIO NUMBER TO THE SELECTED FILE NAME
        MakeNewName = strName & m_intEditScenario.ToString

        'CREATE THE NEW ENVISIONB FEATURE CLASS NAME BASE ON SELECT FILE NAME
        'DETERMINE THE NEXT LAYER NAME
        'Dim pWksFactory As IWorkspaceFactory = Nothing
        'Dim pFeatWks As IFeatureWorkspace
        'Dim pFeatClass As IFeatureClass
        'Dim intCount As Integer
        'Dim strNextLayerName As String = ""

        'intCount = 0
        'm_strProcessing = m_strProcessing & "Determine the next edit layer name should it be need to copy a file: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
        'Do While intCount < 50
        '    intCount = intCount + 1
        '    strNextLayerName = strName & intCount.ToString
        '    Try
        '        pWksFactory = New FileGDBWorkspaceFactory
        '        pFeatWks = pWksFactory.OpenFromFile(m_strFeaturePath, 0)
        '        pFeatClass = pFeatWks.OpenFeatureClass(strNextLayerName)
        '    Catch ex As Exception
        '        MakeNewName = strNextLayerName
        '        Exit Do
        '    End Try
        'Loop
        m_strProcessing = m_strProcessing & "Next Edit Layer name will be " & MakeNewName & ": " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
CleanUp:
        'pWksFactory = Nothing
        'pFeatWks = Nothing
        'pFeatClass = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Function

    Private Function CopyEnvisionLyr(ByVal strSource As String, ByVal strDestination As String) As Boolean
        'COPY THE SELECTED LAYER TO THE ENVISION FILE GEODATABASE
        m_strProcessing = m_strProcessing & "Starting function CopyEnvisionLyr: " & Date.Now.ToString("hh:mm:ss tt")
        CopyEnvisionLyr = True
        'END ANY EDIT SESSION WHICH MIGHT BE IN PROGRESS
        If Not Me.btnEditing.Text = "Start Edit" Then
            EditSession()
        End If

        Dim GP As ESRI.ArcGIS.Geoprocessor.Geoprocessor = New ESRI.ArcGIS.Geoprocessor.Geoprocessor
        Dim pCopy As ESRI.ArcGIS.DataManagementTools.Copy
        Try
            pCopy = New ESRI.ArcGIS.DataManagementTools.Copy
            pCopy.in_data = strSource
            pCopy.out_data = strDestination
            pCopy.data_type = "FeatureClass"
            GP.OverwriteOutput = True
            GP.AddOutputsToMap = True
            GP.TemporaryMapLayers = False
            RunTool(GP, pCopy)
            pCopy = Nothing
            GP = Nothing
            m_strProcessing = m_strProcessing & "Ending function CopyEnvisionLyr: " & Date.Now.ToString("hh:mm:ss tt")
        Catch ex As Exception
            m_strProcessing = m_strProcessing & "Error in CopyEnvisionLyr: " & Date.Now.ToString("hh:mm:ss tt")
            m_strProcessing = m_strProcessing & "Error Message:  " & vbNewLine & ex.Message
            CopyEnvisionLyr = False
            GoTo CleanUp
        End Try
CleanUp:
        GP = Nothing
        pCopy = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Function

    Public Function SelectEnvisionLayer() As ArrayList
        'PROVIDE FORM FOR THE USER TO SELECT AN ENVISION EDIT LAYER
        SelectEnvisionLayer = New ArrayList
        If blnLoading Then
            GoTo CleanUp
        End If
        m_strProcessing = m_strProcessing & "Starting function SelectEnvisionLayer: " & Date.Now.ToString("hh:mm:ss tt")
        Dim strLayerPath As String = ""
        Dim strLayerName As String = ""
        Dim pGxDialog As ESRI.ArcGIS.CatalogUI.IGxDialog
        Dim pGxObject As ESRI.ArcGIS.Catalog.IGxObject = Nothing
        Dim pLyrFilter As ESRI.ArcGIS.Catalog.IGxObjectFilter
        Dim pFGDBFilterColl As ESRI.ArcGIS.Catalog.IGxObjectFilterCollection
        Dim pGxEnumObject As ESRI.ArcGIS.Catalog.IEnumGxObject = Nothing
        Try
            pGxDialog = New GxDialog
            pFGDBFilterColl = CreateObject("esriCatalogUI.GxObjectFiltercollection")
            pLyrFilter = New GxFilterFGDBFeatureClasses
            pFGDBFilterColl.AddFilter(pLyrFilter, True)
            With pGxDialog
                .ObjectFilter = pFGDBFilterColl
                .Title = "Select an Envision Edit Layer for Scenario " & m_intEditScenario.ToString
                .AllowMultiSelect = False
                .RememberLocation = True
                If Not .DoModalOpen(0, pGxEnumObject) Then
                    GoTo CleanUp
                Else
                    pGxObject = pGxEnumObject.Next
                    m_strProcessing = m_strProcessing & "Selected feature layer path: " & pGxObject.Parent.FullName
                    m_strProcessing = m_strProcessing & "Selected feature layer name: " & pGxObject.Name
                    SelectEnvisionLayer.Add(pGxObject.Parent.FullName)
                    SelectEnvisionLayer.Add(pGxObject.Name)
                    GoTo CleanUp
                End If
            End With
            m_strProcessing = m_strProcessing & "Ending function SelectEnvisionLayer: " & Date.Now.ToString("hh:mm:ss tt")
        Catch ex As Exception
            m_strProcessing = m_strProcessing & "Error in SelectEnvisionLayer: " & Date.Now.ToString("hh:mm:ss tt")
            m_strProcessing = m_strProcessing & "Error Message:  " & vbNewLine & ex.Message
            GoTo CleanUp
        End Try
CleanUp:
        strLayerPath = Nothing
        strLayerName = Nothing
        pGxDialog = Nothing
        pGxObject = Nothing
        pLyrFilter = Nothing
        pFGDBFilterColl = Nothing
        pGxEnumObject = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Function

    Public Sub SetScenarioCntrlLabels()
        'SETS THE LABELS TO SCENARIO CONTROLS
        m_strProcessing = m_strProcessing & "Starting function SetScenarioCntrlLabels: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
        Dim intRowCount As Integer = 5
        Dim intRow As Integer
        Dim intScenarioNameFld As Integer
        Dim intLayerNameFld As Integer
        Dim strScenarioName As String
        Dim strLayerName As String
        Dim rowTemp As IRow
        Dim pWksFactory As IWorkspaceFactory = Nothing
        Dim pFeatWks As IFeatureWorkspace
        Dim pFeatClass As IFeatureClass

        Try
            If m_tblScSummary Is Nothing Then
                m_strProcessing = m_strProcessing & "Scenario Summary table object not found" & vbNewLine
                GoTo CleanUp
            End If
            'CHECK TO SEE SELECTED WORKSPACE IS VALID
            Try
                pWksFactory = New FileGDBWorkspaceFactory
                pFeatWks = pWksFactory.OpenFromFile(m_strFeaturePath, 0)
            Catch ex As Exception
                pFeatWks = Nothing
            End Try

            intRowCount = m_tblScSummary.RowCount(Nothing)
            If intRowCount >= 1 Then
                intScenarioNameFld = m_tblScSummary.FindField("SCENARIO")
                intLayerNameFld = m_tblScSummary.FindField("LAYER_NAME")
                For intRow = 1 To intRowCount
                    rowTemp = m_tblScSummary.GetRow(intRow)
                    strScenarioName = "SCENARIO " & intRow.ToString
                    strLayerName = ""
                    If intRow = 1 Then
                        Try
                            If Not pWksFactory Is Nothing Then
                                strScenarioName = CStr(rowTemp.Value(intScenarioNameFld))
                                strLayerName = CStr(rowTemp.Value(intLayerNameFld))
                                pFeatClass = pFeatWks.OpenFeatureClass(strLayerName)
                            End If
                        Catch ex As Exception
                            strLayerName = ""
                        End Try
                        m_dockEnvisionWinForm.itmScenario1.Text = strScenarioName
                        m_dockEnvisionWinForm.itmScenario1.Tag = strLayerName
                        rowTemp.Value(intScenarioNameFld) = strScenarioName
                        rowTemp.Value(intLayerNameFld) = strLayerName
                        rowTemp.Store()
                    ElseIf intRow = 2 Then
                        Try
                            If Not pWksFactory Is Nothing Then
                                strScenarioName = CStr(rowTemp.Value(intScenarioNameFld))
                                strLayerName = CStr(rowTemp.Value(intLayerNameFld))
                                pFeatClass = pFeatWks.OpenFeatureClass(strLayerName)
                            End If
                        Catch ex As Exception
                            strLayerName = ""
                        End Try
                        m_dockEnvisionWinForm.itmScenario2.Text = strScenarioName
                        m_dockEnvisionWinForm.itmScenario2.Tag = strLayerName
                        rowTemp.Value(intScenarioNameFld) = strScenarioName
                        rowTemp.Value(intLayerNameFld) = strLayerName
                        rowTemp.Store()
                    ElseIf intRow = 3 Then
                        Try
                            If Not pWksFactory Is Nothing Then
                                strScenarioName = CStr(rowTemp.Value(intScenarioNameFld))
                                strLayerName = CStr(rowTemp.Value(intLayerNameFld))
                                pFeatClass = pFeatWks.OpenFeatureClass(strLayerName)
                            End If
                        Catch ex As Exception
                            strLayerName = ""
                        End Try
                        m_dockEnvisionWinForm.itmScenario3.Text = strScenarioName
                        m_dockEnvisionWinForm.itmScenario3.Tag = strLayerName
                        rowTemp.Value(intScenarioNameFld) = strScenarioName
                        rowTemp.Value(intLayerNameFld) = strLayerName
                        rowTemp.Store()
                    ElseIf intRow = 4 Then
                        Try
                            If Not pWksFactory Is Nothing Then
                                strScenarioName = CStr(rowTemp.Value(intScenarioNameFld))
                                strLayerName = CStr(rowTemp.Value(intLayerNameFld))
                                pFeatClass = pFeatWks.OpenFeatureClass(strLayerName)
                            End If
                        Catch ex As Exception
                            strLayerName = ""
                        End Try
                        m_dockEnvisionWinForm.itmScenario4.Text = strScenarioName
                        m_dockEnvisionWinForm.itmScenario4.Tag = strLayerName
                        rowTemp.Value(intScenarioNameFld) = strScenarioName
                        rowTemp.Value(intLayerNameFld) = strLayerName
                        rowTemp.Store()
                    ElseIf intRowCount = 5 Then
                        Try
                            If Not pWksFactory Is Nothing Then
                                strScenarioName = CStr(rowTemp.Value(intScenarioNameFld))
                                strLayerName = CStr(rowTemp.Value(intLayerNameFld))
                                pFeatClass = pFeatWks.OpenFeatureClass(strLayerName)
                            End If
                        Catch ex As Exception
                            strLayerName = ""
                        End Try
                        m_dockEnvisionWinForm.itmScenario5.Text = strScenarioName
                        m_dockEnvisionWinForm.itmScenario5.Tag = strLayerName
                        rowTemp.Value(intScenarioNameFld) = strScenarioName
                        rowTemp.Value(intLayerNameFld) = strLayerName
                        rowTemp.Store()
                    End If
                Next
            End If
            GoTo CleanUp
        Catch ex As Exception
            m_strProcessing = m_strProcessing & "Error in SetScenarioCntrlLabels: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
            m_strProcessing = m_strProcessing & "Error Message:  " & vbNewLine & ex.Message & vbNewLine
            GoTo CleanUp
        End Try
CleanUp:
        m_strProcessing = m_strProcessing & "Ending function SetScenarioCntrlLabels: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
        intRowCount = Nothing
        intScenarioNameFld = Nothing
        intLayerNameFld = Nothing
        rowTemp = Nothing
        pFeatWks = Nothing
        pFeatClass = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub

    Private Sub itmSynchronizeFull_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles itmSynchronizeFull.Click
        SynchronizeData("FULL")
    End Sub

    Private Sub itmSynchronizePartial_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles itmSynchronizePartial.Click
        SynchronizeData("QUICK")
    End Sub

    Private Sub SynchronizeData(ByVal strSelection As String)
        'SYNCHRONIZE THE ENVISION DATA LAYER DEVELOPMENT TYPE ACRES WITH THE ENVISION EXCEL DEVELOPMENT TYPE ACRES

        'CHECK FOR ENVISION EDIT LAYER
        If m_pEditFeatureLyr Is Nothing Then
            GoTo CleanUp
        End If

        'CHECK FOR VALID DEVELOPMENT TYPES
        If m_tblDevelopmentTypes.RowCount(Nothing) <= 0 Then
            GoTo CleanUp
        End If

        Dim mxApplication As IMxApplication
        Dim mxDoc As IMxDocument
        Dim pREnv As IEnvelope = Nothing
        Dim pSpatFilter As ISpatialFilter = New SpatialFilter
        Dim pFeatSelection As IFeatureSelection
        Dim pCursor As ICursor = Nothing
        Dim pFeatureCursor As IFeatureCursor
        Dim pFeat As IFeature
        Dim pFeatClass As IFeatureClass = m_pEditFeatureLyr.FeatureClass
        Dim pTable As ITable = CType(pFeatClass, ITable)
        Dim pFeatTable As ITable = CType(pFeatClass, ITable)
        Dim intTotalCount As Integer
        Dim intCount As Integer
        Dim strDevType As String = ""
        Dim dblVacantAcres As Double = 0
        Dim dblDevelopedAcres As Double = 0
        Dim intDevTypeFld As Integer = -1
        Dim intVacantAcresFld As Integer = -1
        Dim intDevelopedAcresFld As Integer = -1
        Dim intRedevFld As Integer = -1
        Dim intNetAcreFld As Integer = -1
        Dim rowDevType As IRow = Nothing
        Dim intDevTypeRow As Integer = -1
        Dim intRow As Integer
        Dim rowTemp As IRow
        Dim intScVacLookUpFld As Integer
        Dim intScDevdLookUpFld As Integer
        Dim rowScLookUp As IRow
        Dim dblScExValue As Double
        Dim intScExField As Integer
        Dim tblScLookUpTbl As ITable = Nothing
        Dim strFromFieldName As String = ""
        Dim strToFieldName As String = ""
        Dim dblCalcAcres As Double = 0
        Dim dblRedevRate As Double = 0
        Dim dblNetAcre As Double = 0
        Dim dblXValue As Double = 0
        Dim dblExistingXValue As Double = 0
        Dim intFldCount As Integer = 0
        Dim pDef As IFeatureLayerDefinition2
        Dim pQFilter As IQueryFilter
        Dim strDefExpression As String = ""
        Dim strQString As String = Nothing
        Dim intScLookUpFld As Integer = 0
        Dim strFieldName As String = ""
        Dim fldSc As IField
        Dim dblPreviousValue As Double
        Dim dblValue As Double
        Dim strDepVarFieldName As String = ""
        Dim dblDepVar As Double = 0

        'CHECK FOR REQUIRED FIELDS
        If pTable Is Nothing Then
            MessageBox.Show("The Envision Edit Layer could not be found in the current view document.  Please reselect the Envision layer.", "Layer Not Found", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo CleanUp
        End If

        'CHECK FOR REQUIRED FIELDS
        EnvisionLyrRequiredFieldsCheck(pTable)
        intVacantAcresFld = pFeatClass.FindField("VAC_ACRE")
        intDevelopedAcresFld = pFeatClass.FindField("DEVD_ACRE")


        '*********************************************************
        'MAKE THE FEATURE SELECTION
        '*********************************************************
        pFeatClass = m_pEditFeatureLyr.FeatureClass
        pDef = m_pEditFeatureLyr
        strDefExpression = pDef.DefinitionExpression
        pQFilter = New QueryFilter
        If strSelection = "QUICK" Then
            strQString = "NOT DEV_TYPE = ''"
            If strDefExpression.Length > 0 Then
                pQFilter.WhereClause = "(" & strQString & ") AND (" & pDef.DefinitionExpression & ")"
                pFeatureCursor = pFeatClass.Search(pQFilter, False)
            Else
                pQFilter.WhereClause = strQString
                pFeatureCursor = pFeatClass.Search(pQFilter, False)
            End If
            intTotalCount = pFeatClass.FeatureCount(pQFilter)
        Else
            If strDefExpression.Length > 0 Then
                pQFilter.WhereClause = pDef.DefinitionExpression
                pFeatureCursor = pFeatClass.Search(pQFilter, False)
                intTotalCount = pFeatClass.FeatureCount(pQFilter)
            Else
                pFeatureCursor = pFeatClass.Search(Nothing, False)
                intTotalCount = pFeatClass.FeatureCount(Nothing)
            End If
        End If
        If Me.btnEditing.Text = "Start Edit" Then
            blnTempEdit = True
            EditSession()
        End If

        'RETRIEVE THE CORRESPONDING SCENARIO LOOKUP TABLE 
        If m_intEditScenario = 1 Then
            tblScLookUpTbl = m_tblSc1
        ElseIf m_intEditScenario = 2 Then
            tblScLookUpTbl = m_tblSc2
        ElseIf m_intEditScenario = 3 Then
            tblScLookUpTbl = m_tblSc3
        ElseIf m_intEditScenario = 4 Then
            tblScLookUpTbl = m_tblSc4
        ElseIf m_intEditScenario = 5 Then
            tblScLookUpTbl = m_tblSc5
        End If
        If tblScLookUpTbl Is Nothing Then
            GoTo CleanUp
        End If

        '*********************************************************
        'CLEAR THE VACANT AND DEVELOPABLE ACRE AND HH CALC COLUMNS
        '*********************************************************
        Dim intRowCount As Integer
        intRowCount = tblScLookUpTbl.RowCount(Nothing)
        For intRow = 1 To intRowCount
            rowTemp = tblScLookUpTbl.GetRow(intRow)
            For intScLookUpFld = 1 To (tblScLookUpTbl.Fields.FieldCount - 1)
                strFieldName = rowTemp.Fields.Field(intScLookUpFld).Name
                fldSc = rowTemp.Fields.Field(intScLookUpFld)
                If fldSc.Type = esriFieldType.esriFieldTypeDouble Then
                    rowTemp.Value(intScLookUpFld) = 0
                End If
            Next
            rowTemp.Store()
        Next
        rowTemp = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()

        '*********************************************************
        'WRITE THE EDIT VALUE(S)
        '*********************************************************
        mxApplication = CType(m_appEnvision, IMxApplication)
        mxDoc = CType(m_appEnvision.Document, IMxDocument)
        Try
            'RETRIEVE THE CORRESPONDING SCENARIO LOOKUP TABLE 
            If m_intEditScenario = 1 Then
                tblScLookUpTbl = m_tblSc1
            ElseIf m_intEditScenario = 2 Then
                tblScLookUpTbl = m_tblSc2
            ElseIf m_intEditScenario = 3 Then
                tblScLookUpTbl = m_tblSc3
            ElseIf m_intEditScenario = 4 Then
                tblScLookUpTbl = m_tblSc4
            ElseIf m_intEditScenario = 5 Then
                tblScLookUpTbl = m_tblSc5
            End If
            If tblScLookUpTbl Is Nothing Then
                GoTo CleanUp
            End If

            pFeat = pFeatureCursor.NextFeature
            Do While Not pFeat Is Nothing
                intCount = intCount + 1
                'RETRIEVE THE DEVELOPMENT TYPES
                intDevTypeFld = pFeat.Fields.FindField("Dev_Type")
                If intDevTypeFld > -1 Then
                    Try
                        strDevType = CStr(pFeat.Value(intDevTypeFld))
                    Catch ex As Exception
                        strDevType = ""
                    End Try
                End If

                'RETRIEVE VACANT AND DEVELOPABLE ACRES
                Try
                    dblVacantAcres = CType(CType((pFeat.Value(intVacantAcresFld).ToString), Double), Double)
                    If dblVacantAcres < 0 Then
                        dblVacantAcres = 0
                    End If
                Catch ex As Exception
                    dblVacantAcres = 0
                End Try
                Try
                    dblDevelopedAcres = CType(CType((pFeat.Value(intDevelopedAcresFld).ToString), Double), Double)
                    If dblDevelopedAcres < 0 Then
                        dblDevelopedAcres = 0
                    End If
                Catch ex As Exception
                    dblDevelopedAcres = 0
                End Try
                dblVacantAcres = RoundNumber(dblVacantAcres, 6)
                dblDevelopedAcres = RoundNumber(dblDevelopedAcres, 6)

                'RETRIEVE THE RECORD NUMBER
                intDevTypeRow = -1
                rowTemp = Nothing
                rowDevType = Nothing
                dblRedevRate = 0
                dblNetAcre = 0
                For intRow = 1 To 50
                    rowTemp = m_tblDevelopmentTypes.GetRow(intRow)
                    Try
                        If CStr(rowTemp.Value(m_tblDevelopmentTypes.FindField("DEVELOPMENT_TYPE"))) = strDevType Then
                            intDevTypeRow = intRow
                            rowDevType = rowTemp
                            intRedevFld = m_tblDevelopmentTypes.FindField("REDEV_RATE")
                            If intRedevFld >= 0 Then
                                dblRedevRate = rowDevType.Value(intRedevFld)
                            End If
                            intNetAcreFld = m_tblDevelopmentTypes.FindField("NET_ACRE")
                            If intNetAcreFld >= 0 Then
                                dblNetAcre = rowDevType.Value(intNetAcreFld)
                            End If
                            Exit For
                        End If
                    Catch ex As Exception
                        GoTo CleanUp
                    End Try
                Next

                rowScLookUp = Nothing
                If intDevTypeRow >= 0 Then
                    'ADD THE ACRES TO THE SCENARIO LOOKUP TABLES
                    For intScLookUpFld = 1 To (tblScLookUpTbl.Fields.FieldCount - 1)
                        strFieldName = tblScLookUpTbl.Fields.Field(intScLookUpFld).Name
                        fldSc = tblScLookUpTbl.Fields.Field(intScLookUpFld)
                        If fldSc.Type = esriFieldType.esriFieldTypeDouble Then
                            If m_pEditFeatureLyr.FeatureClass.FindField(strFieldName) >= 0 Then
                                dblValue = 0
                                Try
                                    dblValue = CType((pFeat.Value(pFeat.Fields.FindField(strFieldName)).ToString), Double)
                                Catch ex As Exception
                                    dblValue = 0
                                End Try

                                'RETRIEVE VACANT AND DEVELOPABLE ACRES AND ADD TO SUMMARY TABLE
                                If strFieldName = "VAC_ACRE" Then
                                    Try
                                        dblVacantAcres = CType((pFeat.Value(intVacantAcresFld).ToString), Double)
                                        If dblVacantAcres < 0 Then
                                            dblVacantAcres = 0
                                        End If
                                    Catch ex As Exception
                                        dblVacantAcres = 0
                                    End Try
                                End If
                                If strFieldName = "DEVD_ACRE" Then
                                    Try
                                        dblDevelopedAcres = CType((pFeat.Value(intDevelopedAcresFld).ToString), Double)
                                        If dblDevelopedAcres < 0 Then
                                            dblDevelopedAcres = 0
                                        End If
                                    Catch ex As Exception
                                        dblDevelopedAcres = 0
                                    End Try
                                End If

                                'ADD VALUES TO SCENARIO TABLE
                                If intDevTypeRow >= 0 Then
                                    dblPreviousValue = 0
                                    rowScLookUp = tblScLookUpTbl.GetRow(intDevTypeRow)
                                    Try
                                        dblPreviousValue = rowScLookUp.Value(intScLookUpFld)
                                    Catch ex As Exception
                                        dblPreviousValue = 0
                                    End Try
                                    If dblPreviousValue >= 0 Then
                                        rowScLookUp.Value(intScLookUpFld) = dblPreviousValue + dblValue
                                    Else
                                        rowScLookUp.Value(intScLookUpFld) = dblValue
                                    End If
                                    rowScLookUp.Store()
                                End If
                            End If
                        End If
                    Next
                End If

                'APPLY THE DEVELOPMENT TYPE ATTTIBUTES IF THE CALCS ARE ON AND THE DEV TYPE RECORD WAS FOUND
                If m_blnRunCalcs Then
                    If intDevTypeRow = -1 Then
                        If m_arrWriteDevTypeFields.Count > 0 Then
                            For intFldCount = 0 To m_arrWriteDevTypeFields.Count - 1
                                strFromFieldName = m_arrWriteDevTypeFields.Item(intFldCount)
                                Try
                                    If pFeatClass.Fields.Field(pFeatClass.FindField(strFromFieldName)).Type = esriFieldType.esriFieldTypeInteger Then
                                        pFeat.Value(pFeatClass.FindField(strFromFieldName)) = 0
                                    ElseIf pFeatClass.Fields.Field(pFeatClass.FindField(strFromFieldName)).Type = esriFieldType.esriFieldTypeDouble Then
                                        pFeat.Value(pFeatClass.FindField(strFromFieldName)) = 0
                                    ElseIf pFeatClass.Fields.Field(pFeatClass.FindField(strFromFieldName)).Type = esriFieldType.esriFieldTypeSingle Then
                                        pFeat.Value(pFeatClass.FindField(strFromFieldName)) = 0
                                    ElseIf pFeatClass.Fields.Field(pFeatClass.FindField(strFromFieldName)).Type = esriFieldType.esriFieldTypeSmallInteger Then
                                        pFeat.Value(pFeatClass.FindField(strFromFieldName)) = 0
                                    ElseIf pFeatClass.Fields.Field(pFeatClass.FindField(strFromFieldName)).Type = esriFieldType.esriFieldTypeString Then
                                        pFeat.Value(pFeatClass.FindField(strFromFieldName)) = ""
                                    End If
                                Catch ex As Exception
                                End Try
                            Next
                        End If

                        If m_arrWriteDevTypeAcresFields.Count > 0 Then
                            For intFldCount = 0 To m_arrWriteDevTypeAcresFields.Count - 1
                                strFromFieldName = m_arrWriteDevTypeAcresFields.Item(intFldCount)
                                strToFieldName = m_arrWriteDevTypeAcresFieldsAltName.Item(intFldCount)
                                If strToFieldName = "" Then
                                    strToFieldName = strFromFieldName
                                End If
                                Try
                                    If intRedevFld <= -1 Then
                                        pFeat.Value(pFeatClass.FindField(strFromFieldName)) = 0
                                    Else
                                        pFeat.Value(pFeatClass.FindField(strToFieldName)) = 0
                                        pFeat.Value(pFeatClass.FindField(strFromFieldName)) = 0
                                    End If
                                Catch ex As Exception
                                End Try
                            Next
                        End If
                    Else
                        If m_arrWriteDevTypeFields.Count > 0 Then
                            For intFldCount = 0 To m_arrWriteDevTypeFields.Count - 1
                                strFromFieldName = m_arrWriteDevTypeFields.Item(intFldCount)
                                Try
                                    pFeat.Value(pFeatClass.FindField(strFromFieldName)) = rowDevType.Value(m_tblDevelopmentTypes.FindField(strFromFieldName))
                                Catch ex As Exception
                                    If pFeatClass.Fields.Field(pFeatClass.FindField(strFromFieldName)).Type = esriFieldType.esriFieldTypeInteger Then
                                        pFeat.Value(pFeatClass.FindField(strFromFieldName)) = 0
                                    ElseIf pFeatClass.Fields.Field(pFeatClass.FindField(strFromFieldName)).Type = esriFieldType.esriFieldTypeDouble Then
                                        pFeat.Value(pFeatClass.FindField(strFromFieldName)) = 0
                                    ElseIf pFeatClass.Fields.Field(pFeatClass.FindField(strFromFieldName)).Type = esriFieldType.esriFieldTypeSingle Then
                                        pFeat.Value(pFeatClass.FindField(strFromFieldName)) = 0
                                    ElseIf pFeatClass.Fields.Field(pFeatClass.FindField(strFromFieldName)).Type = esriFieldType.esriFieldTypeSmallInteger Then
                                        pFeat.Value(pFeatClass.FindField(strFromFieldName)) = 0
                                    ElseIf pFeatClass.Fields.Field(pFeatClass.FindField(strFromFieldName)).Type = esriFieldType.esriFieldTypeString Then
                                        pFeat.Value(pFeatClass.FindField(strFromFieldName)) = ""
                                    End If
                                End Try
                            Next
                        End If

                        If m_arrWriteDevTypeAcresFields.Count > 0 Then
                            For intFldCount = 0 To m_arrWriteDevTypeAcresFields.Count - 1
                                strFromFieldName = m_arrWriteDevTypeAcresFields.Item(intFldCount)
                                strToFieldName = m_arrWriteDevTypeAcresFieldsAltName.Item(intFldCount)
                                strDepVarFieldName = m_arrWriteDevTypeAcres2ndVarFields.Item(intFldCount)
                                If strToFieldName = "" Then
                                    strToFieldName = strFromFieldName
                                End If
                                Try
                                    If intRedevFld <= -1 Then
                                        pFeat.Value(pFeatClass.FindField(strFromFieldName)) = 0
                                    Else
                                        dblXValue = CDbl(rowDevType.Value(m_tblDevelopmentTypes.FindField(strFromFieldName)))
                                        dblExistingXValue = 0
                                        dblCalcAcres = 0
                                        dblDepVar = 0
                                        dblScExValue = 0
                                        Try
                                            'dblCalcAcres = (dblXValue * dblVacantAcres) + (dblXValue * dblRedevRate * dblDevelopedAcres)
                                            'dblCalcAcres = (dblXValue * dblVacantAcres) + ((dblXValue / dblNetAcre) * dblDevelopedAcres * dblRedevRate)
                                            If strDepVarFieldName.Length > 0 Then
                                                Try
                                                    dblDepVar = CDbl(rowDevType.Value(m_tblDevelopmentTypes.FindField(strDepVarFieldName)))
                                                    dblCalcAcres = ((dblVacantAcres * dblNetAcre * dblDepVar) + (dblDevelopedAcres * dblRedevRate * dblDepVar)) * dblXValue
                                                Catch ex As Exception
                                                    dblCalcAcres = 0
                                                End Try
                                            Else
                                                dblCalcAcres = (dblVacantAcres * dblNetAcre * dblXValue) + (dblDevelopedAcres * dblRedevRate * dblXValue)
                                            End If
                                        Catch ex As Exception
                                            dblCalcAcres = 0
                                        End Try

                                        pFeat.Value(pFeatClass.FindField(strToFieldName)) = dblCalcAcres
                                        If m_arrWriteDevTypeAcresFieldsOnly.IndexOf(strToFieldName) <= -1 Then
                                            pFeat.Value(pFeatClass.FindField(strFromFieldName)) = dblXValue
                                        End If

                                        'APPLY CALCULATE VALUE TO THE SCENARIO TABLE
                                        If dblExistingXValue > 0 Then
                                            intScExField = rowScLookUp.Fields.FindField("EX_" & strToFieldName)
                                            If intScExField > -1 Then
                                                Try
                                                    dblScExValue = rowScLookUp.Value(intScExField)
                                                    dblScExValue = RoundNumber(dblScExValue, 6)
                                                Catch ex As Exception
                                                    dblScExValue = 0
                                                End Try
                                                If dblScExValue >= 0 Then
                                                    rowScLookUp.Value(intScExField) = dblScExValue + dblExistingXValue
                                                Else
                                                    rowScLookUp.Value(intScExField) = dblScExValue
                                                End If
                                            End If
                                            rowScLookUp.Store()
                                        End If
                                    End If
                                Catch ex As Exception
                                    pFeat.Value(pFeatClass.FindField(strToFieldName)) = 0
                                    If m_arrWriteDevTypeAcresFieldsOnly.IndexOf(strToFieldName) <= -1 Then
                                        pFeat.Value(pFeatClass.FindField(strFromFieldName)) = 0
                                    End If
                                End Try
                            Next
                        End If
                    End If
                End If

                m_dockEnvisionWinForm.prgBarEnvision.Value = (intCount / intTotalCount) * 100
                m_dockEnvisionWinForm.prgBarEnvision.Refresh()
                pFeat.Store()
                pFeat = pFeatureCursor.NextFeature
            Loop

            'STOP THE EDITTING IF STARTED JUST FOR THE SYNCHRONIZE
            If blnTempEdit Then
                EditSession()
                blnTempEdit = False
            End If

            'UPDATE THE INDICATOR TABLE SELECTION
            UpdateSummaryTableSelection()
            'WRITE VALUES TO ENVISION EXCEL FILE
            Write2EnvisionExcelFile()

            GoTo CleanUp
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Attribute Synchronize Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo CleanUp
        End Try

CleanUp:
        mxApplication = Nothing
        mxDoc = Nothing
        pREnv = Nothing
        pSpatFilter = Nothing
        pFeatSelection = Nothing
        pCursor = Nothing
        pFeatureCursor = Nothing
        pFeat = Nothing
        pFeatClass = Nothing
        pTable = Nothing
        pFeatTable = Nothing
        intTotalCount = Nothing
        intCount = Nothing
        strDevType = Nothing
        dblVacantAcres = Nothing
        dblDevelopedAcres = Nothing
        intDevTypeFld = Nothing
        intVacantAcresFld = Nothing
        intDevelopedAcresFld = Nothing
        intRedevFld = Nothing
        rowDevType = Nothing
        intDevTypeRow = Nothing
        intRow = Nothing
        rowTemp = Nothing
        intScVacLookUpFld = Nothing
        intScDevdLookUpFld = Nothing
        rowScLookUp = Nothing
        dblScExValue = Nothing
        intScExField = Nothing
        tblScLookUpTbl = Nothing
        strFromFieldName = Nothing
        strToFieldName = Nothing
        dblCalcAcres = Nothing
        dblRedevRate = Nothing
        dblXValue = Nothing
        dblExistingXValue = Nothing
        intFldCount = Nothing
        pDef = Nothing
        pQFilter = Nothing
        strDefExpression = Nothing
        strQString = Nothing
        intScLookUpFld = Nothing
        strFieldName = Nothing
        fldSc = Nothing
        dblPreviousValue = Nothing
        dblValue = Nothing
        strDepVarFieldName = Nothing
        dblDepVar = Nothing

        Me.prgBarEnvision.Value = 0
        m_blnSynchronizePrompt = False
        GC.Collect()
        GC.WaitForPendingFinalizers()
    End Sub

    Private Sub itmSynchronizeExisting_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles itmSynchronizeExisting.Click
        'STOP EDIT SESSION IF OPEN
        Dim blnEditting As Boolean = False
        If Not Me.btnEditing.Text = "Start Edit" Then
            blnTempEdit = False
            blnEditting = True
            EditSession()
        End If
        DeleteEnvisionTempSummaryTables(True)
        EnvisionSyncEX_LU()
        EnvisionSyncEX_LU_CreateSummaryTables()
'        'STOP THE EDITTING IF STARTED JUST FOR THE SYNCHRONIZE
        If blnEditting Then
            EditSession()
            blnTempEdit = True
        End If
    End Sub

    Public Sub EnvisionSyncEX_LU_CreateSummaryTables()
        'CHECK FOR ENVISION EDIT LAYER
        If m_pEditFeatureLyr Is Nothing Then
            GoTo CleanUp
        End If

        Dim pLayer As IFeatureLayer = m_pEditFeatureLyr
        Dim pCursor As ICursor = Nothing
        Dim pFeatClass As IFeatureClass = pLayer.FeatureClass
        Dim pTable As ITable = CType(pFeatClass, ITable)
        Dim pFeatureCursor As IFeatureCursor
        Dim intRow As Integer
        Dim intTblRow As Integer
        Dim intFieldRow As Integer = -1
        Dim intStartRow As Integer = -1
        Dim strFldCellValue As String = ""
        Dim intCol As Integer = 0
        Dim shtExisting As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        Dim pDef As IFeatureLayerDefinition2
        Dim pQFilter As IQueryFilter
        Dim strDefExpression As String = ""
        Dim strQString As String = Nothing
        Dim intCount As Integer = 0
        Dim intTotalCount As Integer = 0
        Dim pGeoProc As IBasicGeoprocessor
        Dim strExLUFld As String = ""
        Dim strDevdAcresAllFld As String = ""
        Dim strDevdAcresPaintedFld As String = ""
        Dim intExLUCol As Integer = 0
        Dim intDevdAcresAllCol As Integer = -1
        Dim intDevdAcresPaintedCol As Integer = -1
        Dim pSumTable As ITable = Nothing
        Dim pWkSpName As IName
        Dim pOutTabName As ITableName
        Dim pOutDatasetName As IDatasetName
        Dim pDataSet As IDataset
        Dim pWkSpDS As IDataset
        Dim pFeatSelection As IFeatureSelection
        Dim rowAll As IRow
        Dim rowPainted As IRow
        Dim intExcelFormulaCalc As Integer

        'RETRIEVE OR OPEN THE SELECTED ENVISION EXCEL FILE
        Try
            If m_strEnvisionExcelFile = "" Or m_dockEnvisionWinForm.dgvDevelopmentTypes.Rows.Count <= 0 Then
                'MessageBox.Show("No Envision Excel file has been selected.  Use the 'Open Envision Excel File | Select' option to define an Envision Excel file.", "No Envision Excel File", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
                GoTo CleanUp
            End If
            Try
                If m_xlPaintApp Is Nothing Then
                    m_xlPaintApp = New Microsoft.Office.Interop.Excel.Application
                    m_xlPaintApp.WindowState = Interop.Excel.XlWindowState.xlMinimized
                    m_xlPaintApp.Visible = True
                End If
                If m_xlPaintWB1 Is Nothing Then
                    m_xlPaintWB1 = CType(m_xlPaintApp.Workbooks.Open(m_strEnvisionExcelFile), Microsoft.Office.Interop.Excel.Workbook)
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Opening Envision Excel File Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
                GoTo CleanUp
            End Try
            'RETRIEVE THE EXISTING CONDITIONS TAB
            Try
                shtExisting = m_xlPaintWB1.Sheets("Existing Developed Area")
                If Not TypeOf m_xlPaintWB1.Sheets("Existing Developed Area") Is Microsoft.Office.Interop.Excel.Worksheet Then
                    'MessageBox.Show("The 'Existing Developed Area' tab could not be found in the selected ENVISION Excel file. Please select another file.", "TAB NOT FOUND", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
                    GoTo CleanUp
                Else
                    shtExisting = m_xlPaintWB1.Sheets("Existing Developed Area")
                End If
            Catch ex As Exception
                'MessageBox.Show("The 'Existing Developed Area' tab could not be found in the selected ENVISION Excel file. Please select another file.", "TAB NOT FOUND", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
                GoTo CleanUp
            End Try
        Catch ex As Exception
            CloseEnvisionExcel()
        End Try

        'DETERMINE THE CURRENT FORMULA CALC SETTING TO RESET AFTER FUNCTION EXECUTES
        If m_xlPaintApp.Calculation = Interop.Excel.XlCalculation.xlCalculationAutomatic Then
            intExcelFormulaCalc = 1
        ElseIf m_xlPaintApp.Calculation = Interop.Excel.XlCalculation.xlCalculationManual Then
            intExcelFormulaCalc = 2
        ElseIf m_xlPaintApp.Calculation = Interop.Excel.XlCalculation.xlCalculationSemiautomatic Then
            intExcelFormulaCalc = 3
        End If
        'SET EXCEL FORMULA CALC TO MANUAL
        m_xlPaintApp.Calculation = Interop.Excel.XlCalculation.xlCalculationManual

        'FIND THE STARTING POINT
        For intRow = 1 To 10
            strFldCellValue = CStr(CType(shtExisting.Cells(intRow, 1), Microsoft.Office.Interop.Excel.Range).Value)
            If Not strFldCellValue Is Nothing Then
                If UCase(strFldCellValue) = "EXISTING" Then
                    intFieldRow = intRow
                    intStartRow = intRow + 1
                    Exit For
                End If
            End If
        Next

        If intFieldRow = -1 Then
            'MessageBox.Show("The 'Existing Developed Area' tab does not appear to be formatted correctly.", "Value Entry 1 not Found", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo CleanUp
        End If

        'SELECT ALL FEATURES / APPLY QUERY DEFINITION IF AVAILABLE
        pFeatClass = pLayer.FeatureClass
        pDef = pLayer
        strDefExpression = pDef.DefinitionExpression
        pQFilter = New QueryFilter
        If strDefExpression.Length > 0 Then
            pQFilter.WhereClause = pDef.DefinitionExpression
            pCursor = pFeatClass.Search(pQFilter, False)
        Else
            pCursor = pFeatClass.Search(Nothing, False)
        End If

        'CYCLE THROUGH THE FIRST 25 COULMNS FOR THE INPUT FIELDS, FIRST FOUND IS EXISTING LAND USE FIELD, SECOND IS DEVELOPED ACREAS FOR ALL FEATURES, THIRD IS DEVELOPED ACREAS FOR PAINTED FEATURES
        intCount = 0
        intTotalCount = 25
        For intCol = 1 To 25
            intCount = intCount + 1
            strFldCellValue = CStr(CType(shtExisting.Cells(intFieldRow, intCol), Microsoft.Office.Interop.Excel.Range).Value)
            If Not strFldCellValue Is Nothing Then
                'If pFeatClass.FindField(strFldCellValue) >= 0 Then
                If strExLUFld.Length <= 0 And Not UCase(strFldCellValue) = "EXISTING" And strFldCellValue.Length > 0 Then
                    strExLUFld = strFldCellValue
                    intExLUCol = intCol
                ElseIf strDevdAcresAllFld.Length <= 0 And UCase(strFldCellValue) = "DEVD_ACRE" And strFldCellValue.Length > 0 Then
                    strDevdAcresAllFld = strFldCellValue
                    intDevdAcresAllCol = intCol
                ElseIf strDevdAcresPaintedFld.Length <= 0 And UCase(strFldCellValue) = ("PAINTED" & m_intEditScenario.ToString) And strFldCellValue.Length > 0 Then
                    strDevdAcresPaintedFld = "DEVD_ACRE"
                    intDevdAcresPaintedCol = intCol
                End If
                'End If
                If strExLUFld.Length > 0 And strDevdAcresAllFld.Length > 0 And strDevdAcresPaintedFld.Length > 0 Then
                    Exit For
                End If
            End If
        Next
        If intDevdAcresAllCol = -1 Or intDevdAcresPaintedCol = -1 Then
            'MessageBox.Show("The 'Existing Developed Area' tab does not appear to be formatted correctly.", "Value Entry 2 or 3 not Found", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo CleanUp
        End If

        pTable = CType(pFeatClass, ITable)
        pDataSet = pTable
        pWkSpDS = pDataSet.Workspace
        pWkSpName = pWkSpDS.FullName

        'FIRST PROCESS THE DEV TYPE PAINTED FEATURES ONLY
        Try
            pFeatClass = pLayer.FeatureClass
            pDef = pLayer
            strDefExpression = pDef.DefinitionExpression
            pQFilter = New QueryFilter
            strQString = "NOT DEV_TYPE = ''"
            If strDefExpression.Length > 0 Then
                m_appEnvision.StatusBar.Message(0) = "Existing Developed Area (PAINTED): Search includes Subarea Definition: " & "(" & strQString & ") AND (" & pDef.DefinitionExpression & ")"
                pQFilter.WhereClause = "(" & strQString & ") AND (" & pDef.DefinitionExpression & ")"
                pFeatSelection = CType(pLayer, IFeatureSelection)
                pFeatSelection.SelectFeatures(pQFilter, esriSelectionResultEnum.esriSelectionResultNew, False)
            Else
                m_appEnvision.StatusBar.Message(0) = "Existing Developed Area (PANITED): Full Area "
                pQFilter.WhereClause = strQString
                pFeatSelection = CType(pLayer, IFeatureSelection)
                pFeatSelection.SelectFeatures(pQFilter, esriSelectionResultEnum.esriSelectionResultNew, False)
            End If

            intTotalCount = pFeatSelection.SelectionSet.Count
            pTable = CType(pFeatClass, ITable)


            pOutTabName = Nothing
            pOutDatasetName = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()

            pOutTabName = New TableName
            pOutDatasetName = pOutTabName
            pOutDatasetName.Name = "EXISTING_DEVELOPED_AREA_PAINTED"
            pOutDatasetName.WorkspaceName = pWkSpName

            pGeoProc = New BasicGeoprocessor
            If strExLUFld.Length > 0 And strDevdAcresPaintedFld.Length > 0 Then
                m_appEnvision.StatusBar.Message(0) = "Creating summary table for all PAINTED features."
                pSumTable = pGeoProc.Dissolve(pLayer, True, strExLUFld, ("Minimum." & strExLUFld & ",Sum." & strDevdAcresPaintedFld), pOutDatasetName)
            End If
        Catch ex As Exception

        End Try

        'SECOND PROCESS THE FULL DATASET 
        Try
            'pFeatClass = pLayer.FeatureClass
            'pDef = pLayer
            'strDefExpression = pDef.DefinitionExpression
            'pQFilter = New QueryFilter
            'If strDefExpression.Length > 0 Then
            '    pQFilter.WhereClause = pDef.DefinitionExpression
            '    pFeatureCursor = pFeatClass.Search(pQFilter, False)
            '    intTotalCount = pFeatClass.FeatureCount(pQFilter)
            'Else
            '    pFeatureCursor = pFeatClass.Search(Nothing, False)
            '    intTotalCount = pFeatClass.FeatureCount(Nothing)
            'End If

            pFeatClass = pLayer.FeatureClass
            pDef = pLayer
            strDefExpression = pDef.DefinitionExpression
            pQFilter = New QueryFilter
            If strDefExpression.Length > 0 Then
                m_appEnvision.StatusBar.Message(0) = "Existing Developed Area (ALL): Search includes Subarea Definition: " & "(" & strQString & ") AND (" & pDef.DefinitionExpression & ")"
                pQFilter.WhereClause = pDef.DefinitionExpression
                pFeatSelection = CType(pLayer, IFeatureSelection)
                pFeatSelection.SelectFeatures(pQFilter, esriSelectionResultEnum.esriSelectionResultNew, False)
            Else
                m_appEnvision.StatusBar.Message(0) = "Existing Developed Area (ALL): Full Area "
                pQFilter.WhereClause = strQString
                pFeatSelection = CType(pLayer, IFeatureSelection)
                pFeatSelection.SelectFeatures(pQFilter, esriSelectionResultEnum.esriSelectionResultNew, False)
            End If

            intTotalCount = pFeatSelection.SelectionSet.Count
            pTable = CType(pFeatClass, ITable)

            pOutTabName = New TableName
            pOutDatasetName = pOutTabName
            pOutDatasetName.Name = "EXISTING_DEVELOPED_AREA"
            pOutDatasetName.WorkspaceName = pWkSpName

            pGeoProc = New BasicGeoprocessor
            If strExLUFld.Length > 0 And strDevdAcresAllFld.Length > 0 Then
                m_appEnvision.StatusBar.Message(0) = "Creating summary table for all ALL features."
                m_tblExistingLU = pGeoProc.Dissolve(pLayer, False, strExLUFld, ("Minimum." & strExLUFld & ",Sum." & strDevdAcresAllFld), pOutDatasetName)
            End If
        Catch ex As Exception

        End Try

        pGeoProc = New BasicGeoprocessor
        pGeoProc = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()

        'CLEAR VALUES
        m_appEnvision.StatusBar.Message(0) = "Clearing previous developed acre values from excel table."
        For intRow = intStartRow To (intStartRow + 20)
            strFldCellValue = CStr(CType(shtExisting.Cells(intRow, intExLUCol), Microsoft.Office.Interop.Excel.Range).Value)
            If Not strFldCellValue Is Nothing Then
                'ALL FEATURES
                CType(shtExisting.Cells(intRow, intDevdAcresAllCol), Microsoft.Office.Interop.Excel.Range).Value = ""
                'PAINTED FEATURES
                CType(shtExisting.Cells(intRow, intDevdAcresPaintedCol), Microsoft.Office.Interop.Excel.Range).Value = ""
            End If
        Next

        'CODE TO WRITE THE SUMMARY VALUES TO EXCEL FILE
        For intRow = intStartRow To (intStartRow + 20)
            strFldCellValue = CStr(CType(shtExisting.Cells(intRow, intExLUCol), Microsoft.Office.Interop.Excel.Range).Value)
            m_appEnvision.StatusBar.Message(0) = "Pulling SUM values for LU, " & strFldCellValue & "."
            If Not strFldCellValue Is Nothing Then
                'ALL FEATURES
                If Not m_tblExistingLU Is Nothing Then
                    For intTblRow = 1 To m_tblExistingLU.RowCount(Nothing)
                        Try
                            If strFldCellValue = CStr(m_tblExistingLU.GetRow(intTblRow).Value(1)) Then
                                CType(shtExisting.Cells(intRow, intDevdAcresAllCol), Microsoft.Office.Interop.Excel.Range).Value = m_tblExistingLU.GetRow(intTblRow).Value(2)
                            End If
                        Catch ex As Exception

                        End Try
                    Next
                End If
                'PAINTED FEATURES
                Try
                    If Not pSumTable Is Nothing Then
                        For intTblRow = 1 To pSumTable.RowCount(Nothing)
                            Try
                                If strFldCellValue = CStr(pSumTable.GetRow(intTblRow).Value(1)) Then
                                    CType(shtExisting.Cells(intRow, intDevdAcresPaintedCol), Microsoft.Office.Interop.Excel.Range).Value = pSumTable.GetRow(intTblRow).Value(2)
                                End If
                            Catch ex As Exception

                            End Try
                        Next
                    End If
                Catch ex As Exception
                End Try
                
            End If
        Next

        'RESET FORMULA CALC SETTING
        If intExcelFormulaCalc = 1 Then
            m_xlPaintApp.Calculation = Interop.Excel.XlCalculation.xlCalculationAutomatic
        ElseIf intExcelFormulaCalc = 2 Then
            m_xlPaintApp.Calculation = Interop.Excel.XlCalculation.xlCalculationManual
        ElseIf intExcelFormulaCalc = 3 Then
            m_xlPaintApp.Calculation = Interop.Excel.XlCalculation.xlCalculationSemiautomatic
        End If

        m_xlPaintWB1.Save()

        'SYNCHRONIZE THE ALL AND PAINTED TABLES BY OVERWRITING THE PAINTED WITH ALL NUMBERS
        Try
            If Not m_tblExistingLU Is Nothing Then
                'CLEAR ALL RECORDS OF SUMMARY VALUES
                For intTblRow = 1 To m_tblExistingLU.RowCount(Nothing)
                    rowAll = m_tblExistingLU.GetRow(intTblRow)
                    rowAll.Value(2) = 0
                    rowAll.Store()
                Next
                'LOAD THE PAINTED SUMMARY INTO ALL EX LU TABLE
                For intRow = 1 To pSumTable.RowCount(Nothing)
                    rowPainted = pSumTable.GetRow(intRow)
                    strFldCellValue = rowPainted.Value(1)
                    For intTblRow = 1 To m_tblExistingLU.RowCount(Nothing)
                        rowAll = m_tblExistingLU.GetRow(intTblRow)
                        If strFldCellValue = rowAll.Value(1) Then
                            rowAll.Value(2) = rowPainted.Value(2)
                            rowAll.Store()
                        End If
                    Next
                Next
            End If
        Catch ex As Exception

        End Try

        'DELETE THE PAINTED SUMMARY TABLE
        DeleteEnvisionTempSummaryTables(False)

        GoTo CleanUp

CleanUp:
        m_dockEnvisionWinForm.prgBarEnvision.Value = 0
        m_dockEnvisionWinForm.prgBarEnvision.Refresh()
        pLayer = Nothing
        pCursor = Nothing
        pFeatClass = Nothing
        pTable = Nothing
        pFeatureCursor = Nothing
        intRow = Nothing
        intTblRow = Nothing
        intFieldRow = Nothing
        intStartRow = Nothing
        strFldCellValue = Nothing
        intCol = Nothing
        shtExisting = Nothing
        pDef = Nothing
        pQFilter = Nothing
        strDefExpression = Nothing
        strQString = Nothing
        intCount = Nothing
        intTotalCount = Nothing
        pGeoProc = Nothing
        strExLUFld = Nothing
        strDevdAcresAllFld = Nothing
        strDevdAcresPaintedFld = Nothing
        intExLUCol = Nothing
        intDevdAcresAllCol = Nothing
        intDevdAcresPaintedCol = Nothing
        pSumTable = Nothing
        pWkSpName = Nothing
        pOutTabName = Nothing
        pOutDatasetName = Nothing
        pDataSet = Nothing
        pWkSpDS = Nothing
        pFeatSelection = Nothing
        intExcelFormulaCalc = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub

    Public Sub DeleteEnvisionTempSummaryTables(ByVal blnBoth As Boolean)
        'DELETE THE SUMMARY TABLES
        Dim GP As Geoprocessor
        Dim pDeleteTable As ESRI.ArcGIS.DataManagementTools.Delete
        GP = New Geoprocessor


        'REMOVE PRE-EXISTING INSTANCES OF SUMMARY TABLES
        If m_strFeaturePath.Length > 0 Then
            GC.WaitForPendingFinalizers()
            GC.Collect()
            If blnBoth Then
                'CLEAR THE GLOBAL VARIABLE
                m_tblExistingLU = Nothing
                GC.Collect()
                pDeleteTable = New ESRI.ArcGIS.DataManagementTools.Delete
                pDeleteTable.in_data = m_strFeaturePath & "\" & "EXISTING_DEVELOPED_AREA"
                pDeleteTable.data_type = "Table"
                RunTool(GP, pDeleteTable)
            End If
            pDeleteTable = New ESRI.ArcGIS.DataManagementTools.Delete
            pDeleteTable.in_data = m_strFeaturePath & "\" & "EXISTING_DEVELOPED_AREA_PAINTED"
            pDeleteTable.data_type = "Table"
            RunTool(GP, pDeleteTable)
        End If

        GP = Nothing
        pDeleteTable = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub

    Public Sub EnvisionSyncEX_LU()
        'CHECK FOR ENVISION EDIT LAYER
        If m_pEditFeatureLyr Is Nothing Then
            GoTo CleanUp
        End If

        Dim pLayer As IFeatureLayer = m_pEditFeatureLyr
        Dim pCursor As ICursor = Nothing
        Dim pFeatClass As IFeatureClass = pLayer.FeatureClass
        Dim pTable As ITable = CType(pFeatClass, ITable)
        Dim intRow As Integer
        Dim intFieldRow As Integer = 6
        Dim intStartRow As Integer = 7
        Dim strFldCellValue As String = ""
        Dim intCol As Integer = 0
        Dim shtExisting As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        Dim pData As IDataStatistics
        Dim pDef As IFeatureLayerDefinition2
        Dim pQFilter As IQueryFilter
        Dim strDefExpression As String = ""
        Dim strQString As String = Nothing
        Dim intCount As Integer = 0
        Dim intTotalCount As Integer = 0
        Dim intExcelFormulaCalc As Integer

        'RETRIEVE OR OPEN THE SELECTED ENVISION EXCEL FILE
        Try
            m_appEnvision.StatusBar.Message(0) = "Checking selected excel file for tab, Existing Land Use."
            If m_strEnvisionExcelFile = "" Or m_dockEnvisionWinForm.dgvDevelopmentTypes.Rows.Count <= 0 Then
                'MessageBox.Show("No Envision Excel file has been selected.  Use the 'Open Envision Excel File | Select' option to define an Envision Excel file.", "No Envision Excel File", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
                GoTo CleanUp
            End If
            Try
                If m_xlPaintApp Is Nothing Then
                    m_xlPaintApp = New Microsoft.Office.Interop.Excel.Application
                    m_xlPaintApp.WindowState = Interop.Excel.XlWindowState.xlMinimized
                    m_xlPaintApp.Visible = True
                End If
                If m_xlPaintWB1 Is Nothing Then
                    m_xlPaintWB1 = CType(m_xlPaintApp.Workbooks.Open(m_strEnvisionExcelFile), Microsoft.Office.Interop.Excel.Workbook)
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Opening Envision Excel File Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
                GoTo CloseExcel
            End Try
            'RETRIEVE THE EXISTING CONDITIONS TAB
            Try
                shtExisting = m_xlPaintWB1.Sheets("Existing Land Use")
                If Not TypeOf m_xlPaintWB1.Sheets("Existing Land Use") Is Microsoft.Office.Interop.Excel.Worksheet Then
                    MessageBox.Show("The 'Existing Land Use' tab could not be found in the selected ENVISION Excel file. Please select another file.", "TAB NOT FOUND", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
                    GoTo CloseExcel
                Else
                    shtExisting = m_xlPaintWB1.Sheets("Existing Land Use")
                End If
            Catch ex As Exception
                MessageBox.Show("The 'Existing Land Use' tab could not be found in the selected ENVISION Excel file. Please select another file.", "TAB NOT FOUND", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
                GoTo CloseExcel
            End Try
        Catch ex As Exception
            CloseEnvisionExcel()
            OpenEnvisionExcel()
        End Try

        'DETERMINE THE CURRENT FORMULA CALC SETTING TO RESET AFTER FUNCTION EXECUTES
        If m_xlPaintApp.Calculation = Interop.Excel.XlCalculation.xlCalculationAutomatic Then
            intExcelFormulaCalc = 1
        ElseIf m_xlPaintApp.Calculation = Interop.Excel.XlCalculation.xlCalculationManual Then
            intExcelFormulaCalc = 2
        ElseIf m_xlPaintApp.Calculation = Interop.Excel.XlCalculation.xlCalculationSemiautomatic Then
            intExcelFormulaCalc = 3
        End If
        'SET EXCEL FORMULA CALC TO MANUAL
        m_xlPaintApp.Calculation = Interop.Excel.XlCalculation.xlCalculationManual

        'FIND THE STARTING POINT
        For intRow = 1 To 10
            strFldCellValue = CStr(CType(shtExisting.Cells(intRow, 1), Microsoft.Office.Interop.Excel.Range).Value)
            If Not strFldCellValue Is Nothing Then
                If UCase(strFldCellValue) = "EXISTING" Then
                    intFieldRow = intRow
                    intStartRow = intRow + 1
                    Exit For
                End If
            End If
        Next

        'CHECK FOR REQUIRED FIELDS
        If pTable Is Nothing Then
            MessageBox.Show("The Envision Edit Layer could not be found in the current view document.  Please reselect the Envision layer.", "Layer Not Found", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo CleanUp
        End If

        'SELECT ALL FEATURES / APPLY QUERY DEFINITION IF AVAILABLE
        pFeatClass = pLayer.FeatureClass
        pDef = pLayer
        strDefExpression = pDef.DefinitionExpression
        pQFilter = New QueryFilter

        If strDefExpression.Length > 0 Then
            m_appEnvision.StatusBar.Message(0) = "Selecting ALL features with Subarea Definition"
            pQFilter.WhereClause = pDef.DefinitionExpression
            pCursor = pFeatClass.Search(pQFilter, False)
        Else
            m_appEnvision.StatusBar.Message(0) = "Selecting ALL features"
            pCursor = pFeatClass.Search(Nothing, False)
        End If

        'CYCLE THROUGH THE FIRST 25 COULMNS FOR FIELDS TO RETURN STATISTCS FOR
        intCount = 0
        intTotalCount = 50
        m_dockEnvisionWinForm.prgBarEnvision.Value = (intCount / intTotalCount) * 100
        m_dockEnvisionWinForm.prgBarEnvision.Refresh()
        For intCol = 1 To 50
            intCount = intCount + 1
            strFldCellValue = CStr(CType(shtExisting.Cells(intFieldRow, intCol), Microsoft.Office.Interop.Excel.Range).Value)
            m_appEnvision.StatusBar.Message(0) = "Calculating SUM value for field, " & strFldCellValue & "."
            If Not strFldCellValue Is Nothing Then
                If UCase(strFldCellValue) = "EXISTING" Then
                    Continue For
                End If
                Try
                    If strDefExpression.Length > 0 Then
                        pQFilter.WhereClause = pDef.DefinitionExpression
                        pCursor = pFeatClass.Search(pQFilter, False)
                    Else
                        pCursor = pFeatClass.Search(Nothing, False)
                    End If
                    pData = New DataStatistics
                    pData.Cursor = pCursor
                    pData.Field = strFldCellValue
                    CType(shtExisting.Cells(intStartRow, intCol), Microsoft.Office.Interop.Excel.Range).Value = pData.Statistics.Sum
                Catch ex As Exception
                    CType(shtExisting.Cells(intStartRow, intCol), Microsoft.Office.Interop.Excel.Range).Value = 0
                End Try
            End If
            m_dockEnvisionWinForm.prgBarEnvision.Value = (intCount / intTotalCount) * 100
            m_dockEnvisionWinForm.prgBarEnvision.Refresh()
        Next

        'RESET FORMULA CALC SETTING
        If intExcelFormulaCalc = 1 Then
            m_xlPaintApp.Calculation = Interop.Excel.XlCalculation.xlCalculationAutomatic
        ElseIf intExcelFormulaCalc = 2 Then
            m_xlPaintApp.Calculation = Interop.Excel.XlCalculation.xlCalculationManual
        ElseIf intExcelFormulaCalc = 3 Then
            m_xlPaintApp.Calculation = Interop.Excel.XlCalculation.xlCalculationSemiautomatic
        End If

        m_xlPaintWB1.Save()
        GoTo CleanUp


CloseExcel:
        CloseEnvisionExcel()
        GoTo CleanUp

CleanUp:
        m_dockEnvisionWinForm.prgBarEnvision.Value = 0
        m_dockEnvisionWinForm.prgBarEnvision.Refresh()
        pLayer = Nothing
        pCursor = Nothing
        pFeatClass = Nothing
        pTable = Nothing
        intRow = Nothing
        intFieldRow = Nothing
        intStartRow = Nothing
        strFldCellValue = Nothing
        intCol = Nothing
        shtExisting = Nothing
        pData = Nothing
        pDef = Nothing
        pQFilter = Nothing
        strDefExpression = Nothing
        strQString = Nothing
        intCount = Nothing
        intTotalCount = Nothing
        intExcelFormulaCalc = Nothing

        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub

    Private Sub itmPixelBrush_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles itmPixelBrush.Click
        'Calibrate the brush settings to match the user selection
        Try
            Me.ddbBrushes.Image = Me.itmPixelBrush.Image
            Me.ddbBrushes.ToolTipText = Me.itmPixelBrush.ToolTipText
            m_strEnvisionEditOption = "POINT"
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Envision Tools Pixel Brush Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
        End Try

        Dim pUID As New UID
        Try
            pUID.Value = "Envision_Tools.EnvisionAttributeEditor"
            m_appEnvision.CurrentTool = m_appEnvision.Document.CommandBars.Find(pUID)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Select Button Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
        End Try
    End Sub

    Private Sub itmPolygonBrush_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles itmPolygonBrush.Click
        'Calibrate the brush settings to match the user selection
        Try
            Me.ddbBrushes.Image = Me.itmPolygonBrush.Image
            Me.ddbBrushes.Text = Me.itmPolygonBrush.Text
            Me.ddbBrushes.ToolTipText = Me.itmPolygonBrush.ToolTipText
            m_strEnvisionEditOption = "POLYGON"
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Envision Tools Polygon Brush Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
        End Try

        Dim pUID As New UID
        Try
            pUID.Value = "Envision_Tools.EnvisionAttributeEditor"
            m_appEnvision.CurrentTool = m_appEnvision.Document.CommandBars.Find(pUID)
            'Me.BringToFront()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Select Button Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
        End Try
    End Sub

    Private Sub itmRectBrush_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles itmRectBrush.Click
        'Calibrate the brush settings to match the user selection
        Try
            Me.ddbBrushes.Image = Me.itmRectBrush.Image
            Me.ddbBrushes.Text = Me.itmRectBrush.Text
            Me.ddbBrushes.ToolTipText = Me.itmRectBrush.ToolTipText
            m_strEnvisionEditOption = "RECT"
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Envision Tools Rectangle Brush Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
        End Try

        Dim pUID As New UID
        Try
            pUID.Value = "Envision_Tools.EnvisionAttributeEditor"
            m_appEnvision.CurrentTool = m_appEnvision.Document.CommandBars.Find(pUID)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Select Button Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
        End Try

    End Sub

    Private Sub itmCircleBrush_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles itmCircleBrush.Click
        'Calibrate the brush settings to match the user selection
        Try
            Me.ddbBrushes.Image = Me.itmCircleBrush.Image
            Me.ddbBrushes.Text = Me.itmCircleBrush.Text
            Me.ddbBrushes.ToolTipText = Me.itmCircleBrush.ToolTipText
            m_strEnvisionEditOption = "CIRCLE"
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Envision Tools Circle Brush Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
        End Try
        Dim pUID As New UID

        Try
            pUID.Value = "Envision_Tools.EnvisionAttributeEditor"
            m_appEnvision.CurrentTool = m_appEnvision.Document.CommandBars.Find(pUID)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Select Button Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
        End Try

    End Sub

    Private Sub itmPolylineBrush_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles itmPolylineBrush.Click
        'Calibrate the brush settings to match the user selection
        Try
            Me.ddbBrushes.Image = Me.itmPolylineBrush.Image
            Me.ddbBrushes.Text = Me.itmPolylineBrush.Text
            Me.ddbBrushes.ToolTipText = Me.itmPolylineBrush.ToolTipText
            m_strEnvisionEditOption = "POLYLINE"
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Envision Tools Polyline Brush Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
        End Try

        Dim pUID As New UID
        Try
            pUID.Value = "Envision_Tools.EnvisionAttributeEditor"
            m_appEnvision.CurrentTool = m_appEnvision.Document.CommandBars.Find(pUID)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Select Button Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
        End Try

    End Sub

    Private Sub itmColorSelector_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles itmColorSelector.Click
        'Provide the ESRI Color Selector form for the user to select a Development Type Color
        Dim pInitColor As IRgbColor
        Dim pNewColor As IRgbColor
        Dim pSelector As ESRI.ArcGIS.Framework.IColorSelector
        Dim bColorSet As Boolean
        Dim bmpTemp As Bitmap
        Dim hit As DataGridView.HitTestInfo
        Dim rowDev As IRow
        Try
            'Me.WindowState = FormWindowState.Minimized
            pInitColor = New RgbColor
            ' the dialog will open with red as a default.
            hit = Me.dgvDevelopmentTypes.HitTest(m_pntMouse.X, m_pntMouse.Y)
            pInitColor.Red = CInt(Me.dgvDevelopmentTypes.Rows(hit.RowIndex).Cells(2).Value)
            pInitColor.Green = CInt(Me.dgvDevelopmentTypes.Rows(hit.RowIndex).Cells(4).Value)
            pInitColor.Blue = CInt(Me.dgvDevelopmentTypes.Rows(hit.RowIndex).Cells(4).Value)

            pSelector = New ESRI.ArcGIS.Framework.ColorSelector
            pSelector.Color = pInitColor
            bColorSet = pSelector.DoModal(0)
            If bColorSet Then
                pNewColor = New RgbColor
                pNewColor = CType(pSelector.Color, IRgbColor)
                bmpTemp = CreateColorBMP(pNewColor.Red, pNewColor.Green, pNewColor.Blue)
                If hit.Type = DataGridViewHitTestType.Cell Then
                    Me.dgvDevelopmentTypes.Rows(hit.RowIndex).Cells(0).Value = bmpTemp
                    Me.dgvDevelopmentTypes.Rows(hit.RowIndex).Cells(2).Value = CInt(pNewColor.Red)
                    Me.dgvDevelopmentTypes.Rows(hit.RowIndex).Cells(3).Value = CInt(pNewColor.Green)
                    Me.dgvDevelopmentTypes.Rows(hit.RowIndex).Cells(4).Value = CInt(pNewColor.Blue)
                    Me.dgvDevelopmentTypes.Rows(hit.RowIndex).Cells(0).Selected = True
                    'UPDATE THE LOOKUP TABLE
                    If Not m_tblDevelopmentTypes Is Nothing Then
                        rowDev = m_tblDevelopmentTypes.GetRow(hit.RowIndex)
                        rowDev.Value(m_tblDevelopmentTypes.FindField("RED")) = CInt(pNewColor.Red)
                        rowDev.Value(m_tblDevelopmentTypes.FindField("GREEN")) = CInt(pNewColor.Green)
                        rowDev.Value(m_tblDevelopmentTypes.FindField("BLUE")) = CInt(pNewColor.Blue)
                        rowDev.Store()
                    End If
                    If Me.itmAutoUpdateLegend.Checked Then
                        UpdateEnvisionLyrLegend()
                    End If
                End If
                'UPDATE THE ENVISION EXCEL FILE
                WriteRGB2EnvisionExcelFile()
            End If
            GoTo CleanUp
        Catch ex As Exception
            GoTo CleanUp
        End Try

CleanUp:
        'Me.WindowState = FormWindowState.Normal
        pInitColor = Nothing
        pNewColor = Nothing
        pSelector = Nothing
        bColorSet = Nothing
        bmpTemp = Nothing
        rowDev = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()

    End Sub

    Private Sub itmColorBrowser_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles itmColorBrowser.Click
        'Provide the ESRI Color Browser form for the user to select a Development Type Color
        Dim pInitColor As IRgbColor
        Dim pNewColor As IRgbColor
        Dim pColorBrowser As IColorBrowser
        Dim bColorSet As Boolean
        Dim hit As DataGridView.HitTestInfo
        Dim bmpTemp As Bitmap
        Dim rowDev As IRow

        Try
            'Me.WindowState = FormWindowState.Minimized
            pInitColor = New RgbColor
            hit = Me.dgvDevelopmentTypes.HitTest(m_pntMouse.X, m_pntMouse.Y)
            pInitColor.Red = CInt(Me.dgvDevelopmentTypes.Rows(hit.RowIndex).Cells(2).Value)
            pInitColor.Green = CInt(Me.dgvDevelopmentTypes.Rows(hit.RowIndex).Cells(3).Value)
            pInitColor.Blue = CInt(Me.dgvDevelopmentTypes.Rows(hit.RowIndex).Cells(4).Value)

            pColorBrowser = New ColorBrowser
            pColorBrowser.Color = pInitColor
            bColorSet = pColorBrowser.DoModal(0)
            If bColorSet Then
                pNewColor = New RgbColor
                pNewColor = CType(pColorBrowser.Color, IRgbColor)
                bmpTemp = CreateColorBMP(pNewColor.Red, pNewColor.Green, pNewColor.Blue)
                If hit.Type = DataGridViewHitTestType.Cell Then
                    Me.dgvDevelopmentTypes.Rows(hit.RowIndex).Cells(0).Value = bmpTemp
                    Me.dgvDevelopmentTypes.Rows(hit.RowIndex).Cells(2).Value = CInt(pNewColor.Red)
                    Me.dgvDevelopmentTypes.Rows(hit.RowIndex).Cells(3).Value = CInt(pNewColor.Green)
                    Me.dgvDevelopmentTypes.Rows(hit.RowIndex).Cells(4).Value = CInt(pNewColor.Blue)
                    Me.dgvDevelopmentTypes.Rows(hit.RowIndex).Cells(0).Selected = True
                    'UPDATE THE LOOKUP TABLE
                    If Not m_tblDevelopmentTypes Is Nothing Then
                        rowDev = m_tblDevelopmentTypes.GetRow(hit.RowIndex)
                        rowDev.Value(m_tblDevelopmentTypes.FindField("RED")) = CInt(pNewColor.Red)
                        rowDev.Value(m_tblDevelopmentTypes.FindField("GREEN")) = CInt(pNewColor.Green)
                        rowDev.Value(m_tblDevelopmentTypes.FindField("BLUE")) = CInt(pNewColor.Blue)
                        rowDev.Store()
                    End If
                    If Me.itmAutoUpdateLegend.Checked Then
                        UpdateEnvisionLyrLegend()
                    End If
                    'UPDATE THE ENVISION EXCEL FILE
                    WriteRGB2EnvisionExcelFile()
                End If
                GoTo CleanUp
            End If
        Catch ex As Exception
            GoTo CleanUp
        End Try

CleanUp:
        'Me.WindowState = FormWindowState.Normal
        pInitColor = Nothing
        pNewColor = Nothing
        bColorSet = Nothing
        bmpTemp = Nothing
        rowDev = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub

    Private Sub dgvDevelopmentTypes_DragOver(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles dgvDevelopmentTypes.DragOver
        'Sets the graphic cursor to be displayed
        e.Effect = DragDropEffects.All
    End Sub

    Private Sub dgvDevelopmentTypes_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dgvDevelopmentTypes.MouseDown
        Dim pntMouse As System.Drawing.Point
        Dim pntDGV As System.Drawing.Point = New System.Drawing.Point
        If e.Button = Windows.Forms.MouseButtons.Right Then   'right mouse button was clicked
        End If

        'Make sure only the corrent menu items are enabled
        ' If the user right-clicks a cell, store it for use by the 
        ' shortcut menu.
        Try
            blnShowColor = False
            blnShowDevTypes = False
            Dim hit As DataGridView.HitTestInfo = dgvDevelopmentTypes.HitTest(e.X, e.Y)
            If e.Button = MouseButtons.Right Then
                If hit.Type = DataGridViewHitTestType.Cell Then
                    pntMouse.X = Me.MousePosition.X
                    pntMouse.Y = Me.MousePosition.Y
                    pntDGV.X = e.X
                    pntDGV.Y = e.Y
                    m_pntMouse = pntDGV
                    If hit.ColumnIndex = 0 Then
                        blnShowColor = True
                        Me.DevTypeContextMenu.Show(pntMouse)
                        Me.dgvDevelopmentTypes.Rows(hit.RowIndex).Selected = True
                    End If
                End If
            Else
                If hit.Type = DataGridViewHitTestType.Cell Then
                    pntMouse.X = Me.MousePosition.X
                    pntMouse.Y = Me.MousePosition.Y
                    pntDGV.X = e.X
                    pntDGV.Y = e.Y
                    m_pntMouse = pntDGV
                    If hit.ColumnIndex = 6 Then
                        blnShowDevTypes = True
                        m_intRow = hit.RowIndex
                        Me.DevTypeContextMenu.Show(pntMouse)
                    End If
                End If
            End If
            GoTo CleanUp
        Catch ex As Exception
            GoTo CleanUp
        End Try

CleanUp:
        pntMouse = Nothing
        pntDGV = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub

    Private Sub dgvDevelopmentTypes_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvDevelopmentTypes.SelectionChanged
        If m_blnLoadingDevTypes Then
            Exit Sub
        End If
        'Update the selected Development Type 
        Dim selectedRowCount As Integer = Me.dgvDevelopmentTypes.Rows.GetRowCount(DataGridViewElementStates.Selected)
        Try
            If selectedRowCount > 0 Then
                Dim sb As New System.Text.StringBuilder()
                Dim i As Integer
                For i = 0 To selectedRowCount - 1
                    m_strDWriteValue = CStr(Me.dgvDevelopmentTypes.Rows(Me.dgvDevelopmentTypes.SelectedRows(i).Index).Cells(1).Value)
                    If m_strDWriteValue = "ERASE" Then
                        m_strDWriteValue = ""
                    End If
                    Exit For
                Next i
            Else
                m_strDWriteValue = ""
            End If
            GoTo CleanUp
        Catch ex As Exception
            GoTo CleanUp
        End Try
CleanUp:
        selectedRowCount = Nothing
        GC.Collect()
        GC.WaitForPendingFinalizers()
    End Sub

    Private Sub btnApply_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnApply.Click
        WriteValues(Nothing, Nothing, Nothing)
    End Sub

    Private Sub DevTypeContextMenu_Opening(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles DevTypeContextMenu.Opening
        'Show Only those items needed 
        Me.itmChangeColor.Visible = blnShowColor
        Me.itmDevTypeChar.Visible = blnShowDevTypes
    End Sub

    Private Sub itmAutoUpdateLegend_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If itmAutoUpdateLegend.Checked Then
            UpdateEnvisionLyrLegend()
        End If
    End Sub

    Private Sub btnEditing_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEditing.Click
        EditSession()
    End Sub

    Private Sub btnSaveEdits_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSaveEdits.Click
        'CONTROLS THE EDIT SESSION ON THE ENVISION EDIT LAYER
        Dim pEditor As IEditor
        Dim pID As New UID
        Dim pDataset As IDataset

        Try
            'RETRIEVE THE EDITOR EXTENSION
            pID.Value = "esriEditor.Editor"
            pEditor = m_appEnvision.FindExtensionByCLSID(pID)
            If pEditor.EditState = esriEditState.esriStateEditing Then
                'SAVE THE EDITS BY STOPPING THE EDIT SESSION
                pEditor.StopEditing(True)
                'RETRIEVE DATASET FROM SELECTED ENVISION EDIT FEATURE CLASS
                pDataset = CType(m_pEditFeatureLyr.FeatureClass, IDataset)
                'RESTART EDITING THE DATASET WORKSPACE
                pEditor.StartEditing(pDataset.Workspace)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Save Edit Error")
            GoTo CleanUp
        End Try
CleanUp:
        pEditor = Nothing
        pID = Nothing
        pDataset = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub

    Private Sub UncheckAllScenarios()
        'UNCHECK ALL SCENARIOS MENU ITEMS
        Me.itmScenario1.Checked = False
        Me.itmScenario2.Checked = False
        Me.itmScenario3.Checked = False
        Me.itmScenario4.Checked = False
        Me.itmScenario5.Checked = False
    End Sub

    Private Sub itmScenario_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles itmScenario1.Click, itmScenario2.Click, itmScenario3.Click, itmScenario4.Click, itmScenario5.Click
        'CHANGE SCENARIO EDIT LAYER
        Dim intPrevious As Integer = m_intEditScenario
        Dim pMxDocument As IMxDocument
        Dim mapCurrent As Map
        Dim pFeatLayer As IFeatureLayer
        Dim intLayer As Integer
        Dim pDataset As IDataset = Nothing
        Dim pFeatClass As IFeatureClass = Nothing
        Dim strName As String = ""
        Dim strCurrentName As String = ""

        m_strProcessing = "DEFINE ENVISION EDIT SCENARIO"
        m_strProcessing = m_strProcessing & vbNewLine & "PROCESSING START TIME: " & Date.Now.ToString("M/d/yyyy hh:mm:ss tt") & vbNewLine & "---------------------------------------------------------------------------------------------" & vbNewLine

        Dim itmSelected As ToolStripMenuItem = sender
        UncheckAllScenarios()
        If itmSelected.Name = "itmScenario1" Then
            Me.itmScenario1.Checked = True
            If m_intEditScenario = 1 Then
                'GoTo CleanUp
            Else
                m_intEditScenario = 1
            End If
        ElseIf itmSelected.Name = "itmScenario2" Then
            Me.itmScenario2.Checked = True
            If m_intEditScenario = 2 Then
                'GoTo CleanUp
            Else
                m_intEditScenario = 2
            End If
        ElseIf itmSelected.Name = "itmScenario3" Then
            Me.itmScenario3.Checked = True
            If m_intEditScenario = 3 Then
                'GoTo CleanUp
            Else
                m_intEditScenario = 3
            End If
        ElseIf itmSelected.Name = "itmScenario4" Then
            Me.itmScenario4.Checked = True
            If m_intEditScenario = 4 Then
                'GoTo CleanUp
            Else
                m_intEditScenario = 4
            End If
        ElseIf itmSelected.Name = "itmScenario5" Then
            Me.itmScenario5.Checked = True
            If m_intEditScenario = 5 Then
                'GoTo CleanUp
            Else
                m_intEditScenario = 5
            End If
        End If

        If ChangeEnvisionLayer() Then
            Dim blnSynchronize = True
            'APPLY THE SUBAREA IF ONE IS DEFINED
            Try
                If m_dockEnvisionWinForm.itmcmbSubareaFields.Tag.Length > 0 And m_dockEnvisionWinForm.itmSelectSubarea.Tag.Length > 0 Then
                    If m_dockEnvisionWinForm.itmcmbSubareaFields.FindString(m_dockEnvisionWinForm.itmcmbSubareaFields.Tag) Then
                        m_dockEnvisionWinForm.itmcmbSubareaFields.Text = m_dockEnvisionWinForm.itmcmbSubareaFields.Tag
                        If m_dockEnvisionWinForm.itmSelectSubarea.FindString(m_dockEnvisionWinForm.itmSelectSubarea.Tag) >= 0 Then
                            blnSynchronize = False
                            m_dockEnvisionWinForm.itmSelectSubarea.Text = m_dockEnvisionWinForm.itmSelectSubarea.Tag
                        End If
                    End If
                End If
            Catch ex As Exception
            End Try
             'EXECUTE A SYNCHRONIZE
            If blnSynchronize Then
                SynchronizeData("QUICK")
            End If
        Else
            'SET ENVISION BACK TO THE PREVIOUS SCENARIO
            UncheckAllScenarios()
            pMxDocument = DirectCast(m_appEnvision.Document, IMxDocument)
            mapCurrent = CType(pMxDocument.FocusMap, Map)
            'RETRIEVE THE CURRENT EDIT LAYER NAME
            If Not m_pEditFeatureLyr Is Nothing Then
                pDataset = CType(m_pEditFeatureLyr.FeatureClass, IDataset)
                strCurrentName = pDataset.Name
                pDataset = Nothing
                GC.WaitForPendingFinalizers()
                GC.Collect()
            End If
            If mapCurrent.LayerCount > 0 Then
                For intLayer = 0 To mapCurrent.LayerCount - 1
                    If TypeOf mapCurrent.Layer(intLayer) Is IFeatureLayer Then
                        pFeatLayer = CType(mapCurrent.Layer(intLayer), IFeatureLayer)
                        pFeatClass = pFeatLayer.FeatureClass
                        pDataset = CType(pFeatClass, IDataset)
                        strName = pDataset.BrowseName
                        If pFeatLayer.Name.Contains("<ACTIVE>") Then
                            If Not strCurrentName = pDataset.Name Then
                                pFeatLayer.Name = pDataset.Name
                                pFeatLayer.Visible = False
                            End If
                        End If
                        GC.WaitForPendingFinalizers()
                        GC.Collect()
                    End If
                Next
                pMxDocument.ActiveView.ContentsChanged()
                pMxDocument.UpdateContents()
            End If

        End If
CleanUp:
        m_strProcessing = m_strProcessing & "PROCESSING END TIME: " & Date.Now.ToString("M/d/yyyy hh:mm:ss tt") & vbNewLine & "---------------------------------------------------------------------------------------------" & vbNewLine
        WriteToProcessingFile("DefiningEnvisionScenioeditLayer.txt")
        itmSelected = Nothing
        intPrevious = Nothing
        pMxDocument = Nothing
        mapCurrent = Nothing
        pFeatLayer = Nothing
        intLayer = Nothing
        pDataset = Nothing
        pFeatClass = Nothing
        strName = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()

    End Sub

    Private Sub itmCloseEnvisionExcel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles itmCloseEnvisionExcel.Click
        'CLOSE ENVISION EXCEL IF OPEN
        CloseEnvisionExcel()
        'Me.itmWriteExcelCalcs.Text = "OFF"
    End Sub

    'Private Sub itmAutoUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles itmWriteExcelCalcs.Click
    '    If Me.itmWriteExcelCalcs.Text = "ON" Then
    '        Me.itmWriteExcelCalcs.Text = "OFF"
    '        m_blnAutoUpdate = False
    '        CloseEnvisionExcel()
    '    Else
    '        Me.itmWriteExcelCalcs.Text = "ON"
    '        m_blnAutoUpdate = True
    '        OpenEnvisionExcel()
    '    End If
    'End Sub

    Private Sub itmSaveToExcelFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles itmSaveToExcelFile.Click

        Write2EnvisionExcelFile()

    End Sub

    Private Sub itmSelectEnvisionFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles itmSelectEnvisionFile.Click
        SelectEnvisionExcelFile()
    End Sub

    Private Sub SelectEnvisionExcelFile()
        'WRITE VALUES TO ENVISION EXCEL FILE
        'RETRIEVE OR OPEN THE SELECTED ENVISION EXCEL FILE
        Dim FileDialog1 As New OpenFileDialog
        Dim blnEditSession As Boolean = False
        Try
            'END ANY EDIT SESSIONS
            If Not Me.btnEditing.Text = "Start Edit" Then
                m_strProcessing = m_strProcessing & "An edit session is currently started and needs to be stopped before proceeding: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine & vbNewLine
                blnEditSession = True
                EditSession()
            End If

            FileDialog1.Title = "Select an ENVISION Excel File"
            FileDialog1.Filter = "Excel Files (*.xls,*.xlsx,*.xlsm)|*.xls;*.xlsx;*.xlsm"
            FileDialog1.CheckPathExists = True
            If FileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                'CLOSE ANY EXCEL, WHICH MAY BE OPEN
                If m_blnAutoUpdate Then
                    CloseEnvisionExcel()
                End If
                'APPLY TO MODULE VARIABLE
                m_strEnvisionExcelFile = FileDialog1.FileName
                'SET OTHER CONTROL VARIABLES
                Me.itmSelectEnvisionFile.ToolTipText = FileDialog1.FileName
                Me.itmLoadEnvisionExcelFile.Enabled = True
                Me.itmCloseEnvisionExcel.Enabled = True
                'Me.itmAutoUpdateStatus.Enabled = True
                'OPEN THE EXCEL IF AUTO UPDATE IS ON
                If m_blnAutoUpdate Then
                    OpenEnvisionExcel()
                Else
                    CloseEnvisionExcel()
                End If
                'REVIEW THE LOOKUPTABLES
                LookUpTablesEnvisionCheck(m_strFeaturePath)
            Else
                GoTo cleanup
            End If
            GoTo CleanUp
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Select Envision Excel File Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            CloseEnvisionExcel()
            GoTo CleanUp
        End Try

CleanUp:
        If blnEditSession Then
            EditSession()
        End If
        blnEditSession = Nothing
        FileDialog1 = Nothing
        GC.Collect()
        'GC.WaitForPendingFinalizers()
    End Sub

    Private Sub itmLoadEnvisionExcelFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles itmLoadEnvisionExcelFile.Click
        LoadExcelDevTypes()
    End Sub

    Private Sub RetrieveDevTypeExcelData(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'RETRIEVE THE DEVELOPMENT TYPE NAME, GRID CODE, 
        'RGB COLOR VALUES, AND ACRE SUMS FROM SELECTED ENVISION EXCEL FILE
        Dim shtScenario As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        Dim shtDevType As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        Dim intRow As Integer
        Dim strCellValue As String = ""
        Dim intRed As Integer = 0
        Dim intGreen As Integer = 0
        Dim intBlue As Integer = 0
        Dim bmpTemp As Bitmap

        If m_xlPaintWB1 Is Nothing Or m_xlPaintApp Is Nothing Then
            GoTo CleanUp
        Else
            Try
                If Not TypeOf m_xlPaintWB1.Sheets("SCENARIO") Is Microsoft.Office.Interop.Excel.Worksheet Then

                End If

                m_blnLoadingDevTypes = True
                'RETRIEVE SCENARIO TAB
                If Not TypeOf m_xlPaintWB1.Sheets("SCENARIO") Is Microsoft.Office.Interop.Excel.Worksheet Then
                    MessageBox.Show("The 'SCENARIO' tab could not be found in the selected ENVISION Excel file. Please select another file.", "SCENARIO TAB NOT FOUND", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
                    GoTo CloseExcel
                Else
                    shtScenario = m_xlPaintWB1.Sheets("SCENARIO")
                End If
                If Not TypeOf m_xlPaintWB1.Sheets("Dev Type Attributes") Is Microsoft.Office.Interop.Excel.Worksheet Then
                    MessageBox.Show("The 'Dev Type Attributes' tab could not be found in the selected ENVISION Excel file. Please select another file.", "SCENARIO TAB NOT FOUND", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
                    GoTo CloseExcel
                Else
                    shtDevType = m_xlPaintWB1.Sheets("Dev Type Attributes")
                End If

                'DEVELOPMENT TYPES SHOULD ONLY BE LOCATED IN ROW 5 TO 28
                m_dockEnvisionWinForm.dgvDevelopmentTypes.Rows.Clear()
                'INSERT THE ERASE OPTION
                m_dockEnvisionWinForm.dgvDevelopmentTypes.Rows.Add()
                bmpTemp = CreateColorBMP(255, 255, 255)
                m_dockEnvisionWinForm.dgvDevelopmentTypes.Rows(m_dockEnvisionWinForm.dgvDevelopmentTypes.RowCount - 1).Cells(0).Value = bmpTemp
                m_dockEnvisionWinForm.dgvDevelopmentTypes.Rows(m_dockEnvisionWinForm.dgvDevelopmentTypes.RowCount - 1).Cells(1).Value = "ERASE"
                m_dockEnvisionWinForm.dgvDevelopmentTypes.Rows(m_dockEnvisionWinForm.dgvDevelopmentTypes.RowCount - 1).Cells(2).Value = -1
                m_dockEnvisionWinForm.dgvDevelopmentTypes.Rows(m_dockEnvisionWinForm.dgvDevelopmentTypes.RowCount - 1).Cells(3).Value = 255
                m_dockEnvisionWinForm.dgvDevelopmentTypes.Rows(m_dockEnvisionWinForm.dgvDevelopmentTypes.RowCount - 1).Cells(4).Value = 255
                m_dockEnvisionWinForm.dgvDevelopmentTypes.Rows(m_dockEnvisionWinForm.dgvDevelopmentTypes.RowCount - 1).Cells(5).Value = 255
                For intRow = 7 To 17
                    m_dockEnvisionWinForm.dgvDevelopmentTypes.Rows(m_dockEnvisionWinForm.dgvDevelopmentTypes.RowCount - 1).Cells(intRow).Value = 0
                Next

                m_blnLoadingDevTypes = True
                m_strDWriteValue = ""
                m_strDWriteValue = -1
                'm_dockEnvisionWinForm.itmDevLabelOnly_Click(sender, e)
                m_dockEnvisionWinForm.prgBarEnvision.Value = 0
                For intRow = 37 To 60
                    m_dockEnvisionWinForm.prgBarEnvision.Value = ((intRow / 62) * 100)
                    strCellValue = CStr(CType(shtDevType.Cells(intRow, 1), Microsoft.Office.Interop.Excel.Range).Value)
                    'MAKE SURE THERE IS A DEVELOPMENT TYPE
                    If strCellValue.Length > 0 And (Not strCellValue = "" And Not strCellValue = "0") Then
                        m_dockEnvisionWinForm.dgvDevelopmentTypes.Rows.Add()
                        'COLUMN A = DEVELOPMENT TYPE NAME
                        m_dockEnvisionWinForm.dgvDevelopmentTypes.Rows(m_dockEnvisionWinForm.dgvDevelopmentTypes.RowCount - 1).Cells(1).Value = strCellValue
                        'GRID CODE NUMBER
                        If IsNumeric(CStr(CType(shtDevType.Cells(intRow, 74), Microsoft.Office.Interop.Excel.Range).Value)) Then
                            m_dockEnvisionWinForm.dgvDevelopmentTypes.Rows(m_dockEnvisionWinForm.dgvDevelopmentTypes.RowCount - 1).Cells(2).Value = CInt(CType(shtDevType.Cells(intRow, 74), Microsoft.Office.Interop.Excel.Range).Value)
                        Else
                            m_dockEnvisionWinForm.dgvDevelopmentTypes.Rows(m_dockEnvisionWinForm.dgvDevelopmentTypes.RowCount - 1).Cells(2).Value = 0
                        End If
                        'SYMBOL RED COLOR 
                        If IsNumeric(CStr(CType(shtDevType.Cells(intRow, 70), Microsoft.Office.Interop.Excel.Range).Value)) Then
                            intRed = CInt(CType(shtDevType.Cells(intRow, 70), Microsoft.Office.Interop.Excel.Range).Value)
                        Else
                            intRed = 0
                        End If
                        m_dockEnvisionWinForm.dgvDevelopmentTypes.Rows(m_dockEnvisionWinForm.dgvDevelopmentTypes.RowCount - 1).Cells(3).Value = intRed
                        'SYMBOL GREEN COLOR
                        If IsNumeric(CStr(CType(shtDevType.Cells(intRow, 71), Microsoft.Office.Interop.Excel.Range).Value)) Then
                            intGreen = CInt(CType(shtDevType.Cells(intRow, 71), Microsoft.Office.Interop.Excel.Range).Value)
                        Else
                            intGreen = 0
                        End If
                        m_dockEnvisionWinForm.dgvDevelopmentTypes.Rows(m_dockEnvisionWinForm.dgvDevelopmentTypes.RowCount - 1).Cells(4).Value = intGreen
                        'SYMBOL BLUE COLOR
                        If IsNumeric(CStr(CType(shtDevType.Cells(intRow, 72), Microsoft.Office.Interop.Excel.Range).Value)) Then
                            intBlue = CInt(CType(shtDevType.Cells(intRow, 72), Microsoft.Office.Interop.Excel.Range).Value)
                        Else
                            intBlue = 0
                        End If
                        m_dockEnvisionWinForm.dgvDevelopmentTypes.Rows(m_dockEnvisionWinForm.dgvDevelopmentTypes.RowCount - 1).Cells(5).Value = intBlue
                        'COLOR BITMAP SYMBOL
                        bmpTemp = CreateColorBMP(intRed, intGreen, intBlue)
                        m_dockEnvisionWinForm.dgvDevelopmentTypes.Rows(m_dockEnvisionWinForm.dgvDevelopmentTypes.RowCount - 1).Cells(0).Value = bmpTemp
                    Else
                        Exit For
                    End If
                Next

                m_dockEnvisionWinForm.prgBarEnvision.Value = 100
                'm_dockEnvisionWinForm.itmSymbolWithLabel_Click(sender, e)
                m_dockEnvisionWinForm.dgvDevelopmentTypes.Visible = False

                GoTo CleanUp

            Catch ex As Exception
                MessageBox.Show(ex.Message, "DEV TYPE LOAD ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
                GoTo CloseExcel
            End Try
            GoTo CleanUp
        End If

CloseExcel:
        CloseEnvisionExcel()
        GoTo CleanUp

CleanUp:
        m_blnLoadingDevTypes = False
        shtScenario = Nothing
        shtDevType = Nothing
        intRow = Nothing
        strCellValue = Nothing
        intRed = Nothing
        intGreen = Nothing
        intBlue = Nothing
        bmpTemp = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
        m_dockEnvisionWinForm.dgvDevelopmentTypes.Visible = True
        m_dockEnvisionWinForm.dgvDevelopmentTypes.Refresh()
        m_dockEnvisionWinForm.prgBarEnvision.Value = 0
    End Sub

    Private Sub itmIndicatorShowAll_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)

        'CHECK ALL OF THE SCENARIO OPTIONS
        m_blnScrMenuCheck = True
        m_blnScrMenuCheck = False
        UpdateSummaryTableSelection()

CLeanUp:
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub

    Private Sub itmIndicator1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        UpdateSummaryTableSelection()
    End Sub

    Private Sub itmcmbSubareaFields_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles itmcmbSubareaFields.SelectedIndexChanged, itmcmbSubareaFields.Enter
        lbLoadingSubaraeas = True
        LoadUniqueSubareaValues()
        lbLoadingSubaraeas = False
    End Sub

    Private Sub itmSelectSubarea_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles itmSelectSubarea.SelectedIndexChanged, itmSelectSubarea.Enter
        'SET THE ENVISION LAYER QUERY DEFINITION

        Dim mxApplication As IMxApplication
        Dim pMxDoc As IMxDocument
        Dim pFLDef As IFeatureLayerDefinition
        Dim pField As Field
        Dim pFCur As IFeatureCursor
        Dim pEnv As IEnvelope = Nothing
        Dim pFeat As IFeature

        If lbLoadingSubaraeas Then
            GoTo CleanUp
        End If

        Try
            'RETRIEVE THE CURRENT VIEW DOCUMENT
            mxApplication = CType(m_appEnvision, IMxApplication)
            pMxDoc = CType(m_appEnvision.Document, IMxDocument)

            'SET QUERY DEFINTITON
            pFLDef = m_pEditFeatureLyr
            If m_pEditFeatureLyr.FeatureClass.FindField(m_dockEnvisionWinForm.itmcmbSubareaFields.Text) >= 0 Then
                pField = m_pEditFeatureLyr.FeatureClass.Fields.Field(m_pEditFeatureLyr.FeatureClass.FindField(m_dockEnvisionWinForm.itmcmbSubareaFields.Text))
                m_appEnvision.StatusBar.Message(0) = "Applying subarea, " & m_dockEnvisionWinForm.itmcmbSubareaFields.Text & ", to the Envision layer  (START)"
                If pField.Type = esriFieldType.esriFieldTypeString Then
                    pFLDef.DefinitionExpression = m_dockEnvisionWinForm.itmcmbSubareaFields.Text & " = '" & m_dockEnvisionWinForm.itmSelectSubarea.Text & "'"
                Else
                    pFLDef.DefinitionExpression = m_dockEnvisionWinForm.itmcmbSubareaFields.Text & " = " & m_dockEnvisionWinForm.itmSelectSubarea.Text
                End If
            End If
            m_appEnvision.StatusBar.Message(0) = "Applying subarea, " & m_dockEnvisionWinForm.itmcmbSubareaFields.Text & ", to the Envision layer (COMPLETE)"

            'ZOOM TO THE QUERY DEFINITION AREA
            pFCur = m_pEditFeatureLyr.Search(Nothing, False)
            pFeat = pFCur.NextFeature
            Do While Not pFeat Is Nothing
                If pEnv Is Nothing Then
                    pEnv = New Envelope
                    pEnv = pFeat.Shape.Envelope
                Else
                    pEnv.Union(pFeat.Shape.Envelope)
                End If
                pFeat = pFCur.NextFeature
            Loop
            If Not pEnv Is Nothing Then
                m_appEnvision.StatusBar.Message(0) = "Applying subarea, " & m_dockEnvisionWinForm.itmcmbSubareaFields.Text & ", to the Envision layer (ZOOMING TO SUBAREA)"
                pMxDoc.ActivatedView.Extent = pEnv
                pMxDoc.ActivatedView.Refresh()
                m_appEnvision.StatusBar.Message(0) = "Applying subarea, " & m_dockEnvisionWinForm.itmcmbSubareaFields.Text & ", to the Envision layer (ZOOMING TO SUBAREA) - REFRESH COMPLETE"
            End If

            'RUN THE SYNCHRONIZE WITH ALL THE CALCUALTIONS OFF
            m_blnSynchronizePrompt = True
            m_dockEnvisionWinForm.itmSelectSubarea.Tag = m_dockEnvisionWinForm.itmSelectSubarea.Text
            m_appEnvision.StatusBar.Message(0) = "Applying subarea, " & m_dockEnvisionWinForm.itmcmbSubareaFields.Text & ", to the Envision layer (PARTIAL SYNCHRONIZE STARTED)"
            SynchronizeData("QUICK")
            m_appEnvision.StatusBar.Message(0) = "Applying subarea, " & m_dockEnvisionWinForm.itmcmbSubareaFields.Text & ", to the Envision layer (PARTIAL SYNCHRONIZE FINISHED)"
            m_blnSynchronizePrompt = False
            m_appEnvision.StatusBar.Message(0) = "Subarea Processing complete for subarea: " & m_dockEnvisionWinForm.itmcmbSubareaFields.Text & ", "
            GoTo CleanUp
        Catch ex As Exception
            MessageBox.Show("Error in applying a subarea: " & vbNewLine & ex.Message, "Subarea Process Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo CleanUp
        End Try

CleanUp:
        mxApplication = Nothing
        pMxDoc = Nothing
        pFLDef = Nothing
        pField = Nothing
        pFCur = Nothing
        pEnv = Nothing
        pFeat = Nothing
        GC.Collect()
        GC.WaitForPendingFinalizers()
    End Sub

    Private Sub itmClearSubareaDefinition_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles itmClearSubareaDefinition.Click
        'CLEAR THE ENVISION LAYER QUERY DEFINITION
        Dim mxApplication As IMxApplication
        Dim pMxDoc As IMxDocument
        Dim pFLDef As IFeatureLayerDefinition
        Dim pField As Field
        pFLDef = m_pEditFeatureLyr
        pFLDef.DefinitionExpression = ""
        m_dockEnvisionWinForm.itmcmbSubareaFields.Text = ""
        m_dockEnvisionWinForm.itmSelectSubarea.Text = ""

        'RUN THE SYNCHRONIZE WITH ALL THE CALCUALTIONS OFF
        m_blnSynchronizePrompt = True
        SynchronizeData("QUICK")
        m_blnSynchronizePrompt = False

        'RETRIEVE THE CURRENT VIEW DOCUMENT
        mxApplication = CType(m_appEnvision, IMxApplication)
        pMxDoc = CType(m_appEnvision.Document, IMxDocument)
        pMxDoc.ActivatedView.Extent = m_pEditFeatureLyr.AreaOfInterest
        'pMxDoc.ActiveView.PartialRefresh(esriViewDrawPhase.esriViewGeoSelection, m_pEditFeatureLyr, pMxDoc.ActiveView.Extent)
        pMxDoc.ActiveView.Refresh()

        mxApplication = Nothing
        pMxDoc = Nothing
        pField = Nothing
        pFLDef = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub

    Private Sub ddbBrushes_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ddbBrushes.Click
        Dim pUID As New UID
        Try
            pUID.Value = "Envision_Tools.EnvisionAttributeEditor"
            m_appEnvision.CurrentTool = m_appEnvision.Document.CommandBars.Find(pUID)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Select Button Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
        End Try
    End Sub

    Private Sub ctlEnvisionAttributesEditor_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        'SET THE CALCULATION VARAIBLE - ON
        If m_blnRunCalcs Then
            Me.mnuAltCalcs.Image = Me.itmCalcsOn.Image
            Me.mnuAltCalcs.Text = Me.itmCalcsOn.ToolTipText
            Me.mnuAltCalcs.ToolTipText = Me.itmCalcsOn.ToolTipText
        Else
            Me.mnuAltCalcs.Image = Me.itmCalcsOff.Image
            Me.mnuAltCalcs.Text = Me.itmCalcsOff.ToolTipText
            Me.mnuAltCalcs.ToolTipText = Me.itmCalcsOff.ToolTipText
        End If

        Dim dblValue As Double = 0
        Dim dblValue1 As Double = 0
        Dim dblValue2 As Double = 0
        Dim dblValue3 As Double = 0
        Dim dblValue4 As Double = 0
        Dim strValue As String = ""
        Dim intRow As Integer = 0
        If m_tblScSummary Is Nothing Then
            GoTo cleanup
        End If
        
CleanUp:
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub

    Private Sub btnEndEditNoSave_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEndEditNoSave.Click
        If MessageBox.Show("Would you like to continue to end the current edit session WITHOUT saving edits?", "END EDIT - NO SAVE", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2, MessageBoxOptions.DefaultDesktopOnly, False) = DialogResult.No Then
            Exit Sub
        End If
        'CONTROLS THE EDIT SESSION ON THE ENVISION EDIT LAYER
        Dim pEditor As IEditor
        Dim pID As New UID
        Dim pDataset As IDataset

        Try

            'RETRIEVE THE EDITOR EXTENSION
            pID.Value = "esriEditor.Editor"
            pEditor = m_appEnvision.FindExtensionByCLSID(pID)

            If Me.btnEditing.Text = "Start Edit" Then
                'CHECK FOR EDIT LAYER
                If m_pEditFeatureLyr Is Nothing Then
                    GoTo CleanUp
                Else
                    Try
                        'RETRIEVE DATASET FROM SELECTED ENVISION EDIT FEATURE CLASS
                        pDataset = CType(m_pEditFeatureLyr.FeatureClass, IDataset)
                        'START EDITING THE DATASET WORKSPACE
                        pEditor.StartEditing(pDataset.Workspace)
                        Me.btnEditing.Text = "End Edit" & vbNewLine & " - Save"
                        Me.btnSaveEdits.Visible = False
                        Me.tlsEndEdit.Visible = False
                        Me.sep2.Visible = False
                    Catch ex As Exception
                        MessageBox.Show(ex.Message, "Start Edit Error")
                        GoTo CleanUp
                    End Try
                    Me.btnSaveEdits.Visible = True
                    Me.tlsEndEdit.Visible = True
                    Me.sep2.Visible = True
                    GoTo CleanUp
                End If
            Else
                Try
                    pEditor.StopEditing(False)
                    Me.btnEditing.Text = "Start Edit"
                    Me.btnSaveEdits.Visible = False
                    Me.tlsEndEdit.Visible = False
                    Me.sep2.Visible = False
                    SynchronizeData("FULL")
                Catch ex As Exception
                    MessageBox.Show(ex.Message, "End Edit Error")
                    GoTo CleanUp
                End Try
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Envision Edit Error")
            GoTo CleanUp
        End Try
CleanUp:
        pEditor = Nothing
        pID = Nothing
        pDataset = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub

    Private Sub LoadPolyLayers()
        'CLCYLE THROUGH THE SELECTED ENVISION FILE GDB AND POPULATE THE SCENARIO COMBOX BOXES WITH THE NAME
        'OF ALL AVAILABLE POLYGON LAYERS
        Dim pWSF As IWorkspaceFactory
        Dim pFeatureWS As IFeatureWorkspace
        Dim pWS As IWorkspace
        Dim pEnumDSN As IEnumDatasetName
        Dim pDSName As IDatasetName
        Dim pFeatClass As IFeatureClass
        Dim arrList As ArrayList = New ArrayList
        Dim strName As String
        Dim rowTemp As IRow
        Dim strLayerName As String
        Dim intCount As Integer

        If m_strFeaturePath.Length <= 0 Then
            GoTo CleanUp
        Else
            m_dockEnvisionWinForm.cmbDefineLayer1.Items.Clear()
            m_dockEnvisionWinForm.cmbDefineLayer2.Items.Clear()
            m_dockEnvisionWinForm.cmbDefineLayer3.Items.Clear()
            m_dockEnvisionWinForm.cmbDefineLayer4.Items.Clear()
            m_dockEnvisionWinForm.cmbDefineLayer5.Items.Clear()
        End If

        Try
            pWSF = New FileGDBWorkspaceFactory
            pWS = pWSF.OpenFromFile(m_strFeaturePath, 0)
            pEnumDSN = pWS.DatasetNames(esriDatasetType.esriDTFeatureClass)
            pDSName = pEnumDSN.Next
            Do Until pDSName Is Nothing
                If TypeOf pDSName Is IFeatureClassName Then
                    arrList.Add(pDSName.Name)
                End If
                pDSName = pEnumDSN.Next
            Loop
            pWS = Nothing
            pEnumDSN = Nothing
            pDSName = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()
        Catch ex As Exception
            GoTo CleanUp
        End Try

        If arrList.Count <= 0 Then
            GoTo CleanUp
        Else
            Try
                pWSF = New FileGDBWorkspaceFactory
                pFeatureWS = pWSF.OpenFromFile(m_strFeaturePath, 0)
                'Insert Clear Option
                m_dockEnvisionWinForm.cmbDefineLayer1.Items.Add("")
                m_dockEnvisionWinForm.cmbDefineLayer2.Items.Add("")
                m_dockEnvisionWinForm.cmbDefineLayer3.Items.Add("")
                m_dockEnvisionWinForm.cmbDefineLayer4.Items.Add("")
                m_dockEnvisionWinForm.cmbDefineLayer5.Items.Add("")
                'Load in available polygon layer names as available scenario layers
                For Each strName In arrList
                    pFeatClass = pFeatureWS.OpenFeatureClass(strName)
                    If pFeatClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                        m_dockEnvisionWinForm.cmbDefineLayer1.Items.Add(strName)
                        m_dockEnvisionWinForm.cmbDefineLayer2.Items.Add(strName)
                        m_dockEnvisionWinForm.cmbDefineLayer3.Items.Add(strName)
                        m_dockEnvisionWinForm.cmbDefineLayer4.Items.Add(strName)
                        m_dockEnvisionWinForm.cmbDefineLayer5.Items.Add(strName)
                    End If
                Next
            Catch ex As Exception
            End Try
        End If
        'LOAD THE CURRENTLY SELECTED EDIT LAYER NAMES INTO THE TEXT PROPERTY FOR THE COMBO BOXES
        For intCount = 1 To 5
            rowTemp = m_tblScSummary.GetRow(intCount)
            strLayerName = rowTemp.Value(m_tblScSummary.FindField("LAYER_NAME"))
            If intCount = 1 Then
                cmbDefineLayer1.Text = strLayerName
            ElseIf intCount = 2 Then
                cmbDefineLayer2.Text = strLayerName
            ElseIf intCount = 3 Then
                cmbDefineLayer3.Text = strLayerName
            ElseIf intCount = 4 Then
                cmbDefineLayer4.Text = strLayerName
            ElseIf intCount = 5 Then
                cmbDefineLayer5.Text = strLayerName
            End If
        Next
CleanUp:
        pWSF = Nothing
        pFeatureWS = Nothing
        pWS = Nothing
        pEnumDSN = Nothing
        pDSName = Nothing
        pFeatClass = Nothing
        arrList = Nothing
        rowTemp = Nothing
        strLayerName = Nothing
        intCount = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub

    Private Sub cmbDefineLayer1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbDefineLayer1.SelectedIndexChanged, cmbDefineLayer2.SelectedIndexChanged, cmbDefineLayer3.SelectedIndexChanged, cmbDefineLayer4.SelectedIndexChanged, cmbDefineLayer5.SelectedIndexChanged
        'REVIEW THE SCENARIO TABLE TO CHECK THAT THE SELECTED LAYER IS NOT ALLOCATED TO ANOTHER SCENARIO OTHER THAN SCENARIO 1
        Dim rowScTemp As IRow = Nothing
        Dim strLayerName As String
        Dim arrList As ArrayList = New ArrayList
        Dim blnFound As Boolean = False
        Dim intCount As Integer
        Dim cmbSelected As ToolStripComboBox = Nothing
        Dim strSelectedLayerName As String
        Dim intScenario As Integer

        Dim intFldCount As Integer
        Dim intRow As Integer
        Dim intRowCount As Integer
        Dim intFld As Integer
        Dim rowTemp As IRow
        Dim pField As Field
        Dim strVal As String
        Dim intVal As Integer
        Dim blnStore As Boolean
        Dim pTable As ITable = Nothing

        Dim pMxDocument As IMxDocument
        Dim mapCurrent As Map
        Dim pFeatLayer As IFeatureLayer
        Dim intLayer As Integer
        Dim pDataset As IDataset = Nothing
        Dim pFeatClass As IFeatureClass = Nothing
        Dim strName As String = ""
        Dim strCurrentName As String = ""

        If blnLoading = True Then
            GoTo CleanUp
        End If

        Try
            'RETRIEVE THE SELECTED CONTROL 
            cmbSelected = CType(sender, ToolStripComboBox)
            strSelectedLayerName = cmbSelected.Text
        Catch ex As Exception
            GoTo CleanUp
        End Try

        For intCount = 1 To 5
            rowScTemp = m_tblScSummary.GetRow(intCount)
            strLayerName = rowScTemp.Value(m_tblScSummary.FindField("LAYER_NAME"))
            'IF SELECTION MATCHES THE SCENARIO'S CURRENT LAYER, THEN EXIT
            'RETRIEVE THE ROW SCENARIO NAME AND COMPARED TO THE SELECTED SCENARIO LAYER NAME 
            If intCount = 1 And cmbSelected.Name.Contains("1") Then
                If strLayerName = strSelectedLayerName Then
                    blnFound = True
                End If
                Exit For
            ElseIf intCount = 2 And cmbSelected.Name.Contains("2") Then
                If strLayerName = strSelectedLayerName Then
                    blnFound = True
                End If
                Exit For
            ElseIf intCount = 3 And cmbSelected.Name.Contains("3") Then
                If strLayerName = strSelectedLayerName Then
                    blnFound = True
                End If
                Exit For
            ElseIf intCount = 4 And cmbSelected.Name.Contains("4") Then
                If strLayerName = strSelectedLayerName Then
                    blnFound = True
                End If
                Exit For
            ElseIf intCount = 5 And cmbSelected.Name.Contains("5") Then
                If strLayerName = strSelectedLayerName Then
                    blnFound = True
                End If
                Exit For
            End If
        Next

        'IF LAYER IS AVAILABLE, THEN APPLY TO THE SELECTED SCENARIO
        If cmbSelected.Name.Contains("1") Then
            intScenario = 1
            rowScTemp = m_tblScSummary.GetRow(1)
            pTable = m_tblSc1
        ElseIf cmbSelected.Name.Contains("2") Then
            intScenario = 2
            rowScTemp = m_tblScSummary.GetRow(2)
            pTable = m_tblSc2
        ElseIf cmbSelected.Name.Contains("3") Then
            intScenario = 3
            rowScTemp = m_tblScSummary.GetRow(3)
            pTable = m_tblSc3
        ElseIf cmbSelected.Name.Contains("4") Then
            intScenario = 4
            rowScTemp = m_tblScSummary.GetRow(4)
            pTable = m_tblSc4
        ElseIf cmbSelected.Name.Contains("5") Then
            intScenario = 5
            rowScTemp = m_tblScSummary.GetRow(5)
            pTable = m_tblSc5
        End If
        If Not blnFound Then
            rowScTemp.Value(m_tblScSummary.FindField("LAYER_NAME")) = strSelectedLayerName
            rowScTemp.Store()
            'SET ALL NULL VALUES TO DEFAULTS of "" OR 0
            m_appEnvision.StatusBar.Message(0) = "CHECKING FOR REQUIRED SCENARIO TABLE LOOKUP TABLE FIELDS: CHECKING FOR NULL VALUES"
            intRowCount = pTable.RowCount(Nothing)
            intFldCount = pTable.Fields.FieldCount
            If intRowCount > 0 Then
                For intRow = 1 To intRowCount
                    blnStore = False
                    rowTemp = pTable.GetRow(intRow)
                    For intFld = 1 To intFldCount - 1
                        pField = pTable.Fields.Field(intFld)
                        If pField.Type = esriFieldType.esriFieldTypeString Then
                            Try
                                strVal = rowTemp.Value(intFld)
                            Catch ex As Exception
                                rowTemp.Value(intFld) = ""
                                blnStore = True
                            End Try
                        Else
                            Try
                                intVal = CInt(rowTemp.Value(intFld))
                            Catch ex As Exception
                                rowTemp.Value(intFld) = 0
                                blnStore = True
                            End Try
                        End If
                    Next
                    If blnStore Then
                        rowTemp.Store()
                    End If
                Next
            End If
            'CLEAR EXISTING EXIT LAYER IF USER VOIDS THE SCENARIO
            If intScenario = m_intEditScenario Then
                If strSelectedLayerName = "" Then
                    m_pEditFeatureLyr = Nothing
                End If
            End If
            pMxDocument = DirectCast(m_appEnvision.Document, IMxDocument)
            mapCurrent = CType(pMxDocument.FocusMap, Map)
            'RETRIEVE THE CURRENT EDIT LAYER NAME
            If Not m_pEditFeatureLyr Is Nothing Then
                pDataset = CType(m_pEditFeatureLyr.FeatureClass, IDataset)
                strCurrentName = pDataset.Name
                pDataset = Nothing
                GC.WaitForPendingFinalizers()
                GC.Collect()
                m_strProcessing = m_strProcessing & "The previous edit layer name is " & strCurrentName & ": " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
            End If
            If mapCurrent.LayerCount > 0 Then
                For intLayer = 0 To mapCurrent.LayerCount - 1
                    If TypeOf mapCurrent.Layer(intLayer) Is IFeatureLayer Then
                        pFeatLayer = CType(mapCurrent.Layer(intLayer), IFeatureLayer)
                        pFeatClass = pFeatLayer.FeatureClass
                        pDataset = CType(pFeatClass, IDataset)
                        strName = pDataset.BrowseName
                        If pFeatLayer.Name.Contains("<ACTIVE>") Then
                            If Not strCurrentName = pDataset.Name Then
                                pFeatLayer.Name = pDataset.Name
                                pFeatLayer.Visible = False
                            End If
                        End If
                        GC.WaitForPendingFinalizers()
                        GC.Collect()
                    End If
                Next
                pMxDocument.ActiveView.ContentsChanged()
                pMxDocument.UpdateContents()
            End If
        End If
        'EXECUTE THE FUNCITON TO LOAD THE NEWLY DEFINE SCENARIO EDIT LAYER
        If intScenario = m_intEditScenario Then
            ChangeEnvisionLayer()
        End If
        GoTo CleanUp

CleanUp:
        'blnLoading = False
        rowTemp = Nothing
        strLayerName = Nothing
        arrList = Nothing
        blnFound = Nothing
        intCount = Nothing
        pMxDocument = Nothing
        mapCurrent = Nothing
        pFeatLayer = Nothing
        intLayer = Nothing
        pDataset = Nothing
        pFeatClass = Nothing
        strName = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub

    Private Sub itmSynchronizeCalcNulls_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles itmSynchronizeCalcNulls.Click
        'QUERY ALL REQUIRED FIELDS FOR NULL VALUES AND CALCULATE TO DEFUALT VALUES
        Dim strFields As String = ""
        Dim strFieldName As String = ""
        Dim pFeatClass As IFeatureClass = Nothing
        Dim pTable As ITable = Nothing
        Dim intFld As Integer = 0
        Dim intFldCount As Integer = 0
        Try
            pFeatClass = m_pEditFeatureLyr.FeatureClass
            pTable = CType(pFeatClass, ITable)
            strFields = "VAC_ACRE,DEVD_ACRE,HU,EMP,SF,TH,MF,MH,RET,OFF,IND,NetHUDen,PCT_MF,NetRETDen,WORKER,NW1664,P0004,P0515,NW6500,RET,VO,VMT,DAT,APT,TRT,NMT,P_HH100K,MEDIAN,RegTrans,TransBuff,LANDMIX"
            intFldCount = strFields.Split(",").Length
            For Each strFieldName In strFields.Split(",")
                intFld = intFld + 1
                If pTable.FindField(strFieldName) Then
                    If QueryNulls(pFeatClass, strFieldName) Then
                    End If
                End If
                m_dockEnvisionWinForm.prgBarEnvision.Value = (intFld / intFldCount) * 100
            Next
        Catch ex As Exception

        End Try
CleanUp:
        m_dockEnvisionWinForm.prgBarEnvision.Value = 0
        strFields = Nothing
        strFieldName = Nothing
        pFeatClass = Nothing
        pTable = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub

    Public Function QueryNulls(ByVal pFeatClass As IFeatureClass, ByVal strFieldName As String) As Boolean
        'QUERY THE SELECTED FIELD FOR NULL VALUES
        QueryNulls = False
        Dim pQFilter As IQueryFilter
        Dim strQString As String = Nothing
        Dim pFeatureCursor As FeatureCursor
        Dim intTotalcount As Integer = 0
        Dim pFeatSelection As IFeatureSelection
        Dim pCursor As ICursor
        Dim pCalc As ICalculator
        Try
            If pFeatClass Is Nothing Then
                GoTo CleanUp
            End If
            strQString = strFieldName & " IS null"
            pQFilter = New QueryFilter
            pQFilter.WhereClause = strQString
            pFeatSelection = CType(m_pEditFeatureLyr, IFeatureSelection)
            pFeatSelection.SelectFeatures(pQFilter, esriSelectionResultEnum.esriSelectionResultNew, False)
            If pFeatSelection.SelectionSet.Count > 0 Then
                pCalc = New Calculator
                pCursor = pFeatClass.Update(pQFilter, False)
                With pCalc
                    .Cursor = pCursor
                    .Field = strFieldName
                    .Expression = "0"
                End With
                pCalc.Calculate()
            End If
            pFeatSelection.Clear()
            GoTo CleanUp
        Catch ex As Exception
        End Try
CleanUp:
        pQFilter = Nothing
        strQString = Nothing
        pFeatClass = Nothing
        pFeatureCursor = Nothing
        intTotalcount = Nothing
        pFeatSelection = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Function

    Private Sub itm1toManyExport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles itm1toManyExport.Click
        'OPEN 1 to MANY EXPORT FORM
        If m_frm1toManyExport Is Nothing Then
            m_frm1toManyExport = New frmOnetoManyExport
        End If
        m_frm1toManyExport.ShowDialog()
    End Sub

    Private Sub itmFieldManager_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles itmFieldManager.Click
        'OPEN ENVISION PARCEL ATTRIBUTE MANAGER FORM
        If m_frmAttribManager Is Nothing Then
            m_frmAttribManager = New frmFieldVariables
        End If
        m_frmAttribManager.Show()
    End Sub

    Private Sub itmCalcsOn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles itmCalcsOn.Click
        'SET THE CALCULATION VARAIBLE - ON
        'm_blnRunCalcs = True
        ' m_blnAutoUpdate = True
        Me.mnuAltCalcs.Image = Me.itmCalcsOn.Image
        Me.mnuAltCalcs.Text = Me.itmCalcsOn.ToolTipText
        Me.mnuAltCalcs.ToolTipText = Me.itmCalcsOn.ToolTipText
    End Sub

    Private Sub itmCalcsOff_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles itmCalcsOff.Click
        'SET THE CALCULATION VARAIBLE - OFF
        'm_blnAutoUpdate = False
        Me.mnuAltCalcs.Image = Me.itmCalcsOff.Image
        Me.mnuAltCalcs.Text = Me.itmCalcsOff.ToolTipText
        Me.mnuAltCalcs.ToolTipText = Me.itmCalcsOff.ToolTipText
    End Sub

    Private Sub itmBuffer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles itmBufferSummary.Click
        'OPEN THE FORM TO RUN ACCESS BUFFER SUMMARIES
        If m_frmAccessBufferSummary Is Nothing Then
            m_frmAccessBufferSummary = New frmBufferSum
        End If
        m_frmAccessBufferSummary.ShowDialog()
    End Sub

    Private Sub itmFeaturesummary_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles itmFeatureSummaries.Click
        'OPEN THE FORM TO RUN ACCESS BUFFER SUMMARIES
        If m_frmAccessFeatureSummary Is Nothing Then
            m_frmAccessFeatureSummary = New frmFeatureSum
        End If
        m_frmAccessFeatureSummary.ShowDialog()
    End Sub

    Private Sub itmAggregate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles itmAggregate.Click
        'OPEN THE FORM TO RUN ACCESS BUFFER SUMMARIES
        If m_frmFieldAggregation Is Nothing Then
            m_frmFieldAggregation = New frmAggregation
        End If
        m_frmFieldAggregation.ShowDialog()
    End Sub

    Private Sub itmRedevTiming_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles itmRedevTiming.Click
        'OPENS THE REDEVELOPMENT TIMING FORM
        If m_frmRedevTiming Is Nothing Then
            m_frmRedevTiming = New frmRedevTiming
        End If
        m_frmRedevTiming.Show()
    End Sub

    Private Sub itmJHBalance_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles itmJHBalance.Click
        'OPENS THE LOCAL JOBS-HOUSING BALANCE FORM
        If m_frmLocalJHBalance Is Nothing Then
            m_frmLocalJHBalance = New frmLocalJobsHousingBalance
        End If
        m_frmLocalJHBalance.Show()
    End Sub

    Private Sub itm7DModalCalc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles itm7DModalCalc.Click
        'OPEN THE FORM TO RUN 7D MODAL CALCULATIONS
        If m_frm7DModelCalc Is Nothing Then
            m_frm7DModelCalc = New frm7DModelCalc
        End If
        m_frm7DModelCalc.ShowDialog()
    End Sub

    Private Sub itmSumTransLocations_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles itmSumTransLocations.Click
        'OPEN THE FORM TO RUN 7D MODAL CALCULATIONS
        If m_frmSumTransLocations Is Nothing Then
            m_frmSumTransLocations = New frmSumTransportationPnts
        End If
        m_frmSumTransLocations.ShowDialog()
    End Sub

    Private Sub itmTravelSummaryBuffers_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles itmTravelSummaryBuffers.Click
        'OPEN THE FORM TO RUN JOBS BUFFER SUMMARIES - NETWORK SERVICE AREAS
        If m_frmTravelSummaryBuffers Is Nothing Then
            m_frmTravelSummaryBuffers = New frmTravelSummaryBuffers
        End If
        m_frmTravelSummaryBuffers.ShowDialog()
    End Sub


    Private Sub itmQuickSync_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles itmQuickSync.Click
        'EXECUTE A QUICK SYNCH
        SynchronizeData("QUICK")
    End Sub

    Private Sub itmProximtySummary_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles itmProximtySummary.Click
        'OPEN THE FORM TO PROXIMITY BUFFER SUMMARIES
        If m_frmProximtySummary Is Nothing Then
            m_frmProximtySummary = New frmProximity
        End If
        m_frmProximtySummary.ShowDialog()
    End Sub

    Private Sub itmSumLandUseMix_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles itmSumLandUseMix.Click
        'OPEN THE FORM FOR SUMMARIZING THE LAND USE MIX TO THE NEIGHBORHOOD LEVEL
        If m_frmLUMix Is Nothing Then
            m_frmLUMix = New frmLandUseMix
        End If
        m_frmLUMix.ShowDialog()
    End Sub

    Private Sub dgvDevelopmentTypes_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvDevelopmentTypes.CellContentClick

    End Sub

    Private Sub itmcmbSubareaFields_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles itmcmbSubareaFields.Click

    End Sub
End Class
