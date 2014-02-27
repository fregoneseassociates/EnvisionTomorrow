Imports System
Imports System.Windows.Forms
Imports System.IO
Imports System.Drawing

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

Public Class frmTravelSummaryBuffers
    Dim arrFeatLyr As ArrayList = New ArrayList
    Dim pGPServiceArea As Geoprocessor = New Geoprocessor
    Dim pSA As ESRI.ArcGIS.NetworkAnalystTools.MakeServiceAreaLayer = New ESRI.ArcGIS.NetworkAnalystTools.MakeServiceAreaLayer
    Dim pAddFac As ESRI.ArcGIS.NetworkAnalystTools.AddLocations = New ESRI.ArcGIS.NetworkAnalystTools.AddLocations
    Dim pSASolve As ESRI.ArcGIS.NetworkAnalystTools.Solve = New ESRI.ArcGIS.NetworkAnalystTools.Solve
    Dim pCopy As ESRI.ArcGIS.DataManagementTools.CopyFeatures = New ESRI.ArcGIS.DataManagementTools.CopyFeatures
    Dim pDelete As ESRI.ArcGIS.DataManagementTools.Delete = New ESRI.ArcGIS.DataManagementTools.Delete
    Dim pCreateFeatClass As ESRI.ArcGIS.DataManagementTools.CreateFeatureclass
    Dim pFeatClass10Min As IFeatureClass
    Dim pFeatClass20Min As IFeatureClass
    Dim pFeatClass30Min As IFeatureClass
    Dim pFeatClassQtrMi As IFeatureClass
    Dim pFeatClassHalfMi As IFeatureClass
    Dim pFeatClass1Mi As IFeatureClass
    Dim pFeatClassTransit As IFeatureClass
    Dim pPoly As IGeometry


    Private Sub frmTravelSummaryN_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'CLEAR OUT ALL THE OLD VARIABLE
        m_frmTravelSummaryBuffers = Nothing
        arrFeatLyr = Nothing
        pGPServiceArea = Nothing
        pSA = Nothing
        pAddFac = Nothing
        pSASolve = Nothing
        pCopy = Nothing
        pDelete = Nothing
        pCreateFeatClass = Nothing
        pFeatClass10Min = Nothing
        pFeatClass20Min = Nothing
        pFeatClass30Min = Nothing
        pFeatClassQtrMi = Nothing
        pFeatClassHalfMi = Nothing
        pFeatClass1Mi = Nothing
        pFeatClassTransit = Nothing
        GC.WaitForPendingFinalizers()
        GC.WaitForFullGCComplete()
        GC.Collect()
    End Sub

    Private Sub frmTravelSummaryN_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'LOAD IN POINT AND POLYGON LAYERS FOR USER TO SELECT FROM
        If Not Form_LoadData(sender, e) Then
            m_frmTravelSummaryBuffers.Close()
        End If
    End Sub

    Public Function Form_LoadData(ByVal sender As Object, ByVal e As System.EventArgs) As Boolean
        Form_LoadData = True
        Dim mxApplication As IMxApplication = Nothing
        Dim pMxDocument As IMxDocument = Nothing
        Dim mapCurrent As Map
        Dim uid As ESRI.ArcGIS.esriSystem.IUID = New ESRI.ArcGIS.esriSystem.UIDClass
        Dim enumLayer As ESRI.ArcGIS.Carto.IEnumLayer
        Dim intCount As Integer
        Dim pLyr As ILayer
        Dim pFeatLyr As IFeatureLayer
        Dim pRasterLyr As IRasterLayer
        Dim intLayer As Integer
        Dim intFeatCount As Integer = 0
        Dim pActiveView As IActiveView = Nothing

        Try

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
            m_frmTravelSummaryBuffers.cmbLayers.Items.Clear()
            arrFeatLyr.Clear()
            
            'RETRIEVE THE FEATURE LAYERS TO POPULATE FEATURE LAYER OPTIONS
            m_appEnvision.StatusBar.Message(0) = "Building list of feature layers"
            m_arrFeatureLayers = New ArrayList
            UID.Value = "{E156D7E5-22AF-11D3-9F99-00C04F6BC78E}" '= IGeoFeatureLayer
            enumLayer = mapCurrent.Layers((CType(UID, ESRI.ArcGIS.esriSystem.UID)), True) ' Explicit Cast
            enumLayer.Reset()
            pLyr = enumLayer.Next
            Do While Not (pLyr Is Nothing)
                pFeatLyr = CType(pLyr, IFeatureLayer)
                If pFeatLyr.FeatureClass.ShapeType = ESRI.ArcGIS.Geometry.esriGeometryType.esriGeometryPolygon Or pFeatLyr.FeatureClass.ShapeType = ESRI.ArcGIS.Geometry.esriGeometryType.esriGeometryPoint Then
                    arrFeatLyr.Add(pFeatLyr)
                    Me.cmbLayers.Items.Add(pFeatLyr.Name)
                    intFeatCount = intFeatCount + 1
                End If
                pLyr = enumLayer.Next()
            Loop

            If Me.cmbLayers.Items.Count <= 0 Then
                MessageBox.Show("Please add an input point or polygon layer in the current view document to use this tool.", "No Layer(s) Found", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
                GoTo CloseForm
            End If
            GoTo CleanUp
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Network Service Area Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo CloseForm
        End Try
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
        pRasterLyr = Nothing
        intLayer = Nothing
        intFeatCount = Nothing
        pActiveView = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Function

    Private Sub rdbNetwork_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdbNetwork.CheckedChanged, rdbStandard.CheckedChanged
        'SET THE ENABLED STATUS OF THE NETWORK LAYER INPUTS BASED ON USER SELECTIONS
        Me.lblStNetowrk.Enabled = (Me.rdbNetwork.Checked)
        Me.tbxNetworkLyr.Enabled = (Me.rdbNetwork.Checked)
        Me.btnNetworkLyr.Enabled = (Me.rdbNetwork.Checked)
    End Sub

    Private Sub cmbLayers_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbLayers.SelectedIndexChanged
        Dim pFeatSelection As IFeatureSelection
        Dim pCursor As ICursor = Nothing
        Dim pFeatLyr As IFeatureLayer
        Dim pFeatureClass As IFeatureClass = Nothing
        Dim intFld As Integer = 0
        Dim pField As IField
        Dim strFld As String

        Try
            'PARCEL SELECTED
            Try
                pFeatLyr = arrFeatLyr.Item(Me.cmbLayers.SelectedIndex)
                pFeatureClass = pFeatLyr.FeatureClass
                pFeatSelection = CType(pFeatLyr, IFeatureSelection)
                pFeatSelection.SelectionSet.Search(Nothing, False, pCursor)
                If pFeatSelection.SelectionSet.Count > 0 Then
                    Me.chkUseSelected.Enabled = True
                    Me.chkUseSelected.Text = "Selected Features (" & pFeatSelection.SelectionSet.Count.ToString & ")"
                Else
                    Me.chkUseSelected.Enabled = False
                    Me.chkUseSelected.Text = "Selected Features"
                End If

                'FIELDS
                'LOAD IN ALL INTEGER AND DOUBLE FIELDS INTO THE ACRES AND SQ MI CONTROLS
                If pFeatureClass Is Nothing Then
                    Me.lblFieldsId.Enabled = False
                    Me.cmbFieldsID.Enabled = False
                    Me.cmbFieldsID.Items.Clear()
                Else
                    For intFld = 0 To pFeatureClass.Fields.FieldCount - 1
                        pField = pFeatureClass.Fields.Field(intFld)
                        If pField.Type = esriFieldType.esriFieldTypeInteger Then
                            Me.cmbFieldsID.Items.Add(pField.Name)
                        End If
                    Next

                    If Me.cmbFieldsID.Items.Count > 0 Then
                        Me.lblFieldsId.Enabled = True
                        Me.cmbFieldsID.Enabled = True
                    End If
                End If
            Catch ex As Exception

            End Try

            GoTo CleanUp
        Catch ex As Exception
            GoTo CleanUp
        End Try

CleanUp:
        pFeatSelection = Nothing
        pCursor = Nothing
        pFeatLyr = Nothing
        pFeatureClass = Nothing
        intFld = Nothing
        pField = Nothing
        strFld = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub

    Private Sub btnNetworkLyr_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNetworkLyr.Click
        'SELECT THE STREET NETWORK TO CREATE THE SERVICE AREAS FCROM
        Dim pGxDialog As ESRI.ArcGIS.CatalogUI.IGxDialog
        Dim pGxObject As ESRI.ArcGIS.Catalog.IGxObject = Nothing
        Dim pLyrFilter As ESRI.ArcGIS.Catalog.IGxObjectFilter
        Dim pFGDBFilterColl As ESRI.ArcGIS.Catalog.IGxObjectFilterCollection
        Dim pGxEnumObject As ESRI.ArcGIS.Catalog.IEnumGxObject = Nothing
        Dim strPath As String
        Dim strName As String

        Try
            'CHECK TO SEE IF THE EDIT OPTION HAS BEEN ENABLED
            pGxDialog = New GxDialog
            pFGDBFilterColl = CreateObject("esriCatalogUI.GxObjectFiltercollection")
            pLyrFilter = New GxFilterNetworkDatasets
            pFGDBFilterColl.AddFilter(pLyrFilter, True)

            With pGxDialog
                .ObjectFilter = pFGDBFilterColl
                .Title = "Select an Street Network"
                .AllowMultiSelect = False
                .RememberLocation = True
                If Not .DoModalOpen(0, pGxEnumObject) Then
                    GoTo CleanUp
                Else
                    Try
                        pGxObject = pGxEnumObject.Next
                        strPath = pGxObject.Parent.FullName
                        strName = pGxObject.Name
                        Me.tbxNetworkLyr.Text = strPath & "\" & strName
                        pGxObject = Nothing
                        pGxEnumObject = Nothing
                        pGxEnumObject = Nothing
                        GC.WaitForPendingFinalizers()
                        GC.Collect()


                        ''RETRIEVE THE LOOKUP TABLES FOR THE SELECTED FILE GEODATABASE
                        'If LookUpTablesEnvisionCheck(strPath & "\" & strName) Then
                        '    Me.itmOpenEnvisionFGDB.ToolTipText = strPath & "\" & strName
                        '    m_strFeaturePath = strPath & "\" & strName
                        '    If m_strFeaturePath.Contains("\\") Then
                        '        m_strFeaturePath = m_strFeaturePath.Replace("\\", "\")
                        '    End If
                        '    SetScenarioCntrlLabels()
                        'Else
                        '    GoTo CleanUp
                        'End If

                    Catch ex As Exception
                        m_strProcessing = m_strProcessing & "Error 1 in sub itmOpenEnvisionLyr_Click: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
                        m_strProcessing = m_strProcessing & "Error Message:  " & vbNewLine & ex.Message & vbNewLine
                        MessageBox.Show(m_strProcessing)
                        GoTo CleanUp
                    End Try
                End If
                GoTo CleanUp
            End With
        Catch ex As Exception
            GoTo CleanUp
        End Try

CleanUp:
        pGxDialog = Nothing
        pGxObject = Nothing
        pLyrFilter = Nothing
        pFGDBFilterColl = Nothing
        pGxEnumObject = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub

    Private Sub btnWorkspace_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnWorkspace.Click
        'SELECT THE FILE GEODATABASE CONTAINING THE FEATURE BUFFER LAYERS
        Dim pGxDialog As ESRI.ArcGIS.CatalogUI.IGxDialog
        Dim pGxObject As ESRI.ArcGIS.Catalog.IGxObject = Nothing
        Dim pLyrFilter As ESRI.ArcGIS.Catalog.IGxObjectFilter
        Dim pFGDBFilterColl As ESRI.ArcGIS.Catalog.IGxObjectFilterCollection
        Dim pGxEnumObject As ESRI.ArcGIS.Catalog.IEnumGxObject = Nothing
        Dim strPath As String
        Dim strName As String
        Dim strBufferFileGDB As String

        Try
            'CHECK TO SEE IF THE EDIT OPTION HAS BEEN ENABLED
            pGxDialog = New GxDialog
            pFGDBFilterColl = CreateObject("esriCatalogUI.GxObjectFiltercollection")
            pLyrFilter = New GxFilterFileGeodatabases
            pFGDBFilterColl.AddFilter(pLyrFilter, True)

            With pGxDialog
                .ObjectFilter = pFGDBFilterColl
                .Title = "Select a Buffer File Geodatabase"
                .AllowMultiSelect = False
                .RememberLocation = True
                If Not .DoModalOpen(0, pGxEnumObject) Then
                    GoTo CleanUp
                Else
                    Try
                        pGxObject = pGxEnumObject.Next
                        strPath = pGxObject.Parent.FullName
                        strName = pGxObject.Name
                        pGxObject = Nothing
                        pGxEnumObject = Nothing
                        pGxEnumObject = Nothing
                        strBufferFileGDB = (strPath & "\" & strName)
                        Me.tbxWorkspace.Text = strBufferFileGDB
                    Catch ex As Exception
                        MessageBox.Show("Error in Opening selected geodatabase." & vbNewLine & ex.Message, "Opening Geodatabase Error", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
                        GoTo CleanUp
                    End Try
                End If

                GoTo CleanUp
            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Open Buffer File Geodatabase Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo CleanUp
        End Try

CleanUp:
        pGxDialog = Nothing
        pGxObject = Nothing
        pLyrFilter = Nothing
        pFGDBFilterColl = Nothing
        pGxEnumObject = Nothing
        strPath = Nothing
        strName = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()

        '        'WORKSPACE DIRECTORY
        '        Dim MyDialog As New FolderBrowserDialog
        '        MyDialog.Description = "Select a Workspace Directory:"
        '        If (MyDialog.ShowDialog() = Windows.Forms.DialogResult.OK) Then
        '            Me.tbxWorkspace.Text = MyDialog.SelectedPath.ToString
        '        End If
        '        tbxProjectName_TextChanged(sender, e)
        '        GoTo CleanUp
        'CleanUp:
        '        MyDialog = Nothing
        '        GC.WaitForPendingFinalizers()
        '        GC.Collect()
    End Sub

    Private Sub btnRun_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRun.Click
        Me.barStatusRun.Value = 1
        Me.barStatusRun.Visible = True
        'pGPServiceArea.AddOutputsToMap = False

        Dim blnCloseForm As Boolean = True
        Dim pFeatLyr As IFeatureLayer
        Dim pFeatClass As IFeatureClass
        Dim pTable As ITable
        Dim pField As IField
        Dim pSpatRef As ISpatialReference
        Dim mxApplication As IMxApplication = Nothing
        Dim pMxDocument As IMxDocument = Nothing
        Dim mapCurrent As Map
        Dim uid As ESRI.ArcGIS.esriSystem.IUID = New ESRI.ArcGIS.esriSystem.UIDClass
        Dim enumLayer As ESRI.ArcGIS.Carto.IEnumLayer
        Dim pLyr As ILayer
        Dim pTempFeatLyr As IFeatureLayer
        Dim pPntFeatClass As IFeatureClass = Nothing
        Dim pWksFactory As IWorkspaceFactory = Nothing
        Dim pFeatWks As ESRI.ArcGIS.Geodatabase.IFeatureWorkspace
        Dim pFeatSelection As IFeatureSelection = Nothing
        Dim pCursor As ICursor = Nothing
        Dim pDef As IFeatureLayerDefinition2
        Dim pQFilter As IQueryFilter
        Dim strDefExpression As String = ""
        Dim strQString As String = Nothing
        Dim pFeatureCursor As IFeatureCursor = Nothing
        Dim intTotalCount As Integer
        Dim pFeat As IFeature
        Dim arrIDs As ArrayList = New ArrayList
        Dim intIdFld As Integer
        Dim intVal As Integer
        Dim intCount As Integer

        Me.barStatusRun.Value = 1
        Me.barStatusRun.Visible = True

        'CHECK FOR INPUT STREET NETWORK DATASET
        If Me.tbxNetworkLyr.Text.Length <= 0 And Me.rdbNetwork.Checked Then
            MessageBox.Show("Please select a Street Network to build service area buffers.", "Network Input Required", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            blnCloseForm = False
            GoTo CleanUp
        End If

        'CHECK FOR THE ID FIELD
        If Me.cmbFieldsID.Text.Length <= 0 Or cmbFieldsID.Items.IndexOf(Me.cmbFieldsID.Text) <= -1 Then
            MessageBox.Show("Please select a valid id field.", "Input Required", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            blnCloseForm = False
            GoTo CleanUp
        End If

        'EXIT IF THE DESTINATION FILE GEODATABASE IS MISSING OR NOT DESIGNATED
        If Me.tbxWorkspace.Text.Length <= 0 Then
            MessageBox.Show("A file geodatabase has not be defined to store the service area feature classes.", "Input Required", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo CleanUp
        End If

        'RETRIEVE INPUT FEATURE CLASS
        Try
            pFeatLyr = arrFeatLyr(Me.cmbLayers.SelectedIndex)
            pFeatClass = pFeatLyr.FeatureClass
            pTable = CType(pFeatClass, ITable)
            If pFeatClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                'CREATE A POINT FEATURE CLASS FOR POLYGON CENTROIDS
                Try
                    pWksFactory = New FileGDBWorkspaceFactory
                    pFeatWks = pWksFactory.OpenFromFile(Me.tbxWorkspace.Text, 0)
                    Try
                        pPntFeatClass = pFeatWks.OpenFeatureClass(Me.cmbLayers.Text & "_CENTROIDS")
                    Catch ex As Exception
                        pCreateFeatClass = New ESRI.ArcGIS.DataManagementTools.CreateFeatureclass
                        pCreateFeatClass.geometry_type = "POINT"
                        pCreateFeatClass.out_path = Me.tbxWorkspace.Text
                        pCreateFeatClass.out_name = Me.cmbLayers.Text & "_CENTROIDS"
                        pField = pFeatClass.Fields.Field(pFeatClass.FindField("Shape"))
                        pSpatRef = pField.GeometryDef.SpatialReference
                        pCreateFeatClass.spatial_reference = pSpatRef
                        pGPServiceArea.AddOutputsToMap = True
                        RunTool(pGPServiceArea, pCreateFeatClass)
                        pGPServiceArea.AddOutputsToMap = False

                        pMxDocument = DirectCast(m_appEnvision.Document, IMxDocument)
                        mapCurrent = CType(pMxDocument.FocusMap, Map)
                        uid.Value = "{E156D7E5-22AF-11D3-9F99-00C04F6BC78E}" '= IGeoFeatureLayer
                        enumLayer = mapCurrent.Layers((CType(uid, ESRI.ArcGIS.esriSystem.UID)), True) ' Explicit Cast
                        enumLayer.Reset()
                        pLyr = enumLayer.Next
                        Do While Not (pLyr Is Nothing)
                            pTempFeatLyr = CType(pLyr, IFeatureLayer)
                            If pTempFeatLyr.Name = (Me.cmbLayers.Text & "_CENTROIDS") Then
                                pPntFeatClass = pTempFeatLyr.FeatureClass
                                pTable = CType(pPntFeatClass, ITable)
                                AddEnvisionField(pTable, Me.cmbFieldsID.Text, "INTEGER", 16, 0)
                                CreateCentroidFeatures(pFeatLyr, pPntFeatClass)
                                pFeatLyr = pTempFeatLyr
                                Exit Do
                            End If
                            pLyr = enumLayer.Next()
                        Loop
                    End Try

                    'TURN OFF THE USE SELECTED OPTION IF POINT WERE CREATED
                    Me.chkUseSelected.Checked = False

                Catch ex As Exception
                    MessageBox.Show("An error occured in attempting to create a centroid point layer from the selected polygon layer." & vbNewLine & ex.Message, "Create Centroid Layer Error", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
                    GoTo CleanUp
                End Try
            Else
                pPntFeatClass = pFeatLyr.FeatureClass
            End If
        Catch ex As Exception
            MessageBox.Show("Error in retrieving selected feature layer." & vbNewLine & ex.Message, "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo CleanUp
        End Try


        If Me.rdbStandard.Checked Then
            CreateStandardBuffers(pFeatClass, pPntFeatClass)
            GoTo CleanUp
        Else

            ''CREATE THE REQUIRED OUTPUT FEATURE CLASSES
            'CreateSAFeatClasses()

            ''RETRIEVE THE SERVICE AREA FEATURE CLASSES
            'pWksFactory = New FileGDBWorkspaceFactory
            'pFeatWks = pWksFactory.OpenFromFile(Me.tbxWorkspace.Text, 0)
            'Try
            '    pFeatClass10Min = pFeatWks.OpenFeatureClass("SA_10_MINUTE_AUTO")
            'Catch ex As Exception
            '    pFeatClass10Min = Nothing
            'End Try
            'Try
            '    pFeatClass20Min = pFeatWks.OpenFeatureClass("SA_20_MINUTE_AUTO")
            'Catch ex As Exception
            '    pFeatClass20Min = Nothing
            'End Try
            'Try
            '    pFeatClass30Min = pFeatWks.OpenFeatureClass("SA_30_MINUTE_AUTO")
            'Catch ex As Exception
            '    pFeatClass30Min = Nothing
            'End Try
            'Try
            '    pFeatClassQtrMi = pFeatWks.OpenFeatureClass("SA_QUARTER_MILE")
            'Catch ex As Exception
            '    pFeatClassQtrMi = Nothing
            'End Try
            'Try
            '    pFeatClassHalfMi = pFeatWks.OpenFeatureClass("SA_HALF_MILE")
            'Catch ex As Exception
            '    pFeatClassHalfMi = Nothing
            'End Try
            'Try
            '    pFeatClass1Mi = pFeatWks.OpenFeatureClass("SA_ONE_MILE")
            'Catch ex As Exception
            '    pFeatClass1Mi = Nothing
            'End Try
            'Try
            '    pFeatClassTransit = pFeatWks.OpenFeatureClass("SA_30_MINUTE_TRANSIT")
            'Catch ex As Exception
            '    pFeatClassTransit = Nothing
            'End Try

            'DEFINE WHAT WILL BE PROCESSED IN THE SELECTED NEIGHBORHOOD LAYER (ALL OR SELECTED FEATURES)
            Try
                If pFeatClass.ShapeType = esriGeometryType.esriGeometryPoint And Me.chkUseSelected.Checked Then
                    pFeatSelection = CType(pFeatLyr, IFeatureSelection)
                    pFeatSelection.SelectionSet.Search(Nothing, False, pCursor)
                    pFeatureCursor = DirectCast(pCursor, IFeatureCursor)
                    intTotalCount = pFeatSelection.SelectionSet.Count
                Else
                    pDef = pFeatLyr
                    strDefExpression = pDef.DefinitionExpression
                    pQFilter = New QueryFilter
                    pFeatSelection = CType(pFeatLyr, IFeatureSelection)
                    If strDefExpression.Length > 0 Then
                        pQFilter.WhereClause = pDef.DefinitionExpression
                        pFeatureCursor = pPntFeatClass.Search(pQFilter, False)
                        intTotalCount = pPntFeatClass.FeatureCount(pQFilter)
                    Else
                        pFeatureCursor = pPntFeatClass.Search(Nothing, False)
                        intTotalCount = pPntFeatClass.FeatureCount(Nothing)
                    End If
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try

            ''BUILD LIST OF THE FEATURE IDs TO PROCESS
            'intIdFld = pPntFeatClass.FindField(Me.cmbFieldsID.Text)
            'If intIdFld >= 0 Then
            '    pFeat = pFeatureCursor.NextFeature
            '    Do While Not pFeat Is Nothing
            '        arrIDs.Add(pFeat.Value(intIdFld))
            '        pFeat = pFeatureCursor.NextFeature
            '    Loop
            'End If

            '------------------------------------------------------------------------------------------------------------------------------------
            If Me.chkTransit.Checked And Not pFeatClassTransit Is Nothing Then
                Me.barStatusRun.Value = 0
                Me.lblStatus.Text = "Processing Transit Service Areas"

                'CYCLE THROUGH VALUE LIST
                intCount = 0
                'PROCESS THE TRANIST SERVICE AREA BUFFERS
                For Each intVal In arrIDs
                    pQFilter = New QueryFilter
                    strQString = Me.cmbFieldsID.Text & " = " & intVal.ToString
                    pQFilter.WhereClause = strQString

                    pFeatSelection.Clear()
                    pFeatSelection.SelectFeatures(pQFilter, esriSelectionResultEnum.esriSelectionResultNew, True)
                    If pFeatSelection.SelectionSet.Count > 0 Then
                      

                    End If
                    pFeatSelection.Clear()
                    intCount = intCount + 1
                    Me.barStatusRun.Value = (intCount / arrIDs.Count) * 100
                Next
            End If

            pFeatSelection.Clear()
            intCount = intCount + 1
            Me.barStatusRun.Value = (intCount / arrIDs.Count) * 100
        End If

        
        blnCloseForm = True
        GoTo CleanUp

CleanUp:
        Me.barStatusRun.Value = 0
        Me.barStatusRun.Visible = False
        Me.lblStatus.Text = ""
        mxApplication = Nothing
        pMxDocument = Nothing
        mapCurrent = Nothing
        pLyr = Nothing
        pTempFeatLyr = Nothing
        blnCloseForm = Nothing
        pFeatLyr = Nothing
        pFeatClass = Nothing
        pTable = Nothing
        pWksFactory = Nothing
        pFeatWks = Nothing
        pSpatRef = Nothing
        pField = Nothing
        pPntFeatClass = Nothing
        pFeatSelection = Nothing
        pCursor = Nothing
        pDef = Nothing
        pQFilter = Nothing
        strDefExpression = Nothing
        strQString = Nothing
        pFeatureCursor = Nothing
        intTotalCount = Nothing
        pFeat = Nothing
        arrIDs = Nothing
        intIdFld = Nothing
        intVal = Nothing
        intCount = Nothing
        uid = Nothing
        enumLayer = Nothing

        If blnCloseForm Then
            Me.Close()
        End If
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub


    Public Function CreateServiceAreasLyr(ByVal strImpedence As String, ByVal strBreakValues As String, ByVal pFeatClass As IFeatureClass, ByVal strIdFieldName As String, ByVal strOutLyaerName As String) As Boolean
        'Dim TempFeatClass As IFeatureClass
        'Dim pInsertFeatureBuffer As IFeatureBuffer
        'Dim pInsertFeatureCursor As IFeatureCursor = Nothing
        'Dim pPoly As IPolygon

        'pWksFactory = New FileGDBWorkspaceFactory
        'pFeatWks = pWksFactory.OpenFromFile(Me.tbxWorkspace.Text, 0)
        'RUNSCRIPT TO BUILD SERVIE AREA POLYGONS
        CreateServiceAreasLyr = True
        pSA = New ESRI.ArcGIS.NetworkAnalystTools.MakeServiceAreaLayer
        pAddFac = New ESRI.ArcGIS.NetworkAnalystTools.AddLocations
        pSASolve = New ESRI.ArcGIS.NetworkAnalystTools.Solve
        pCopy = New ESRI.ArcGIS.DataManagementTools.CopyFeatures
        pGPServiceArea.AddOutputsToMap = True

        Try
            'CREATE A SERVICE AREA LAYER FOR QUARTER MILE SERVICE AREAS
            pSA.in_network_dataset = Me.tbxNetworkLyr.Text
            'Impedance attribute is Length or Time
            pSA.impedance_attribute = strImpedence  '"Time"
            pSA.merge = "NO_MERGE"
            'Breake vales are seperated by a comma and are the length or time values
            pSA.default_break_values = strBreakValues    '"30"
            pSA.output_layer = "TEST_SA_LAYER"
            pSA.polygon_type = "SIMPLE_POLYS"
            pSA.polygon_trim = "FALSE"
            pSA.hierarchy = "TRUE"
            pSA.out_network_analysis_layer = "ET_SA_TEMP"
            RunTool(pGPServiceArea, pSA)
        Catch ex As Exception
        End Try

        Try
            'ADD LOCATIONS, IF LAYER HAS A SELECTION, THEN ONLY SELECTED FEATURES ARE ADDED
            pAddFac.in_network_analysis_layer = "ET_SA_TEMP"
            'pAddFac.in_table = pFeatLyr.Name
            pAddFac.sub_layer = "Facilities"
            pAddFac.field_mappings = "Name " & strIdFieldName ' & " #;CurbApproach # 0;Attr_Length # 0;Attr_Time # 0;Breaks_Length # #;Breaks_Time # #"
            pAddFac.search_tolerance = "1000 Meters"
            pAddFac.sort_field = ""
            pAddFac.search_criteria = "'SDC Edge Source' SHAPE"
            pAddFac.match_type = "MATCH_TO_CLOSEST"
            pAddFac.append = "CLEAR"
            pAddFac.snap_to_position_along_network = "SNAP"
            pAddFac.snap_offset = "5 Meters"
            pAddFac.exclude_restricted_elements = "INCLUDE"
            pAddFac.search_query = "'SDC Edge Source' #"
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Service Area Processing Error")
            GoTo CleanUp
        End Try
        RunTool(pGPServiceArea, pAddFac)

        Try
            'SOLVE FOR CURRENT SERVICE AREA(S)
            pSASolve.in_network_analysis_layer = "ET_SA_TEMP"
            pSASolve.ignore_invalids = "SKIP"
            pSASolve.terminate_on_solve_error = "TERMINATE"
            RunTool(pGPServiceArea, pSASolve)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Service Area Processing Error")
            CreateServiceAreasLyr = False
            GoTo CleanUp
        End Try
        Me.barStatusRun.Value = 25

        'COPY RESULTS TO FILE GEODATABASE
        pCopy = New ESRI.ArcGIS.DataManagementTools.CopyFeatures
        pCopy.in_features = "POLYGONS"
        pCopy.out_feature_class = Me.tbxWorkspace.Text & "\" & strOutLyaerName '& intCount.ToString
        pGPServiceArea.AddOutputsToMap = False
        RunTool(pGPServiceArea, pCopy)
        pCopy = Nothing
        GC.WaitForFullGCComplete()
        GC.WaitForPendingFinalizers()
        GC.Collect()

        ''RETRIEVE THE TEMPORARY FEAT CLASS
        'TempFeatClass = pFeatWks.OpenFeatureClass("ET_TEMP_SA")
        'pFeatureCursor = TempFeatClass.Search(Nothing, False)
        'intTotalCount = TempFeatClass.FeatureCount(Nothing)
        'If intTotalCount >= 1 Then
        '    pFeat = pFeatureCursor.NextFeature
        '    Do While Not pFeat Is Nothing
        '        pPoly = pFeat.Shape
        '        pInsertFeatureCursor = pFeatClassTransit.Insert(True)
        '        pInsertFeatureBuffer = pFeatClassTransit.CreateFeatureBuffer
        '        pInsertFeatureBuffer.Shape = pPoly
        '        pInsertFeatureBuffer.Value(pFeatClassTransit.FindField(Me.cmbFieldsID.Text)) = intVal
        '        pInsertFeatureCursor.InsertFeature(pInsertFeatureBuffer)
        '        pInsertFeatureCursor.Flush()
        '        pPoly = Nothing
        '        pFeat = pFeatureCursor.NextFeature
        '    Loop
        'End If

        'TempFeatClass = Nothing
        'pFeatureCursor = Nothing
        'intTotalCount = Nothing
        'pFeat = Nothing
        'pInsertFeatureCursor = Nothing
        'pInsertFeatureBuffer = Nothing
        'pPoly = Nothing
        GC.Collect()
        GC.WaitForFullGCComplete()
        GC.WaitForPendingFinalizers()


        ''DELETE THE TEMPORARY FEATURE LAYER
        'pDelete = New ESRI.ArcGIS.DataManagementTools.Delete
        'pDelete.in_data = Me.tbxWorkspace.Text & "\ET_TEMP_SA" '& intCount.ToString
        'pDelete.data_type = "FeatureClass"
        'RunTool(pGPServiceArea, pDelete)
        'pDelete = Nothing

        Me.barStatusRun.Value = 35

CleanUp:
        pSA = Nothing
        pAddFac = Nothing
        pSASolve = Nothing
        pCopy = Nothing
        'TempFeatClass = Nothing
        'pInsertFeatureBuffer = Nothing
        'pInsertFeatureCursor = Nothing
        'pPoly = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Function

    Private Function CalcIdFld(ByVal pTable As ITable) As Boolean
        CalcIdFld = True

        Dim pCursor As ICursor
        Dim pCalc As ICalculator
        Try
            pCalc = New Calculator
            pCursor = pTable.Update(Nothing, False)
            With pCalc
                .Cursor = pCursor
                .Field = Me.cmbFieldsID.Text
                .Expression = "SPLIT( [NAME]," & """" & " : " & """" & ")(0)"
            End With
            pCalc.Calculate()
            GoTo CleanUp
        Catch ex As Exception
            CalcIdFld = False
            'MessageBox.Show(ex.Message, "Calculate Field Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo CleanUp
        End Try

CleanUp:
        pCalc = Nothing
        pCursor = Nothing
        pTable = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Function

    Private Sub CreateStandardBuffers(ByVal pFeatClass As IFeatureClass, ByVal pPntFeatClass As IFeatureClass)
        Me.barStatusRun.Value = 1
        Me.barStatusRun.Visible = True
        pGPServiceArea.AddOutputsToMap = True


        Dim arrBufferLayers As ArrayList = New ArrayList
        Dim arrBufferDistances As ArrayList = New ArrayList
        Dim arrBufferType As ArrayList = New ArrayList
        Dim pBuffer As ESRI.ArcGIS.AnalysisTools.Buffer = New ESRI.ArcGIS.AnalysisTools.Buffer
        Dim pCreateTable As ESRI.ArcGIS.DataManagementTools.CreateTable
        Dim pCreateFeatClass As ESRI.ArcGIS.DataManagementTools.CreateFeatureclass
        Dim strIdFieldName As String = Me.cmbFieldsID.Text
        Dim pQueryFilter As IQueryFilter = New QueryFilter
        Dim blnCloseForm As Boolean = True
        Dim pWksFactory As IWorkspaceFactory = Nothing
        Dim pFeatWks As ESRI.ArcGIS.Geodatabase.IFeatureWorkspace
        Dim pSpatRef As ISpatialReference
        Dim pField As IField
        Dim strLyrName As String = ""
        Dim intLyrCount As Integer = 0
        Dim intPercent As Integer = 0
        Dim pTable As ITable

        'BUILD LIST OF BUFFER LAYERS AND DISTANCES
        If Me.chkQtrMiAuto.Checked Then
            arrBufferLayers.Add("SA_QUARTER_MILE")
            arrBufferDistances.Add("1")
            arrBufferType.Add("Miles")
        End If
        If Me.chkHalfMiAuto.Checked Then
            arrBufferLayers.Add("SA_HALF_MILE")
            arrBufferDistances.Add("0.5")
            arrBufferType.Add("Miles")
        End If
        If Me.chkOneMiAuto.Checked Then
            arrBufferLayers.Add("SA_ONE_MILE")
            arrBufferDistances.Add("0.25")
            arrBufferType.Add("Miles")
        End If
        If Me.chk10min.Checked Then
            arrBufferLayers.Add("SA_10_MINUTE_AUTO")
            arrBufferDistances.Add("4.5")
            arrBufferType.Add("Miles")
        End If
        If Me.chk20min.Checked Then
            arrBufferLayers.Add("SA_20_MINUTE_AUTO")
            arrBufferDistances.Add("9")
            arrBufferType.Add("Miles")
        End If
        If Me.chk30min.Checked Then
            arrBufferLayers.Add("SA_30_MINUTE_AUTO")
            arrBufferDistances.Add("13.5")
            arrBufferType.Add("Miles")
        End If
        If Me.chkTransit.Checked Then
            arrBufferLayers.Add("SA_30_MINUTE_TRANSIT")
            arrBufferDistances.Add("10")
            arrBufferType.Add("Miles")
        End If

        Me.barStatusRun.Value = 5

        'EXIT IF NOT BUFFERS ARE CHECKED
        If arrBufferLayers.Count <= 0 Then
            MessageBox.Show("No buffer layer(s) were selected to be created.", "Input Required", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            Me.barStatusRun.Value = 100
            GoTo CleanUp
        End If

        'EXIT IF NOT BUFFERS ARE CHECKED
        If Me.tbxWorkspace.Text.Length <= 0 Then
            MessageBox.Show("Please select the File geodatabase to have output layers written too.", "Input Required", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            Me.barStatusRun.Value = 100
            GoTo CleanUp
        End If

        '***************************************************************************************************************************
        'CYCLE THROUGH BUFFER LAYER NAMES AND DISTANCES TO CREATE REQUIRED LAYERS
        '***************************************************************************************************************************
        For Each strLyrName In arrBufferLayers
            Try
                'CREATE A SERVICE AREA LAYER
                pBuffer.buffer_distance_or_field = 1
                If pFeatClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                    pBuffer.in_features = pPntFeatClass
                Else
                    If Me.chkUseSelected.Checked Then
                        pBuffer.in_features = Me.cmbLayers.Text
                    Else
                        pBuffer.in_features = pFeatClass
                    End If
                End If
                pBuffer.line_end_type = "ROUND"
                pBuffer.line_side = "FULL"
                pBuffer.buffer_distance_or_field = arrBufferDistances(intLyrCount) & " " & arrBufferType(intLyrCount)
                pBuffer.dissolve_option = "LIST"
                pBuffer.dissolve_field = Me.cmbFieldsID.Text
                pBuffer.out_feature_class = Me.tbxWorkspace.Text & "\" & strLyrName
                RunTool(pGPServiceArea, pBuffer)
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Standard Buffer Processing Error")
                GoTo CleanUp
            End Try
            intLyrCount = intLyrCount + 1
            Me.barStatusRun.Value = (intLyrCount / arrBufferLayers.Count) * 100
        Next

        GoTo CleanUp

CleanUp:
        Me.barStatusRun.Value = 0
        Me.barStatusRun.Visible = False
        pBuffer = Nothing
        pCreateTable = Nothing
        pCreateFeatClass = Nothing
        strIdFieldName = Nothing
        pQueryFilter = Nothing
        blnCloseForm = Nothing
        pFeatClass = Nothing
        pTable = Nothing
        pWksFactory = Nothing
        pFeatWks = Nothing
        pSpatRef = Nothing
        pField = Nothing
        pPntFeatClass = Nothing
        strLyrName = Nothing
        If blnCloseForm Then
            Me.Close()
        End If
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub

    Private Function CreateSAFeatClasses() As Boolean
        'RETRIEVE INPUT FEATURE CLASS
        Dim arrBufferLayers As ArrayList = New ArrayList
        Dim strLyrName As String = ""
        Dim pWksFactory As IWorkspaceFactory = Nothing
        Dim pFeatWks As ESRI.ArcGIS.Geodatabase.IFeatureWorkspace
        Dim pSpatRef As ISpatialReference
        Dim pField As IField
        Dim intLyrCount As Integer = 0
        Dim pFeatLyr As IFeatureLayer
        Dim pFeatClass As IFeatureClass
        Dim pTempFeatClass As IFeatureClass
        Dim pTable As ITable

        'BUILD LIST OF BUFFER LAYERS AND DISTANCES
        If Me.chkQtrMiAuto.Checked Then
            arrBufferLayers.Add("SA_QUARTER_MILE")
        End If
        If Me.chkHalfMiAuto.Checked Then
            arrBufferLayers.Add("SA_HALF_MILE")
        End If
        If Me.chkOneMiAuto.Checked Then
            arrBufferLayers.Add("SA_ONE_MILE")
        End If
        If Me.chk10min.Checked Then
            arrBufferLayers.Add("SA_10_MINUTE_AUTO")
        End If
        If Me.chk20min.Checked Then
            arrBufferLayers.Add("SA_20_MINUTE_AUTO")
        End If
        If Me.chk30min.Checked Then
            arrBufferLayers.Add("SA_30_MINUTE_AUTO")
        End If
        If Me.chkTransit.Checked Then
            arrBufferLayers.Add("SA_30_MINUTE_TRANSIT")
        End If

        'EXIT IF THE DESTINATION FILE GEODATABASE IS MISSING OR NOT DESIGNATED
        If Me.tbxWorkspace.Text.Length <= 0 Then
            MessageBox.Show("A file geodatabase has not be defined to store the service area feature classes.", "Input Required", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo CleanUp
        End If

        'CREATE A POINT FEATURE CLASS FOR POLYGON CENTROIDS
        Try
            pFeatLyr = arrFeatLyr(Me.cmbLayers.SelectedIndex)
            pFeatClass = pFeatLyr.FeatureClass
            pField = pFeatClass.Fields.Field(pFeatClass.FindField("Shape"))
            pSpatRef = pField.GeometryDef.SpatialReference
            pWksFactory = New FileGDBWorkspaceFactory
            pFeatWks = pWksFactory.OpenFromFile(Me.tbxWorkspace.Text, 0)

            If arrBufferLayers.Count > 0 Then
                For Each strLyrName In arrBufferLayers
                    Try
                        pCreateFeatClass = New ESRI.ArcGIS.DataManagementTools.CreateFeatureclass
                        pCreateFeatClass.geometry_type = "POLYGON"
                        pCreateFeatClass.out_path = Me.tbxWorkspace.Text
                        pCreateFeatClass.out_name = strLyrName
                        pCreateFeatClass.spatial_reference = pSpatRef
                        RunTool(pGPServiceArea, pCreateFeatClass)
                        pCreateFeatClass = Nothing

                        pTempFeatClass = pFeatWks.OpenFeatureClass(strLyrName)
                        pTable = CType(pTempFeatClass, ITable)
                        AddEnvisionField(pTable, Me.cmbFieldsID.Text, "INTEGER", 16, 0)
                        pTempFeatClass = Nothing
                        pTable = Nothing
                    Catch ex As Exception
                        MessageBox.Show(ex.Message)
                    End Try

                    intLyrCount = intLyrCount + 1
                    Me.barStatusRun.Value = (intLyrCount / arrBufferLayers.Count) * 100
                Next
            End If

            GoTo CleanUp
        Catch ex As Exception
            MessageBox.Show("An error occured in attempting to create a service area layer (" & strLyrName & ") from the selected polygon layer." & vbNewLine & ex.Message, "Create Service Area Layer Error", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo CleanUp
        End Try

CleanUp:
        arrBufferLayers = Nothing
        strLyrName = Nothing
        pWksFactory = Nothing
        pFeatWks = Nothing
        pSpatRef = Nothing
        pField = Nothing
        intLyrCount = Nothing
        pFeatLyr = Nothing
        pFeatClass = Nothing
        pTempFeatClass = Nothing
        pTable = Nothing
        GC.WaitForPendingFinalizers()
        GC.WaitForFullGCComplete()
        GC.Collect()

    End Function

    Private Function CreateCentroidFeatures(ByVal pLayer As IFeatureLayer, ByVal pOutFeatClass As IFeatureClass) As Boolean
        'GENERATE THE GRID CELL POLYGONS
        CreateCentroidFeatures = True

        Dim intIdFld As Integer = pOutFeatClass.FindField(Me.cmbFieldsID.Text)
        Dim intFromIdFld As Integer = pLayer.FeatureClass.FindField(Me.cmbFieldsID.Text)
        Dim pFeatSelection As IFeatureSelection = CType(pLayer, IFeatureSelection)
        Dim pCursor As ICursor = Nothing
        Dim pFeatureCursor As IFeatureCursor
        Dim intTotalCount As Integer = 0
        Dim pInFeatClass As IFeatureClass = pLayer.FeatureClass
        Dim pFeat As IFeature
        Dim intRecCount As Integer = 0
        Dim pArea As IArea
        Dim pntCentroid As ESRI.ArcGIS.Geometry.IPoint = New ESRI.ArcGIS.Geometry.Point
        Dim intId As Integer
        Dim pInsertFeatureBuffer As IFeatureBuffer
        Dim pInsertFeatureCursor As IFeatureCursor = Nothing
        Dim intNewFeatureCount As Integer = 0

        Try
            If intIdFld = -1 Or intFromIdFld <= -1 Then
                CreateCentroidFeatures = False
                GoTo CleanUp
            End If

            'ALL FEATURES OR SELECTED FEAGTURE(S)
            If Me.chkUseSelected.Checked Then
                pFeatSelection.SelectionSet.Search(Nothing, False, pCursor)
                pFeatureCursor = DirectCast(pCursor, IFeatureCursor)
                intTotalCount = pFeatSelection.SelectionSet.Count
            Else
                pFeatureCursor = pInFeatClass.Search(Nothing, False)
                intTotalCount = pInFeatClass.FeatureCount(Nothing)
            End If

            pFeat = pFeatureCursor.NextFeature
            Do While Not pFeat Is Nothing
                intRecCount = intRecCount + 1
                pArea = pFeat.Shape
                pntCentroid = pArea.Centroid
                Try
                    intId = pFeat.Value(intFromIdFld)
                Catch ex As Exception
                    intId = 0
                End Try

                'FEATURE BUFFER SETUP
                pInsertFeatureCursor = pOutFeatClass.Insert(True)
                pInsertFeatureBuffer = pOutFeatClass.CreateFeatureBuffer
                pInsertFeatureBuffer.Shape = pntCentroid
                pInsertFeatureBuffer.Value(intIdFld) = intId
                pInsertFeatureCursor.InsertFeature(pInsertFeatureBuffer)
                intNewFeatureCount = intNewFeatureCount + 1
                If intNewFeatureCount = 100 Then
                    pInsertFeatureCursor.Flush()
                    intNewFeatureCount = 0
                End If

                pArea = Nothing
                pntCentroid = Nothing
                intId = 0
                GC.WaitForPendingFinalizers()
                GC.Collect()
                pFeat = pFeatureCursor.NextFeature
            Loop
            If Not pInsertFeatureCursor Is Nothing Then
                pInsertFeatureCursor.Flush()
            End If

        Catch ex As Exception
            CreateCentroidFeatures = False
            GoTo CleanUp
        End Try

CleanUp:
        'Try
        '    System.Runtime.InteropServices.Marshal.ReleaseComObject(pLayer)
        '    System.Runtime.InteropServices.Marshal.ReleaseComObject(pOutFeatClass)
        '    System.Runtime.InteropServices.Marshal.ReleaseComObject(pInsertFeatureCursor)
        '    System.Runtime.InteropServices.Marshal.ReleaseComObject(pFeatureCursor)
        'Catch ex As Exception

        'End Try
        intIdFld = Nothing
        intFromIdFld = Nothing
        pFeatSelection = Nothing
        pCursor = Nothing
        pFeatureCursor = Nothing
        intTotalCount = Nothing
        pInFeatClass = Nothing
        pFeat = Nothing
        intRecCount = Nothing
        pArea = Nothing
        pntCentroid = Nothing
        intId = Nothing
        pInsertFeatureBuffer = Nothing
        pInsertFeatureCursor = Nothing
        intNewFeatureCount = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()


    End Function

    Private Sub btnCheckAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCheckAll.Click
        'SET THE CHECKED STATUS OF EACH BUFFER INPUT CONTROL
        SelectBuffers(True)
    End Sub

    Private Sub btnUnselect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUnselect.Click
        'SET THE CHECKED STATUS OF EACH BUFFER INPUT CONTROL
        SelectBuffers(False)
    End Sub

    Private Sub SelectBuffers(ByVal blnStatus As Boolean)
        'SET THE BUFFER STATUS BASE UPON INPUT BOOLEAN VALUE
        Me.chk10min.Checked = blnStatus
        Me.chk20min.Checked = blnStatus
        Me.chk30min.Checked = blnStatus
        Me.chkQtrMiAuto.Checked = blnStatus
        Me.chkHalfMiAuto.Checked = blnStatus
        Me.chkOneMiAuto.Checked = blnStatus
        Me.chkTransit.Checked = blnStatus
    End Sub
End Class