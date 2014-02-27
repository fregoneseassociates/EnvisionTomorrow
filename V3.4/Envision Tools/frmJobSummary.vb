'MICROSOFT REFERENCES
Imports System
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Runtime.InteropServices
Imports System.Collections
Imports System.IO
Imports System.Drawing.Printing
Imports System.Data
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
Imports ESRI.ArcGIS.AnalysisTools

Public Class frmJobSummary
    Dim blnOpenForm As Boolean = True
    Dim arrFeatLyr As ArrayList = New ArrayList
    Dim blnLoading As Boolean = True
    Dim pGPServiceArea As Geoprocessor = New Geoprocessor

    Private Sub frmJobSummary_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        m_frmSumJobs = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub

    Private Sub frmJobSummary_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'LOAD IN POINT AND POLYGON LAYERS FOR USER TO SELECT FROM
        blnOpenForm = True
        If Not Form_LoadData(sender, e) Then
            m_frmSumJobs.Close()
            blnOpenForm = False
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
        Dim pRasterLyr As IRasterLayer
        Dim intLayer As Integer
        Dim intFeatCount As Integer = 0
        Dim pActiveView As IActiveView = Nothing

        Try
            If m_pEditFeatureLyr Is Nothing Then
                MessageBox.Show("An Envision Parcel layer has not been defined. Select FILE | Open Envision File Geodatabase, to define a parcel layer.", "Envision Layer Not Defined", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
                GoTo CloseForm
            Else
                'CEHCK THE ENVISION LAYER FOR THE REQUIRED 'EMP' FIELD
                If m_pEditFeatureLyr.FeatureClass.FindField("EMP") <= -1 Then
                    MessageBox.Show("The required field, EMP, could not be found in the selected Envision data layer.", "Missing Field, EMP", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
                    GoTo CloseForm
                End If
            End If

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
            m_frmSumJobs.cmbLayers.Items.Clear()
            arrFeatLyr.Clear()
            If mapCurrent.LayerCount > 0 Then
                For intLayer = 0 To mapCurrent.LayerCount - 1
                    pLyr = CType(mapCurrent.Layer(intLayer), ILayer)
                    If TypeOf mapCurrent.Layer(intLayer) Is IFeatureLayer Then
                        pFeatLyr = CType(mapCurrent.Layer(intLayer), IFeatureLayer)
                        If pFeatLyr.FeatureClass.ShapeType = ESRI.ArcGIS.Geometry.esriGeometryType.esriGeometryPolygon Or pFeatLyr.FeatureClass.ShapeType = ESRI.ArcGIS.Geometry.esriGeometryType.esriGeometryPoint Then
                            arrFeatLyr.Add(pFeatLyr)
                            Me.cmbLayers.Items.Add(pFeatLyr.Name)
                            intFeatCount = intFeatCount + 1
                        End If
                        pFeatLyr = Nothing
                    End If
                Next
                GoTo CleanUp
            Else
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

    Private Sub cmbLayers_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbLayers.SelectedIndexChanged
        Dim pFeatSelection As IFeatureSelection
        Dim pCursor As ICursor = Nothing
        Dim pFeatLyr As IFeatureLayer
        Dim pFeatureClass As IFeatureClass = Nothing
        Dim intFld As Integer = 0
        Dim pField As IField

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
        pFeatLyr = Nothing
        pFeatSelection = Nothing
        pCursor = Nothing
        pFeatureClass = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub

    Private Sub btnRun_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRun.Click
        Dim pET_FeatClass As IFeatureClass
        Dim intET_EmpFld As Integer = -1
        Dim pInputFeatLyr As IFeatureLayer
        Dim pInputFeatClass As IFeatureClass
        Dim pInputFeatSelection As IFeatureSelection
        Dim pInputCursor As ICursor = Nothing
        Dim pInputFeatureCursor As IFeatureCursor = Nothing
        Dim intInputIdFld As Integer = -1
        Dim intInputJobsFld As Integer = -1
        Dim intInputTotalCount As Integer = 0
        Dim pInputFeat As IFeature
        Dim pBufferFeatClass As IFeatureClass
        Dim strBufferLayerNames As String = "SA_QUARTER_MILE,SA_HALF_MILE,SA_ONE_MILE,SA_10_MINUTE_AUTO,SA_20_MINUTE_AUTO,SA_30_MINUTE_AUTO,SA_30_MINUTE_TRANSIT"
        Dim strBufferLyrName As String = ""
        Dim intBufferIdFld As Integer
        Dim intBufferJobsFld As Integer
        Dim lyrBuffer As IFeatureLayer
        Dim pBufferCursor As ICursor = Nothing
        Dim pBufferFeat As IFeature
        Dim pBufferFeatSelection As IFeatureSelection
        Dim pBufferFeatureCursor As IFeatureCursor
        Dim pWksFactory As IWorkspaceFactory = Nothing
        Dim pFeatWks As ESRI.ArcGIS.Geodatabase.IFeatureWorkspace
        Dim intBuff_ID As Integer
        Dim pBuffer_QueryFilter As IQueryFilter
        Dim strBuffer_WhereClause As String
        Dim pQBLayer As IQueryByLayer
        Dim pET_FeatSelection As IFeatureSelection
        Dim pET_Cursor As ICursor = Nothing
        Dim pET_FeatureCursor As IFeatureCursor
        Dim pET_Feature As IFeature
        Dim pET_EmpFld As Integer
        Dim intValue As Integer = 0
        Dim intFailed As Integer = 0
        Dim intSummaryTotal As Integer = 0
        Dim intCount As Integer = 0
        Dim pTable As ITable

        'FIRST CHECK THAT THE ENVISION LAYER IS DEFINED TO PULL JOB NUMBERS
        If m_pEditFeatureLyr Is Nothing Then
            MessageBox.Show("An Envision Parcel layer has not been defined. Select FILE | Open Envision File Geodatabase, to define a parcel layer.", "Envision Layer Not Defined", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo CleanUp
        Else
            'CEHCK THE ENVISION LAYER FOR THE REQUIRED 'EMP' FIELD
            If m_pEditFeatureLyr.FeatureClass.FindField("EMP") <= -1 Then
                MessageBox.Show("The required field, EMP, could not be found in the selected Envision data layer.", "Missing Field, EMP", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
                GoTo CleanUp
            Else
                pET_EmpFld = m_pEditFeatureLyr.FeatureClass.FindField("EMP")
            End If
        End If

        pET_FeatClass = m_pEditFeatureLyr.FeatureClass
        intET_EmpFld = pET_FeatClass.FindField("EMP")
        pET_FeatSelection = CType(m_pEditFeatureLyr, IFeatureSelection)

        'RETRIEVE THE INPUT LAYER AND ID FIELD
        pInputFeatLyr = arrFeatLyr.Item(Me.cmbLayers.SelectedIndex)
        pInputFeatClass = pInputFeatLyr.FeatureClass
        pInputFeatSelection = CType(pInputFeatLyr, IFeatureSelection)
        intInputIdFld = pInputFeatClass.FindField(Me.cmbFieldsID.Text)

        Try
            'ADD MISSING FIELDS IF NEEDED
            For Each strBufferLyrName In strBufferLayerNames.Split(",")
                'CHECK FOR THE WRITE FIELD, ADD IF MISSING
                If pInputFeatClass.FindField(strBufferLyrName) = -1 Then
                    pTable = CType(pInputFeatClass, ITable)
                    AddEnvisionField(pTable, (strBufferLyrName), "DOUBLE", 16, 6)
                End If
            Next
        Catch ex As Exception
        End Try

        'CYCLE THROUGH ALL OR SELECTED INPUT FEATURES
        'ALL FEATURES OR SELECTED FEAGTURE(S)
        If Me.chkUseSelected.Checked Then
            pInputFeatSelection.SelectionSet.Search(Nothing, False, pInputCursor)
            pInputFeatureCursor = DirectCast(pInputCursor, IFeatureCursor)
            intInputTotalCount = pInputFeatSelection.SelectionSet.Count
        Else
            pInputFeatureCursor = pInputFeatClass.Search(Nothing, False)
            intInputTotalCount = pInputFeatClass.FeatureCount(Nothing)
        End If

        Try
            Me.lblStatus.Visible = True
            pInputFeat = pInputFeatureCursor.NextFeature
            Do While Not pInputFeat Is Nothing
                intCount = intCount + 1
                Me.lblStatus.Text = intCount.ToString & " of " & intInputTotalCount.ToString
                Me.Refresh()
                Try
                    intBuff_ID = pInputFeat.Value(intInputIdFld)
                    'CYCLE TROUGH AND RETRIEVE THE SERVICE AREA BUFFER LAYERS
                    For Each strBufferLyrName In strBufferLayerNames.Split(",")
                        Try
                            pWksFactory = New FileGDBWorkspaceFactory
                            pFeatWks = pWksFactory.OpenFromFile(Me.tbxWorkspace.Text, 0)
                            pBufferFeatClass = pFeatWks.OpenFeatureClass(strBufferLyrName)
                            lyrBuffer = New FeatureLayer
                            lyrBuffer.FeatureClass = pBufferFeatClass
                            intBufferIdFld = pBufferFeatClass.FindField(Me.cmbFieldsID.Text)
                            'CHECK FOR THE WRITE FIELD, ADD IF MISSING
                            If pBufferFeatClass.FindField(("JOBS_SCENARIO_" & CStr(m_intEditScenario))) = -1 Then
                                pTable = CType(pBufferFeatClass, ITable)
                                AddEnvisionField(pTable, ("JOBS_SCENARIO_" & CStr(m_intEditScenario)), "DOUBLE", 16, 6)
                            End If
                            intBufferJobsFld = pBufferFeatClass.FindField(("JOBS_SCENARIO_" & CStr(m_intEditScenario)))
                            If intBufferIdFld <= -1 Or intBufferJobsFld <= -1 Then
                                Continue For
                            End If
                            pBufferFeatSelection = CType(lyrBuffer, IFeatureSelection)

                            Try
                                'If strDevType.Length > 0 Then
                                pBuffer_QueryFilter = New QueryFilter
                                strBuffer_WhereClause = Me.cmbFieldsID.Text & " = " & intBuff_ID.ToString
                                pBuffer_QueryFilter.WhereClause = strBuffer_WhereClause
                                pBufferFeatSelection.Clear()
                                pBufferFeatSelection.SelectFeatures(pBuffer_QueryFilter, esriSelectionResultEnum.esriSelectionResultNew, True)
                                pBufferFeatSelection.SelectionSet.Search(Nothing, False, pBufferCursor)
                                pBufferFeatureCursor = DirectCast(pBufferCursor, IFeatureCursor)
                                pBufferFeat = pBufferFeatureCursor.NextFeature
                                intSummaryTotal = 0
                                If pBufferFeatSelection.SelectionSet.Count >= 1 Then
                                    pQBLayer = New QueryByLayer
                                    With pQBLayer
                                        .ByLayer = lyrBuffer
                                        .FromLayer = pET_FeatSelection
                                        .BufferDistance = 0
                                        .BufferUnits = esriUnits.esriMiles
                                        .ResultType = esriSelectionResultEnum.esriSelectionResultNew
                                        .LayerSelectionMethod = esriLayerSelectionMethod.esriLayerSelectIntersect
                                        .UseSelectedFeatures = True
                                        pET_FeatSelection.SelectionSet = .Select
                                    End With

                                    pET_FeatSelection.SelectionSet.Search(Nothing, False, pET_Cursor)
                                    pET_FeatureCursor = DirectCast(pET_Cursor, IFeatureCursor)
                                    If pET_FeatSelection.SelectionSet.Count > 0 Then
                                        pET_Feature = pET_FeatureCursor.NextFeature
                                        Do While Not pET_Feature Is Nothing
                                            Try
                                                intValue = pET_Feature.Value(pET_EmpFld)
                                            Catch ex As Exception
                                                intValue = 0
                                            End Try
                                            intSummaryTotal = intSummaryTotal + intValue
                                            pET_Feature = pET_FeatureCursor.NextFeature
                                        Loop
                                    End If
                                End If

                                pBufferFeat.Value(intBufferJobsFld) = intSummaryTotal
                                pBufferFeat.Store()

                                'SAVE THE TOTAL JOBS TO THE BUFFER FEATURE
                                Try
                                    'SAVE THE TOTAL JOBS TO THE BUFFER FEATURE
                                    If pInputFeatClass.FindField(strBufferLyrName) >= 0 Then
                                        pInputFeat.Value(pInputFeat.Fields.FindField(strBufferLyrName)) = intSummaryTotal
                                    End If
                                    pInputFeat.Store()
                                Catch ex As Exception
                                End Try
                            Catch ex As Exception
                                intFailed = intFailed + 1
                            End Try
                        Catch ex As Exception
                        End Try
                    Next  'END BUFFER LAYER LOOP
                Catch ex As Exception
                End Try

                GC.WaitForPendingFinalizers()
                GC.Collect()
                pInputFeat = pInputFeatureCursor.NextFeature
            Loop
        Catch ex As Exception

        End Try

CleanUp:
        pET_FeatClass = Nothing
        intET_EmpFld = Nothing
        pInputFeatLyr = Nothing
        pInputFeatClass = Nothing
        pInputFeatSelection = Nothing
        pInputCursor = Nothing
        pInputFeatureCursor = Nothing
        intInputIdFld = Nothing
        intInputJobsFld = Nothing
        intInputTotalCount = Nothing
        pInputFeat = Nothing
        pBufferFeatClass = Nothing
        strBufferLayerNames = Nothing
        strBufferLyrName = Nothing
        intBufferIdFld = Nothing
        intBufferJobsFld = Nothing
        lyrBuffer = Nothing
        pBufferCursor = Nothing
        pBufferFeat = Nothing
        pBufferFeatSelection = Nothing
        pBufferFeatureCursor = Nothing
        pWksFactory = Nothing
        pFeatWks = Nothing
        intBuff_ID = Nothing
        pBuffer_QueryFilter = Nothing
        strBuffer_WhereClause = Nothing
        pQBLayer = Nothing
        pET_FeatSelection = Nothing
        pET_Cursor = Nothing
        pET_FeatureCursor = Nothing
        pET_Feature = Nothing
        pET_EmpFld = Nothing
        intValue = Nothing
        intFailed = Nothing
        intSummaryTotal = Nothing
        intCount = Nothing
        pTable = Nothing
        Me.lblStatus.Visible = False
        Me.Close()
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub

    Private Sub btnWorkspace_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnWorkspace.Click

        'SELECT THE FILE GEODATABASE TO EDIT
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
            pLyrFilter = New GxFilterFileGeodatabases
            pFGDBFilterColl.AddFilter(pLyrFilter, True)

            With pGxDialog
                .ObjectFilter = pFGDBFilterColl
                .Title = "Select an Buffer File Geodatabase"
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
                        GC.WaitForPendingFinalizers()
                        GC.Collect()

                        Me.tbxWorkspace.Text = strPath & "\" & strName

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
            m_strProcessing = m_strProcessing & "Error 2 in sub itmOpenEnvisionLyr_Click: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
            m_strProcessing = m_strProcessing & "Error Message:  " & vbNewLine & ex.Message & vbNewLine
            MessageBox.Show(ex.Message, "Open File Geodatabase Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
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
End Class