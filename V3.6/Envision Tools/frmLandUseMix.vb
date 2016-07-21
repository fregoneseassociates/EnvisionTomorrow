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
Imports ESRI.ArcGIS.Catalog
Imports ESRI.ArcGIS.CatalogUI

Public Class frmLandUseMix
    Public m_arrPolyFeatureLayers As New ArrayList
    Dim m_arrDevTypes As ArrayList = New ArrayList
    Dim m_arrLUMIXCoulmnNames As ArrayList = New ArrayList
    Dim m_strMissingLyrs As String = ""
    Dim m_intMissingLyrs As Integer = 0
    Dim pSA_QUARTER_MILELyr As IFeatureLayer = Nothing
    Dim pSA_HALF_MILELyr As IFeatureLayer = Nothing
    Dim pSA_ONE_MILELyr As IFeatureLayer = Nothing
    Dim m_intDevTypeFld As Integer
    Dim m_arrLUColumns As ArrayList = New ArrayList
    Dim m_dblTotalAcres As Double = 0
    Dim m_dblExTotalAcres As Double = 0
    Dim m_arrEX_LU As ArrayList = New ArrayList
    Dim m_arrMIX_LU As ArrayList = New ArrayList
    Dim m_intBucketCount As Integer

    Private Sub frmLandUseMix_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        m_arrPolyFeatureLayers = Nothing
        m_frmLUMix = Nothing
        m_arrDevTypes = Nothing
        m_arrLUMIXCoulmnNames = Nothing
        m_strMissingLyrs = Nothing
        m_intMissingLyrs = Nothing
        m_intDevTypeFld = Nothing
        m_arrLUColumns = Nothing
    End Sub

    Private Sub frmLandUseMix_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim mxApplication As IMxApplication = Nothing
        Dim pMxDocument As IMxDocument = Nothing
        Dim mapCurrent As Map
        Dim uid As ESRI.ArcGIS.esriSystem.IUID = New ESRI.ArcGIS.esriSystem.UIDClass
        Dim enumLayer As ESRI.ArcGIS.Carto.IEnumLayer
        Dim pLyr As ILayer
        Dim pFeatLyr As IFeatureLayer
        Dim intFeatCount As Integer = 0
        Dim pActiveView As IActiveView = Nothing
        Dim pFeatClass As IFeatureClass
        Dim intCount As Integer
        Dim intRow As Integer
        Dim strValue As String
        Dim dblValue As Double
        Dim pField As IField
        Dim rowTemp As IRow
        Dim colTemp As System.Windows.Forms.DataGridViewColumn
        Dim arrBucketList As ArrayList = New ArrayList


        'RETRIEVE CURRENT VIEW DOCUMENT TO OBTAIN LIST OF CURRENT LAYER(S)
        If Not TypeOf m_appEnvision Is IApplication Then
            GoTo CleanUp
        Else
            pMxDocument = DirectCast(m_appEnvision.Document, IMxDocument)
            mapCurrent = CType(pMxDocument.FocusMap, Map)
            pActiveView = CType(pMxDocument.FocusMap, IActiveView)
        End If

        If Not m_dockEnvisionWinForm.btnEditing.Text = "Start Edit" Then
            m_strProcessing = m_strProcessing & "An edit session is currently started and needs to be stopped before proceeding: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine & vbNewLine
            Me.Close()
            GoTo CleanUp
        End If

        'EXIT IF THE USER HAS NOT DEFINED A PARCEL LAYER
        If m_pEditFeatureLyr Is Nothing Then
            MessageBox.Show("Please select a Parcel layer to use this tool.", "Parcel Layer Not Defined", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo CleanUp
        End If

        'CYCLE THROUGH LAYERS TO POPULATE ARRAYS AND CONTROLS
        'RETRIEVE THE FEATURE LAYERS TO POPULATE FEATURE LAYER OPTIONS
        m_arrPolyFeatureLayers.Clear()
        m_arrPolyFeatureLayers = New ArrayList
        uid.Value = "{E156D7E5-22AF-11D3-9F99-00C04F6BC78E}" '= IGeoFeatureLayer
        enumLayer = mapCurrent.Layers((CType(uid, ESRI.ArcGIS.esriSystem.UID)), True) ' Explicit Cast
        enumLayer.Reset()
        pLyr = enumLayer.Next
        Do While Not (pLyr Is Nothing)
            pFeatLyr = CType(pLyr, IFeatureLayer)
            pFeatClass = pFeatLyr.FeatureClass
            If pFeatClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                If Not pFeatLyr.Name = m_pEditFeatureLyr.Name Then
                    m_arrPolyFeatureLayers.Add(pFeatLyr)
                    Me.cmbLayers.Items.Add(pFeatLyr.Name)
                    intFeatCount = intFeatCount + 1
                End If
            End If
            pLyr = enumLayer.Next()
        Loop
        'CLOSE IF NO POLYGON LAYERS FOUND
        If m_arrPolyFeatureLayers.Count <= 0 Then
            MessageBox.Show("No polygon feature layer(s) could be found in the current view document.  Please load in a neighnorhood polygon layer to utilize these functions.", "No Polygon Layers Found", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            Me.Close()
            GoTo CleanUp
        End If

        'BUILD LIST OF EXISTING LAND USE AND MIX USE CROSS REFERENCE
        m_arrEX_LU.Clear()
        m_arrMIX_LU.Clear()
        If Not m_tblEXLURef Is Nothing Then
            For intCount = 0 To 20
                Try
                    rowTemp = m_tblEXLURef.GetRow((intCount + 1))
                    strValue = rowTemp.Value(rowTemp.Fields.FindField("EX_LU"))
                    m_arrEX_LU.Add(strValue)
                    strValue = rowTemp.Value(rowTemp.Fields.FindField("EX_LU_MIX_TYPE"))
                    m_arrMIX_LU.Add(strValue)
                    'BUILD LIST OF UNIQUE LU MIX TYPES
                    If strValue.Length > 0 Then
                        If arrBucketList.IndexOf(strValue) = -1 Then
                            arrBucketList.Add(strValue)
                        End If
                    End If
                Catch ex As Exception

                End Try
            Next
        End If
        intCount = 0

        'THE NUMBER OF UNIQUE LU MIX TYPES WILL DETERMINE THE BUCKET LIST COUNT
        m_intBucketCount = arrBucketList.Count

        'RETRIEVE DEVELOPMENT TYPES AND LAND USE MIX 
        If m_tblDevelopmentTypes Is Nothing Then
            MessageBox.Show("Please load in a neighnorhood polygon layer to utilize these functions.", "No Dev Types table Found", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            Me.Close()
            GoTo CleanUp
        Else
            'RETRIEVE DEV TYPE FIELD NUMBER
            m_intDevTypeFld = m_tblDevelopmentTypes.FindField("DEVELOPMENT_TYPE")

            'BUILD LIST OF LU FIELDS
            For intCount = 0 To m_tblDevelopmentTypes.Fields.FieldCount - 1
                pField = m_tblDevelopmentTypes.Fields.Field(intCount)
                strValue = pField.Name
                If strValue.Contains("LU_MIX_") Then
                    m_arrLUColumns.Add(intCount)
                    m_arrLUMIXCoulmnNames.Add(strValue)
                    colTemp = Me.dgvLUMIX.Columns.Item(m_arrLUColumns.Count)
                    colTemp.HeaderText = strValue
                    colTemp.Visible = True
                    colTemp = Me.dgvLUMIX.Columns.Item(m_arrLUColumns.Count + 6)
                    colTemp.HeaderText = strValue & "_BUCKET"
                    colTemp.Visible = True
                    colTemp = Me.dgvLUMIX.Columns.Item(m_arrLUColumns.Count + 13)
                    colTemp.HeaderText = strValue & "_PREBUCKET"
                    colTemp.Visible = True
                End If
            Next

            If m_arrLUColumns.Count <= 0 Then
                MessageBox.Show("Please select Excel Envision file with Land Use Mix fields defined and Load Development Types to use this function.", "No Land Use Miex Fields Found", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
                Me.Close()
                GoTo CleanUp
            End If

            'RETRIEVE THE RECORD NUMBER
            Try
                For intRow = 1 To m_intDevTypeMax
                    rowTemp = m_tblDevelopmentTypes.GetRow(intRow)
                    Me.dgvLUMIX.Rows.Add()
                    strValue = CStr(rowTemp.Value(m_intDevTypeFld))
                    Me.dgvLUMIX.Rows(Me.dgvLUMIX.RowCount - 1).Cells(0).Value = strValue
                    Try
                        dblValue = CDbl(rowTemp.Value(m_tblDevelopmentTypes.FindField("REDEV_RATE")))
                    Catch ex As Exception
                        dblValue = 0
                    End Try
                    Me.dgvLUMIX.Rows(Me.dgvLUMIX.RowCount - 1).Cells(Me.dgvLUMIX.ColumnCount - 1).Value = dblValue
                    m_arrDevTypes.Add(strValue)

                    For intCount = 0 To m_arrLUColumns.Count - 1
                        dblValue = CDbl(rowTemp.Value(m_arrLUColumns.Item(intCount)))
                        Me.dgvLUMIX.Rows(Me.dgvLUMIX.RowCount - 1).Cells(intCount + 1).Value = dblValue
                    Next
                Next
            Catch ex As Exception
                MessageBox.Show(ex.Message, "issue")
            End Try


            'ADD ROW FOR EXISTING DEVELOPMENT
            Me.dgvLUMIX.Rows.Add()
            Me.dgvLUMIX.Rows(Me.dgvLUMIX.RowCount - 1).Cells(0).Value = "Developed Unpainted"
            For intCount = 1 To 13
                Me.dgvLUMIX.Rows(Me.dgvLUMIX.RowCount - 1).Cells(intCount).Value = 0
            Next
        End If

CleanUp:
        mxApplication = Nothing
        pMxDocument = Nothing
        mapCurrent = Nothing
        uid = Nothing
        enumLayer = Nothing
        pLyr = Nothing
        pFeatLyr = Nothing
        intFeatCount = Nothing
        pActiveView = Nothing
        pFeatClass = Nothing
        intCount = Nothing
        intRow = Nothing
        strValue = Nothing
        dblValue = Nothing
        pField = Nothing
        rowTemp = Nothing
        colTemp = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub


    Private Sub cmbLayers_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbLayers.SelectedIndexChanged
        Dim pFeatLyr As IFeatureLayer
        Dim pFeatureClass As IFeatureClass = Nothing
        Dim pFeatSelection As IFeatureSelection
        Dim pCursor As ICursor = Nothing
        Dim intFld As Integer = 0
        Dim pField As IField
        Dim pTable As ITable

        'RETRIEVE SELECTED NEIGHBORHOOD LAYER
        If Me.cmbLayers.Text.Length > 0 Then
            m_strNeighborhoodLayerName = Me.cmbLayers.Text
            pFeatLyr = m_arrPolyFeatureLayers.Item(Me.cmbLayers.FindString(Me.cmbLayers.Text))
            pFeatureClass = pFeatLyr.FeatureClass
            pFeatSelection = CType(pFeatLyr, IFeatureSelection)
            pFeatSelection.SelectionSet.Search(Nothing, False, pCursor)
            'Default to ALL features
            If pFeatSelection.SelectionSet.Count > 0 Then
                Me.chkUseSelected.Visible = True
                Me.chkUseSelected.Enabled = True
                Me.chkUseSelected.Text = "Use " & pFeatSelection.SelectionSet.Count.ToString & " Selected Features"
            Else
                Me.chkUseSelected.Visible = False
            End If
        Else
            GoTo CleanUp
        End If

        For intFld = 0 To pFeatureClass.Fields.FieldCount - 1
            pField = pFeatureClass.Fields.Field(intFld)
            'DON'T ADD OUTPUT FIELD NAMES IN LIST OF ID OR ACRES FIELD CONTROLS
            If (pField.Type = esriFieldType.esriFieldTypeInteger Or pField.Type = esriFieldType.esriFieldTypeSmallInteger) And Not pField.Type = esriFieldType.esriFieldTypeOID Then
                Me.cmbFieldId.Items.Add(pField.Name)
            End If
        Next

        If Me.cmbFieldId.Items.Count > 0 Then
            Me.cmbFieldId.Enabled = True
            Me.lblId.Enabled = True
        Else
            Me.cmbFieldId.Enabled = False
            Me.lblId.Enabled = False
        End If

        'ADD OUTPUT FIELDS IF MISSING
        pTable = CType(pFeatureClass, ITable)
        If pFeatureClass.FindField("TOT_LU_MIX_0Mi") = -1 Then
            AddEnvisionField(pTable, "TOT_LU_MIX_0Mi", "DOUBLE", 16, 6)
        End If
        If pFeatureClass.FindField("EX_PRE_LU_MIX_0Mi") = -1 Then
            AddEnvisionField(pTable, "EX_PRE_LU_MIX_0Mi", "DOUBLE", 16, 6)
        End If
        If pFeatureClass.FindField("TOT_LU_MIX_QTRMi") = -1 Then
            AddEnvisionField(pTable, "TOT_LU_MIX_QTRMi", "DOUBLE", 16, 6)
        End If
        If pFeatureClass.FindField("EX_PRE_LU_MIX_QTRMi") = -1 Then
            AddEnvisionField(pTable, "EX_PRE_LU_MIX_QTRMi", "DOUBLE", 16, 6)
        End If
        If pFeatureClass.FindField("TOT_LU_MIX_HALFMi") = -1 Then
            AddEnvisionField(pTable, "TOT_LU_MIX_HALFMi", "DOUBLE", 16, 6)
        End If
        If pFeatureClass.FindField("EX_PRE_LU_MIX_HALFMi") = -1 Then
            AddEnvisionField(pTable, "EX_PRE_LU_MIX_HALFMi", "DOUBLE", 16, 6)
        End If
        If pFeatureClass.FindField("TOT_LU_MIX_1Mi") = -1 Then
            AddEnvisionField(pTable, "TOT_LU_MIX_1Mi", "DOUBLE", 16, 6)
        End If
        If pFeatureClass.FindField("EX_PRE_LU_MIX_1Mi") = -1 Then
            AddEnvisionField(pTable, "EX_PRE_LU_MIX_1Mi", "DOUBLE", 16, 6)
        End If
        GoTo CleanUp

CleanUp:
        pFeatLyr = Nothing
        pFeatureClass = Nothing
        pFeatSelection = Nothing
        pCursor = Nothing
        intFld = Nothing
        pField = Nothing
        pTable = Nothing
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
                        CheckForBufferLayers(strBufferFileGDB)
                        If m_intMissingLyrs > 0 Then
                            If m_intMissingLyrs = m_strBufferLayerNames.Split(",").Length Then
                                MessageBox.Show("The selected geodatabase does not contain any of the required buffer layers.", "Buffer Layers Missing", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
                                GoTo CleanUp
                            End If
                            If MessageBox.Show(m_intMissingLyrs.ToString & " buffer layers were missing.  Would you like to continue using this geodatabase? Missing Layer(s):" & vbNewLine & m_strMissingLyrs, "Buffer Layers Missing", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2, MessageBoxOptions.DefaultDesktopOnly, False) = False Then
                                GoTo CleanUp
                            End If
                        End If

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
    End Sub

    Private Sub CheckForBufferLayers(ByVal strBufferFileGDB As String)
        'CHECK TO ENSURE ALL BUFFER LAYERS ARE PRESENT
        Dim pWksFactory As ESRI.ArcGIS.Geodatabase.IWorkspaceFactory = Nothing
        Dim pFeatWks As ESRI.ArcGIS.Geodatabase.IFeatureWorkspace
        Dim pFeatClass As IFeatureClass
        Dim strLyrName As String
        Dim strBufferLayerNames As String = "SA_QUARTER_MILE,SA_HALF_MILE,SA_ONE_MILE"
        m_strMissingLyrs = ""
        m_intMissingLyrs = 0

        Try
            pWksFactory = New FileGDBWorkspaceFactory
            pFeatWks = pWksFactory.OpenFromFile(strBufferFileGDB, 0)
            For Each strLyrName In strBufferLayerNames.Split(",")
                Try
                    pFeatClass = pFeatWks.OpenFeatureClass(strLyrName)
                Catch ex As Exception
                    m_strMissingLyrs = m_strMissingLyrs & strLyrName & vbNewLine
                    m_intMissingLyrs = m_intMissingLyrs + 1
                End Try
            Next
        Catch ex As Exception

        End Try
CLeanUp:
        pWksFactory = Nothing
        pFeatWks = Nothing
        pFeatClass = Nothing
        strLyrName = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub

    Private Sub OpenBufferLayers()
        'RETRIEVE EACH OF THE 3 BUFFER FEATURE LAYERS FROM THE SELECTED BUFFER WORKSPACE AND HOLD IN MEMORY
        Dim pWksFactory As IWorkspaceFactory = Nothing
        Dim pFeatWks As IFeatureWorkspace
        Dim pFeatClass As IFeatureClass

        Try
            pWksFactory = New FileGDBWorkspaceFactory
            pFeatWks = pWksFactory.OpenFromFile(Me.tbxWorkspace.Text, 0)
        Catch ex As Exception
            GoTo CleanUp
        End Try
        'OPEN EACH OF THE BUFFER LAYERS
        Try
            pFeatClass = pFeatWks.OpenFeatureClass("SA_QUARTER_MILE")
            If pFeatClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                pSA_QUARTER_MILELyr = New FeatureLayer
                pSA_QUARTER_MILELyr.FeatureClass = pFeatClass
            Else
                pSA_QUARTER_MILELyr = Nothing
            End If
        Catch ex As Exception
            GoTo CleanUp
        End Try
        Try
            pFeatClass = pFeatWks.OpenFeatureClass("SA_HALF_MILE")
            If pFeatClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                pSA_HALF_MILELyr = New FeatureLayer
                pSA_HALF_MILELyr.FeatureClass = pFeatClass
            Else
                pSA_HALF_MILELyr = Nothing
            End If
        Catch ex As Exception
            GoTo CleanUp
        End Try
        Try
            pFeatClass = pFeatWks.OpenFeatureClass("SA_ONE_MILE")
            If pFeatClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                pSA_ONE_MILELyr = New FeatureLayer
                pSA_ONE_MILELyr.FeatureClass = pFeatClass
            Else
                pSA_ONE_MILELyr = Nothing
            End If
        Catch ex As Exception
            GoTo CleanUp
        End Try

CleanUp:
        pWksFactory = Nothing
        pFeatWks = Nothing
        pFeatClass = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub

    Private Sub btnRun_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRun.Click
        'EXECUTE THE SUBS TO SUMMARIZE THE NEEDED VARIABLES
        Dim pFeatSelection As IFeatureSelection
        Dim pCursor As ICursor = Nothing
        Dim pDef As IFeatureLayerDefinition2
        Dim pNFeatLyr As IFeatureLayer
        Dim pNFeatClass As IFeatureClass
        Dim intFld As Integer = 0
        Dim strFld As String
        Dim arrReqFlds As ArrayList = New ArrayList
        Dim pQFilter As IQueryFilter
        Dim strDefExpression As String = ""
        Dim strQString As String = Nothing
        Dim pFeatureCursor As IFeatureCursor = Nothing
        Dim intTotalCount As Integer
        Dim intObjFld As Integer
        Dim pFeat As IFeature
        Dim intObjId As Integer
        Dim pBuffFeatSelection As IFeatureSelection
        Dim pDataStatistics As DataStatistics
        Dim dblTotalAcres As Double


        'RETREIVE THE SELECTED NEIGHBORHOOD LAYER
        m_appEnvision.StatusBar.Message(0) = "Retrieving selected Neighborhood layer"
        Try
            pNFeatLyr = m_arrPolyFeatureLayers.Item(Me.cmbLayers.SelectedIndex)
            pNFeatClass = pNFeatLyr.FeatureClass
            pDef = pNFeatLyr
            pFeatSelection = CType(pNFeatLyr, IFeatureSelection)
        Catch ex As Exception
            MessageBox.Show("Error in retrieving the selected neighborhood layer." & vbNewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo CleanUp
        End Try

        'CHECK THAT THE ID FIELD IS SELECTED
        If Not Me.cmbFieldId.Items.Contains(Me.cmbFieldId.Text) Then
            MessageBox.Show("Please selected an uniquie ID field.", "ID Field Required", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo CleanUp
        End If

        'RETRIEVE THE BUFFER LAYERS
        OpenBufferLayers()

        'DEFINE WHAT WILL BE PROCESSED IN THE SELECTED NEIGHBORHOOD LAYER (ALL OR SELECTED FEATURES)
        strDefExpression = pDef.DefinitionExpression
        pQFilter = New QueryFilter
        If Not Me.chkUseSelected.Checked Then
            If strDefExpression.Length > 0 Then
                pQFilter.WhereClause = pDef.DefinitionExpression
                pFeatureCursor = pNFeatClass.Search(pQFilter, False)
                intTotalCount = pNFeatClass.FeatureCount(pQFilter)
            Else
                pFeatureCursor = pNFeatClass.Search(Nothing, False)
                intTotalCount = pNFeatClass.FeatureCount(Nothing)
            End If
        Else
            pFeatSelection.SelectionSet.Search(Nothing, False, pCursor)
            pFeatureCursor = DirectCast(pCursor, IFeatureCursor)
            intTotalCount = pFeatSelection.SelectionSet.Count
        End If

        intObjFld = pNFeatClass.FindField(Me.cmbFieldId.Text)
        pFeat = pFeatureCursor.NextFeature
        Do While Not pFeat Is Nothing
            intObjId = pFeat.Value(intObjFld)

            Try
                pQFilter = New QueryFilter
                strQString = Me.cmbFieldId.Text & " = " & intObjId.ToString
                pQFilter.WhereClause = strQString

                'Feature(ONLY)
                m_dblTotalAcres = 0
                pBuffFeatSelection = CType(pNFeatLyr, IFeatureSelection)
                pBuffFeatSelection.Clear()
                pBuffFeatSelection.SelectFeatures(pQFilter, esriSelectionResultEnum.esriSelectionResultNew, True)
                If pBuffFeatSelection.SelectionSet.Count > 0 Then
                    m_appEnvision.StatusBar.Message(0) = "Feature " & intObjId.ToString & " Selection"
                    MakeFeatSelection(pNFeatLyr, m_pEditFeatureLyr)
                    CalcLandUseMix(dblTotalAcres, "TOT_LU_MIX_0Mi", pFeat)
                    CalcPRELandUseMix(m_dblExTotalAcres, "EX_PRE_LU_MIX_0Mi", pFeat)
                End If

                'QUARTER MILE BUFFER
                m_dblTotalAcres = 0
                pBuffFeatSelection = CType(pSA_QUARTER_MILELyr, IFeatureSelection)
                pBuffFeatSelection.Clear()
                pBuffFeatSelection.SelectFeatures(pQFilter, esriSelectionResultEnum.esriSelectionResultNew, True)
                If pBuffFeatSelection.SelectionSet.Count > 0 Then
                    m_appEnvision.StatusBar.Message(0) = "Quarter Mile Buffer Selection"
                    MakeFeatSelection(pSA_QUARTER_MILELyr, m_pEditFeatureLyr)
                    CalcLandUseMix(dblTotalAcres, "TOT_LU_MIX_QTRMi", pFeat)
                    CalcPRELandUseMix(m_dblExTotalAcres, "EX_PRE_LU_MIX_QTRMi", pFeat)
                End If

                'HALF MILE BUFFER
                m_dblTotalAcres = 0
                pBuffFeatSelection = CType(pSA_HALF_MILELyr, IFeatureSelection)
                pBuffFeatSelection.Clear()
                pBuffFeatSelection.SelectFeatures(pQFilter, esriSelectionResultEnum.esriSelectionResultNew, True)
                If pBuffFeatSelection.SelectionSet.Count > 0 Then
                    m_appEnvision.StatusBar.Message(0) = "Half Mile Buffer Selection"
                    MakeFeatSelection(pSA_HALF_MILELyr, m_pEditFeatureLyr)
                    CalcLandUseMix(dblTotalAcres, "TOT_LU_MIX_HALFMi", pFeat)
                    CalcPRELandUseMix(m_dblExTotalAcres, "EX_PRE_LU_MIX_HALFMi", pFeat)
                End If

                'ONE MILE BUFFER
                m_dblTotalAcres = 0
                pBuffFeatSelection = CType(pSA_ONE_MILELyr, IFeatureSelection)
                pBuffFeatSelection.Clear()
                pBuffFeatSelection.SelectFeatures(pQFilter, esriSelectionResultEnum.esriSelectionResultNew, True)
                If pBuffFeatSelection.SelectionSet.Count > 0 Then
                    m_appEnvision.StatusBar.Message(0) = "One Mile Buffer Selection"
                    MakeFeatSelection(pSA_ONE_MILELyr, m_pEditFeatureLyr)
                    CalcLandUseMix(dblTotalAcres, "TOT_LU_MIX_1Mi", pFeat)
                    CalcPRELandUseMix(m_dblExTotalAcres, "EX_PRE_LU_MIX_1Mi", pFeat)
                End If
            Catch ex As Exception
                m_appEnvision.StatusBar.Message(0) = ex.Message
            End Try
            pFeat = pFeatureCursor.NextFeature
            m_appEnvision.StatusBar.Message(0) = ""
        Loop

CleanUp:
        MessageBox.Show("Land Use Mix Summary processing has completed.", "End Processing")
        ' MessageBox.Show("Land Use Mix Summary processing has completed." & vbNewLine & m_dblTotalAcres.ToString & vbNewLine & m_dblExTotalAcres.ToString, "End Processing")
        'Me.Close()
        pFeatSelection = Nothing
        pCursor = Nothing
        pDef = Nothing
        pNFeatLyr = Nothing
        pNFeatClass = Nothing
        intFld = Nothing
        strFld = Nothing
        arrReqFlds = Nothing
        pQFilter = Nothing
        strDefExpression = Nothing
        strQString = Nothing
        pFeatureCursor = Nothing
        intTotalCount = Nothing
        intObjFld = Nothing
        pFeat = Nothing
        intObjId = Nothing
        pBuffFeatSelection = Nothing
        pDataStatistics = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub

    Private Function MakeFeatSelection(ByVal pByFeatLayer As IFeatureLayer2, ByVal pSelLayer As IFeatureLayer) As Boolean

        Dim pQBLayer As ESRI.ArcGIS.Carto.IQueryByLayer
        Dim pFSelOther As IFeatureSelection = Nothing
        Dim intFeatCount As Integer
        Dim pCursor As ICursor = Nothing
        Dim pFeatureCursor As IFeatureCursor
        Dim strDevType As String
        Dim intDevTypeRow As Integer
        Dim intCol As Integer
        Dim dblLUPercent As Double
        Dim dblVacAcres As Double
        Dim dblDevdAcres As Double
        Dim dblExDevdAcres As Double = 0
        Dim dblValue As Double
        Dim intEX_LUfld As Integer
        Dim strEX_LU As String
        Dim dblRedevRate As Double
        Dim pFeat As IFeature
        Dim intCount As Integer
        Dim intLUMIX As Integer

        'CLEAR THE DATA FORM BUCKETS
        EmptyBuckets()
        m_dblTotalAcres = 0
        m_dblExTotalAcres = 0


        'SELECT THE CORRESPONDING BUFFER FOR THE SELECTED NEIGHBORHOOD FEATURE
        Try
            pFSelOther = CType(pSelLayer, IFeatureSelection)
            intFeatCount = pFSelOther.SelectionSet.Count
            pQBLayer = New QueryByLayer
            With pQBLayer
                .ByLayer = pByFeatLayer
                .FromLayer = pSelLayer
                .ResultType = esriSelectionResultEnum.esriSelectionResultNew
                .BufferDistance = 0
                .BufferUnits = ESRI.ArcGIS.esriSystem.esriUnits.esriMiles
                .LayerSelectionMethod = esriLayerSelectionMethod.esriLayerSelectHaveTheirCenterIn
                .UseSelectedFeatures = True
                pFSelOther.SelectionSet = .Select
            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error")
            GoTo CleanUp
        End Try

        Try
            'WRITE THE COUNT TO THE INPUT FEATURE
            intFeatCount = pFSelOther.SelectionSet.Count
            intEX_LUfld = pSelLayer.FeatureClass.FindField("EX_LU")
            If intFeatCount > 0 Then
                intcount = 0
                pFSelOther.SelectionSet.Search(Nothing, False, pCursor)
                pFeatureCursor = DirectCast(pCursor, IFeatureCursor)
                pFeat = pFeatureCursor.NextFeature
                Do While Not pFeat Is Nothing
                    Try
                        strDevType = pFeat.Value(pFeat.Fields.FindField("DEV_TYPE"))
                    Catch ex As Exception
                        strDevType = ""
                    End Try
                    Try
                        dblVacAcres = pFeat.Value(pFeat.Fields.FindField("VAC_ACRE"))
                    Catch ex As Exception
                        dblVacAcres = 0
                    End Try
                    Try
                        dblDevdAcres = pFeat.Value(pFeat.Fields.FindField("DEVD_ACRE"))
                    Catch ex As Exception
                        dblDevdAcres = 0
                    End Try
                    'RETRIEVE EX LU FROM FEATURE
                    Try
                        strEX_LU = pFeat.Value(intEX_LUfld)
                    Catch ex As Exception
                        strEX_LU = ""
                    End Try
                    'FIND THE MIX LU MATCH
                    intLUMIX = m_arrEX_LU.IndexOf(strEX_LU)
                    If intLUMIX >= 0 Then
                        Try
                            strEX_LU = m_arrMIX_LU.Item(intLUMIX)
                        Catch ex As Exception
                            strEX_LU = ""
                        End Try
                    Else
                        strEX_LU = ""
                    End If

                    intDevTypeRow = m_arrDevTypes.IndexOf(strDevType)
                    If (intDevTypeRow >= 0 And m_arrLUColumns.Count >= 1) And Not (strDevType = "") Then
                        For intCol = 1 To 6
                            Try
                                dblLUPercent = CDbl(Me.dgvLUMIX.Rows(intDevTypeRow).Cells(intCol).Value)
                            Catch ex As Exception
                                dblLUPercent = 0
                            End Try
                            Try
                                dblValue = CDbl(Me.dgvLUMIX.Rows(intDevTypeRow).Cells(intCol + 6).Value)
                            Catch ex As Exception
                                dblValue = 0
                            End Try
                            Try
                                dblRedevRate = CDbl(Me.dgvLUMIX.Rows(intDevTypeRow).Cells(Me.dgvLUMIX.ColumnCount - 1).Value)
                            Catch ex As Exception
                                dblRedevRate = 0
                            End Try
                            If strDevType = "" Then
                                dblRedevRate = 1
                            End If
                            'APPLY THE ACRES TO THE DEVELOPMENT TYPE
                            Try
                                Me.dgvLUMIX.Rows(intDevTypeRow).Cells(intCol + 6).Value = (dblValue + (dblLUPercent * (dblDevdAcres * dblRedevRate)) + (dblLUPercent * dblVacAcres))
                                m_dblTotalAcres = m_dblTotalAcres + (dblLUPercent * (dblDevdAcres * dblRedevRate)) + (dblLUPercent * dblVacAcres)
                            Catch ex As Exception
                            End Try
                            'ADD DEVELOPMENT ACRES TO EXISITNG LU
                            If dblDevdAcres > 0 Then
                                If m_arrLUMIXCoulmnNames.IndexOf("LU_MIX_" & strEX_LU) >= 0 Then
                                    Try
                                        dblValue = CDbl(Me.dgvLUMIX.Rows(Me.dgvLUMIX.RowCount - 1).Cells(m_arrLUMIXCoulmnNames.IndexOf("LU_MIX_" & strEX_LU) + 7).Value)
                                    Catch ex As Exception
                                        dblValue = 0
                                    End Try

                                    Me.dgvLUMIX.Rows(Me.dgvLUMIX.RowCount - 1).Cells(m_arrLUMIXCoulmnNames.IndexOf("LU_MIX_" & strEX_LU) + 7).Value = (dblValue + (dblLUPercent * (dblDevdAcres * (1 - dblRedevRate))))
                                    m_dblTotalAcres = m_dblTotalAcres + (dblLUPercent * (dblDevdAcres * (1 - dblRedevRate)))
                                End If
                            End If
                        Next
                    Else
                        If m_arrLUMIXCoulmnNames.IndexOf("LU_MIX_" & strEX_LU) >= 0 Then
                            Try
                                dblValue = CDbl(Me.dgvLUMIX.Rows(Me.dgvLUMIX.RowCount - 1).Cells(m_arrLUMIXCoulmnNames.IndexOf("LU_MIX_" & strEX_LU) + 7).Value)
                            Catch ex As Exception
                                dblValue = 0
                            End Try

                            If strDevType = "" Then
                                m_dblTotalAcres = m_dblTotalAcres + dblDevdAcres
                                Me.dgvLUMIX.Rows(Me.dgvLUMIX.RowCount - 1).Cells(m_arrLUMIXCoulmnNames.IndexOf("LU_MIX_" & strEX_LU) + 7).Value = dblValue + dblDevdAcres
                            Else
                                m_dblTotalAcres = m_dblTotalAcres + (dblDevdAcres + dblVacAcres)
                                Me.dgvLUMIX.Rows(Me.dgvLUMIX.RowCount - 1).Cells(m_arrLUMIXCoulmnNames.IndexOf("LU_MIX_" & strEX_LU) + 7).Value = dblValue + dblDevdAcres + dblVacAcres
                            End If
                        End If
                    End If

                    'POPULATE FOR JUST THE EXISTING LAND USE MIX
                    If m_arrLUMIXCoulmnNames.IndexOf("LU_MIX_" & strEX_LU) >= 0 Then
                        Try
                            dblValue = CDbl(Me.dgvLUMIX.Rows(Me.dgvLUMIX.RowCount - 1).Cells(m_arrLUMIXCoulmnNames.IndexOf("LU_MIX_" & strEX_LU) + 14).Value)
                        Catch ex As Exception
                            dblValue = 0
                        End Try

                        Me.dgvLUMIX.Rows(Me.dgvLUMIX.RowCount - 1).Cells(m_arrLUMIXCoulmnNames.IndexOf("LU_MIX_" & strEX_LU) + 14).Value = dblValue + dblDevdAcres
                        m_dblExTotalAcres = m_dblExTotalAcres + dblDevdAcres
                    End If
                    intCount = intCount + 1
                    m_frmLUMix.barStatus.Value = (intCount / intFeatCount) * 100

                    pFeat = pFeatureCursor.NextFeature
                Loop
            End If

        Catch ex As Exception
            GoTo CleanUp
        End Try

CleanUp:
        pByFeatLayer = Nothing
        pQBLayer = Nothing
        pFSelOther = Nothing
        pFSelOther = Nothing
        pSelLayer = Nothing
        pFeat = Nothing
        intFeatCount = Nothing
        pCursor = Nothing
        intcount = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()

    End Function

    Private Sub CalcLandUseMix(ByVal dblTotalAcres As Double, ByVal strWriteFldName As String, ByVal pFeat As IFeature)
        'CALCULATES THE LAND USE MIX BASED UPON INPUT FROM THE FORM DATA GRID VIEW BUCKETS
        Dim dblTotalLU_MIX As Double = 0
        Dim dblBucket1 As Double = 0
        Dim dblBucket2 As Double = 0
        Dim dblBucket3 As Double = 0
        Dim dblBucket4 As Double = 0
        Dim dblBucket5 As Double = 0
        Dim dblBucket6 As Double = 0
        Dim intRow As Integer
        Dim dblValue As Double
        Dim intBucketCount As Integer
        Dim dblPart1 As Double = 0
        Dim dblPart2 As Double = 0
        Dim dblPart3 As Double = 0
        Dim dblPart4 As Double = 0
        Dim dblPart5 As Double = 0
        Dim dblPart6 As Double = 0
        For intRow = 0 To m_intDevTypeMax
            'BUCKET 1
            Try
                dblValue = CDbl(Me.dgvLUMIX.Rows(intRow).Cells(7).Value)
            Catch ex As Exception
                dblValue = 0
            End Try
            dblBucket1 = dblBucket1 + dblValue
            'BUCKET 2
            Try
                dblValue = CDbl(Me.dgvLUMIX.Rows(intRow).Cells(8).Value)
            Catch ex As Exception
                dblValue = 0
            End Try
            dblBucket2 = dblBucket2 + dblValue
            'BUCKET 3
            Try
                dblValue = CDbl(Me.dgvLUMIX.Rows(intRow).Cells(9).Value)
            Catch ex As Exception
                dblValue = 0
            End Try
            dblBucket3 = dblBucket3 + dblValue
            'BUCKET 4
            Try
                dblValue = CDbl(Me.dgvLUMIX.Rows(intRow).Cells(10).Value)
            Catch ex As Exception
                dblValue = 0
            End Try
            dblBucket4 = dblBucket4 + dblValue
            'BUCKET 5
            Try
                dblValue = CDbl(Me.dgvLUMIX.Rows(intRow).Cells(11).Value)
            Catch ex As Exception
                dblValue = 0
            End Try
            dblBucket5 = dblBucket5 + dblValue
            'BUCKET 6
            Try
                dblValue = CDbl(Me.dgvLUMIX.Rows(intRow).Cells(12).Value)
            Catch ex As Exception
                dblValue = 0
            End Try
            dblBucket6 = dblBucket6 + dblValue
        Next

        If m_intBucketCount > 1 Then
            Try
                dblPart1 = (dblBucket1 / m_dblTotalAcres)
                dblPart2 = (dblBucket2 / m_dblTotalAcres)
                dblPart3 = (dblBucket3 / m_dblTotalAcres)
                dblPart4 = (dblBucket4 / m_dblTotalAcres)
                dblPart5 = (dblBucket5 / m_dblTotalAcres)
                dblPart6 = (dblBucket6 / m_dblTotalAcres)

                If dblBucket1 > 0 Then
                    dblPart1 = ((dblBucket1 / m_dblTotalAcres) * Log(dblBucket1 / m_dblTotalAcres))
                Else
                    dblPart1 = 0
                End If
                If dblBucket2 > 0 Then
                    dblPart2 = ((dblBucket2 / m_dblTotalAcres) * Log(dblBucket2 / m_dblTotalAcres))
                Else
                    dblPart2 = 0
                End If
                If dblBucket3 > 0 Then
                    dblPart3 = ((dblBucket3 / m_dblTotalAcres) * Log(dblBucket3 / m_dblTotalAcres))
                Else
                    dblPart3 = 0
                End If
                If dblBucket4 > 0 Then
                    dblPart4 = ((dblBucket4 / m_dblTotalAcres) * Log(dblBucket4 / m_dblTotalAcres))
                Else
                    dblPart4 = 0
                End If
                If dblBucket5 > 0 Then
                    dblPart5 = ((dblBucket5 / m_dblTotalAcres) * Log(dblBucket5 / m_dblTotalAcres))
                Else
                    dblPart5 = 0
                End If
                If dblBucket6 > 0 Then
                    dblPart6 = ((dblBucket6 / m_dblTotalAcres) * Log(dblBucket6 / m_dblTotalAcres))
                Else
                    dblPart6 = 0
                End If
                dblTotalLU_MIX = (dblPart1 + dblPart2 + dblPart3 + dblPart4 + dblPart5 + dblPart6) / Log(m_intBucketCount)
                If dblTotalLU_MIX < 0 Then
                    dblTotalLU_MIX = dblTotalLU_MIX * -1
                End If
            Catch ex As Exception
                dblTotalLU_MIX = 0
            End Try
        Else
            dblTotalLU_MIX = 0
        End If

        Try
            If pFeat.Fields.FindField(strWriteFldName) >= 0 Then
                pFeat.Value(pFeat.Fields.FindField(strWriteFldName)) = dblTotalLU_MIX
                pFeat.Store()
            End If
        Catch ex As Exception
            GoTo CleanUp
        End Try
        GoTo CleanUp

CleanUp:
        dblTotalLU_MIX = Nothing
        dblBucket1 = Nothing
        dblBucket2 = Nothing
        dblBucket3 = Nothing
        dblBucket4 = Nothing
        dblBucket5 = Nothing
        dblBucket6 = Nothing
        intRow = Nothing
        dblValue = Nothing
        intBucketCount = Nothing
        dblPart1 = Nothing
        dblPart2 = Nothing
        dblPart3 = Nothing
        dblPart4 = Nothing
        dblPart5 = Nothing
        dblPart6 = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()

    End Sub

    Private Sub CalcPRELandUseMix(ByVal dblTotalAcres As Double, ByVal strWriteFldName As String, ByVal pFeat As IFeature)
        'CALCULATES THE PRE DEVELOPMENT LAND USE MIX BASED UPON INPUT FROM THE FORM DATA GRID VIEW BUCKETS
        Dim dblTotalLU_MIX As Double = 0
        Dim dblBucket1 As Double = 0
        Dim dblBucket2 As Double = 0
        Dim dblBucket3 As Double = 0
        Dim dblBucket4 As Double = 0
        Dim dblBucket5 As Double = 0
        Dim dblBucket6 As Double = 0
        Dim intRow As Integer
        Dim dblValue As Double
        Dim dblPart1 As Double = 0
        Dim dblPart2 As Double = 0
        Dim dblPart3 As Double = 0
        Dim dblPart4 As Double = 0
        Dim dblPart5 As Double = 0
        Dim dblPart6 As Double = 0
        For intRow = 0 To m_intDevTypeMax
            'BUCKET 1
            Try
                dblValue = CDbl(Me.dgvLUMIX.Rows(intRow).Cells(13).Value)
            Catch ex As Exception
                dblValue = 0
            End Try
            dblBucket1 = dblBucket1 + dblValue
            'BUCKET 2
            Try
                dblValue = CDbl(Me.dgvLUMIX.Rows(intRow).Cells(14).Value)
            Catch ex As Exception
                dblValue = 0
            End Try
            dblBucket2 = dblBucket2 + dblValue
            'BUCKET 3
            Try
                dblValue = CDbl(Me.dgvLUMIX.Rows(intRow).Cells(15).Value)
            Catch ex As Exception
                dblValue = 0
            End Try
            dblBucket3 = dblBucket3 + dblValue
            'BUCKET 4
            Try
                dblValue = CDbl(Me.dgvLUMIX.Rows(intRow).Cells(16).Value)
            Catch ex As Exception
                dblValue = 0
            End Try
            dblBucket4 = dblBucket4 + dblValue
            'BUCKET 5
            Try
                dblValue = CDbl(Me.dgvLUMIX.Rows(intRow).Cells(17).Value)
            Catch ex As Exception
                dblValue = 0
            End Try
            dblBucket5 = dblBucket5 + dblValue
            'BUCKET 6
            Try
                dblValue = CDbl(Me.dgvLUMIX.Rows(intRow).Cells(18).Value)
            Catch ex As Exception
                dblValue = 0
            End Try
            dblBucket6 = dblBucket6 + dblValue
        Next

        If m_intBucketCount > 1 Then
            Try
                dblPart1 = (dblBucket1 / m_dblTotalAcres)
                dblPart2 = (dblBucket2 / m_dblTotalAcres)
                dblPart3 = (dblBucket3 / m_dblTotalAcres)
                dblPart4 = (dblBucket4 / m_dblTotalAcres)
                dblPart5 = (dblBucket5 / m_dblTotalAcres)
                dblPart6 = (dblBucket6 / m_dblTotalAcres)

                If dblBucket1 > 0 Then
                    dblPart1 = ((dblBucket1 / m_dblExTotalAcres) * Log(dblBucket1 / m_dblExTotalAcres))
                Else
                    dblPart1 = 0
                End If
                If dblBucket2 > 0 Then
                    dblPart2 = ((dblBucket2 / m_dblExTotalAcres) * Log(dblBucket2 / m_dblExTotalAcres))
                Else
                    dblPart2 = 0
                End If
                If dblBucket3 > 0 Then
                    dblPart3 = ((dblBucket3 / m_dblExTotalAcres) * Log(dblBucket3 / m_dblExTotalAcres))
                Else
                    dblPart3 = 0
                End If
                If dblBucket4 > 0 Then
                    dblPart4 = ((dblBucket4 / m_dblExTotalAcres) * Log(dblBucket4 / m_dblExTotalAcres))
                Else
                    dblPart4 = 0
                End If
                If dblBucket5 > 0 Then
                    dblPart5 = ((dblBucket5 / m_dblExTotalAcres) * Log(dblBucket5 / m_dblExTotalAcres))
                Else
                    dblPart5 = 0
                End If
                If dblBucket6 > 0 Then
                    dblPart6 = ((dblBucket6 / m_dblExTotalAcres) * Log(dblBucket6 / m_dblExTotalAcres))
                Else
                    dblPart6 = 0
                End If
                dblTotalLU_MIX = (dblPart1 + dblPart2 + dblPart3 + dblPart4 + dblPart5 + dblPart6) / Log(m_intBucketCount)
                If dblTotalLU_MIX < 0 Then
                    dblTotalLU_MIX = dblTotalLU_MIX * -1
                End If
            Catch ex As Exception
                dblTotalLU_MIX = 0
            End Try
        Else
            dblTotalLU_MIX = 0
        End If

        Try
            If pFeat.Fields.FindField(strWriteFldName) >= 0 Then
                pFeat.Value(pFeat.Fields.FindField(strWriteFldName)) = dblTotalLU_MIX
                pFeat.Store()
            End If
        Catch ex As Exception
            GoTo CleanUp
        End Try
        GoTo CleanUp

CleanUp:
        dblTotalLU_MIX = Nothing
        dblBucket1 = Nothing
        dblBucket2 = Nothing
        dblBucket3 = Nothing
        dblBucket4 = Nothing
        dblBucket5 = Nothing
        dblBucket6 = Nothing
        intRow = Nothing
        dblValue = Nothing
        dblPart1 = Nothing
        dblPart2 = Nothing
        dblPart3 = Nothing
        dblPart4 = Nothing
        dblPart5 = Nothing
        dblPart6 = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()

    End Sub

    Private Sub EmptyBuckets()
        'CLEAR OUT THE BUCKETS TO STORE ARCE VALUES
        Dim intRow As Integer
        Dim intCount As Integer
        'RETRIEVE THE RECORD NUMBER
        For intRow = 0 To m_intDevTypeMax
            For intCount = 6 To 18
                Me.dgvLUMIX.Rows(intRow).Cells(intCount).Value = 0
            Next
        Next
        intRow = Nothing
        intCount = Nothing
    End Sub
End Class
