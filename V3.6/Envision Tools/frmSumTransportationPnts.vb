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

'Imports ESRI.ArcGIS.ADF
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.CartoUI
Imports ESRI.ArcGIS.Catalog
Imports ESRI.ArcGIS.CatalogUI
Imports ESRI.ArcGIS.DataSourcesGDB
'Imports ESRI.ArcGIS.DataSourcesFile
Imports ESRI.ArcGIS.Display
'Imports ESRI.ArcGIS.Editor
'Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.Geodatabase
'Imports ESRI.ArcGIS.GeoDatabaseUI
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.Geoprocessor
Imports ESRI.ArcGIS.SystemUI
'Imports ESRI.ArcGIS.AnalysisTools

Public Class frmSumTransportationPnts

    Dim arrPolyLyrs As ArrayList = New ArrayList
    Dim arrPntLyrs As ArrayList = New ArrayList
    Dim strBufferFileGDB As String = ""
    Dim m_strMissingLyrs As String = ""
    Dim m_intMissingLyrs As Integer = 0
    Dim pNFeatLyr As IFeatureLayer
    Dim pInersectionsLyr As IFeatureLayer = Nothing
    Dim p4WayLyr As IFeatureLayer = Nothing
    Dim pTransitStopLyr As IFeatureLayer = Nothing
    Dim pRailStopLyr As IFeatureLayer = Nothing
    Dim pSA_QUARTER_MILELyr As IFeatureLayer = Nothing
    Dim pSA_HALF_MILELyr As IFeatureLayer = Nothing
    Dim pSA_ONE_MILELyr As IFeatureLayer = Nothing
    Dim pSA_10_MINUTE_AUTOLyr As IFeatureLayer = Nothing
    Dim pSA_20_MINUTE_AUTOLyr As IFeatureLayer = Nothing
    Dim pSA_30_MINUTE_AUTOLyr As IFeatureLayer = Nothing
    Dim pSA_30_MINUTE_TRANSITLyr As IFeatureLayer = Nothing
    Dim m_intFeatCount As Integer = 0
    Dim m_intTotalEmp As Integer = 0
    Dim m_intPRETotalEmp As Integer = 0

    Private Sub frmSumTransportationPnts_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        arrPolyLyrs = Nothing
        arrPntLyrs = Nothing
        strBufferFileGDB = Nothing
        m_strMissingLyrs = Nothing
        m_intMissingLyrs = Nothing
        pInersectionsLyr = Nothing
        p4WayLyr = Nothing
        pTransitStopLyr = Nothing
        pRailStopLyr = Nothing
        pSA_QUARTER_MILELyr = Nothing
        pSA_HALF_MILELyr = Nothing
        pSA_ONE_MILELyr = Nothing
        pSA_10_MINUTE_AUTOLyr = Nothing
        pSA_20_MINUTE_AUTOLyr = Nothing
        pSA_30_MINUTE_AUTOLyr = Nothing
        pSA_30_MINUTE_TRANSITLyr = Nothing
        m_intFeatCount = Nothing
        m_intTotalEmp = Nothing
        m_frmSumTransLocations = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub

    Private Sub frmSumTransportationPnts_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'LOAD IN POLYGON AND POINT LAYERS INTO FORM CONTROLS
        If Not Form_LoadData(sender, e) Then
            m_frmSumTransLocations.Close()
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
        Dim intPolyFeatCount As Integer = 0
        Dim intPntFeatCount As Integer = 0
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
        m_frmSumTransLocations.cmbLayers.Items.Clear()
        m_frmSumTransLocations.cmbAllIntersections.Items.Clear()
        m_frmSumTransLocations.cmb4WayIntersections.Items.Clear()
        m_frmSumTransLocations.cmbTransitStops.Items.Clear()
        m_frmSumTransLocations.cmbRailStops.Items.Clear()
        arrPolyLyrs.Clear()
        arrPntLyrs.Clear()

        If mapCurrent.LayerCount > 0 Then
            For intLayer = 0 To mapCurrent.LayerCount - 1
                pLyr = CType(mapCurrent.Layer(intLayer), ILayer)
                If TypeOf mapCurrent.Layer(intLayer) Is IFeatureLayer Then
                    pFeatLyr = CType(mapCurrent.Layer(intLayer), IFeatureLayer)
                    If pFeatLyr.FeatureClass.ShapeType = ESRI.ArcGIS.Geometry.esriGeometryType.esriGeometryPolygon Then
                        arrPolyLyrs.Add(pFeatLyr)
                        Me.cmbLayers.Items.Add(pFeatLyr.Name)
                        intPolyFeatCount = intPolyFeatCount + 1
                    ElseIf pFeatLyr.FeatureClass.ShapeType = ESRI.ArcGIS.Geometry.esriGeometryType.esriGeometryPoint Or pFeatLyr.FeatureClass.ShapeType = ESRI.ArcGIS.Geometry.esriGeometryType.esriGeometryMultipoint Then
                        arrPntLyrs.Add(pFeatLyr)
                        Me.cmbAllIntersections.Items.Add(pFeatLyr.Name)
                        Me.cmb4WayIntersections.Items.Add(pFeatLyr.Name)
                        Me.cmbTransitStops.Items.Add(pFeatLyr.Name)
                        Me.cmbRailStops.Items.Add(pFeatLyr.Name)
                        intPntFeatCount = intPntFeatCount + 1
                    End If
                    pFeatLyr = Nothing
                End If
            Next
        Else
            MessageBox.Show("Please add an input point or polygon layers in the current view document to use this tool.", "No Layer(s) Found", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo CloseForm
        End If

        If intPolyFeatCount <= 0 Then
            MessageBox.Show("No polygon layers were found in the current active map document.  Please add a Neighborhood polygon layer before accessing this tool.", "Polygon Layer Required", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo CloseForm
        ElseIf intPntFeatCount <= 0 Then
            MessageBox.Show("No point layers were found in the current active map document.  Please add a transportation location point layers before accessing this tool.", "Point Layers Required", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo CloseForm
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
        intPolyFeatCount = Nothing
        intPntFeatCount = Nothing
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
        Dim pField As IField
        Dim intFld As Integer = 0
        Dim strFld As String

        'NEIGHBORHOOD LAYER SELECTED
        Try
            pFeatLyr = arrPolyLyrs.Item(Me.cmbLayers.SelectedIndex)
            pFeatureClass = pFeatLyr.FeatureClass
            pTable = CType(pFeatureClass, ITable)
            pFeatSelection = CType(pFeatLyr, IFeatureSelection)
            pFeatSelection.SelectionSet.Search(Nothing, False, pCursor)
            Me.chkUseSelected.Visible = False
            If pFeatSelection.SelectionSet.Count > 0 Then
                Me.chkUseSelected.Visible = True
                Me.chkUseSelected.Text = "Selected Features (" & pFeatSelection.SelectionSet.Count.ToString & ")"
            End If

            'FIND AVAILABLE ID FIELD(S) 
            Me.cmbFieldId.Enabled = False
            Me.cmbArea.Enabled = False
            Me.cmbEmpFld.Enabled = False
            Me.cmbFieldId.Items.Clear()
            Me.cmbArea.Items.Clear()
            Me.cmbEmpFld.Items.Clear()
            For intFld = 0 To pFeatureClass.Fields.FieldCount - 1
                pField = pFeatureClass.Fields.Field(intFld)
                'DON'T ADD OUTPUT FIELD NAMES IN LIST OF ID OR ACRES FIELD CONTROLS
                If (pField.Type = esriFieldType.esriFieldTypeInteger Or pField.Type = esriFieldType.esriFieldTypeSmallInteger) And Not pField.Type = esriFieldType.esriFieldTypeOID Then
                    Me.cmbFieldId.Items.Add(pField.Name)
                End If
                If (pField.Type = esriFieldType.esriFieldTypeInteger Or pField.Type = esriFieldType.esriFieldTypeSmallInteger Or pField.Type = esriFieldType.esriFieldTypeDouble) And Not pField.Type = esriFieldType.esriFieldTypeOID Then
                    Me.cmbArea.Items.Add(pField.Name)
                    Me.cmbEmpFld.Items.Add(pField.Name)
                End If
            Next

            If Me.cmbFieldId.Items.Count > 0 Then
                Me.cmbFieldId.Enabled = True
            End If
            If Me.cmbArea.Items.Count > 0 Then
                Me.cmbArea.Enabled = True
                Me.cmbEmpFld.Enabled = True
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
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub

    Private Sub cmbFieldId_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbAllIntersections.SelectedIndexChanged
        'RETRIEVE THE INTERSECTIONS LAYER
        If Me.cmbAllIntersections.Items.Contains(Me.cmbAllIntersections.Text) Then
            pInersectionsLyr = arrPntLyrs.Item(Me.cmbAllIntersections.SelectedIndex)
        End If
    End Sub

    Private Sub cmb4WayIntersections_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmb4WayIntersections.SelectedIndexChanged
        'RETRIEVE THE 4 WAY INTERSECTIONS LAYER
        If Me.cmb4WayIntersections.Items.Contains(Me.cmb4WayIntersections.Text) Then
            p4WayLyr = arrPntLyrs.Item(Me.cmb4WayIntersections.SelectedIndex)
        End If
    End Sub

    Private Sub cmbTransitStops_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbTransitStops.SelectedIndexChanged
        'RETRIEVE THE TRANSIT STPS LAYER
        If Me.cmbTransitStops.Items.Contains(Me.cmbTransitStops.Text) Then
            pTransitStopLyr = arrPntLyrs.Item(Me.cmbTransitStops.SelectedIndex)
        End If
    End Sub

    Private Sub cmbRailStops_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbRailStops.SelectedIndexChanged
        'RETRIEVE THE RAIL STOPS LAYER
        If Me.cmbRailStops.Items.Contains(Me.cmbRailStops.Text) Then
            pRailStopLyr = arrPntLyrs.Item(Me.cmbRailStops.SelectedIndex)
        End If
    End Sub

    Private Sub OpenBufferLayers()
        'RETRIEVE EACH OF THE 7 BUFFER FEATURE LAYERS FROM THE SELECTED BUFFER WORKSPACE AND HOLD IN MEMORY
        Dim pWksFactory As IWorkspaceFactory = Nothing
        Dim pFeatWks As IFeatureWorkspace
        Dim pFeatClass As IFeatureClass
        Dim blnNotFound As Boolean = False
        Dim strMissingFieldLyrs As String = ""

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
                If pSA_QUARTER_MILELyr.FeatureClass.FindField(Me.cmbFieldId.Text) <= -1 Then
                    blnNotFound = True
                    If strMissingFieldLyrs.Length <= 0 Then
                        strMissingFieldLyrs = strMissingFieldLyrs & "SA_QUARTER_MILE"
                    Else
                        strMissingFieldLyrs = strMissingFieldLyrs & ", SA_QUARTER_MILE"
                    End If
                    pSA_QUARTER_MILELyr = Nothing
                End If
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
                If pSA_HALF_MILELyr.FeatureClass.FindField(Me.cmbFieldId.Text) <= -1 Then
                    blnNotFound = True
                    If strMissingFieldLyrs.Length <= 0 Then
                        strMissingFieldLyrs = strMissingFieldLyrs & "SA_HALF_MILE"
                    Else
                        strMissingFieldLyrs = strMissingFieldLyrs & ", SA_HALF_MILE"
                    End If
                    pSA_HALF_MILELyr = Nothing
                End If
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
                If pSA_ONE_MILELyr.FeatureClass.FindField(Me.cmbFieldId.Text) <= -1 Then
                    blnNotFound = True
                    If strMissingFieldLyrs.Length <= 0 Then
                        strMissingFieldLyrs = strMissingFieldLyrs & "SA_ONE_MILE"
                    Else
                        strMissingFieldLyrs = strMissingFieldLyrs & ", SA_ONE_MILE"
                    End If
                    pSA_ONE_MILELyr = Nothing
                End If
            Else
                pSA_ONE_MILELyr = Nothing
            End If
        Catch ex As Exception
            GoTo CleanUp
        End Try
        Try
            pFeatClass = pFeatWks.OpenFeatureClass("SA_10_MINUTE_AUTO")
            If pFeatClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                pSA_10_MINUTE_AUTOLyr = New FeatureLayer
                pSA_10_MINUTE_AUTOLyr.FeatureClass = pFeatClass
                If pSA_10_MINUTE_AUTOLyr.FeatureClass.FindField(Me.cmbFieldId.Text) <= -1 Then

                    blnNotFound = True
                    If strMissingFieldLyrs.Length <= 0 Then
                        strMissingFieldLyrs = strMissingFieldLyrs & "SA_10_MINUTE_AUTO"
                    Else
                        strMissingFieldLyrs = strMissingFieldLyrs & ", SA_10_MINUTE_AUTO"
                    End If
                    pSA_10_MINUTE_AUTOLyr = Nothing
                End If
            Else
                pSA_10_MINUTE_AUTOLyr = Nothing
            End If
        Catch ex As Exception
            GoTo CleanUp
        End Try
        Try
            pFeatClass = pFeatWks.OpenFeatureClass("SA_20_MINUTE_AUTO")
            If pFeatClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                pSA_20_MINUTE_AUTOLyr = New FeatureLayer
                pSA_20_MINUTE_AUTOLyr.FeatureClass = pFeatClass
                If pSA_20_MINUTE_AUTOLyr.FeatureClass.FindField(Me.cmbFieldId.Text) <= -1 Then

                    blnNotFound = True
                    If strMissingFieldLyrs.Length <= 0 Then
                        strMissingFieldLyrs = strMissingFieldLyrs & "SA_20_MINUTE_AUTO"
                    Else
                        strMissingFieldLyrs = strMissingFieldLyrs & ", SA_20_MINUTE_AUTO"
                    End If
                    pSA_20_MINUTE_AUTOLyr = Nothing
                End If
            Else
                pSA_20_MINUTE_AUTOLyr = Nothing
            End If
        Catch ex As Exception
            GoTo CleanUp
        End Try
        Try
            pFeatClass = pFeatWks.OpenFeatureClass("SA_30_MINUTE_AUTO")
            If pFeatClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                pSA_30_MINUTE_AUTOLyr = New FeatureLayer
                pSA_30_MINUTE_AUTOLyr.FeatureClass = pFeatClass
                If pSA_30_MINUTE_AUTOLyr.FeatureClass.FindField(Me.cmbFieldId.Text) <= -1 Then
                    blnNotFound = True
                    If strMissingFieldLyrs.Length <= 0 Then
                        strMissingFieldLyrs = strMissingFieldLyrs & "SA_30_MINUTE_AUTO"
                    Else
                        strMissingFieldLyrs = strMissingFieldLyrs & ", SA_30_MINUTE_AUTO"
                    End If
                    pSA_30_MINUTE_AUTOLyr = Nothing
                End If
            Else
                pSA_30_MINUTE_AUTOLyr = Nothing
            End If
        Catch ex As Exception
            GoTo CleanUp
        End Try
        Try
            pFeatClass = pFeatWks.OpenFeatureClass("SA_30_MINUTE_TRANSIT")
            If pFeatClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                pSA_30_MINUTE_TRANSITLyr = New FeatureLayer
                pSA_30_MINUTE_TRANSITLyr.FeatureClass = pFeatClass
                If pSA_30_MINUTE_TRANSITLyr.FeatureClass.FindField(Me.cmbFieldId.Text) <= -1 Then
                    blnNotFound = True
                    If strMissingFieldLyrs.Length <= 0 Then
                        strMissingFieldLyrs = strMissingFieldLyrs & "SA_30_MINUTE_TRANSIT"
                    Else
                        strMissingFieldLyrs = strMissingFieldLyrs & ", SA_30_MINUTE_TRANSIT"
                    End If
                    pSA_30_MINUTE_TRANSITLyr = Nothing
                End If
            Else
                pSA_30_MINUTE_TRANSITLyr = Nothing
            End If

            If strMissingFieldLyrs.Length > 0 Then
                MessageBox.Show("The field, " & Me.cmbFieldId.Text & ", could not be found in following service area layer(s): (" & strMissingFieldLyrs & ")", "Id Field Not Found", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)

            End If
            GoTo CleanUp

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

    Private Sub btnWorkspace_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnWorkspace.Click
        'SELECT THE FILE GEODATABASE CONTAINING THE FEATURE BUFFER LAYERS
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
        m_strMissingLyrs = ""
        m_intMissingLyrs = 0

        Try
            pWksFactory = New FileGDBWorkspaceFactory
            pFeatWks = pWksFactory.OpenFromFile(strBufferFileGDB, 0)
            For Each strLyrName In m_strBufferLayerNames.Split(",")
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

    Private Sub btnRun_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRun.Click
        'EXECUTE THE SUBS TO SUMMARIZE THE NEEDED VARIABLES
        Dim pFeatSelection As IFeatureSelection
        Dim pCursor As ICursor = Nothing
        Dim pDef As IFeatureLayerDefinition2
        Dim pNFeatClass As IFeatureClass
        Dim pTable As ITable
        Dim intFld As Integer = 0
        Dim strFld As String
        'Dim strReqFlds1 As String = "ActDen,ActDenQtrMi,ActDenHalfMi,ActDen1Mi,JobPopBal,JobPopBalQtrMi,JobPopBalHalfMi,JobPopBal1Mi,PctRegEmp10MinA,Emp10MinA,PctRegEmp20MinA,Emp20MinA,PctRegEmp30MinA,Emp30MinA,PctRegEmp30MinT,IntDen,IntDenQtrMi,IntDenHalfMi,IntDen1Mi,TransStopDen,TransStopDenQtrMi,TransStopDenHalfMi,TransStopDen1Mi,RailStopDen,RailStopDenHalfMi"
        'Dim strReqFlds2 As String = "Pct4WayIntQtrMi,Pct4WayIntHalfMi,Pct4WayInt1Mi"
        'Dim strPreFixes = "EX_,TOT_"
        Dim strOutputMainA = "ActDen,JobPopBal,IntDen,Pct4WayInt,TransStopDen,RailStopDen"
        Dim strOutputMainB = "Emp,PctRegEmp"
        Dim strSuffixesA As String = ",QtrMi,HalfMi,1Mi"
        Dim strSuffixesB As String = ",10MinA,20MinA,30MinA,30MinT"
        Dim strFieldPart1 As String
        Dim strFieldPart2 As String
        Dim strFieldPart3 As String
        Dim pQFilter As IQueryFilter
        Dim strDefExpression As String = ""
        Dim strQString As String = Nothing
        Dim pFeatureCursor As IFeatureCursor = Nothing
        Dim intTotalCount As Integer
        Dim intObjFld As Integer
        Dim pFeat As IFeature
        Dim intCount As Integer = 0
        Dim intObjId As Integer
        Dim pBuffFeatSelection As IFeatureSelection
        Dim pDataStatistics As DataStatistics


        'RESET THE PREFIX LIST
        If Me.rdbExisting.Checked Then
            strFieldPart1 = "EX_PRE_"
        Else
            strFieldPart1 = "TOT_"
        End If
        'RETREIVE THE SELECTED NEIGHBORHOOD LAYER
        m_appEnvision.StatusBar.Message(0) = "Retrieving selected Neighborhood layer"
        Try
            pNFeatLyr = arrPolyLyrs.Item(Me.cmbLayers.SelectedIndex)
            pNFeatClass = pNFeatLyr.FeatureClass
            pDef = pNFeatLyr
            pTable = CType(pNFeatClass, ITable)
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

        'CHECK THAT THE TOTAL EMPLOYMENT FIELD IS SELECTED
        If Not Me.cmbEmpFld.Items.Contains(Me.cmbEmpFld.Text) Then
            If MessageBox.Show("A valid Total Employment field was not defined.  Would you like to continue?", "Total Employment Field Not Found", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2, MessageBoxOptions.DefaultDesktopOnly, False) = Windows.Forms.DialogResult.No Then
                GoTo CleanUp
            End If
        End If


        'BUILD LIST OF REQUIRED FIELD(S) FOR NON-EMP VARIABLE, ADD IF MISSING FROM NEIGHBORHOOD LAYER
        For Each strFieldPart2 In Split(strOutputMainA, ",")
            'For Each strFieldPart1 In Split(strPreFixes, ",")
            For Each strFieldPart3 In Split(strSuffixesA, ",")
                If pNFeatClass.FindField(strFieldPart1 & strFieldPart2 & strFieldPart3) = -1 Then
                    m_appEnvision.StatusBar.Message(0) = "Checking for required value fields:  Adding missing field: " & (strFieldPart1 & strFieldPart2 & strFieldPart3)
                    AddEnvisionField(pTable, (strFieldPart1 & strFieldPart2 & strFieldPart3), "DOUBLE", 16, 6)
                End If
            Next
        Next

        'BUILD LIST OF REQUIRED FIELD(S) FOR NON-EMP VARIABLE, ADD IF MISSING FROM NEIGHBORHOOD LAYER
        For Each strFieldPart2 In Split(strOutputMainB, ",")
            'For Each strFieldPart1 In Split(strPreFixes, ",")
            For Each strFieldPart3 In Split(strSuffixesB, ",")
                If pNFeatClass.FindField(strFieldPart1 & strFieldPart2 & strFieldPart3) = -1 Then
                    m_appEnvision.StatusBar.Message(0) = "Checking for required value fields:  Adding missing field: " & (strFieldPart1 & strFieldPart2 & strFieldPart3)
                    AddEnvisionField(pTable, (strFieldPart1 & strFieldPart2 & strFieldPart3), "DOUBLE", 16, 6)
                End If
            Next
        Next


        'RETRIEVE THE BUFFER LAYERS
        OpenBufferLayers()


        'RETRIEVE THE NEIGHBORHOOD "TOT_EMP" FIELD AND SUM VALUE 
        If pNFeatLyr.FeatureClass.FindField(Me.cmbEmpFld.Text) >= 0 Then
            Try
                pCursor = CType(pTable.Search(Nothing, False), ICursor)
                pDataStatistics = New DataStatistics
                pDataStatistics.Field = Me.cmbEmpFld.Text
                pDataStatistics.Cursor = pCursor
                m_intTotalEmp = pDataStatistics.Statistics.Sum
            Catch ex As Exception
                m_intTotalEmp = 0
            End Try
        End If

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
            intCount = 0
            intObjId = pFeat.Value(intObjFld)

            Try
                pQFilter = New QueryFilter
                strQString = Me.cmbFieldId.Text & " = " & intObjId.ToString
                pQFilter.WhereClause = strQString
                '----------------------------------------------------------------------------------------------------------------------------------
                'NO BUFFER DISTANCE DENSITIES
                pBuffFeatSelection = CType(pNFeatLyr, IFeatureSelection)
                pBuffFeatSelection.Clear()
                pBuffFeatSelection.SelectFeatures(pQFilter, esriSelectionResultEnum.esriSelectionResultNew, True)
                If pNFeatLyr Is Nothing Or pBuffFeatSelection.SelectionSet.Count <= 0 Then
                    pFeat = pFeatureCursor.NextFeature
                    Continue Do
                End If


                If Not pNFeatLyr Is Nothing And pBuffFeatSelection.SelectionSet.Count > 0 Then
                    'INTERSECTION SUMMARY
                    If Not pInersectionsLyr Is Nothing Then
                        m_intFeatCount = 0
                        m_appEnvision.StatusBar.Message(0) = "Zero Mile Buffer Selection with All Intersections"
                        CalcDensity(Nothing, pInersectionsLyr, "IntDen", pFeat, True)
                    End If
                    intCount = intCount + 1
                    Me.barStatus.Value = (intCount / 23) * 100
                    Me.Refresh()
                    pBuffFeatSelection = CType(pNFeatLyr, IFeatureSelection)
                    pBuffFeatSelection.Clear()
                    pBuffFeatSelection.SelectFeatures(pQFilter, esriSelectionResultEnum.esriSelectionResultNew, True)
                    '4-WAY INTERSECTION SUMMARY
                    If Not p4WayLyr Is Nothing Then
                        m_appEnvision.StatusBar.Message(0) = "Zero Mile Buffer Selection with 4-Way Intersections"
                        If Me.rdbExisting.Checked Then
                            CalcPercentage(Nothing, p4WayLyr, "EX_PRE_Pct4WayInt", pFeat, True)
                        Else
                            CalcPercentage(Nothing, p4WayLyr, "TOT_Pct4WayInt", pFeat, True)
                        End If
                    End If
                    intCount = intCount + 1
                    Me.barStatus.Value = (intCount / 23) * 100
                    Me.Refresh()
                    pBuffFeatSelection = CType(pNFeatLyr, IFeatureSelection)
                    pBuffFeatSelection.Clear()
                    pBuffFeatSelection.SelectFeatures(pQFilter, esriSelectionResultEnum.esriSelectionResultNew, True)
                    'TRANSIT STOP SUMMARY
                    If Not pTransitStopLyr Is Nothing Then
                        m_appEnvision.StatusBar.Message(0) = "Zero Mile Buffer Selection with Transit Stops"
                        CalcDensity(Nothing, pTransitStopLyr, "TransStopDen", pFeat, True)
                        intCount = intCount + 1
                        Me.barStatus.Value = (intCount / 23) * 100
                        Me.Refresh()
                        pBuffFeatSelection = CType(pNFeatLyr, IFeatureSelection)
                        pBuffFeatSelection.Clear()
                        pBuffFeatSelection.SelectFeatures(pQFilter, esriSelectionResultEnum.esriSelectionResultNew, True)
                        'RAIL STOP SUMMARY
                        If Not pTransitStopLyr Is Nothing Then
                            m_appEnvision.StatusBar.Message(0) = "Zero Mile Buffer Selection with Rail Stops"
                            CalcDensity(Nothing, pRailStopLyr, "RailStopDen", pFeat, True)
                        End If
                        intCount = intCount + 1
                        Me.barStatus.Value = (intCount / 23) * 100
                        Me.Refresh()
                    End If
                End If

                '----------------------------------------------------------------------------------------------------------------------------------
                'QUARTER MILE BUFFER
                pBuffFeatSelection = CType(pSA_QUARTER_MILELyr, IFeatureSelection)
                pBuffFeatSelection.Clear()
                pBuffFeatSelection.SelectFeatures(pQFilter, esriSelectionResultEnum.esriSelectionResultNew, True)
                If Not pSA_QUARTER_MILELyr Is Nothing And pBuffFeatSelection.SelectionSet.Count > 0 Then
                    'INTERSECTION SUMMARY
                    If Not pInersectionsLyr Is Nothing Then
                        m_intFeatCount = 0
                        m_appEnvision.StatusBar.Message(0) = "Quarter Mile Buffer Selection with All Intersections"
                        CalcDensity(pSA_QUARTER_MILELyr, pInersectionsLyr, "IntDenQtrMi", pFeat, False)
                    End If
                    intCount = intCount + 1
                    Me.barStatus.Value = (intCount / 23) * 100
                    Me.Refresh()
                    '4-WAY INTERSECTION SUMMARY
                    pBuffFeatSelection = CType(pSA_QUARTER_MILELyr, IFeatureSelection)
                    pBuffFeatSelection.Clear()
                    pBuffFeatSelection.SelectFeatures(pQFilter, esriSelectionResultEnum.esriSelectionResultNew, True)
                    If Not p4WayLyr Is Nothing Then
                        m_appEnvision.StatusBar.Message(0) = "Quarter Mile Buffer Selection with 4-Way Intersections"
                        If Me.rdbExisting.Checked Then
                            CalcPercentage(pSA_QUARTER_MILELyr, p4WayLyr, "EX_PRE_Pct4WayIntQtrMi", pFeat, False)
                        Else
                            CalcPercentage(pSA_QUARTER_MILELyr, p4WayLyr, "TOT_Pct4WayIntQtrMi", pFeat, False)
                        End If
                    End If
                    intCount = intCount + 1
                    Me.barStatus.Value = (intCount / 23) * 100
                    Me.Refresh()
                    'TRANSIT STOP SUMMARY
                    pBuffFeatSelection = CType(pSA_QUARTER_MILELyr, IFeatureSelection)
                    pBuffFeatSelection.Clear()
                    pBuffFeatSelection.SelectFeatures(pQFilter, esriSelectionResultEnum.esriSelectionResultNew, True)
                    If Not pTransitStopLyr Is Nothing Then
                        m_appEnvision.StatusBar.Message(0) = "Quarter Mile Buffer Selection with Transit Stops"
                        CalcDensity(pSA_QUARTER_MILELyr, pTransitStopLyr, "TransStopDenQtrMi", pFeat, False)
                    End If
                    intCount = intCount + 1
                    Me.barStatus.Value = (intCount / 23) * 100
                    Me.Refresh()
                    'RAIL STOP SUMMARY
                    pBuffFeatSelection = CType(pSA_QUARTER_MILELyr, IFeatureSelection)
                    pBuffFeatSelection.Clear()
                    pBuffFeatSelection.SelectFeatures(pQFilter, esriSelectionResultEnum.esriSelectionResultNew, True)
                    If Not pTransitStopLyr Is Nothing Then
                        m_appEnvision.StatusBar.Message(0) = "Quarter Mile Buffer Selection with Rail Stops"
                        CalcDensity(pSA_QUARTER_MILELyr, pRailStopLyr, "RailStopDenQtrMi", pFeat, False)
                    End If
                    intCount = intCount + 1
                    Me.barStatus.Value = (intCount / 23) * 100
                    Me.Refresh()
                End If
                intCount = 3

                '----------------------------------------------------------------------------------------------------------------------------------
                'HALF MILE BUFFER
                pBuffFeatSelection = CType(pSA_HALF_MILELyr, IFeatureSelection)
                pBuffFeatSelection.Clear()
                pBuffFeatSelection.SelectFeatures(pQFilter, esriSelectionResultEnum.esriSelectionResultNew, True)
                If Not pSA_HALF_MILELyr Is Nothing And pBuffFeatSelection.SelectionSet.Count > 0 Then
                    'INTERSECTION SUMMARY
                    If Not pInersectionsLyr Is Nothing Then
                        m_intFeatCount = 0
                        m_appEnvision.StatusBar.Message(0) = "Half Mile Buffer Selection with All Intersections"
                        CalcDensity(pSA_HALF_MILELyr, pInersectionsLyr, "IntDenHalfMi", pFeat, False)
                    End If
                    intCount = intCount + 1
                    Me.barStatus.Value = (intCount / 23) * 100
                    Me.Refresh()
                    '4-WAY INTERSECTION SUMMARY
                    If Not p4WayLyr Is Nothing Then
                        m_appEnvision.StatusBar.Message(0) = "Half Mile Buffer Selection with 4-Way Instersections"
                        If Me.rdbExisting.Checked Then
                            CalcPercentage(pSA_HALF_MILELyr, p4WayLyr, "EX_PRE_Pct4WayIntHalfMi", pFeat, False)
                        Else
                            CalcPercentage(pSA_HALF_MILELyr, p4WayLyr, "TOT_Pct4WayIntHalfMi", pFeat, False)
                        End If
                    End If
                    intCount = intCount + 1
                    Me.barStatus.Value = (intCount / 23) * 100
                    Me.Refresh()
                    'TRANSIT STOP SUMMARY
                    If Not pTransitStopLyr Is Nothing Then
                        m_appEnvision.StatusBar.Message(0) = "Half Mile Buffer Selection with Transit Stops"
                        CalcDensity(pSA_HALF_MILELyr, pTransitStopLyr, "TransStopDenHalfMi", pFeat, False)
                    End If
                    intCount = intCount + 1
                    Me.barStatus.Value = (intCount / 23) * 100
                    Me.Refresh()
                    'RAIL STOP SUMMARY
                    If Not pRailStopLyr Is Nothing Then
                        m_appEnvision.StatusBar.Message(0) = "Half Mile Buffer Selection with Rail Stops"
                        CalcDensity(pSA_HALF_MILELyr, pRailStopLyr, "RailStopDenHalfMi", pFeat, False)
                    End If
                    intCount = intCount + 1
                    Me.barStatus.Value = (intCount / 23) * 100
                    Me.Refresh()
                End If
                intCount = 7

                '----------------------------------------------------------------------------------------------------------------------------------
                'ONE MILE BUFFER
                pBuffFeatSelection = CType(pSA_ONE_MILELyr, IFeatureSelection)
                pBuffFeatSelection.Clear()
                pBuffFeatSelection.SelectFeatures(pQFilter, esriSelectionResultEnum.esriSelectionResultNew, True)
                If Not pSA_ONE_MILELyr Is Nothing And pBuffFeatSelection.SelectionSet.Count > 0 Then
                    'INTERSECTION SUMMARY
                    If Not pInersectionsLyr Is Nothing Then
                        m_intFeatCount = 0
                        m_appEnvision.StatusBar.Message(0) = "One Mile Buffer Selection with All Instersections"
                        CalcDensity(pSA_ONE_MILELyr, pInersectionsLyr, "IntDen1Mi", pFeat, False)

                        intCount = intCount + 1
                        Me.barStatus.Value = (intCount / 23) * 100
                        Me.Refresh()
                        pBuffFeatSelection = CType(pSA_ONE_MILELyr, IFeatureSelection)
                        pBuffFeatSelection.Clear()
                        pBuffFeatSelection.SelectFeatures(pQFilter, esriSelectionResultEnum.esriSelectionResultNew, True)
                        '4-WAY INTERSECTION SUMMARY
                        If Not p4WayLyr Is Nothing Then
                            m_appEnvision.StatusBar.Message(0) = "One Mile Buffer Selection with 4-Way Intersections"
                            If Me.rdbExisting.Checked Then
                                CalcPercentage(pSA_ONE_MILELyr, p4WayLyr, "EX_PRE_Pct4WayInt1Mi", pFeat, False)
                            Else
                                CalcPercentage(pSA_ONE_MILELyr, p4WayLyr, "TOT_Pct4WayInt1Mi", pFeat, False)
                            End If
                        End If
                        intCount = intCount + 1
                        Me.barStatus.Value = (intCount / 23) * 100
                        Me.Refresh()
                        pBuffFeatSelection = CType(pSA_ONE_MILELyr, IFeatureSelection)
                        pBuffFeatSelection.Clear()
                        pBuffFeatSelection.SelectFeatures(pQFilter, esriSelectionResultEnum.esriSelectionResultNew, True)
                        'TRANSIT STOP SUMMARY
                        If Not pTransitStopLyr Is Nothing Then
                            m_appEnvision.StatusBar.Message(0) = "One Mile Buffer Selection with Transit Stops"
                            CalcDensity(pSA_ONE_MILELyr, pTransitStopLyr, "TransStopDen1Mi", pFeat, False)
                        End If
                        pBuffFeatSelection = CType(pSA_ONE_MILELyr, IFeatureSelection)
                        pBuffFeatSelection.Clear()
                        pBuffFeatSelection.SelectFeatures(pQFilter, esriSelectionResultEnum.esriSelectionResultNew, True)
                        'RAIL STOP SUMMARY
                        intCount = intCount + 1
                        Me.barStatus.Value = (intCount / 23) * 100
                        Me.Refresh()
                        If Not pRailStopLyr Is Nothing Then
                            m_appEnvision.StatusBar.Message(0) = "Quarter Mile Buffer Selection with Rail Stops"
                            CalcDensity(pSA_ONE_MILELyr, pRailStopLyr, "RailStopDen1Mi", pFeat, False)
                        End If
                    End If
                End If

                '----------------------------------------------------------------------------------------------------------------------------------
                'SUMMARY OF ACTIVITY DENSITY AND JOBS/HOUSING BALANCE NUMBERS BY BUFFERS
                'ZERO MILE BUFFER
                pBuffFeatSelection = CType(pNFeatLyr, IFeatureSelection)
                pBuffFeatSelection.Clear()
                pBuffFeatSelection.SelectFeatures(pQFilter, esriSelectionResultEnum.esriSelectionResultNew, True)
                If pBuffFeatSelection.SelectionSet.Count > 0 Then
                    m_intFeatCount = 0
                    m_appEnvision.StatusBar.Message(0) = "Zero Mile Buffer for Activity Density and Jobs-Pop Balance"
                    CalcActDenAndJobPopBal(Nothing, pNFeatLyr, pFeat, True, "ZERO")
                End If
                intCount = intCount + 1
                Me.barStatus.Value = (intCount / 23) * 100
                Me.Refresh()
                '1/4 MILE BUFFER
                pBuffFeatSelection = CType(pSA_QUARTER_MILELyr, IFeatureSelection)
                pBuffFeatSelection.Clear()
                pBuffFeatSelection.SelectFeatures(pQFilter, esriSelectionResultEnum.esriSelectionResultNew, True)
                If pBuffFeatSelection.SelectionSet.Count > 0 Then
                    m_intFeatCount = 0
                    m_appEnvision.StatusBar.Message(0) = "Quarter Mile Buffer for Activity Density and Jobs-Pop Balance"
                    CalcActDenAndJobPopBal(pSA_QUARTER_MILELyr, pNFeatLyr, pFeat, False, "QUARTER")
                End If
                intCount = intCount + 1
                Me.barStatus.Value = (intCount / 23) * 100
                Me.Refresh()
                '1/2 MILE BUFFER
                pBuffFeatSelection = CType(pSA_QUARTER_MILELyr, IFeatureSelection)
                pBuffFeatSelection.Clear()
                pBuffFeatSelection.SelectFeatures(pQFilter, esriSelectionResultEnum.esriSelectionResultNew, True)
                If pBuffFeatSelection.SelectionSet.Count > 0 Then
                    m_intFeatCount = 0
                    m_appEnvision.StatusBar.Message(0) = "Half Mile Buffer for Activity Density and Jobs-Pop Balance"
                    CalcActDenAndJobPopBal(pSA_HALF_MILELyr, pNFeatLyr, pFeat, False, "HALF")
                End If
                intCount = intCount + 1
                Me.barStatus.Value = (intCount / 23) * 100
                Me.Refresh()
                '1 MILE BUFFER
                pBuffFeatSelection = CType(pSA_QUARTER_MILELyr, IFeatureSelection)
                pBuffFeatSelection.Clear()
                pBuffFeatSelection.SelectFeatures(pQFilter, esriSelectionResultEnum.esriSelectionResultNew, True)
                If pBuffFeatSelection.SelectionSet.Count > 0 Then
                    m_intFeatCount = 0
                    m_appEnvision.StatusBar.Message(0) = "One Mile Buffer for Activity Density and Jobs-Pop Balance"
                    CalcActDenAndJobPopBal(pSA_ONE_MILELyr, pNFeatLyr, pFeat, False, "ONE")
                End If
                intCount = intCount + 1
                Me.barStatus.Value = (intCount / 23) * 100
                Me.Refresh()

                '----------------------------------------------------------------------------------------------------------------------------------
                'SUMMARY OF EMPLOYMENT NUMBERS BY BUFFERS
                If m_intTotalEmp > 0 And pNFeatLyr.FeatureClass.FindField(Me.cmbEmpFld.Text) >= 0 Then
                    '10 MINUTE AUTO BUFFER
                    pBuffFeatSelection = CType(pNFeatLyr, IFeatureSelection)
                    pBuffFeatSelection.Clear()
                    pBuffFeatSelection.SelectFeatures(pQFilter, esriSelectionResultEnum.esriSelectionResultNew, True)
                    If pBuffFeatSelection.SelectionSet.Count > 0 Then
                        m_intFeatCount = 0
                        m_appEnvision.StatusBar.Message(0) = "10 Minute Auto Buffer with Envision Parcel Emp"
                        CalcEmp(Nothing, pNFeatLyr, "PctRegEmp", "PRE_Emp", pFeat, True)
                    End If
                    '10 MINUTE AUTO BUFFER
                    pBuffFeatSelection = CType(pSA_10_MINUTE_AUTOLyr, IFeatureSelection)
                    pBuffFeatSelection.Clear()
                    pBuffFeatSelection.SelectFeatures(pQFilter, esriSelectionResultEnum.esriSelectionResultNew, True)
                    If pBuffFeatSelection.SelectionSet.Count > 0 Then
                        m_intFeatCount = 0
                        m_appEnvision.StatusBar.Message(0) = "10 Minute Auto Buffer with Envision Parcel Emp"
                        CalcEmp(pSA_10_MINUTE_AUTOLyr, pNFeatLyr, "PctRegEmp10MinA", "Emp10MinA", pFeat, False)
                    End If
                    intCount = intCount + 1
                    Me.barStatus.Value = (intCount / 23) * 100
                    Me.Refresh()
                    '20 MINUTE AUTO BUFFER
                    If Not pSA_20_MINUTE_AUTOLyr Is Nothing Then
                        pBuffFeatSelection = CType(pSA_20_MINUTE_AUTOLyr, IFeatureSelection)
                        pBuffFeatSelection.Clear()
                        pBuffFeatSelection.SelectFeatures(pQFilter, esriSelectionResultEnum.esriSelectionResultNew, True)
                        If pBuffFeatSelection.SelectionSet.Count > 0 Then
                            'INTERSECTION SUMMARY
                            m_appEnvision.StatusBar.Message(0) = "20 Minute Auto Buffer with Envision Parcel Emp"
                            m_intFeatCount = 0
                            CalcEmp(pSA_20_MINUTE_AUTOLyr, pNFeatLyr, "PctRegEmp20MinA", "Emp20MinA", pFeat, False)
                        End If
                    End If
                    intCount = intCount + 1
                    Me.barStatus.Value = (intCount / 23) * 100
                    Me.Refresh()
                    '30 MINUTE AUTO BUFFER
                    If Not pSA_30_MINUTE_AUTOLyr Is Nothing Then
                        pBuffFeatSelection = CType(pSA_30_MINUTE_AUTOLyr, IFeatureSelection)
                        pBuffFeatSelection.Clear()
                        pBuffFeatSelection.SelectFeatures(pQFilter, esriSelectionResultEnum.esriSelectionResultNew, True)
                        If pBuffFeatSelection.SelectionSet.Count > 0 Then
                            'INTERSECTION SUMMARY
                            m_appEnvision.StatusBar.Message(0) = "30 Minute Auto Buffer with Envision Parcel Emp"
                            m_intFeatCount = 0
                            CalcEmp(pSA_30_MINUTE_AUTOLyr, pNFeatLyr, "PctRegEmp30MinA", "Emp30MinA", pFeat, False)
                        End If
                    End If
                    intCount = intCount + 1
                    Me.barStatus.Value = (intCount / 23) * 100
                    Me.Refresh()
                    '30 MINUTE TRANSIT BUFFER
                    If Not pSA_30_MINUTE_TRANSITLyr Is Nothing Then
                        pBuffFeatSelection = CType(pSA_30_MINUTE_TRANSITLyr, IFeatureSelection)
                        pBuffFeatSelection.Clear()
                        pBuffFeatSelection.SelectFeatures(pQFilter, esriSelectionResultEnum.esriSelectionResultNew, True)
                        If pBuffFeatSelection.SelectionSet.Count > 0 Then
                            'INTERSECTION SUMMARY
                            m_appEnvision.StatusBar.Message(0) = "30 Minute Transit Buffer with Envision Parcel Emp"
                            m_intFeatCount = 0
                            CalcEmp(pSA_30_MINUTE_TRANSITLyr, pNFeatLyr, "PctRegEmp30MinT", "Emp30MinT", pFeat, False)
                        End If
                    End If
                    intCount = intCount + 1
                    Me.barStatus.Value = (intCount / 23) * 100
                    Me.Refresh()
                End If
            Catch ex As Exception
                m_appEnvision.StatusBar.Message(0) = ex.Message
            End Try
            intCount = 14
            Me.barStatus.Value = (intCount / 23) * 100
            Me.Refresh()

            pFeat = pFeatureCursor.NextFeature
            m_appEnvision.StatusBar.Message(0) = ""
        Loop

CleanUp:
        MessageBox.Show("Transportation Locations Summary processing has completed.", "End Processing")
        Me.Close()
        pFeatSelection = Nothing
        pCursor = Nothing
        pDef = Nothing
        pNFeatLyr = Nothing
        pNFeatClass = Nothing
        pTable = Nothing
        intFld = Nothing
        strFld = Nothing
        pQFilter = Nothing
        strDefExpression = Nothing
        strQString = Nothing
        pFeatureCursor = Nothing
        intTotalCount = Nothing
        intObjFld = Nothing
        pFeat = Nothing
        intCount = Nothing
        intObjId = Nothing
        pBuffFeatSelection = Nothing
        pDataStatistics = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub

    Private Function CalcEmp(ByVal pByFeatLayer As IFeatureLayer2, ByVal pSelLayer As IFeatureLayer2, ByVal strPrecentFld As String, ByVal strSumFld As String, ByVal pFeat As IFeature, ByVal blnIsZero As Boolean) As Boolean

        Dim pQBLayer As ESRI.ArcGIS.Carto.IQueryByLayer
        Dim pFSelOther As IFeatureSelection = Nothing
        Dim pNSelOther As IFeatureSelection = Nothing
        Dim intFeatCount As Integer
        Dim pDataStats As IDataStatistics
        Dim dblEmpTOT As Double = 0
        Dim dblNumeratorEmpTOT As Double = 0
        Dim pCursor As ICursor = Nothing
        Dim pTable As ITable
        Dim strFldPrefix As String

        'DEFINE IF EXSTING OR SCENARIO TOTALS ARE TO BE STORED
        If Me.rdbExisting.Checked Then
            strFldPrefix = "EX_PRE_"
        Else
            strFldPrefix = "TOT_"
        End If

        'FIRST SELECT THE NEIGHBORHOOD FEATURE(S) THAT HAVE THEIR CENTER WITHIN THE SELECTED BUFFER FEATURE
        Try
            pFSelOther = CType(pSelLayer, IFeatureSelection)
            pNSelOther = CType(pNFeatLyr, IFeatureSelection)
            intFeatCount = pFSelOther.SelectionSet.Count
            If Not blnIsZero Then
                pQBLayer = New QueryByLayer
                With pQBLayer
                    .ByLayer = pByFeatLayer
                    .FromLayer = pNSelOther
                    .ResultType = esriSelectionResultEnum.esriSelectionResultNew
                    .BufferDistance = 0
                    .BufferUnits = ESRI.ArcGIS.esriSystem.esriUnits.esriMiles
                    .LayerSelectionMethod = esriLayerSelectionMethod.esriLayerSelectHaveTheirCenterIn
                    .UseSelectedFeatures = True
                    pFSelOther.SelectionSet = .Select
                End With
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, pSelLayer.Name.ToString)
            GoTo CleanUp
        End Try

        intFeatCount = pNSelOther.SelectionSet.Count
        If intFeatCount <= 0 Then
            If pFeat.Fields.FindField(strFldPrefix & strPrecentFld) >= 0 Then
                pFeat.Value(pFeat.Fields.FindField(strFldPrefix & strPrecentFld)) = 0
            End If
            If pFeat.Fields.FindField(strFldPrefix & strSumFld) >= 0 Then
                pFeat.Value(pFeat.Fields.FindField(strFldPrefix & strSumFld)) = 0
            End If
            pFeat.Store()
            GoTo CleanUp
        End If


        Try
            'WRITE THE COUNT TO THE INPUT FEATURE
            pTable = pNFeatLyr.FeatureClass
            intFeatCount = pFSelOther.SelectionSet.Count
            dblEmpTOT = 0
            dblNumeratorEmpTOT = 0
            Try
                If intFeatCount > 0 Then
                    pNSelOther.SelectionSet.Search(Nothing, False, pCursor)
                    pDataStats = New DataStatistics
                    pDataStats.Cursor = pCursor
                    pDataStats.Field = Me.cmbEmpFld.Text
                    dblNumeratorEmpTOT = pDataStats.Statistics.Sum
                    If m_intTotalEmp > 0 Then
                        dblEmpTOT = dblNumeratorEmpTOT / m_intTotalEmp
                    End If
                End If
                If pTable.Fields.FindField(strFldPrefix & strSumFld) >= 0 Then
                    pFeat.Value(pTable.Fields.FindField(strFldPrefix & strSumFld)) = dblNumeratorEmpTOT
                End If
                If pTable.Fields.FindField(strFldPrefix & strPrecentFld) >= 0 Then
                    pFeat.Value(pTable.Fields.FindField(strFldPrefix & strPrecentFld)) = dblEmpTOT
                End If
            Catch ex As Exception
                If pTable.Fields.FindField(strFldPrefix & strPrecentFld) >= 0 Then
                    pFeat.Value(pTable.Fields.FindField(strFldPrefix & strPrecentFld)) = 0
                End If
                If pTable.Fields.FindField(strFldPrefix & strSumFld) >= 0 Then
                    pFeat.Value(pTable.Fields.FindField(strFldPrefix & strSumFld)) = 0
                End If
            End Try

            
            pFeat.Store()
            pFSelOther.Clear()
            GoTo CleanUp

        Catch ex As Exception
            GoTo CleanUp
        End Try

CleanUp:
        pByFeatLayer = Nothing
        pFSelOther = Nothing
        intFeatCount = Nothing
        pDataStats = Nothing
        dblEmpTOT = Nothing
        dblNumeratorEmpTOT = Nothing
        pCursor = Nothing
        pTable = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()

    End Function

    Private Function CalcActDenAndJobPopBal(ByVal pByFeatLayer As IFeatureLayer2, ByVal pSelLayer As IFeatureLayer2, ByVal pFeat As IFeature, ByVal blnIsZero As Boolean, ByVal strType As String) As Boolean
        Dim pQBLayer As ESRI.ArcGIS.Carto.IQueryByLayer
        Dim pFSelOther As IFeatureSelection = Nothing
        Dim intFeatCount As Integer
        Dim dblPercent As Double = 0
        Dim pDataStats As IDataStatistics
        Dim pCursor As ICursor = Nothing
        Dim dblTOT_POP As Double = 0
        Dim dblTOT_EMP As Double = 0
        Dim dblTOT_AREA As Double = 0
        Dim dblACTIVITYVALUE As Double
        Dim dblBALANCEVALUE As Double

        'SELECT THE CORRESPONDING BUFFER FOR THE SELECTED NEIGHBORHOOD FEATURE
        Try
            pFSelOther = CType(pSelLayer, IFeatureSelection)
            intFeatCount = pFSelOther.SelectionSet.Count
            If Not blnIsZero Then
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
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, pSelLayer.Name.ToString)
            GoTo CleanUp
        End Try

        Try
            'WRITE THE COUNT TO THE INPUT FEATURE
            If pFSelOther.SelectionSet.Count > 0 Then
                Try
                    'TOTAL AREA
                    pFSelOther.SelectionSet.Search(Nothing, False, pCursor)
                    pDataStats = New DataStatistics
                    pDataStats.Cursor = pCursor
                    pDataStats.Field = Me.cmbArea.Text
                    dblTOT_AREA = CDbl(pDataStats.Statistics.Sum)
                    If Me.rdbTotal.Checked Then
                        'TOTAL POPULATION
                        pFSelOther.SelectionSet.Search(Nothing, False, pCursor)
                        pDataStats = New DataStatistics
                        pDataStats.Cursor = pCursor
                        pDataStats.Field = "TOT_POP"
                        dblTOT_POP = CDbl(pDataStats.Statistics.Sum)
                        'TOTAL EMPLOYMENT
                        pFSelOther.SelectionSet.Search(Nothing, False, pCursor)
                        pDataStats = New DataStatistics
                        pDataStats.Cursor = pCursor
                        pDataStats.Field = "TOT_EMP"
                        dblTOT_EMP = CDbl(pDataStats.Statistics.Sum)
                    Else
                        'PRE POPULATION
                        If pSelLayer.FeatureClass.FindField("EX_PRE_POP") >= 0 Then
                            pFSelOther.SelectionSet.Search(Nothing, False, pCursor)
                            pDataStats = New DataStatistics
                            pDataStats.Cursor = pCursor
                            pDataStats.Field = "EX_PRE_POP"
                            dblTOT_POP = CDbl(pDataStats.Statistics.Sum)
                        Else
                            dblTOT_POP = 0
                        End If
                        If pSelLayer.FeatureClass.FindField("EX_PRE_EMP") >= 0 Then
                            'PRE EMPLOYMENT
                            pFSelOther.SelectionSet.Search(Nothing, False, pCursor)
                            pDataStats = New DataStatistics
                            pDataStats.Cursor = pCursor
                            pDataStats.Field = "EX_PRE_EMP"
                            dblTOT_EMP = CDbl(pDataStats.Statistics.Sum)
                        Else
                            dblTOT_EMP = 0
                        End If
                    End If

                    'CALCUALTE TOTAL VALUES FOR ACTIVITY DENSITY AND JOBS BALANCE
                    If dblTOT_AREA > 0 Then
                        dblACTIVITYVALUE = (dblTOT_POP + dblTOT_EMP) / dblTOT_AREA
                    Else
                        dblACTIVITYVALUE = 0
                    End If
                    Try
                        dblBALANCEVALUE = 1 - (Abs(dblTOT_EMP - 0.2 * dblTOT_POP) / (dblTOT_EMP + 0.2 * dblTOT_POP))
                    Catch ex As Exception
                        dblBALANCEVALUE = 0
                    End Try

                    If Me.rdbTotal.Checked Then
                        Try
                            If strType = "QUARTER" Then
                                pFeat.Value(pFeat.Fields.FindField("TOT_ActDenQtrMi")) = dblACTIVITYVALUE
                                pFeat.Value(pFeat.Fields.FindField("TOT_JobPopBalQtrMi")) = dblBALANCEVALUE
                            ElseIf strType = "HALF" Then
                                pFeat.Value(pFeat.Fields.FindField("TOT_ActDenHalfMi")) = dblACTIVITYVALUE
                                pFeat.Value(pFeat.Fields.FindField("TOT_JobPopBalHalfMi")) = dblBALANCEVALUE
                            ElseIf strType = "ONE" Then
                                pFeat.Value(pFeat.Fields.FindField("TOT_ActDen1Mi")) = dblACTIVITYVALUE
                                pFeat.Value(pFeat.Fields.FindField("TOT_JobPopBal1Mi")) = dblBALANCEVALUE
                            ElseIf strType = "ZERO" Then
                                pFeat.Value(pFeat.Fields.FindField("TOT_ActDen")) = dblACTIVITYVALUE
                                pFeat.Value(pFeat.Fields.FindField("TOT_JobPopBal")) = dblBALANCEVALUE
                            End If
                        Catch ex As Exception
                        End Try
                    Else
                        Try
                            If strType = "QUARTER" Then
                                pFeat.Value(pFeat.Fields.FindField("EX_PRE_ActDenQtrMi")) = dblACTIVITYVALUE
                                pFeat.Value(pFeat.Fields.FindField("EX_PRE_JobPopBalQtrMi")) = dblBALANCEVALUE
                            ElseIf strType = "HALF" Then
                                pFeat.Value(pFeat.Fields.FindField("EX_PRE_ActDenHalfMi")) = dblACTIVITYVALUE
                                pFeat.Value(pFeat.Fields.FindField("EX_PRE_JobPopBalHalfMi")) = dblBALANCEVALUE
                            ElseIf strType = "ONE" Then
                                pFeat.Value(pFeat.Fields.FindField("EX_PRE_ActDen1Mi")) = dblACTIVITYVALUE
                                pFeat.Value(pFeat.Fields.FindField("EX_PRE_JobPopBal1Mi")) = dblBALANCEVALUE
                            ElseIf strType = "ZERO" Then
                                pFeat.Value(pFeat.Fields.FindField("EX_PRE_ActDen")) = dblACTIVITYVALUE
                                pFeat.Value(pFeat.Fields.FindField("EX_PRE_JobPopBal")) = dblBALANCEVALUE
                            End If
                        Catch ex As Exception
                        End Try
                    End If

                    pFeat.Store()
                    pFSelOther.Clear()
                Catch ex As Exception
                    'MessageBox.Show(ex.Message)
                End Try
            End If
            GoTo CleanUp
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
        dblPercent = Nothing
        pDataStats = Nothing
        pCursor = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()

    End Function











    Private Sub CalcPercentage(ByVal pByFeatLayer As IFeatureLayer2, ByVal pSelLayer As IFeatureLayer2, ByVal strFld As String, ByVal pFeat As IFeature, ByVal blnIsZero As Boolean)

        Dim pNSelOther As IFeatureSelection = Nothing
        Dim p4WaySelOther As IFeatureSelection = Nothing
        Dim pIntSelOther As IFeatureSelection = Nothing
        Dim pQBLayer As ESRI.ArcGIS.Carto.IQueryByLayer
        Dim intFeatCount As Integer
        Dim pDataStats As IDataStatistics
        Dim pCursor As ICursor = Nothing
        Dim intNumeratorCount As Integer
        Dim intDenominatorCount As Integer
        Dim dblPercent As Double = 0


        'RETRIEVE THE INTERSECTION LAYER
        If pInersectionsLyr Is Nothing Then
            pFeat.Value(pFeat.Fields.FindField(strFld)) = 0
            pFeat.Store()
            GoTo CleanUp
        Else
            pIntSelOther = CType(pInersectionsLyr, IFeatureSelection)
        End If


        'FIRST SELECT THE NEIGHBORHOOD FEATURE(S) THAT HAVE THEIR CENTER WITHIN THE SELECTED BUFFER FEATURE, NOT REQUIRED IF NEIGHBORHOOD FEATURE ONLY
        pNSelOther = CType(pNFeatLyr, IFeatureSelection)
        If Not blnIsZero Then
            Try
                pQBLayer = New QueryByLayer
                With pQBLayer
                    .ByLayer = pByFeatLayer
                    .FromLayer = pNSelOther
                    .ResultType = esriSelectionResultEnum.esriSelectionResultNew
                    .BufferDistance = 0
                    .BufferUnits = ESRI.ArcGIS.esriSystem.esriUnits.esriMiles
                    .LayerSelectionMethod = esriLayerSelectionMethod.esriLayerSelectHaveTheirCenterIn
                    .UseSelectedFeatures = True
                    pNSelOther.SelectionSet = .Select
                End With
            Catch ex As Exception
                'MessageBox.Show(ex.Message, pSelLayer.Name.ToString)
                GoTo CleanUp
            End Try
        End If

        'MAKE SURE AT LEAST ONE FEATURE IS SELECTED
        intFeatCount = pNSelOther.SelectionSet.Count
        If intFeatCount <= 0 Then
            pFeat.Value(pFeat.Fields.FindField(strFld)) = 0
            pFeat.Store()
            GoTo CleanUp
        End If

        'SELECT THE 4-WAY I FEATURES WITHIN THE NEIGHBORHOOD SELECTION SET
        Try
            p4WaySelOther = CType(p4WayLyr, IFeatureSelection)
            pQBLayer = New QueryByLayer
            With pQBLayer
                .ByLayer = pNSelOther
                .FromLayer = p4WaySelOther
                .ResultType = esriSelectionResultEnum.esriSelectionResultNew
                .BufferDistance = 0
                .BufferUnits = ESRI.ArcGIS.esriSystem.esriUnits.esriMiles
                .LayerSelectionMethod = esriLayerSelectionMethod.esriLayerSelectHaveTheirCenterIn
                .UseSelectedFeatures = True
                p4WaySelOther.SelectionSet = .Select
            End With
        Catch ex As Exception
            'MessageBox.Show(ex.Message, pSelLayer.Name.ToString)
            GoTo CleanUp
        End Try

        'SELECT THE TOTAL INTERSECTION FEATURES WITHIN THE NEIGHBORHOOD SELECTION SET
        Try
            pQBLayer = New QueryByLayer
            With pQBLayer
                .ByLayer = pNSelOther
                .FromLayer = pIntSelOther
                .ResultType = esriSelectionResultEnum.esriSelectionResultNew
                .BufferDistance = 0
                .BufferUnits = ESRI.ArcGIS.esriSystem.esriUnits.esriMiles
                .LayerSelectionMethod = esriLayerSelectionMethod.esriLayerSelectHaveTheirCenterIn
                .UseSelectedFeatures = True
                pIntSelOther.SelectionSet = .Select
            End With
        Catch ex As Exception
            'MessageBox.Show(ex.Message, pSelLayer.Name.ToString)
            GoTo CleanUp
        End Try

        Try
            'WRITE THE PERCENTAGE TO THE INPUT FEATURE
            intNumeratorCount = p4WaySelOther.SelectionSet.Count
            intDenominatorCount = pIntSelOther.SelectionSet.Count
            If intNumeratorCount > 0 And intDenominatorCount > 0 Then
                dblPercent = intNumeratorCount / intDenominatorCount
            Else
                dblPercent = 0
            End If
            pFeat.Value(pFeat.Fields.FindField(strFld)) = dblPercent
            pFeat.Store()
            p4WaySelOther.Clear()
            pIntSelOther.Clear()
            pNSelOther.Clear()
            GoTo CleanUp

        Catch ex As Exception
            GoTo CleanUp
        End Try

CleanUp:
        pNSelOther = Nothing
        p4WaySelOther = Nothing
        pIntSelOther = Nothing
        pQBLayer = Nothing
        intFeatCount = Nothing
        pDataStats = Nothing
        pCursor = Nothing
        intNumeratorCount = Nothing
        intDenominatorCount = Nothing
        dblPercent = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub

    Private Sub CalcDensity(ByVal pByFeatLayer As IFeatureLayer2, ByVal pSelLayer As IFeatureLayer2, ByVal strFld As String, ByVal pFeat As IFeature, ByVal blnIsZero As Boolean)
        Dim pNSelOther As IFeatureSelection = Nothing
        Dim pFSelOther As IFeatureSelection = Nothing
        Dim pQBLayer As ESRI.ArcGIS.Carto.IQueryByLayer
        Dim intFeatCount As Integer
        Dim pDataStats As IDataStatistics
        Dim pCursor As ICursor = Nothing
        Dim dblAreaTOT As Double = 0
        Dim dblDensityTOT As Double = 0
        Dim strFldPrefix As String = "TOT_"

        'DEFINE THE WRITE FIELD PREFIX
        If Me.rdbExisting.Checked Then
            strFldPrefix = "EX_PRE_"
        End If

        'FIRST SELECT THE NEIGHBORHOOD FEATURE(S) THAT HAVE THEIR CENTER WITHIN THE SELECTED BUFFER FEATURE, NOT REQUIRED IF NEIGHBORHOOD FEATURE ONLY
        pNSelOther = CType(pNFeatLyr, IFeatureSelection)
        If Not blnIsZero Then
            Try
                pFSelOther = CType(pSelLayer, IFeatureSelection)
                intFeatCount = pFSelOther.SelectionSet.Count
                pQBLayer = New QueryByLayer
                With pQBLayer
                    .ByLayer = pByFeatLayer
                    .FromLayer = pNSelOther
                    .ResultType = esriSelectionResultEnum.esriSelectionResultNew
                    .BufferDistance = 0
                    .BufferUnits = ESRI.ArcGIS.esriSystem.esriUnits.esriMiles
                    .LayerSelectionMethod = esriLayerSelectionMethod.esriLayerSelectHaveTheirCenterIn
                    .UseSelectedFeatures = True
                    pNSelOther.SelectionSet = .Select
                End With
            Catch ex As Exception
                'MessageBox.Show(ex.Message, pSelLayer.Name.ToString)
                GoTo CleanUp
            End Try
        End If

        'MAKE SURE AT LEAST ONE FEATURE IS SELECTED
        intFeatCount = pNSelOther.SelectionSet.Count
        'If intFeatCount <= 0 Then
        '    GoTo CleanUp
        'End If

        'SELECT THE POINT FEATURES WITHIN THE NEIGHBORHOOD SELECTION SET
        If intFeatCount > 0 Then
            Try
                pFSelOther = CType(pSelLayer, IFeatureSelection)
                pQBLayer = New QueryByLayer
                With pQBLayer
                    .ByLayer = pNSelOther
                    .FromLayer = pSelLayer
                    .ResultType = esriSelectionResultEnum.esriSelectionResultNew
                    .BufferDistance = 0
                    .BufferUnits = ESRI.ArcGIS.esriSystem.esriUnits.esriMiles
                    .LayerSelectionMethod = esriLayerSelectionMethod.esriLayerSelectHaveTheirCenterIn
                    .UseSelectedFeatures = True
                    pFSelOther.SelectionSet = .Select
                End With
            Catch ex As Exception
                'MessageBox.Show(ex.Message, pSelLayer.Name.ToString)
                GoTo CleanUp
            End Try
        End If


        Try
            'WRITE THE COUNT TO THE INPUT FEATURE
            intFeatCount = pFSelOther.SelectionSet.Count

            'WRITE THE DENSITY TO THE INPUT FEATURE
            If strFld.Length > 0 And Me.cmbArea.Text.Length Then
                If pFeat.Fields.FindField(strFldPrefix & strFld) > -1 And pFeat.Fields.FindField(Me.cmbArea.Text) > -1 Then
                    If pFSelOther.SelectionSet.Count <= 0 Then
                        pFeat.Value(pFeat.Fields.FindField(strFldPrefix & strFld)) = 0
                    Else
                        Try
                            ' MessageBox.Show(pNSelOther.SelectionSet.Count.ToString, pNFeatLyr.FeatureClass.FindField(Me.cmbArea.Text).ToString)
                            pNSelOther.SelectionSet.Search(Nothing, False, pCursor)
                            pDataStats = New DataStatistics
                            pDataStats.Cursor = pCursor
                            pDataStats.Field = Me.cmbArea.Text
                            dblAreaTOT = pDataStats.Statistics.Sum
                            dblDensityTOT = intFeatCount / dblAreaTOT
                            pFeat.Value(pFeat.Fields.FindField(strFldPrefix & strFld)) = dblDensityTOT
                        Catch ex As Exception
                            ' MessageBox.Show(ex.Message)
                        End Try
                    End If
                End If
            End If

            pFeat.Store()
            pFSelOther.Clear()
            GoTo CleanUp

        Catch ex As Exception
            GoTo CleanUp
        End Try

CleanUp:
        pByFeatLayer = Nothing
        pFSelOther = Nothing
        intFeatCount = Nothing
        pDataStats = Nothing
        dblAreaTOT = Nothing
        dblDensityTOT = Nothing
        pCursor = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub

    Private Sub cmbEmpFld_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbEmpFld.SelectedIndexChanged
        'CALCULATE THE TOTAL EMPLOYMENT VALUE
        Dim pNFeatLyr As IFeatureLayer
        Dim pNFeatClass As IFeatureClass
        Dim pFeatSelection As IFeatureSelection
        Dim pDataStatistics As DataStatistics
        Dim pCursor As ICursor
        Dim pTable As ITable

        Try
            pNFeatLyr = arrPolyLyrs.Item(Me.cmbLayers.SelectedIndex)
            pNFeatClass = pNFeatLyr.FeatureClass
            pTable = CType(pNFeatClass, ITable)
        Catch ex As Exception
            MessageBox.Show("Error in retrieving the selected neighborhood layer." & vbNewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo CleanUp
        End Try


CLeanUp:
        pNFeatLyr = Nothing
        pNFeatClass = Nothing
        pFeatSelection = Nothing
        pDataStatistics = Nothing
        pCursor = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub

End Class