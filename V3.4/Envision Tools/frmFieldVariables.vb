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

Public Class frmFieldVariables

    Private blnStatusUpdate As Boolean = False

    Private Sub btnCheckAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCheckAll.Click
        CheckStatus(1)
    End Sub

    Private Sub btnUncheckAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUncheckAll.Click
        CheckStatus(0)
    End Sub

    Private Sub CheckStatus(ByVal intCheck As Integer)
        'SET STATUS OF THE OPTIONS
        Dim iRow As Integer
        For iRow = 0 To Me.dgvFields.RowCount - 1
            m_frmAttribManager.dgvFields.Rows(iRow).Cells(0).Value = intCheck
        Next
    End Sub

    Private Sub btnApply_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnApply.Click
        'OPEN THE ENVISION FIELD TRACKING TABLE AND LOAD OPTIONS
        'FIELD TRACKING TABLE
        Dim rowTemp As IRow
        Dim intUse As Integer
        Dim strFieldName1 As String = ""
        Dim strFieldAlias1 As String = ""
        Dim strFieldName2 As String = ""
        Dim strFieldAlias2 As String = ""
        Dim strSource As String = ""
        Dim intCalcByAcres As Integer = 0
        Dim intCalcByDepVar As Integer = 0
        Dim strCalcFieldName As String = ""
        Dim strDepVarFieldName As String = ""
        Dim intCalcByAcresOnly As Integer
        Dim pCursor As ICursor
        Dim iRow As Integer
        Dim pFeatClass As IFeatureClass = m_pEditFeatureLyr.FeatureClass
        Dim pTable As ITable = CType(pFeatClass, ITable)


        If Not m_dockEnvisionWinForm.btnEditing.Text = "Start Edit" Then
            MessageBox.Show("Please end the current edit session before applying attribute tracking.  Require fields can not be added during an edit session.", "Edit Session Found", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo CleanUp
        End If
        Try
            'CLEAR THE EXISTING FIELD ARRAY LIST 
            m_arrWriteDevTypeFields.Clear()
            m_arrWriteDevTypeAcresFields.Clear()
            m_arrWriteDevTypeAcresFieldsOnly.Clear()
            m_arrWriteDevTypeAcresFieldsAltName.Clear()

            'CYCLE THROUGH FIELD LIST
            For iRow = 0 To Me.dgvFields.RowCount - 1
                intUse = CInt(Me.dgvFields.Rows(iRow).Cells(0).Value)
                strFieldName2 = Me.dgvFields.Rows(iRow).Cells(1).Value
                strFieldAlias2 = Me.dgvFields.Rows(iRow).Cells(2).Value
                intCalcByAcres = CInt(Me.dgvFields.Rows(iRow).Cells(3).Value)
                intCalcByDepVar = CInt(Me.dgvFields.Rows(iRow).Cells(4).Value)
                strCalcFieldName = CStr(Me.dgvFields.Rows(iRow).Cells(5).Value)
                intCalcByAcresOnly = CInt(Me.dgvFields.Rows(iRow).Cells(6).Value)
                strDepVarFieldName = CStr(Me.dgvFields.Rows(iRow).Cells(7).Value)
                If intUse = -1 Then
                    intUse = 1
                End If

                'CYCLE THROUGH DEV ATTRIBUTE FIELD TABLE TO WRITE VALUES
                pCursor = m_tblAttribFields.Search(Nothing, False)
                rowTemp = pCursor.NextRow
                Do Until rowTemp Is Nothing
                    strFieldName1 = ""
                    strFieldAlias1 = ""
                    Try
                        strFieldName1 = CStr(rowTemp.Value(rowTemp.Fields.FindField("FIELD_NAME")))
                    Catch ex As Exception
                    End Try
                    Try
                        strFieldAlias1 = CStr(rowTemp.Value(rowTemp.Fields.FindField("FIELD_ALIAS")))
                    Catch ex As Exception
                    End Try
                    'SET THE CALC ACRES TO TRUE IF THE ACRES ONLY IS CHECKED.
                    If intCalcByAcresOnly = 1 And intCalcByAcres = 0 Then
                        intCalcByAcres = 1
                    End If

                    If strFieldAlias1 = strFieldAlias2 And strFieldName1 = strFieldName2 Then
                        rowTemp.Value(rowTemp.Fields.FindField("USE")) = intUse
                        If intCalcByAcres = 1 Then
                            rowTemp.Value(rowTemp.Fields.FindField("CALC_BY_ACRES")) = 1
                        End If
                        If intCalcByDepVar = 1 Then
                            rowTemp.Value(rowTemp.Fields.FindField("CALC_BY_ACRES")) = 2
                        End If
                        rowTemp.Value(rowTemp.Fields.FindField("CALC_ONLY")) = intCalcByAcresOnly
                        rowTemp.Store()
                        Exit Do
                    End If
                    rowTemp = pCursor.NextRow
                Loop
                Me.prgStatus.Value = (iRow / Me.dgvFields.RowCount) * 100
            Next
            'EXECUTE FUNTION TO CHECK ENVISION LAYER FOR REQUIRED FIELD(S)
            EnvisionLyrRequiredFieldsCheck(pTable)
            m_frmAttribManager.WindowState = FormWindowState.Minimized
            MessageBox.Show("Settings were successfully applied.", "Envision Field Management", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            m_frmAttribManager.Close()
            GoTo CleanUp
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "Envision Field Management Error", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo CleanUp
        End Try


CleanUp:
        strFieldName1 = Nothing
        strFieldAlias1 = Nothing
        strFieldName2 = Nothing
        strFieldAlias2 = Nothing
        strSource = Nothing
        intUse = Nothing
        pCursor = Nothing
        iRow = Nothing
        pTable = Nothing
        pFeatClass = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()

    End Sub

    Private Sub frmFieldVariables_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        m_frmAttribManager.WindowState = Windows.Forms.FormWindowState.Minimized
        m_frmAttribManager = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub

    Private Sub frmFieldVariables_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'OPEN THE ENVISION FIELD TRACKING TABLE AND LOAD OPTIONS
        'FIELD TRACKING TABLE
        Dim GP As Geoprocessor
        Dim pCreateTable As ESRI.ArcGIS.DataManagementTools.CreateTable
        Dim pWksFactory As ESRI.ArcGIS.Geodatabase.IWorkspaceFactory = Nothing
        Dim pFeatWks As ESRI.ArcGIS.Geodatabase.IFeatureWorkspace = Nothing
        Dim rowTemp As IRow
        Dim strFieldName As String = ""
        Dim strFieldAlias As String = ""
        Dim intUseField As Integer = 0
        Dim intCalcAcres As Integer = 0
        Dim intCalcByDepVar As Integer = 0
        Dim strCalcDepVar As String = ""
        Dim strCalcFieldName As String = ""
        Dim intCalcAcresOnly As Integer = 0
        Dim pCursor As ICursor

        'CLEAR OUT PREVIOUS FIELD LISTING
        m_frmAttribManager.dgvFields.Rows.Clear()
        blnStatusUpdate = True

        'LOAD THE DEVELOPMENT TYPE ATTRIBUTES TABLE
        If m_tblAttribFields Is Nothing Then
            Try
                pWksFactory = New FileGDBWorkspaceFactory
                pFeatWks = pWksFactory.OpenFromFile(m_strFeaturePath, 0)
                m_tblAttribFields = pFeatWks.OpenTable("ENVISION_DEVTYPE_FIELD_TRACKING")
            Catch ex As Exception
                m_strProcessing = m_strProcessing & "Missing table, ENVISION_DEVTYPE_FIELD_TRACKING, executing geoprocessing tool to create table: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
                GP = New Geoprocessor
                pCreateTable = New ESRI.ArcGIS.DataManagementTools.CreateTable
                pCreateTable.out_name = "ENVISION_DEVTYPE_FIELD_TRACKING"
                pCreateTable.out_path = m_strFeaturePath
                RunTool(GP, pCreateTable)
                m_tblAttribFields = pFeatWks.OpenTable("ENVISION_DEVTYPE_FIELD_TRACKING")
            End Try
            If TypeOf m_tblAttribFields Is ITable Then
                LookUpTablesEnvisionAttributeFieldTrackingTblCheck(m_tblAttribFields)
            End If
        End If

        pCursor = m_tblAttribFields.Search(Nothing, False)
        rowTemp = pCursor.NextRow
        Do Until rowTemp Is Nothing
            strFieldName = ""
            strFieldAlias = ""
            intUseField = 0
            intCalcAcres = 0
            intCalcByDepVar = 0
            intCalcAcresOnly = 0
            strCalcDepVar = ""
            strCalcFieldName = ""
            Try
                strFieldName = CStr(rowTemp.Value(rowTemp.Fields.FindField("FIELD_NAME")))
            Catch ex As Exception
            End Try
            Try
                strFieldAlias = CStr(rowTemp.Value(rowTemp.Fields.FindField("FIELD_ALIAS")))
            Catch ex As Exception
            End Try
            Try
                intUseField = CInt(rowTemp.Value(rowTemp.Fields.FindField("USE")))
            Catch ex As Exception
            End Try
            Try
                intCalcAcres = CInt(rowTemp.Value(rowTemp.Fields.FindField("CALC_BY_ACRES")))
            Catch ex As Exception
            End Try
            Try
                strCalcDepVar = CStr(rowTemp.Value(rowTemp.Fields.FindField("DEP_VAR_FIELD_NAME")))
            Catch ex As Exception
            End Try
            Try
                strCalcFieldName = CStr(rowTemp.Value(rowTemp.Fields.FindField("CALC_FIELD_NAME")))
            Catch ex As Exception
            End Try
            Try
                intCalcAcresOnly = CInt(rowTemp.Value(rowTemp.Fields.FindField("CALC_ONLY")))
            Catch ex As Exception
            End Try
            If strFieldName.Length > 0 Then
                If Not UCase(strFieldName) = "DEV_TYPE" And Not UCase(strFieldName) = "VAC_ACRES" And Not UCase(strFieldName) = "DEVD_ACRES" _
                    And Not UCase(strFieldName) = "CONSTRAINED_ACRE" And Not UCase(strFieldName) = "RED" And Not UCase(strFieldName) = "GREEN" _
                    And Not UCase(strFieldName) = "BLUE" Then
                    m_frmAttribManager.dgvFields.Rows.Add()
                    m_frmAttribManager.dgvFields.Rows(m_frmAttribManager.dgvFields.RowCount - 1).Cells(0).Value = intUseField
                    m_frmAttribManager.dgvFields.Rows(m_frmAttribManager.dgvFields.RowCount - 1).Cells(1).Value = strFieldName
                    m_frmAttribManager.dgvFields.Rows(m_frmAttribManager.dgvFields.RowCount - 1).Cells(2).Value = strFieldAlias
                    If intCalcAcres = 1 Then
                        m_frmAttribManager.dgvFields.Rows(m_frmAttribManager.dgvFields.RowCount - 1).Cells(3).Value = 1
                        m_frmAttribManager.dgvFields.Rows(m_frmAttribManager.dgvFields.RowCount - 1).Cells(4).Value = 0
                    End If
                    If intCalcAcres = 2 Then
                        m_frmAttribManager.dgvFields.Rows(m_frmAttribManager.dgvFields.RowCount - 1).Cells(3).Value = 0
                        m_frmAttribManager.dgvFields.Rows(m_frmAttribManager.dgvFields.RowCount - 1).Cells(4).Value = 1
                    End If
                    m_frmAttribManager.dgvFields.Rows(m_frmAttribManager.dgvFields.RowCount - 1).Cells(5).Value = strCalcFieldName
                    m_frmAttribManager.dgvFields.Rows(m_frmAttribManager.dgvFields.RowCount - 1).Cells(6).Value = intCalcAcresOnly
                    m_frmAttribManager.dgvFields.Rows(m_frmAttribManager.dgvFields.RowCount - 1).Cells(7).Value = strCalcDepVar
                End If
            End If
            rowTemp = pCursor.NextRow
        Loop

        blnStatusUpdate = False
CleanUp:
        GP = Nothing
        pCreateTable = Nothing
        pWksFactory = Nothing
        pFeatWks = Nothing
        rowTemp = Nothing
        strFieldName = Nothing
        strFieldAlias = Nothing
        intUseField = Nothing
        intCalcAcres = Nothing
        strCalcDepVar = Nothing
        strCalcFieldName = Nothing
        intCalcAcresOnly = Nothing
        pCursor = Nothing
        GC.Collect()
        'GC.WaitForPendingFinalizers()
    End Sub

    Private Sub chkLandUse_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkLandUse.CheckedChanged
        'SET THE CHECKED STATUS OF THE LAND USE ATRIBUTE FIELDS
        'SET STATUS OF THE OPTIONS
        Dim iRow As Integer
        Dim arrTemp As ArrayList = New ArrayList

        'VARIABLE TO HALT CELL CONTENT CLICK SUB 
        blnStatusUpdate = True

        'BUILD ARRAY OF ATTRIBUTE FIELD NAMES
        arrTemp.Add("HU")
        arrTemp.Add("EMP")
        arrTemp.Add("SF")
        arrTemp.Add("TH")
        arrTemp.Add("MF")
        arrTemp.Add("MH")
        arrTemp.Add("RET")
        arrTemp.Add("OFF")
        arrTemp.Add("IND")
        arrTemp.Add("PUB")
        arrTemp.Add("EDU")
        arrTemp.Add("HOTEL")
        Try
            For iRow = 0 To m_frmAttribManager.dgvFields.RowCount - 1
                If arrTemp.Contains(m_frmAttribManager.dgvFields.Rows(iRow).Cells(1).Value) Or arrTemp.Contains(m_frmAttribManager.dgvFields.Rows(iRow).Cells(4).Value) Then
                    m_frmAttribManager.dgvFields.Rows(iRow).Cells(0).Value = m_frmAttribManager.chkLandUse.Checked
                End If
            Next
        Catch ex As Exception
        End Try

       blnStatusUpdate = False
    End Sub

End Class