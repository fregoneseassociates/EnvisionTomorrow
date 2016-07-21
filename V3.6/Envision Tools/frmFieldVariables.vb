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
    Dim strCalcFieldName As String = ""
    Dim intCalcByAcresOnly As Integer
    Dim calcTotFieldIndex As Int32
    Dim totalFieldNameIndex As Int32
    Dim calcTotal As Int16
    Dim totFldName As String
    Dim pCursor As ICursor
    Dim iRow As Integer
    Dim pFeatClass As IFeatureClass = m_pEditFeatureLyr.FeatureClass
    Dim pTable As ITable = CType(pFeatClass, ITable)


    If Not m_dockEnvisionWinForm.btnEditing.Text = "Start Edit" Then
      MessageBox.Show(Me, "Please end the current edit session before applying attribute tracking.  Require fields can not be added during an edit session.", "Edit Session Found", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1)
      GoTo CleanUp
    End If
    Try
      'CLEAR THE EXISTING FIELD ARRAY LIST 
      m_arrWriteDevTypeFields.Clear()
      m_arrWriteDevTypeAcresFields.Clear()
      m_arrWriteDevTypeAcresFieldsOnly.Clear()
      m_arrWriteDevTypeAcresFieldsAltName.Clear()
      m_arrWriteDevTypeTotalFields.Clear()
      m_arrWriteDevTypeTotalFieldName.Clear()

      'CYCLE THROUGH FIELD LIST
      calcTotFieldIndex = m_tblAttribFields.FindField("Calc_Total")
      totalFieldNameIndex = m_tblAttribFields.FindField("Total_Field_Name")
      For iRow = 0 To Me.dgvFields.RowCount - 1
        intUse = CInt(Me.dgvFields.Rows(iRow).Cells(0).Value)
        strFieldName2 = Me.dgvFields.Rows(iRow).Cells(1).Value
        strFieldAlias2 = Me.dgvFields.Rows(iRow).Cells(2).Value
        strCalcFieldName = CStr(Me.dgvFields.Rows(iRow).Cells(3).Value)
        totFldName = CStr(Me.dgvFields.Rows(iRow).Cells(4).Value)
        calcTotal = CInt(Me.dgvFields.Rows(iRow).Cells(5).Value)
        intCalcByAcresOnly = CInt(Me.dgvFields.Rows(iRow).Cells(6).Value)
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

          If strFieldAlias1 = strFieldAlias2 And strFieldName1 = strFieldName2 Then
            rowTemp.Value(rowTemp.Fields.FindField("USE")) = intUse
            rowTemp.Value(rowTemp.Fields.FindField("CALC_ONLY")) = intCalcByAcresOnly
            rowTemp.Value(totalFieldNameIndex) = totFldName
            rowTemp.Value(calcTotFieldIndex) = calcTotal
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
      MessageBox.Show(Me, "Settings were successfully applied.", "Envision Field Management", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
      m_frmAttribManager.Close()
      GoTo CleanUp
    Catch ex As Exception
      MessageBox.Show(Me, "Error: " & ex.Message, "Envision Field Management Error", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
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
    Dim GP As ESRI.ArcGIS.Geoprocessor.Geoprocessor
    Dim pCreateTable As ESRI.ArcGIS.DataManagementTools.CreateTable
    Dim pWksFactory As ESRI.ArcGIS.Geodatabase.IWorkspaceFactory = Nothing
    Dim pFeatWks As ESRI.ArcGIS.Geodatabase.IFeatureWorkspace = Nothing
    Dim rowTemp As IRow
    Dim strFieldName As String = ""
    Dim strFieldAlias As String = ""
    Dim intUseField As Integer = 0
    Dim strCalcFieldName As String = ""
    Dim intCalcAcresOnly As Integer = 0
    Dim intCalcAcres As Int16
    Dim totFldName As String
    Dim calcTotalFldIndex As Int32
    Dim calcTotal As Int16
    Dim pCursor As ICursor
    Dim exFldName As String

    'CLEAR OUT PREVIOUS FIELD LISTING
    m_frmAttribManager.dgvFields.Rows.Clear()

    'LOAD THE DEVELOPMENT TYPE ATTRIBUTES TABLE
    If m_tblAttribFields Is Nothing Then
      Try
        pWksFactory = New FileGDBWorkspaceFactory
        pFeatWks = pWksFactory.OpenFromFile(m_strFeaturePath, 0)
        m_tblAttribFields = pFeatWks.OpenTable("ENVISION_DEVTYPE_FIELD_TRACKING")
      Catch ex As Exception
        m_strProcessing = m_strProcessing & "Missing table, ENVISION_DEVTYPE_FIELD_TRACKING, executing geoprocessing tool to create table: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
        GP = New ESRI.ArcGIS.Geoprocessor.Geoprocessor
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

    calcTotalFldIndex = m_tblAttribFields.FindField("Calc_Total")
    pCursor = m_tblAttribFields.Search(Nothing, False)
    rowTemp = pCursor.NextRow
    Do Until rowTemp Is Nothing
      strFieldName = ""
      strFieldAlias = ""
      intUseField = 0
      intCalcAcresOnly = 0
      intCalcAcres = 0
      calcTotal = 0
      strCalcFieldName = ""
      totFldName = String.Empty
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
        strCalcFieldName = CStr(rowTemp.Value(rowTemp.Fields.FindField("CALC_FIELD_NAME")))
      Catch ex As Exception
      End Try
      Try
        intCalcAcresOnly = CInt(rowTemp.Value(rowTemp.Fields.FindField("CALC_ONLY")))
      Catch ex As Exception
      End Try
      Try
        intCalcAcres = CInt(rowTemp.Value(rowTemp.Fields.FindField("CALC_BY_ACRES")))
      Catch ex As Exception
      End Try
      Try
        totFldName = rowTemp.Value(rowTemp.Fields.FindField("Total_Field_Name")).ToString
      Catch ex As Exception
      End Try
      If calcTotalFldIndex <> -1 Then
        If rowTemp.Value(calcTotalFldIndex) IsNot DBNull.Value Then calcTotal = CInt(rowTemp.Value(calcTotalFldIndex))
      End If
      If strFieldName.Length > 0 Then
        If Not UCase(strFieldName) = "DEV_TYPE" And Not UCase(strFieldName) = "VAC_ACRES" And Not UCase(strFieldName) = "DEVD_ACRES" _
            And Not UCase(strFieldName) = "CONSTRAINED_ACRE" And Not UCase(strFieldName) = "RED" And Not UCase(strFieldName) = "GREEN" _
            And Not UCase(strFieldName) = "BLUE" Then
          m_frmAttribManager.dgvFields.Rows.Add()
          m_frmAttribManager.dgvFields.Rows(m_frmAttribManager.dgvFields.RowCount - 1).Cells(0).Value = intUseField
          m_frmAttribManager.dgvFields.Rows(m_frmAttribManager.dgvFields.RowCount - 1).Cells(1).Value = strFieldName
          m_frmAttribManager.dgvFields.Rows(m_frmAttribManager.dgvFields.RowCount - 1).Cells(2).Value = strFieldAlias
          m_frmAttribManager.dgvFields.Rows(m_frmAttribManager.dgvFields.RowCount - 1).Cells(3).Value = strCalcFieldName
          m_frmAttribManager.dgvFields.Rows(m_frmAttribManager.dgvFields.RowCount - 1).Cells(5).Value = calcTotal
          m_frmAttribManager.dgvFields.Rows(m_frmAttribManager.dgvFields.RowCount - 1).Cells(6).Value = intCalcAcresOnly
          If intCalcAcres Then
            exFldName = strCalcFieldName
          Else
            exFldName = strFieldName
          End If
          If m_pEditFeatureLyr.FeatureClass.Fields.FindField("EX_" & exFldName) = -1 Then
            With m_frmAttribManager.dgvFields.Rows(m_frmAttribManager.dgvFields.RowCount - 1).Cells(5)
              .ReadOnly = True
              .Style.BackColor = Color.LightGray
              .Style.ForeColor = Color.DarkGray
            End With
          Else
            If totFldName = String.Empty Then
              If intCalcAcres Then
                totFldName = "TOT_" & strCalcFieldName
              Else
                totFldName = "TOT_" & strFieldName
              End If
            End If
          End If
          m_frmAttribManager.dgvFields.Rows(m_frmAttribManager.dgvFields.RowCount - 1).Cells(4).Value = totFldName
        End If
      End If
      rowTemp = pCursor.NextRow
    Loop

CleanUp:
    GP = Nothing
    pCreateTable = Nothing
    pWksFactory = Nothing
    pFeatWks = Nothing
    rowTemp = Nothing
    strFieldName = Nothing
    strFieldAlias = Nothing
    intUseField = Nothing
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
    Dim arrLandUse As ArrayList = New ArrayList

    'BUILD ARRAY OF ATTRIBUTE FIELD NAMES
    arrLandUse.Add("HU")
    arrLandUse.Add("EMP")
    arrLandUse.Add("SF")
    arrLandUse.Add("TH")
    arrLandUse.Add("MF")
    arrLandUse.Add("MH")
    arrLandUse.Add("RET")
    arrLandUse.Add("OFF")
    arrLandUse.Add("IND")
    arrLandUse.Add("PUB")
    arrLandUse.Add("EDU")
    arrLandUse.Add("HOTEL")
    Try
      For iRow = 0 To m_frmAttribManager.dgvFields.RowCount - 1
        If arrLandUse.Contains(m_frmAttribManager.dgvFields.Rows(iRow).Cells(1).Value) Or arrLandUse.Contains(m_frmAttribManager.dgvFields.Rows(iRow).Cells(3).Value) Then
          m_frmAttribManager.dgvFields.Rows(iRow).Cells(0).Value = m_frmAttribManager.chkLandUse.Checked
        End If
      Next
    Catch ex As Exception
    End Try
  End Sub

  Private Sub chkTravelBehavior_CheckedChanged(sender As Object, e As EventArgs) Handles chkTravelBehavior.CheckedChanged
    'SET THE CHECKED STATUS OF THE TRAVEL BEHAVIOR ATTRIBUTE FIELDS
    'SET STATUS OF THE OPTIONS
    Dim iRow As Integer
    Dim arrTemp As ArrayList = New ArrayList

    'BUILD ARRAY OF ATTRIBUTE FIELD NAMES
    arrTemp.Add("HH")
    arrTemp.Add("HU")
    arrTemp.Add("EMP")
    arrTemp.Add("POP")
    arrTemp.Add("AVG_HH_SIZE")
    arrTemp.Add("AVG_HH_INCOME")
    arrTemp.Add("AVG_HH_WORKERS")
    Try
      For iRow = 0 To m_frmAttribManager.dgvFields.RowCount - 1
        If arrTemp.Contains(m_frmAttribManager.dgvFields.Rows(iRow).Cells(1).Value) Or arrTemp.Contains(m_frmAttribManager.dgvFields.Rows(iRow).Cells(3).Value) Then
          m_frmAttribManager.dgvFields.Rows(iRow).Cells(0).Value = m_frmAttribManager.chkTravelBehavior.Checked
        End If
      Next
    Catch ex As Exception
    End Try
  End Sub

  Private Sub chkPublicHealth_CheckedChanged(sender As Object, e As EventArgs) Handles chkPublicHealth.CheckedChanged
    'SET THE CHECKED STATUS OF THE PUBLIC HEALTH ATTRIBUTE FIELDS
    'SET STATUS OF THE OPTIONS
    Dim iRow As Integer
    Dim arrTemp As ArrayList = New ArrayList

    'BUILD ARRAY OF ATTRIBUTE FIELD NAMES
    arrTemp.Add("POP")
    arrTemp.Add("HH")
    arrTemp.Add("EMP")
    arrTemp.Add("WORKERS")
    arrTemp.Add("ppl_acre")
    arrTemp.Add("NetEMPDen")
    arrTemp.Add("RET")
    arrTemp.Add("EMP_MIX")
    arrTemp.Add("Int_Den_Mi")
    arrTemp.Add("AVG_HH_SIZE")
    arrTemp.Add("Pct_Med_HH_INC")
    arrTemp.Add("Pct_Low_HH_INC")
    arrTemp.Add("Pct_Hi_HH_INC")
    Try
      For iRow = 0 To m_frmAttribManager.dgvFields.RowCount - 1
        If arrTemp.Contains(m_frmAttribManager.dgvFields.Rows(iRow).Cells(1).Value) Or arrTemp.Contains(m_frmAttribManager.dgvFields.Rows(iRow).Cells(3).Value) Then
          m_frmAttribManager.dgvFields.Rows(iRow).Cells(0).Value = m_frmAttribManager.chkPublicHealth.Checked
        End If
      Next
    Catch ex As Exception
    End Try
  End Sub
End Class