Option Explicit On
Option Strict On

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


Module modEnvisionToolbar
  '********************************************
  'ENVISION TOOLBAR GLOBAL VARIABLES
  '********************************************

  'GENERAL
  Public m_appEnvision As IApplication
  Public m_strProcessing As String
  Public m_intDevTypeMax As Integer = 0

  'DOCK WINDOW VARIABLES
  Public m_pEnvisionDockWinMgr As IDockableWindowManager
  Public m_pDevTypeDockWinMgr As IDockableWindowManager
  Public m_pScIndicatorDockWinMgr As IDockableWindowManager
  Public m_pEnvisionDockWin As IDockableWindow
  Public m_pDevTypeDockWin As IDockableWindow
  Public m_pScIndicatorDockWin As IDockableWindow
  Public m_blnEnvisionEditingFormIsOpening As Boolean = False

  'ENVISON 1 TO MANY EXPORT 
  Public m_frm1toManyExport As frmOnetoManyExport = New frmOnetoManyExport

  'ENVISON ATTRIBUTE EDITOR 
  Public m_frmAttribManager As frmFieldVariables = New frmFieldVariables


  'ENVISION ATTRIBUTE EDITING FORM
  Public m_pEditFeatureLyr As IFeatureLayer
  Public m_dockEnvisionWinForm As New ctlEnvisionAttributesEditor
  Public m_pntMouse As System.Drawing.Point
  Public m_strEnvisionEditOption As String = "POLYGON"
  Public m_blnEnvisionMenuStatus As Boolean = False
  Public m_strFeaturePath As String = ""
  Public m_strFeatureName As String = ""
  Public m_dblAcresConversion As Double = 1
  Public m_strDWriteValue As String = ""
  Public m_blnAutoUpdate As Boolean = True
  Public m_blnEditProcessing As Boolean = False
  Public m_tblDevelopmentTypes As ITable = Nothing
  Public m_tblAttribFields As ITable = Nothing
  Public m_tblDevTypeFields As ITable = Nothing
  Public m_tblScSummary As ITable = Nothing
  Public m_tblExistingLU As ITable = Nothing
  Public m_tblSc1 As ITable = Nothing
  Public m_tblSc2 As ITable = Nothing
  Public m_tblSc3 As ITable = Nothing
  Public m_tblSc4 As ITable = Nothing
  Public m_tblSc5 As ITable = Nothing
  Public m_tblTravel As ITable = Nothing
  Public m_tblEXLURef As ITable = Nothing
  Public m_intRow As Integer = -1
  Public m_intEditScenario As Integer = 1
  Public m_blnScrMenuCheck As Boolean = False
  Public m_blnRunCalcs As Boolean = True
  Public m_blnSynchronizePrompt As Boolean = False
  Public blnLoadDevTypes As Boolean = False
  Public m_envSelect As IEnvelope = Nothing

  'ENVISION EXCEL FILE AND ASSOCIATED VARIABLES
  Public m_strEnvisionExcelFile As String = ""
  Public m_xlPaintApp As Microsoft.Office.Interop.Excel.Application = Nothing
  Public m_xlPaintWB1 As Microsoft.Office.Interop.Excel.Workbook = Nothing
  Public m_blnLoadingDevTypes As Boolean = False
  Public m_arrWriteCalcFields As ArrayList = New ArrayList
  Public m_arrWriteDevTypeFields As ArrayList = New ArrayList
  Public m_arrWriteDevTypeAcresFields As ArrayList = New ArrayList
  Public m_arrWriteDevTypeAcresFieldsAltName As ArrayList = New ArrayList
  Public m_arrWriteDevTypeAcresFieldsOnly As ArrayList = New ArrayList
  Public m_arrWriteDevTypeAcres2ndVarFields As ArrayList = New ArrayList
  Public m_arrWriteDevTypeTotalFields As ArrayList = New ArrayList
  Public m_arrWriteDevTypeTotalFieldName As ArrayList = New ArrayList

  'ENVISION SLOPE AND HILLSHADE LAYER CREATION
  Public m_frmEnvisionSlope As frmEnvisionSlope
  Public m_intSlopeGridCellSize As Integer = -1
  Public m_strDEMLayerName As String = ""
  Public m_pSlopeLyr As IRasterLayer = Nothing

  'ENVISION 7DS 
  Public m_frmEnvisionProjectSetup As frmEnvisionProjectSetup
  Public m_frmAccessBufferSummary As frmBufferSum
  Public m_frmAccessFeatureSummary As frmFeatureSum
  Public m_frmTravelSummaryBuffers As frmTravelSummaryBuffers
  Public m_frmProximtySummary As frmProximity
  Public m_frmSumJobs As frmJobSummary
  Public m_frmFieldAggregation As frmAggregation
  Public m_frmLUMix As frmLandUseMix
  Public m_frmRedevTiming As frmRedevTiming
  Public m_frmDevFeasibility As frmDevFeasibility
  Public m_frmLocationEfficiency As frmLocationEfficiency
  Public m_frmLocalJHBalance As frmLocalJobsHousingBalance
  Public m_strNeighborhoodLayerName As String = ""
  Public m_frm7DModelCalc As frm7DModelCalc
  Public m_frmSumTransLocations As frmSumTransportationPnts
  Public m_strBufferLayerNames As String = "SA_QUARTER_MILE,SA_HALF_MILE,SA_ONE_MILE,SA_10_MINUTE_AUTO,SA_20_MINUTE_AUTO,SA_30_MINUTE_AUTO,SA_30_MINUTE_TRANSIT"


  'ENVISION SETUP GLOBAL VARIABLES
  Public m_blnETSetupFormIsOpening As Boolean = False
  Public m_pETSetupExtentEnv As IEnvelope = Nothing
  Public m_arrRasterLayers As ArrayList = Nothing
  Public m_arrFeatureLayers As ArrayList = Nothing
  Public m_pSpatRefProject As ISpatialReference

  Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
 (ByVal hwnd As Int32, ByVal nIndex As Int32, ByVal dwNewLong As Int32) As Int32
  Private Const GWL_HWNDPARENT As Int32 = -8

  ''' -----------------------------------------------------------------------------
  ''' <summary>
  ''' Returns an ILayer object from a layer file.
  ''' </summary>
  ''' <param name="pathToLayerFile"></param>
  ''' <returns></returns>
  ''' <remarks>
  ''' </remarks>
  ''' -----------------------------------------------------------------------------
  Public Function GetLayerFromLayerFile(ByVal pathToLayerFile As String) As ILayer
    ' Make sure the pathToLayerFile has a .lyr extension.
    If pathToLayerFile.LastIndexOf(".") = -1 Then
      pathToLayerFile = pathToLayerFile.Insert(pathToLayerFile.Length, ".lyr")
    End If

    ' Verify that the layer file exists on disk.
    If File.Exists(pathToLayerFile) = False Then
      Throw New Exception("Layer file not found at " & pathToLayerFile)
    End If

    ' Open the layerfile.
    Dim layerFile As ILayerFile = New LayerFile
    Dim layer As ILayer
    layerFile.Open(pathToLayerFile)
    If layerFile.IsLayerFile(pathToLayerFile) AndAlso layerFile.IsPresent(pathToLayerFile) Then
      layer = layerFile.Layer
    Else
      Throw New Exception("The layerfile " & pathToLayerFile & " is invalid.")
    End If

    ' Return the layer
    Return layer
  End Function

  ''' -----------------------------------------------------------------------------
  ''' <summary>
  ''' Disconnects a featurelayer from it's datasource.
  ''' </summary>
  ''' <param name="featureLayer"></param>
  ''' <remarks>
  ''' </remarks>
  ''' -----------------------------------------------------------------------------
  Public Sub DisconnectFeatureLayerFromDatasource(ByVal featureLayer As IFeatureLayer)

    ' Don't bother disconnecting if the layer isn't valid.
    If Not featureLayer.Valid Then
      Return
    End If

    ' Get the idatalayer interface of the layer.
    Dim dataLayer As IDataLayer2 = DirectCast(featureLayer, IDataLayer2)

    ' Disconnect the layer from it's datasource.
    If dataLayer.DataSourceName.NameString <> String.Empty Then dataLayer.Disconnect()
  End Sub

  ''' <summary>
  ''' Gets the spatial analyst extension object
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function GetSpatialAnalystExtension() As IExtensionConfig
    Dim factoryType As Type = Type.GetTypeFromProgID("esriSystem.ExtensionManager")
    Dim extensionManager As IExtensionManager = CType(Activator.CreateInstance(factoryType), IExtensionManager)

    Dim uid As IUID = New UIDClass()
    uid.Value = "esriSpatialAnalystUI.SAExtension"
    Dim extension As IExtension = extensionManager.FindExtension(uid)
    Dim extensionConfig As IExtensionConfig = CType(extension, IExtensionConfig)
    Return extensionConfig
  End Function

  ''' <summary>
  ''' Determine if the spatial analyst extension is alreayd enabled
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks>
  ''' True if the extension is enabled
  ''' </remarks>
  Public Function IsSpatialAnalystExtensionEnabled() As Boolean
    Dim extensionConfig As IExtensionConfig = GetSpatialAnalystExtension()
    Dim isEnabled As Boolean = (extensionConfig.State = esriExtensionState.esriESEnabled)
    Return isEnabled
  End Function

  ''' <summary>
  ''' Disable the spatial analyst extension
  ''' </summary>
  ''' <remarks></remarks>
  Public Sub DisableSpatialAnalystExtension()
    Dim extensionConfig As IExtensionConfig = GetSpatialAnalystExtension()
    If Not extensionConfig.State = esriExtensionState.esriESUnavailable Then
      extensionConfig.State = esriExtensionState.esriESDisabled
    End If
  End Sub

  ''' <summary>
  ''' Enable the spatial analyst extension
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks>
  ''' True if the extension was enabled or was already enabled
  ''' </remarks>
  Public Function EnableSpatialAnalystExtension() As Boolean
    Dim extensionConfig As IExtensionConfig = GetSpatialAnalystExtension()

    Dim isEnabled As Boolean = (extensionConfig.State = esriExtensionState.esriESEnabled)

    Dim wasEnabled As Boolean
    If Not isEnabled Then
      If Not (extensionConfig.State = esriExtensionState.esriESUnavailable) Then
        extensionConfig.State = esriExtensionState.esriESEnabled
        wasEnabled = True
      End If
    End If
    Return wasEnabled
  End Function

  Public Sub WriteToProcessingFile(ByVal strFilename As String)
    m_strProcessing = m_strProcessing & vbNewLine & "Attempting to write processing text to file: " & strFilename & vbNewLine
    'Writes the processing message to the current Envision File Geodatabase
    'SAVE COPIES OF THE ENVISON LICENSE FILES
    Dim datWriteLC As StreamWriter
    Dim datReadPy As StreamReader
    Dim strReadLine As String
    Dim reader As System.IO.DirectoryInfo

    Try
      'CHECK TO SEE IF THE DIRECTORY IS VALID
      m_strProcessing = m_strProcessing & "Check to see if storage directory extist " & vbNewLine
      If Not Directory.Exists((m_strFeaturePath & "\ENVISION")) Then
        m_strProcessing = m_strProcessing & "Directory not found: " & (m_strFeaturePath & "\ENVISION") & vbNewLine
        m_strProcessing = m_strProcessing & "Exiting sub function" & vbNewLine
        GoTo CleanUp
      End If

      'CREATE AND WRITE OUT LICENSE FILE
      m_strProcessing = m_strProcessing & "Writing processing text to file" & vbNewLine
      datWriteLC = New StreamWriter((m_strFeaturePath & "\ENVISION\" & strFilename))
      datWriteLC.Write(m_strProcessing)
      datWriteLC.Close()
      GoTo CleanUp
    Catch ex As Exception
      m_strProcessing = m_strProcessing & "Error in sub WriteToProcessingFile: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
      m_strProcessing = m_strProcessing & "Error Message:  " & vbNewLine & ex.Message & vbNewLine
      MessageBox.Show("Error occured in attempting to create the Envision processing file." & vbNewLine & ex.Message, "Write File Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
      GoTo CleanUp
    End Try

CleanUp:
    datWriteLC = Nothing
    datReadPy = Nothing
    strReadLine = Nothing
    reader = Nothing
    GC.WaitForPendingFinalizers()
    GC.Collect()
  End Sub

  Public Function EnvisionLyrRequiredFieldsCheck(ByVal pTable As ITable) As Boolean

    'FIELDS REQUIRED TO CONDUCT ENVISION PROCESSES
    Dim strFields As String
    Dim intFld As Integer = 0
    Dim intFldCount As Integer = 20
    Dim rowTemp As IRow
    Dim strFieldName As String = ""
    Dim strFieldType As String = ""
    Dim intFieldWidth As Integer = 0
    Dim intFieldDecimal As Integer = 0
    Dim strFieldAlias As String = ""
    Dim intUseField As Integer = 0
    Dim intCalcAcres As Integer = 0
    Dim strCalcFieldName As String = ""
    Dim strDepVarFieldName As String = ""
    Dim totFldName As String
    Dim exFldName As String
    Dim intCalcOnly As Integer = 0
    Dim calcTotal As Int16
    Dim pCursor As ICursor

    'SET THE ARCGIS VIEW STATUS MESSAGE
    m_appEnvision.StatusBar.Message(0) = "CHECKING FOR REQUIRED ENVISION EDIT LAYER TABLE FIELDS AND FIELD CALCULATIONS"
    m_dockEnvisionWinForm.prgBarEnvision.Value = 0

    Try
      'ACRES FIELDS
      strFields = "VAC_ACRE,DEVD_ACRE,CONSTRAINED_ACRE"
      intFldCount = 3
      intFld = 0
      For Each strFieldName In strFields.Split(","c)
        m_appEnvision.StatusBar.Message(0) = "CHECKING FOR REQUIRED ENVISION EDIT LAYER TABLE FIELDS AND FIELD CALCULATIONS: FIELD (" & strFieldName & ")"
        If pTable.FindField(strFieldName) >= 0 Then
          Continue For
        End If
        If Not AddEnvisionField(pTable, strFieldName, "DOUBLE", 16, 6) Then
          Return False
        End If
        intFld = intFld + 1
        m_dockEnvisionWinForm.prgBarEnvision.Value = CType((intFld / intFldCount) * 100, Int32)
      Next

      'DEVELOPMENT  TYPE FIELDS
      strFields = "DEV_TYPE"
      intFldCount = 7
      intFld = 0
      For Each strFieldName In strFields.Split(","c)
        m_appEnvision.StatusBar.Message(0) = "CHECKING FOR REQUIRED ENVISION EDIT LAYER TABLE FIELDS AND FIELD CALCULATIONS: FIELD (" & strFieldName & ")"
        If pTable.FindField(strFieldName) >= 0 Then
          Continue For
        End If
        If AddEnvisionField(pTable, strFieldName, "STRING", 50, 0) Then
          CalcFldValues(pTable, strFieldName, """""", "STRING")
        Else
          Return False
        End If
        intFld = intFld + 1
        m_dockEnvisionWinForm.prgBarEnvision.Value = CType((intFld / intFldCount) * 100, Int32)
      Next

      'CHECK TO SEE IF ANY DEV TYPE ATTRIBUTE VALUE FIELDS ARE REQUIRED
      m_arrWriteDevTypeAcresFields.Clear()
      m_arrWriteDevTypeAcresFieldsAltName.Clear()
      m_arrWriteDevTypeAcresFieldsOnly.Clear()
      m_arrWriteDevTypeFields.Clear()
      m_arrWriteDevTypeAcres2ndVarFields.Clear()
      m_arrWriteDevTypeTotalFields.Clear()
      m_arrWriteDevTypeTotalFieldName.Clear()
      If Not m_tblAttribFields Is Nothing Then
        Dim calcTotalFldIndex As Int32 = m_tblAttribFields.FindField("Calc_Total")
        pCursor = m_tblAttribFields.Search(Nothing, False)
        rowTemp = pCursor.NextRow
        Do Until rowTemp Is Nothing
          strFieldName = ""
          strFieldAlias = ""
          intUseField = 0
          intCalcAcres = 0
          strCalcFieldName = ""
          strDepVarFieldName = ""
          intCalcOnly = 0
          calcTotal = 0
          totFldName = String.Empty
          exFldName = String.Empty
          Try
            strFieldName = CStr(rowTemp.Value(rowTemp.Fields.FindField("FIELD_NAME")))
          Catch ex As Exception
          End Try
          Try
            intUseField = CInt(rowTemp.Value(rowTemp.Fields.FindField("USE")))
          Catch ex As Exception
          End Try
          Try
            intCalcAcres = CInt(rowTemp.Value(rowTemp.Fields.FindField("CALC_BY_ACRES")))
            If intCalcAcres = -1 Then
              intCalcAcres = 1
            End If
          Catch ex As Exception
          End Try
          Try
            strCalcFieldName = CStr(rowTemp.Value(rowTemp.Fields.FindField("CALC_FIELD_NAME")))
          Catch ex As Exception
          End Try
          Try
            strDepVarFieldName = CStr(rowTemp.Value(rowTemp.Fields.FindField("DEP_VAR_FIELD_NAME")))
          Catch ex As Exception
          End Try
          Try
            totFldName = rowTemp.Value(rowTemp.Fields.FindField("Total_FIELD_NAME")).ToString
          Catch ex As Exception
          End Try
          If intCalcAcres = 1 And strCalcFieldName.Length <= 0 Then
            strCalcFieldName = strFieldName
          End If
          Try
            intCalcOnly = CInt(rowTemp.Value(rowTemp.Fields.FindField("CALC_ONLY")))
          Catch ex As Exception
          End Try
          If calcTotalFldIndex <> -1 Then
            If rowTemp.Value(calcTotalFldIndex) IsNot DBNull.Value Then
              calcTotal = CType(rowTemp.Value(calcTotalFldIndex), Int16)
            End If
          End If
          If intUseField = 1 Then
            If intCalcAcres >= 1 Then
              'ADD TO TRACKING FIELD LIST FOR DEV TYPE VALUES THAT ARE CACLED BY PARCEL ACRES
              m_arrWriteDevTypeAcresFields.Add(strFieldName)
              m_arrWriteDevTypeAcresFieldsAltName.Add(strCalcFieldName)
              If intCalcOnly = 1 Then
                m_arrWriteDevTypeAcresFieldsOnly.Add(strCalcFieldName)
              End If
              If intCalcAcres = 2 Then
                m_arrWriteDevTypeAcres2ndVarFields.Add(strDepVarFieldName)
              Else
                m_arrWriteDevTypeAcres2ndVarFields.Add("")
              End If
              exFldName = strCalcFieldName
            Else
              'ADD TO TRACKING FIELD LIST FOR DEV TYPE VALUES
              m_arrWriteDevTypeFields.Add(strFieldName)
              exFldName = strFieldName
            End If
            Try
              If calcTotal <> 0 Then
                m_arrWriteDevTypeTotalFields.Add(exFldName)
                If totFldName = String.Empty Then totFldName = GetDefaultTotalFieldName(intCalcAcres, strCalcFieldName, strFieldName, m_pEditFeatureLyr.FeatureClass)
                If m_arrWriteDevTypeTotalFieldName.Contains(totFldName) Then Throw New Exception("The field name '" & totFldName & "' is not unique.  Please review the attribute field manager and make corrections.")
                m_arrWriteDevTypeTotalFieldName.Add(totFldName)
              End If
            Catch ex As Exception
            End Try
          End If

          If CType(intUseField, Boolean) Then
            Try
              strFieldType = CStr(rowTemp.Value(rowTemp.Fields.FindField("FIELD_TYPE")))
            Catch ex As Exception
            End Try
            Try
              intFieldWidth = CInt(rowTemp.Value(rowTemp.Fields.FindField("FIELD_WIDTH")))
            Catch ex As Exception
            End Try
            Try
              intFieldDecimal = CInt(rowTemp.Value(rowTemp.Fields.FindField("FIELD_DECIMAL")))
            Catch ex As Exception
            End Try
            If pTable.FindField(strFieldName) <= -1 And Not intCalcOnly = 1 Then
              AddEnvisionField(pTable, strFieldName, strFieldType, intFieldWidth, intFieldDecimal)
            End If
            If UCase(strFieldName) = "HU" Or UCase(strFieldName) = "EMP" Or UCase(strFieldName) = "SF" Or UCase(strFieldName) = "TH" Or UCase(strFieldName) = "MF" Or UCase(strFieldName) = "MH" Or UCase(strFieldName) = "RET" Or UCase(strFieldName) = "OFF" Or UCase(strFieldName) = "IND" Then
              If pTable.FindField("EX_" & strFieldName) <= -1 Then
                AddEnvisionField(pTable, "EX_" & strFieldName, strFieldType, intFieldWidth, intFieldDecimal)
              End If
            End If

            If intCalcAcres >= 1 And pTable.FindField(strCalcFieldName) <= -1 Then
              AddEnvisionField(pTable, strCalcFieldName, strFieldType, intFieldWidth, intFieldDecimal)
            End If
            If intCalcAcres = 1 OrElse intCalcAcres = 2 Then
              If UCase(strCalcFieldName) = "HU" Or UCase(strCalcFieldName) = "EMP" Or UCase(strCalcFieldName) = "SF" Or UCase(strCalcFieldName) = "TH" Or UCase(strCalcFieldName) = "MF" Or UCase(strCalcFieldName) = "MH" Or UCase(strCalcFieldName) = "RET" Or UCase(strCalcFieldName) = "OFF" Or UCase(strCalcFieldName) = "IND" Then
                If pTable.FindField("EX_" & strCalcFieldName) <= -1 Then
                  AddEnvisionField(pTable, "EX_" & strCalcFieldName, strFieldType, intFieldWidth, intFieldDecimal)
                End If
              End If
              If calcTotal <> 0 AndAlso totFldName <> String.Empty Then
                If pTable.FindField("EX_" & strCalcFieldName) <> -1 Then
                  If pTable.FindField(totFldName) = -1 Then
                    AddEnvisionField(pTable, totFldName, strFieldType, intFieldWidth, intFieldDecimal)
                  End If
                End If
              End If
            Else
              If calcTotal <> 0 AndAlso totFldName <> String.Empty Then
                If pTable.FindField("EX_" & strFieldName) <> -1 Then
                  If pTable.FindField(totFldName) = -1 Then
                    AddEnvisionField(pTable, totFldName, strFieldType, intFieldWidth, intFieldDecimal)
                  End If
                End If
              End If
            End If
          End If
          rowTemp = pCursor.NextRow
        Loop
      End If

      GoTo CleanUp
    Catch ex As Exception
      MessageBox.Show(ex.Message, "EnvisionLyrREquiredFieldsCheck Function ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
      GoTo CleanUp
    End Try

CleanUp:
    strFields = Nothing
    intFld = Nothing
    intFldCount = Nothing
    rowTemp = Nothing
    strFieldName = Nothing
    strFieldType = Nothing
    intFieldWidth = Nothing
    intFieldDecimal = Nothing
    strFieldAlias = Nothing
    intUseField = Nothing
    intCalcAcres = Nothing
    strCalcFieldName = Nothing
    intCalcOnly = Nothing
    pCursor = Nothing
    GC.Collect()
    GC.WaitForPendingFinalizers()
    m_appEnvision.StatusBar.Message(0) = ""
    m_dockEnvisionWinForm.prgBarEnvision.Value = 0
  End Function

  '  Public Sub EnvisionLyrFieldTrackingCheck(ByVal pTable As ITable)

  '    'FIELDS REQUIRED TO CONDUCT ENVISION PROCESSES
  '    Dim pWksFactory As ESRI.ArcGIS.Geodatabase.IWorkspaceFactory = Nothing
  '    Dim pFeatWks As ESRI.ArcGIS.Geodatabase.IFeatureWorkspace = Nothing
  '    Dim pFieldTable As ITable
  '    Dim GP As ESRI.ArcGIS.Geoprocessor.Geoprocessor
  '    Dim pCreateTable As ESRI.ArcGIS.DataManagementTools.CreateTable
  '    Dim strFields As String
  '    Dim intRow As Integer = 0
  '    Dim intRowCount As Integer = 0
  '    Dim rowTemp As IRow
  '    Dim strFieldName As String = ""
  '    Dim strFieldAlias As String = ""
  '    Dim strFieldType As String = ""
  '    Dim intFieldWidth As Integer = 0
  '    Dim intFieldDecimal As Integer = 0
  '    Dim intCalc As Integer = 0
  '    Dim strCalcFieldName As String = ""
  '    Dim strUseField As String = ""
  '    Dim calcTotal As Int16


  '    'SET THE ARCGIS VIEW STATUS MESSAGE
  '    m_appEnvision.StatusBar.Message(0) = "CHECKING FOR REQUIRED ENVISION EDIT LAYER TABLE FIELDS AND FIELD CALCULATIONS"
  '    m_dockEnvisionWinForm.prgBarEnvision.Value = 0

  '    'FIELD TRACKING TABLE
  '    Try
  '      pWksFactory = New FileGDBWorkspaceFactory
  '      pFeatWks = DirectCast(pWksFactory.OpenFromFile(m_strFeaturePath, 0), IFeatureWorkspace)
  '      pFieldTable = pFeatWks.OpenTable("ENVISION_DEVTYPE_FIELD_TRACKING")
  '    Catch ex As Exception
  '      m_strProcessing = m_strProcessing & "Missing table, ENVISION_FIELD_TRACKING, executing geoprocessing tool to create table: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
  '      GP = New ESRI.ArcGIS.Geoprocessor.Geoprocessor
  '      pCreateTable = New ESRI.ArcGIS.DataManagementTools.CreateTable
  '      pCreateTable.out_name = "ENVISION_DEVTYPE_FIELD_TRACKING"
  '      pCreateTable.out_path = m_strFeaturePath
  '      RunTool(GP, pCreateTable)
  '      pFieldTable = pFeatWks.OpenTable("ENVISION_DEVTYPE_FIELD_TRACKING")
  '      LookUpTablesEnvisionAttributeFieldTrackingTblCheck(pFieldTable)
  '    End Try

  '    'CYCLE THROUGH LIST TO SEE WHAT FIELDS ARE SELECTED FOR PROCESSING 
  '    intRowCount = pFieldTable.RowCount(Nothing)
  '    Dim calcTotalFldIndex As Int32 = pFieldTable.FindField("Calc_Total")
  '    For intRow = 0 To intRowCount - 1
  '      rowTemp = pFieldTable.GetRow(intRow)
  '      strFieldName = rowTemp.Value(pFieldTable.FindField("FIELD_NAME")).ToString
  '      strFieldAlias = rowTemp.Value(pFieldTable.FindField("FIELD_ALIAS")).ToString
  '      strFieldType = rowTemp.Value(pFieldTable.FindField("FIELD_TYPE")).ToString
  '      intFieldWidth = CType(rowTemp.Value(pFieldTable.FindField("FIELD_WIDTH")), Int32)
  '      intFieldDecimal = CType(rowTemp.Value(pFieldTable.FindField("FIELD_DECIMAL")), Int32)
  '      intCalc = CType(rowTemp.Value(pFieldTable.FindField("CALC_BY_ACRES")).ToString, Int32)
  '      strCalcFieldName = rowTemp.Value(pFieldTable.FindField("CALC_FIELD_NAME")).ToString
  '      strUseField = rowTemp.Value(pFieldTable.FindField("USE")).ToString
  '      calcTotal = 0
  '      If calcTotalFldIndex <> -1 Then
  '        If rowTemp.Value(calcTotalFldIndex) IsNot DBNull.Value Then calcTotal = CType(rowTemp.Value(calcTotalFldIndex), Int16)
  '      End If
  '      If strUseField = "YES" Then
  '        If pTable.FindField(strFieldName) = -1 Then
  '          AddEnvisionField(pTable, strFieldName, strFieldType, intFieldWidth, intFieldDecimal)
  '        End If
  '        If intCalc >= 1 And pTable.FindField(strCalcFieldName) = -1 Then
  '          AddEnvisionField(pTable, strCalcFieldName, strFieldType, intFieldWidth, intFieldDecimal)
  '        End If
  '        If calcTotal = 1 Then
  '          AddEnvisionField(pTable, "TOT_" & strCalcFieldName, strFieldType, intFieldWidth, intFieldDecimal)
  '        End If
  '      End If
  '    Next


  'CleanUp:
  '    pWksFactory = Nothing
  '    pFeatWks = Nothing
  '    pFieldTable = Nothing
  '    GP = Nothing
  '    pCreateTable = Nothing
  '    strFields = Nothing
  '    intRow = Nothing
  '    intRowCount = Nothing
  '    rowTemp = Nothing
  '    strFieldName = Nothing
  '    strFieldAlias = Nothing
  '    strFieldType = Nothing
  '    intFieldWidth = Nothing
  '    intFieldDecimal = Nothing
  '    strCalcFieldName = Nothing
  '    strUseField = Nothing
  '    GC.Collect()
  '    GC.WaitForPendingFinalizers()
  '    m_appEnvision.StatusBar.Message(0) = ""
  '    m_dockEnvisionWinForm.prgBarEnvision.Value = 0
  '  End Sub

  Public Function CalcFldValues(ByVal pTable As ITable, ByVal strFieldName As String, ByVal strValue As String, ByVal strType As String) As Boolean
    CalcFldValues = True

    Dim pCursor As ICursor
    Dim pCalc As ICalculator
    Try
      pCalc = New Calculator
      pCursor = pTable.Update(Nothing, False)
      With pCalc
        .Cursor = pCursor
        .Field = strFieldName
        If UCase(strType) = "STRING" Then
          'I NEED TO REVIEW THIS ONE, MIGHT NEED DOUBLE QUOUTES ADDED HERE?
          .Expression = strValue
        Else
          .Expression = strValue
        End If
      End With
      pCalc.Calculate()
      GoTo CleanUp
    Catch ex As Exception
      CalcFldValues = False
      MessageBox.Show(ex.Message, "Calculate Field Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
      GoTo CleanUp
    End Try

CleanUp:
    pCalc = Nothing
    pCursor = Nothing
    pTable = Nothing
    GC.WaitForPendingFinalizers()
    GC.Collect()
  End Function

  Public Function AddEnvisionField(ByVal pTable As ITable, ByVal strFieldName As String, ByVal strFieldType As String, ByVal intFieldWidth As Integer, ByVal intFieldDecimals As Integer) As Boolean
    'ADD FIELD TO GIVEN TABLE BASED UPON USER INPUTS
    AddEnvisionField = True
    Dim pField As IField
    Dim pFieldEdit As IFieldEdit
    Dim intField As Integer

    If Not m_dockEnvisionWinForm.btnEditing.Text = "Start Edit" Then
      'A FIELD CAN NOT BE ADDED DURING AN EDIT SESSION
      GoTo CleanUp
    End If

    Try
      'FIRST CHECK IF FIELD EXISTS
      intField = pTable.FindField(strFieldName)
      If intField = -1 Then
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)
        pFieldEdit = New Field
        pFieldEdit.Editable_2 = True
        With pFieldEdit
          If UCase(strFieldType) = "INTEGER" Then
            .Type_2 = esriFieldType.esriFieldTypeInteger
            .DefaultValue_2 = 0
            If strFieldName = "PRCNT_DEV" Then
              .DefaultValue_2 = 100
            ElseIf strFieldName = "DVLPBL" Then
              .DefaultValue_2 = 1
            End If
            .Precision_2 = intFieldWidth
          End If
          If UCase(strFieldType) = "STRING" Then
            .Type_2 = esriFieldType.esriFieldTypeString
            .DefaultValue_2 = ""
            .Length_2 = intFieldWidth
          End If
          If UCase(strFieldType) = "DOUBLE" Then
            .Type_2 = esriFieldType.esriFieldTypeDouble
            .DefaultValue_2 = 0.0000000001
            .Scale_2 = intFieldDecimals
            .Precision_2 = intFieldWidth
            .Length_2 = intFieldWidth
          End If
          .Editable_2 = True
          .IsNullable_2 = True
          .Name_2 = strFieldName
          .Required_2 = False
        End With
        pTable.AddField(pFieldEdit)
      Else
        GoTo CleanUp
      End If
    Catch ex As Exception
      AddEnvisionField = False
      MessageBox.Show("Unable to add field, " & strFieldName & "." & vbNewLine & ex.Message, "Add Field Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign, False)
    End Try

CleanUp:
    pTable = Nothing
    strFieldName = Nothing
    strFieldType = Nothing
    intFieldWidth = Nothing
    intFieldDecimals = Nothing
    pField = Nothing
    pFieldEdit = Nothing
    GC.WaitForPendingFinalizers()
    GC.Collect()
  End Function

  Public Sub CloseEnvisionExcel()
    'CLOSE THE ENVISION EXCEL APPLICATION
    Try
      If Not m_xlPaintApp Is Nothing Then
        m_xlPaintApp.Quit()
        Marshal.FinalReleaseComObject(m_xlPaintApp)
      End If
      m_xlPaintApp = Nothing

      If Not m_xlPaintWB1 Is Nothing Then
        m_xlPaintWB1.Close()
        Marshal.FinalReleaseComObject(m_xlPaintWB1)
      End If
      m_xlPaintWB1 = Nothing
      GoTo CleanUp
    Catch ex As Exception
      'MessageBox.Show("Error in closing the Envision Excel file.  You may need to terminate the application using the Task Manager." & vbNewLine & ex.Message, "Envision Excel Closing Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
      GoTo CleanUp
    End Try
CleanUp:
    m_xlPaintWB1 = Nothing
    m_xlPaintApp = Nothing
    GC.WaitForPendingFinalizers()
    GC.Collect()
  End Sub

  Public Function OpenEnvisionExcel() As Boolean
    'OPEN SELECTED ENVISION EXCEL FILE
    OpenEnvisionExcel = True
    Dim arrList As ArrayList = New ArrayList
    Dim strTabNum As String

    'A NUMBER FOR EACH REQUIRED SCENARIO TAB
    arrList.Add("1")
    arrList.Add("2")
    arrList.Add("3")
    arrList.Add("4")
    arrList.Add("5")

    Try
      If m_strEnvisionExcelFile.Length <= 0 Then
        OpenEnvisionExcel = False
        GoTo CleanUp
      End If
      m_blnLoadingDevTypes = True
      If m_xlPaintApp Is Nothing Then
        m_xlPaintApp = New Microsoft.Office.Interop.Excel.Application
        m_xlPaintApp.DisplayAlerts = False
        m_xlPaintApp.WindowState = Interop.Excel.XlWindowState.xlMinimized
        m_xlPaintApp.Visible = True
      End If
      If m_xlPaintWB1 Is Nothing Then
        m_xlPaintWB1 = CType(m_xlPaintApp.Workbooks.Open(m_strEnvisionExcelFile), Microsoft.Office.Interop.Excel.Workbook)
      End If
      'CHECK FOR DEV TYPE ATTRIBUTES TAB
      If Not TypeOf m_xlPaintWB1.Sheets("Dev Type Attributes") Is Microsoft.Office.Interop.Excel.Worksheet Then
        MessageBox.Show("The 'Dev Type Attributes' tab could not be found in the selected Excel file. Please select another file.", "DEV TYPE ATTRIBUTES TAB NOT FOUND", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
        CloseEnvisionExcel()
        OpenEnvisionExcel = False
        GoTo CleanUp
      Else
        If Not m_blnAutoUpdate Then
          GoTo cleanup
        End If
      End If
      'CHECK FOR EACH SCENARIO TAB
      For Each strTabNum In arrList
        If Not TypeOf m_xlPaintWB1.Sheets("SCENARIO" & strTabNum) Is Microsoft.Office.Interop.Excel.Worksheet Then
          MessageBox.Show("The 'SCENARIO" & strTabNum & "' tab could not be found in the selected Excel file. Please select another file.", "SCENARIO TAB NOT FOUND", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
          CloseEnvisionExcel()
          OpenEnvisionExcel = False
          GoTo CleanUp
        Else
          If Not m_blnAutoUpdate Then
            GoTo cleanup
          End If
        End If
      Next
    Catch ex As Exception
      MessageBox.Show(ex.Message & vbNewLine & "Stack Trace: " & ex.StackTrace, "Opening Envision Excel File Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
      OpenEnvisionExcel = False
      CloseEnvisionExcel()
      GoTo CleanUp
    End Try

CleanUp:
    m_blnLoadingDevTypes = False
    GC.WaitForPendingFinalizers()
    GC.Collect()
  End Function

  '  Public Function LoadEnvisionEditLayer(ByVal strPath As String, ByVal strName As String, ByVal blnPrompt As Boolean) As Boolean
  '    LoadEnvisionEditLayer = True

  '    Dim pMxDocument As IMxDocument
  '    Dim mapCurrent As Map
  '    Dim blnLayerFound As Boolean = False
  '    Dim pFeatLayer As IFeatureLayer = Nothing
  '    Dim pWksFactory As IWorkspaceFactory = Nothing
  '    Dim pFeatWks As IFeatureWorkspace
  '    Dim pFeatClass As IFeatureClass
  '    Dim pDataset As IDataset
  '    Dim intLayer As Integer
  '    Dim pTable As ITable
  '    Dim strLyrPath As String = ""
  '    Dim strLyrName As String = ""
  '    Dim pLayerAffects As ILayerEffects

  '    pMxDocument = DirectCast(m_appEnvision.Document, IMxDocument)
  '    mapCurrent = CType(pMxDocument.FocusMap, Map)
  '    Try
  '      'DISABLE THE CONTROLS WHICH ARE DEPENDANT UPON THE ENVISION LAYER FILE
  '      m_dockEnvisionWinForm.btnEditing.Enabled = True
  '      m_dockEnvisionWinForm.itmEnvisionExcel.Enabled = True
  '      m_dockEnvisionWinForm.itmSaveToExcelFile.Enabled = True
  '      m_dockEnvisionWinForm.itmSynchronize.Enabled = True
  '      m_dockEnvisionWinForm.itmAccessFunctions.Enabled = True

  '      'REVIEW CURRENT LAYERS FOR ENVISION LAYER
  '      If mapCurrent.LayerCount > 0 Then
  '        For intLayer = 0 To mapCurrent.LayerCount - 1
  '          If TypeOf mapCurrent.Layer(intLayer) Is IFeatureLayer Then
  '            pFeatLayer = CType(mapCurrent.Layer(intLayer), IFeatureLayer)
  '            pFeatClass = pFeatLayer.FeatureClass
  '            pDataset = CType(pFeatClass, IDataset)
  '            strLyrPath = pDataset.Workspace.PathName
  '            strLyrName = pDataset.BrowseName
  '            If strPath = strLyrPath And (strName = strLyrName Or strName = strLyrName & ".shp") And pFeatClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
  '              blnLayerFound = True
  '              m_pEditFeatureLyr = pFeatLayer
  '              m_strFeaturePath = strLyrPath
  '              If m_strFeaturePath.Contains("\\") Then
  '                m_strFeaturePath = m_strFeaturePath.Replace("\\", "\")
  '              End If
  '              m_strFeatureName = strLyrName
  '              pFeatLayer.Name = pFeatLayer.Name & "<ACTIVE>"""
  '              pTable = CType(pFeatClass, Table)
  '              EnvisionLyrFieldTrackingCheck(pTable)
  '            Else
  '              If pFeatLayer.Name.Contains("<ACTIVE>") Then
  '                pFeatLayer.Name = pDataset.Name
  '              End If
  '            End If
  '            GC.WaitForPendingFinalizers()
  '            GC.Collect()
  '          End If
  '        Next
  '      End If
  '    Catch ex As Exception
  '    End Try

  '    Try
  '      If Not blnLayerFound Then
  '        pWksFactory = New FileGDBWorkspaceFactory
  '        pFeatWks = DirectCast(pWksFactory.OpenFromFile(strPath, 0), IFeatureWorkspace)
  '        pFeatClass = pFeatWks.OpenFeatureClass(strName)
  '        If pFeatClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
  '          m_pEditFeatureLyr = New FeatureLayer
  '          m_pEditFeatureLyr.FeatureClass = pFeatClass
  '          m_pEditFeatureLyr.Name = pFeatLayer.Name & "<ACTIVE>"
  '          pLayerAffects = DirectCast(m_pEditFeatureLyr, ILayerEffects)
  '          pLayerAffects.Transparency = 35
  '          m_strFeaturePath = strPath
  '          If m_strFeaturePath.Contains("\\") Then
  '            m_strFeaturePath = m_strFeaturePath.Replace("\\", "\")
  '          End If
  '          m_strFeatureName = strName
  '          m_dockEnvisionWinForm.itmOpenEnvisionFGDB.ToolTipText = m_strFeaturePath & "\" & m_strFeatureName
  '          mapCurrent.AddLayer(m_pEditFeatureLyr)
  '          pTable = CType(pFeatClass, Table)
  '          EnvisionLyrFieldTrackingCheck(pTable)
  '        End If
  '        GC.WaitForPendingFinalizers()
  '        GC.Collect()
  '      End If

  '      'LOAD THE DEVELOPMENT TYPES
  '      RetrieveDevTypeData()

  '      If m_dockEnvisionWinForm.dgvDevelopmentTypes.Rows.Count > 0 Then
  '        'UPDATE THE LEGEND
  '        UpdateEnvisionLyrLegend()
  '      End If

  '      'ENABLE THE CONTROLS WHICH ARE DEPENDANT UPON THE ENVISION LAYER FILE
  '      m_dockEnvisionWinForm.btnEditing.Enabled = True
  '      m_dockEnvisionWinForm.itmEnvisionExcel.Enabled = True
  '      m_dockEnvisionWinForm.itmSaveToExcelFile.Enabled = True
  '      m_dockEnvisionWinForm.itmSynchronize.Enabled = True
  '      m_dockEnvisionWinForm.itmAccessFunctions.Enabled = True

  '      'UPDATE THE INDICATOR TABLE SELECTION
  '      Try
  '        UpdateSummaryTableSelection()
  '      Catch ex As Exception
  '        MessageBox.Show(ex.Message, "Open Layer Error: UpdateSummaryTableSelection Sub Section", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
  '        GoTo CleanUp
  '      End Try

  '      'LOAD THE INDICATOR GRAPH IMAGES
  '      Try
  '        TrackIndicatorGraphs()
  '      Catch ex As Exception
  '        MessageBox.Show(ex.Message, "Open Layer Error: TrackIndicatorGraphs Sub Section", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
  '        GoTo CleanUp
  '      End Try

  '      'LOAD THE SELECTED ENVISION EDIT LAYERS FIELDS
  '      Try
  '        RetrieveEvisionFields()
  '      Catch ex As Exception
  '        MessageBox.Show(ex.Message, "Open Layer Error: RetrieveEvisionFields Sub Section", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
  '        GoTo CleanUp
  '      End Try

  '      GoTo CleanUp

  '    Catch ex As Exception
  '      If m_pEditFeatureLyr Is Nothing Then
  '        m_dockEnvisionWinForm.itmOpenEnvisionFGDB.ToolTipText = ""
  '      End If
  '      LoadEnvisionEditLayer = False
  '      'CloseEnvisionExcel()
  '      MessageBox.Show("The Envision Layer would not load." & vbNewLine & strPath & "\" & strName & vbNewLine & ex.Message, "Envision Load Layer Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
  '      GoTo CleanUp
  '    End Try

  'CleanUp:
  '    pMxDocument = Nothing
  '    mapCurrent = Nothing
  '    blnLayerFound = Nothing
  '    pFeatLayer = Nothing
  '    pWksFactory = Nothing
  '    pFeatWks = Nothing
  '    pFeatClass = Nothing
  '    pDataset = Nothing
  '    intLayer = Nothing
  '    pTable = Nothing
  '    pLayerAffects = Nothing
  '    GC.WaitForPendingFinalizers()
  '    GC.Collect()
  '  End Function

  Public Function WriteRGB2EnvisionExcelFile() As Boolean
    'SET THE RGB COLOR VALUES FOR SYMBOLS TO THE SELECTED ENVISION EXCEL FILE
    Dim arrVariables As ArrayList = New ArrayList
    Dim shtScenario As Microsoft.Office.Interop.Excel.Worksheet = Nothing
    Dim shtDevType As Microsoft.Office.Interop.Excel.Worksheet = Nothing
    Dim shtReport As Microsoft.Office.Interop.Excel.Worksheet = Nothing
    Dim intRow As Integer
    Dim strCellValue As String = ""
    Dim strLyrPath As String = ""
    Dim strLyrName As String = ""
    Dim rowTemp As IRow
    Dim intValue As Integer = 0
    Dim dblValue As Double = 0
    Dim tblsc As ITable = Nothing
    Dim intCol As Integer = 0
    Dim intRedFld As Integer = -1
    Dim intGreenFld As Integer = -1
    Dim intBlueFld As Integer = -1
    Dim strFldCellValue As String = ""
    Dim intRed As Integer = -1
    Dim intGreen As Integer = -1
    Dim intBlue As Integer = -1
    Dim intStartRow As Integer = 5
    Dim intExcelFormulaCalc As Integer

    'CHECK FOR EXISTING DEV TYPES
    If m_dockEnvisionWinForm.dgvDevelopmentTypes.RowCount <= 0 Then
      GoTo CleanUp
    End If

    'RETRIEVE OR OPEN THE SELECTED ENVISION EXCEL FILE
    Try
      If m_strEnvisionExcelFile = "" Or m_dockEnvisionWinForm.dgvDevelopmentTypes.Rows.Count <= 0 Then
        WriteRGB2EnvisionExcelFile = False
        Exit Function
      End If
      Try
        If Not TypeOf m_xlPaintWB1.Sheets("SCENARIO" & m_intEditScenario.ToString) Is Microsoft.Office.Interop.Excel.Worksheet Then
        End If
      Catch ex As Exception
        CloseEnvisionExcel()
        OpenEnvisionExcel()
      End Try
      If m_xlPaintApp Is Nothing Then
        m_xlPaintApp = New Microsoft.Office.Interop.Excel.Application
        m_xlPaintApp.DisplayAlerts = False
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

    If m_xlPaintWB1 Is Nothing Or m_xlPaintApp Is Nothing Then
      GoTo CloseExcel
    Else
      Try
        'RETRIEVE DEVELOPMENT TYPES TAB
        Try
          If Not TypeOf m_xlPaintWB1.Sheets("Dev Type Attributes") Is Microsoft.Office.Interop.Excel.Worksheet Then
            MessageBox.Show("The 'Devt' tab could not be found in the selected ENVISION Excel file. Please select another file.", "SCENARIO NOT FOUND", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo CloseExcel
          Else
            shtDevType = DirectCast(m_xlPaintWB1.Sheets("Dev Type Attributes"), Microsoft.Office.Interop.Excel.Worksheet)
          End If
        Catch ex As Exception
          MessageBox.Show("The 'Dev Type Attributes' tab could not be found in the selected ENVISION Excel file. Please select another file.", "SCENARIO NOT FOUND", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
          GoTo CloseExcel
        End Try

        'REVIEW DEV TYPES TAB ROWS TO DETERMINE THE START ROW
        For intRow = 1 To 35
          Try
            strFldCellValue = CStr(CType(shtDevType.Cells(intRow, 1), Microsoft.Office.Interop.Excel.Range).Value)
            If strFldCellValue.Contains(",") Then
              strFldCellValue = CStr(strFldCellValue.Split(","c)(0))
            End If
          Catch ex As Exception
            strFldCellValue = ""
          End Try
          If UCase(strFldCellValue) = "DEV_TYPE" Then
            intStartRow = intRow
            Exit For
          End If
        Next


        'POPULATE THE EXCEL FROM DATA GRID....ASSUMES THE DEVELOPMENT TYPE ORDER REMAINS STATIC
        'FIND THE SYMBOL COLOR FIELDS
        intRedFld = -1
        intGreenFld = -1
        intBlueFld = -1
        For intCol = 1 To 400
          Try
            strFldCellValue = CStr(CType(shtDevType.Cells(intStartRow, intCol), Microsoft.Office.Interop.Excel.Range).Value)
            If strFldCellValue = String.Empty Then Continue For
            If strFldCellValue.Contains(",") Then
              strFldCellValue = CStr(strFldCellValue.Split(","c)(0))
            End If
          Catch ex As Exception

          End Try
          If UCase(strFldCellValue) = "RED" Then
            intRedFld = intCol
          End If
          If UCase(strFldCellValue) = "GREEN" Then
            intGreenFld = intCol
          End If
          If UCase(strFldCellValue) = "BLUE" Then
            intBlueFld = intCol
          End If
          If intRedFld >= 0 And intGreenFld >= 0 And intBlueFld >= 0 Then
            Exit For
          End If
        Next

        For intRow = 1 To (m_dockEnvisionWinForm.dgvDevelopmentTypes.RowCount - 1)
          If Not m_tblDevelopmentTypes Is Nothing Then
            rowTemp = m_tblDevelopmentTypes.GetRow(intRow)
            If intRedFld >= 0 Then
              intRed = CInt(rowTemp.Value(m_tblDevelopmentTypes.FindField("RED")))
              CType(shtDevType.Cells(intRow + intStartRow, intRedFld), Microsoft.Office.Interop.Excel.Range).Value = intRed
            End If
            If intGreenFld >= 0 Then
              intGreen = CInt(rowTemp.Value(m_tblDevelopmentTypes.FindField("GREEN")))
              CType(shtDevType.Cells(intRow + intStartRow, intGreenFld), Microsoft.Office.Interop.Excel.Range).Value = intGreen
            End If
            If intBlueFld >= 0 Then
              intBlue = CInt(rowTemp.Value(m_tblDevelopmentTypes.FindField("BLUE")))
              CType(shtDevType.Cells(intRow + intStartRow, intBlueFld), Microsoft.Office.Interop.Excel.Range).Value = intBlue
            End If
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
        'If Not m_blnAutoUpdate Then
        '    GoTo CloseExcel
        'End If
        GoTo CleanUp
      Catch ex As Exception
        MessageBox.Show(ex.Message, "SAVE TO ENVISION EXCEL FILE ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
        GoTo CloseExcel
      End Try
    End If

CloseExcel:
    CloseEnvisionExcel()
    GoTo CleanUp

CleanUp:
    arrVariables = Nothing
    shtScenario = Nothing
    shtDevType = Nothing
    shtReport = Nothing
    intRow = Nothing
    strCellValue = Nothing
    strLyrPath = Nothing
    strLyrName = Nothing
    rowTemp = Nothing
    intValue = Nothing
    dblValue = Nothing
    tblsc = Nothing
    intCol = Nothing
    intRedFld = Nothing
    intGreenFld = Nothing
    intBlueFld = Nothing
    strFldCellValue = Nothing
    intRed = Nothing
    intGreen = Nothing
    intExcelFormulaCalc = Nothing
    GC.WaitForPendingFinalizers()
    GC.Collect()
  End Function

  Public Function Write2EnvisionExcelFile() As Boolean
    'SET THE DEVELOPMENT TYPE NAME, GRID CODE, 
    'RGB COLOR VALUES, AND ACRE SUMS TO THE SELECTED ENVISION EXCEL FILE
    m_appEnvision.StatusBar.Message(0) = "Executing function to WRITE Dev Type values to the selected Envision Excel file"
    Dim arrVariables As ArrayList = New ArrayList
    Dim shtScenario As Microsoft.Office.Interop.Excel.Worksheet = Nothing
    Dim intRec As Integer
    Dim strCellValue As String = ""
    Dim strLyrPath As String = ""
    Dim strLyrName As String = ""
    Dim rowTemp As IRow = Nothing
    Dim intValue As Integer = 0
    Dim dblValue As Double = 0
    Dim tblsc As ITable = Nothing
    Dim intRow As Integer
    Dim intCol As Integer = 0
    Dim strFldCellValue As String = ""
    Dim intScTabStartRow As Integer = 4
    Dim intScTabFieldRow As Integer = -1
    Dim intRowCount As Integer = 0
    Dim strDevType As String = ""
    Dim strFieldName As String = ""
    Dim arrFields As ArrayList = New ArrayList
    Dim intTblRow As Integer
    Dim intFieldRow As Integer = -1
    Dim intStartRow As Integer = -1
    Dim shtExisting As Microsoft.Office.Interop.Excel.Worksheet = Nothing
    Dim strDevdAcresAllFld As String = ""
    Dim strDevdAcresPaintedFld As String = ""
    Dim intExLUCol As Integer = 2
    Dim intDevdAcresAllCol As Integer = -1
    Dim intDevdAcresPaintedCol As Integer = -1
    Dim intExcelFormulaCalc As Integer = 0

    ''RETRIEVE OR OPEN THE SELECTED ENVISION EXCEL FILE
    'Try
    '    m_appEnvision.StatusBar.Message(0) = "Retrieving Envision Excel File"
    '    If m_strEnvisionExcelFile = "" Or m_dockEnvisionWinForm.dgvDevelopmentTypes.Rows.Count <= 0 Then
    '        Write2EnvisionExcelFile = False
    '        Exit Function
    '    End If
    '    Try
    '        If Not TypeOf m_xlPaintWB1.Sheets("SCENARIO" & m_intEditScenario.ToString) Is Microsoft.Office.Interop.Excel.Worksheet Then
    '        End If
    '    Catch ex As Exception
    '        CloseEnvisionExcel()
    '        OpenEnvisionExcel()
    '    End Try
    '    If m_xlPaintApp Is Nothing Then
    '        m_xlPaintApp = New Microsoft.Office.Interop.Excel.Application
    '        m_xlPaintApp.DisplayAlerts = false
    '        m_xlPaintApp.WindowState = Interop.Excel.XlWindowState.xlMinimized
    '        m_xlPaintApp.Visible = True
    '    End If
    '    If m_xlPaintWB1 Is Nothing Then
    '        m_xlPaintWB1 = CType(m_xlPaintApp.Workbooks.Open(m_strEnvisionExcelFile), Microsoft.Office.Interop.Excel.Workbook)
    '    End If
    'Catch ex As Exception
    '    MessageBox.Show(ex.Message, "Opening Envision Excel File Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
    '    GoTo CleanUp
    'End Try

    If m_xlPaintWB1 Is Nothing Or m_xlPaintApp Is Nothing Then
      GoTo CleanUp
    Else
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

      Try
        'RETRIEVE SCENARIO TAB
        m_appEnvision.StatusBar.Message(0) = "Retrieving current Scenario tab: " & "SCENARIO" & m_intEditScenario.ToString
        Try
          If Not TypeOf m_xlPaintWB1.Sheets("SCENARIO" & m_intEditScenario.ToString) Is Microsoft.Office.Interop.Excel.Worksheet Then
            MessageBox.Show("The 'SCENARIO" & m_intEditScenario.ToString & "' tab could not be found in the selected ENVISION Excel file. Please select another file.", "SCENARIO TAB NOT FOUND", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo CleanUp
          Else
            shtScenario = DirectCast(m_xlPaintWB1.Sheets("SCENARIO" & m_intEditScenario.ToString), Microsoft.Office.Interop.Excel.Worksheet)
          End If
        Catch ex As Exception
          MessageBox.Show("The 'SCENARIO" & m_intEditScenario.ToString & "' tab could not be found in the selected ENVISION Excel file. Please select another file.", "SCENARIO TAB NOT FOUND", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
          GoTo CleanUp
        End Try

        'RETRIEVE THE CORRESPONDING SCENARIO LOOKUP TABLE 
        If m_intEditScenario = 1 Then
          tblsc = m_tblSc1
        ElseIf m_intEditScenario = 2 Then
          tblsc = m_tblSc2
        ElseIf m_intEditScenario = 3 Then
          tblsc = m_tblSc3
        ElseIf m_intEditScenario = 4 Then
          tblsc = m_tblSc4
        ElseIf m_intEditScenario = 5 Then
          tblsc = m_tblSc5
        End If
        If tblsc Is Nothing Then
          GoTo CleanUp
        Else
          intRowCount = tblsc.RowCount(Nothing)
        End If

        'MAKE SURE THE PROCESSING ROW COUNT DOES NOT EXCEED THE MXIMUM DEV TYPE COUNT
        If intRowCount > m_intDevTypeMax Then
          intRowCount = m_intDevTypeMax
        End If

        'REVIEW SCENARIO TAB ROWS TO DETERMINE THE START ROW AND FIELD ROW (IF AVAILABLE)
        For intRow = 1 To m_intDevTypeMax
          Try
            strFldCellValue = CStr(CType(shtScenario.Cells(intRow, 1), Microsoft.Office.Interop.Excel.Range).Value)
            If strFldCellValue.Contains(",") Then
              strFldCellValue = CStr(strFldCellValue.Split(","c)(0))
            End If
          Catch ex As Exception
            strFldCellValue = ""
          End Try
          If UCase(strFldCellValue) = "DEV_TYPE" Then
            intScTabFieldRow = intRow
            intScTabStartRow = intRow + 1
            Exit For
          End If
        Next

        'BUILD LIST OF COLUMNS TO PROCESS
        If intScTabFieldRow <> -1 Then
          For intCol = 2 To 400
            Try
              strFieldName = CStr(CType(shtScenario.Cells(intScTabFieldRow, intCol), Microsoft.Office.Interop.Excel.Range).Value)
              If strFieldName <> "" Then
                If tblsc.FindField(strFieldName) <> -1 Then
                  arrFields.Add(intCol)
                End If
              Else
                Continue For
              End If
            Catch ex As Exception
            End Try
          Next

          'POPULATE THE EXCEL FROM DATA GRID....ASSUMES THE DEVELOPMENT TYPE ORDER REMAINS STATIC
          m_appEnvision.StatusBar.Message(0) = "Writing values to Scenario tab: " & "SCENARIO" & m_intEditScenario.ToString
          For intRec = 1 To intRowCount
            rowTemp = tblsc.GetRow(intRec)
            'ASSUMPTION IS THE DEV TYPES ARE IN THE SAME ORDER
            For Each intCol In arrFields
              dblValue = 0
              Try
                strFieldName = CStr(CType(shtScenario.Cells(intScTabFieldRow, intCol), Microsoft.Office.Interop.Excel.Range).Value)
                If Not strFieldName = "" Then
                  If tblsc.FindField(strFieldName) >= 0 Then
                    dblValue = CDbl(rowTemp.Value(tblsc.FindField(strFieldName)))
                    dblValue = Round(dblValue, 6)
                    CType(shtScenario.Cells(intScTabStartRow, intCol), Microsoft.Office.Interop.Excel.Range).Value = dblValue
                  End If
                Else
                  Continue For
                End If
              Catch ex As Exception
                CType(shtScenario.Cells(intScTabStartRow, intCol), Microsoft.Office.Interop.Excel.Range).Value = 0
              End Try
              intCol = intCol + 1
            Next
            intScTabStartRow = intScTabStartRow + 1
          Next



        Else


          'OLD ENVISION EXCEL WRITE METHOD
          m_appEnvision.StatusBar.Message(0) = "Populating VAC_ACRE and DEVD_ACRE columns OLD Envision Excel file methodology"
          intScTabStartRow = 5
          For intRec = 1 To intRowCount
            rowTemp = tblsc.GetRow(intRec)
            dblValue = 0
            Try
              If tblsc.FindField("VAC_ACRE") >= 0 Then
                dblValue = CDbl(rowTemp.Value(tblsc.FindField("VAC_ACRE")))
                dblValue = Round(dblValue, 6)
                CType(shtScenario.Cells(intScTabStartRow, 2), Microsoft.Office.Interop.Excel.Range).Value = dblValue
              End If
            Catch ex As Exception
              CType(shtScenario.Cells(intScTabStartRow, 2), Microsoft.Office.Interop.Excel.Range).Value = 0
            End Try
            dblValue = 0
            Try
              If tblsc.FindField("DEVD_ACRE") >= 0 Then
                dblValue = CDbl(rowTemp.Value(tblsc.FindField("DEVD_ACRE")))
                dblValue = Round(dblValue, 6)
                CType(shtScenario.Cells(intScTabStartRow, 3), Microsoft.Office.Interop.Excel.Range).Value = dblValue
              End If
            Catch ex As Exception
              CType(shtScenario.Cells(intScTabStartRow, 3), Microsoft.Office.Interop.Excel.Range).Value = 0
            End Try
            intScTabStartRow = intScTabStartRow + 1
          Next
        End If

        'RETRIEVE OR OPEN THE SELECTED ENVISION EXCEL FILE
        m_appEnvision.StatusBar.Message(0) = "Populating Envision Existing Developed Area tab"
        Try
          If m_xlPaintWB1 Is Nothing Then
            GoTo CleanUp
          End If
          Try
            shtExisting = DirectCast(m_xlPaintWB1.Sheets("Existing Developed Area"), Microsoft.Office.Interop.Excel.Worksheet)

            If shtExisting Is Nothing Then
              GoTo CleanUp
            End If

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
              GoTo CleanUp
            End If
            'CYCLE THROUGH THE FIRST 25 COULMNS FOR THE INPUT FIELDS, FIRST FOUND IS EXISTING LAND USE FIELD, SECOND IS DEVELOPED ACREAS FOR ALL FEATURES, THIRD IS DEVELOPED ACREAS FOR PAINTED FEATURES
            For intCol = 1 To 25
              strFldCellValue = CStr(CType(shtExisting.Cells(intFieldRow, intCol), Microsoft.Office.Interop.Excel.Range).Value)
              If Not strFldCellValue Is Nothing Then
                If UCase(strFldCellValue) = ("PAINTED" & m_intEditScenario.ToString) And strFldCellValue.Length > 0 Then
                  intDevdAcresPaintedCol = intCol
                  'Exit For
                End If
                If UCase(strFldCellValue) = ("EX_LU") Then
                  intExLUCol = intCol
                  'Exit For
                End If
              End If
            Next
            If intDevdAcresPaintedCol = -1 Then
              GoTo CleanUp
            End If

            'CLEAR VALUES
            m_appEnvision.StatusBar.Message(0) = "Clearing Painted / Redeveloped Acre values for Existing Developed Areas"
            For intRow = intStartRow To (intStartRow + 20)
              CType(shtExisting.Cells(intRow, intDevdAcresPaintedCol), Microsoft.Office.Interop.Excel.Range).Value = ""
            Next

            'CODE TO WRITE THE SUMMARY VALUES TO EXCEL FILE
            m_appEnvision.StatusBar.Message(0) = "Writing Painted / Redeveloped Acre values for Existing Developed Areas"
            Try
              For intRow = intStartRow To (intStartRow + 20)
                strFldCellValue = CStr(CType(shtExisting.Cells(intRow, intExLUCol), Microsoft.Office.Interop.Excel.Range).Value)
                If Not strFldCellValue Is Nothing Then
                  'PAINTED FEATURES
                  If Not m_tblExistingLU Is Nothing Then
                    For intTblRow = 1 To m_tblExistingLU.RowCount(Nothing)
                      Try
                        If strFldCellValue = CStr(m_tblExistingLU.GetRow(intTblRow).Value(1)) Then
                          CType(shtExisting.Cells(intRow, intDevdAcresPaintedCol), Microsoft.Office.Interop.Excel.Range).Value = m_tblExistingLU.GetRow(intTblRow).Value(2)
                        End If
                      Catch ex As Exception

                      End Try
                    Next
                  End If
                End If
              Next
            Catch ex As Exception
              GoTo CleanUp
            End Try
          Catch ex As Exception
            GoTo CleanUp
          End Try
        Catch ex As Exception
          GoTo CleanUp
        End Try

        'RESET FORMULA CALC SETTING
        If intExcelFormulaCalc = 1 Then
          m_xlPaintApp.Calculation = Interop.Excel.XlCalculation.xlCalculationAutomatic
        ElseIf intExcelFormulaCalc = 2 Then
          m_xlPaintApp.Calculation = Interop.Excel.XlCalculation.xlCalculationManual
        ElseIf intExcelFormulaCalc = 3 Then
          m_xlPaintApp.Calculation = Interop.Excel.XlCalculation.xlCalculationSemiautomatic
        End If

        'm_xlPaintWB1.Save()

        GoTo CleanUp
      Catch ex As Exception
        MessageBox.Show(ex.Message, "SAVE TO ENVISION EXCEL FILE ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
        GoTo CleanUp
      End Try
    End If


CleanUp:
    arrVariables = Nothing
    arrVariables = Nothing
    shtScenario = Nothing
    intRec = Nothing
    strCellValue = Nothing
    strLyrPath = Nothing
    strLyrName = Nothing
    rowTemp = Nothing
    intValue = Nothing
    dblValue = Nothing
    tblsc = Nothing
    intRow = Nothing
    intCol = Nothing
    strFldCellValue = Nothing
    intScTabStartRow = Nothing
    intScTabFieldRow = Nothing
    intRowCount = Nothing
    strDevType = Nothing
    strFieldName = Nothing
    arrFields = Nothing
    intTblRow = Nothing
    intFieldRow = Nothing
    intStartRow = Nothing
    shtExisting = Nothing
    strDevdAcresAllFld = Nothing
    strDevdAcresPaintedFld = Nothing
    intExLUCol = Nothing
    intDevdAcresAllCol = Nothing
    intDevdAcresPaintedCol = Nothing
    intExcelFormulaCalc = Nothing
    GC.Collect()
    GC.WaitForPendingFinalizers()
  End Function

  Public Function WriteValues(ByVal pPoint As IPoint, ByVal pPolygon As IPolygon, ByVal pPolyline As IPolyline) As Boolean
    WriteValues = True
    m_blnEditProcessing = True

    If m_pEditFeatureLyr Is Nothing Then
      WriteValues = False
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
    Dim pArea As IArea
    Dim pFeatClass As IFeatureClass = m_pEditFeatureLyr.FeatureClass
    Dim pTable As ITable = CType(pFeatClass, ITable)
    Dim dblShpAcres As Double
    Dim strAcres As String
    Dim pTopoOp As ITopologicalOperator

    Dim mapCurrent As Map
    Dim pFeatLayer As IFeatureLayer
    Dim intLayer As Integer

    '*********************************************************
    'MAKE THE FEATURE SELECTION
    '*********************************************************
    mxApplication = CType(m_appEnvision, IMxApplication)
    mxDoc = CType(m_appEnvision.Document, IMxDocument)
    With pSpatFilter
      .GeometryField = m_pEditFeatureLyr.FeatureClass.ShapeFieldName
      .SpatialRel = esriSpatialRelEnum.esriSpatialRelIntersects
      .SubFields = m_pEditFeatureLyr.FeatureClass.ShapeFieldName
    End With
    If Not pPoint Is Nothing Or Not pPolygon Is Nothing Or Not pPolyline Is Nothing Then
      If Not pPoint Is Nothing Then
        pSpatFilter.Geometry = pPoint
        pREnv = pPoint.Envelope
      End If
      If Not pPolygon Is Nothing Then
        pPolygon.Densify(30, 0)
        pPolygon.SimplifyPreserveFromTo()
        pTopoOp = DirectCast(pPolygon, ITopologicalOperator)
        pTopoOp.Simplify()
        pSpatFilter.Geometry = pPolygon
        pREnv = pPolygon.Envelope
        pArea = DirectCast(pPolygon, IArea)
        dblShpAcres = ReturnAcres(pArea.Area, mxDoc.FocusMap.MapUnits.ToString)
        strAcres = dblShpAcres.ToString("#.####")
        If mxDoc.FocusMap.MapUnits = esriUnits.esriUnknownUnits Then
          m_appEnvision.StatusBar.Message(0) = "Selected Feature Area (" + strAcres + ")"
        Else
          m_appEnvision.StatusBar.Message(0) = "Selected Feature Area (" + strAcres + " acres)"
        End If
      End If
      If Not pPolyline Is Nothing Then
        pSpatFilter.Geometry = pPolyline
        pREnv = pPolyline.Envelope
      End If
    End If

    pFeatSelection = CType(m_pEditFeatureLyr, IFeatureSelection)
    Dim queryFilter As IQueryFilter = New QueryFilter
    queryFilter.SubFields = m_pEditFeatureLyr.FeatureClass.ShapeFieldName
    If Not pPoint Is Nothing Or Not pPolygon Is Nothing Or Not pPolyline Is Nothing Then
      pFeatSelection.SelectFeatures(pSpatFilter, esriSelectionResultEnum.esriSelectionResultNew, False)
      pFeatSelection.SelectionSet.Search(queryFilter, False, pCursor)
    Else
      pFeatSelection.SelectionSet.Search(queryFilter, False, pCursor)
    End If
    pFeatureCursor = DirectCast(pCursor, IFeatureCursor)

    Dim dblTotalAreas As Double = 0
    If pFeatSelection.SelectionSet.Count > 0 Then
      pFeat = pFeatureCursor.NextFeature
      pREnv = Nothing
      pREnv = pFeat.Shape.Envelope
      Do While Not pFeat Is Nothing
        pArea = DirectCast(pFeat.Shape, IArea)
        dblTotalAreas = dblTotalAreas + pArea.Area
        pFeat = pFeatureCursor.NextFeature
        If Not pFeat Is Nothing Then
          pREnv.Union(pFeat.Shape.Envelope)
        End If
      Loop
    End If

    If WriteSubtractDevType() Then
      WriteAddDevType()
    Else
      GoTo CleanUp
    End If

    'WRITE THE VALUES TO THE SELECTED ENVISION EXCEL FILE
    If m_blnAutoUpdate Then
      Write2EnvisionExcelFile()
    End If

    If Not m_envSelect Is Nothing Then
      mxDoc.ActiveView.PartialRefresh(esriViewDrawPhase.esriViewGeography Or esriViewDrawPhase.esriViewGraphics, Nothing, m_envSelect.Envelope)
    End If
    m_envSelect = pREnv

    mxDoc.ActiveView.Refresh()

    dblShpAcres = ReturnAcres(dblTotalAreas, mxDoc.FocusMap.MapUnits.ToString)
    strAcres = dblShpAcres.ToString
    strAcres = dblShpAcres.ToString("#.####")
    If mxDoc.FocusMap.MapUnits = esriUnits.esriUnknownUnits Then
      m_appEnvision.StatusBar.Message(0) = "Selected Feature Area (" + strAcres + ")"
    Else
      m_appEnvision.StatusBar.Message(0) = "Selected Feature Area (" + strAcres + " acres)"
    End If

CleanUp:
    mxApplication = Nothing
    mxDoc = Nothing
    pREnv = Nothing
    pSpatFilter = Nothing
    pFeatSelection = Nothing
    pCursor = Nothing
    pFeatureCursor = Nothing
    pFeat = Nothing
    pArea = Nothing
    pFeatClass = Nothing
    pTable = Nothing
    mapCurrent = Nothing
    pFeatLayer = Nothing
    intLayer = Nothing
    m_dockEnvisionWinForm.prgBarEnvision.Value = 0
    GC.WaitForPendingFinalizers()
    GC.Collect()

  End Function

  Public Function WriteSubtractDevType() As Boolean
    m_appEnvision.StatusBar.Message(0) = "Executing function to SUBTRACT previously assigned Dev Type to selected feature(s)"
    System.Windows.Forms.Application.DoEvents()
    Dim pCursor As ICursor = Nothing
    Dim cursorEX_LU As ICursor = Nothing
    Dim cursorScLookup As ICursor = Nothing
    WriteSubtractDevType = True
    If Not m_blnRunCalcs Then
      GoTo CleanUp
    End If

    Dim pFeatSelection As IFeatureSelection
    Dim pFeatureCursor As IFeatureCursor
    Dim pFeat As IFeature
    Dim pFeatClass As IFeatureClass = m_pEditFeatureLyr.FeatureClass
    Dim pFeatTable As ITable = CType(pFeatClass, ITable)
    Dim intTotalCount As Integer
    Dim intCount As Integer
    Dim dblDevelopedAcres As Double
    Dim intDevTypeFld As Integer = -1
    Dim intVacantAcresFld As Integer = -1
    Dim intDevelopedAcresFld As Integer = -1
    Dim intEXLUFeatFld As Integer = -1
    Dim intEXLUFld As Integer = -1
    Dim intSumDEVDACREfld As Integer = -1
    Dim strEX_LU As String = ""
    Dim dblEX_LUAcres As Double = 0
    Dim rowEX_LU As IRow
    Dim intDevTypeRow As Integer = -1
    Dim intRow As Integer
    'Dim rowTemp As IRow
    Dim tblScLookUpTbl As ITable = Nothing
    Dim intScLookUpFld As Integer = 0
    Dim fldSc As IField
    'Dim strFieldName As String = ""
    Dim dblPreviousValue As Double = 0
    Dim dblValue As Double 'Dim intScVacLookUpFld As Integer
    Dim rowScLookUp As IRow
    Dim existingLUDict As Dictionary(Of String, Double) = New Dictionary(Of String, Double)
    Dim devTypeDict As Dictionary(Of String, Dictionary(Of String, Double)) = New Dictionary(Of String, Dictionary(Of String, Double))

    intVacantAcresFld = pFeatClass.FindField("VAC_ACRE")
    intDevelopedAcresFld = pFeatClass.FindField("DEVD_ACRE")
    intDevTypeFld = pFeatClass.FindField("Dev_Type")
    intEXLUFeatFld = pFeatClass.FindField("EX_LU")
    If m_tblExistingLU IsNot Nothing Then
      intEXLUFld = m_tblExistingLU.FindField("EX_LU")
      intSumDEVDACREfld = m_tblExistingLU.FindField("Sum_DEVD_ACRE")
    End If

    '*********************************************************
    'MAKE THE FEATURE SELECTION
    '*********************************************************
    pFeatSelection = CType(m_pEditFeatureLyr, IFeatureSelection)
    pFeatSelection.SelectionSet.Search(Nothing, True, pCursor)
    pFeatureCursor = DirectCast(pCursor, IFeatureCursor)

    ' Get the row count for the existing LU table
    Dim existingLUCount As Int32
    If m_tblExistingLU IsNot Nothing Then existingLUCount = m_tblExistingLU.RowCount(Nothing)

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
      WriteSubtractDevType = False
      GoTo CleanUp
    End If

    intTotalCount = pFeatSelection.SelectionSet.Count
    If intTotalCount > 0 Then
      Try
        ' Make a list of scenario fields
        Dim scenarioFldList As List(Of String) = New List(Of String)
        For intScLookUpFld = 1 To (tblScLookUpTbl.Fields.FieldCount - 1)
          fldSc = tblScLookUpTbl.Fields.Field(intScLookUpFld)
          If fldSc.Type = esriFieldType.esriFieldTypeDouble Then
            If m_pEditFeatureLyr.FeatureClass.FindField(fldSc.Name) <> -1 Then
              scenarioFldList.Add(fldSc.Name)
            End If
          End If
        Next

        ' Loop through the features
        pFeat = pFeatureCursor.NextFeature
        Do While Not pFeat Is Nothing
          intCount += 1

          'RETRIEVE THE DEVELOPMENT TYPES
          Dim strDevType As String = String.Empty
          If intDevTypeFld > -1 Then
            Try
              strDevType = CStr(pFeat.Value(intDevTypeFld)).Trim

              'SKIP OTHER SCRIPTING IF NO DEV TYPE ASSIGNED
              If strDevType = String.Empty Then
                pFeat = pFeatureCursor.NextFeature
                Continue Do
              End If
            Catch ex As Exception
              pFeat = pFeatureCursor.NextFeature
              Continue Do
            End Try
          End If

          'RETRIEVE DEVELOPED ACRES
          Try
            dblDevelopedAcres = CType((pFeat.Value(intDevelopedAcresFld).ToString), Double)
            If dblDevelopedAcres < 0 Then
              dblDevelopedAcres = 0
            End If
          Catch ex As Exception
            dblDevelopedAcres = 0
          End Try
          dblDevelopedAcres = Round(dblDevelopedAcres, 6)

          ' Add to the dictionary by LU
          If m_tblExistingLU IsNot Nothing AndAlso intEXLUFeatFld <> -1 Then
            Try
              strEX_LU = pFeat.Value(intEXLUFeatFld).ToString
              If strEX_LU.Length > 0 Then
                If existingLUDict.ContainsKey(strEX_LU) Then
                  existingLUDict(strEX_LU) = existingLUDict(strEX_LU) + dblDevelopedAcres
                Else
                  existingLUDict.Add(strEX_LU, dblDevelopedAcres)
                End If
              End If
            Catch ex As Exception
            End Try
          End If

          ' Create a new dev type dictionary item or get the existing one
          Dim valsDict As Dictionary(Of String, Double)
          If Not devTypeDict.ContainsKey(strDevType) Then
            valsDict = New Dictionary(Of String, Double)
            For Each fldName As String In scenarioFldList
              valsDict.Add(fldName, 0)
            Next
            devTypeDict.Add(strDevType, valsDict)
          Else
            valsDict = devTypeDict(strDevType)
          End If

          ' Add the values to the dev type dictionary item
          For Each fldName As String In scenarioFldList
            If pFeat.Value(pFeat.Fields.FindField(fldName)) IsNot DBNull.Value Then
              dblValue = CType(pFeat.Value(pFeat.Fields.FindField(fldName)), Double)
            Else
              dblValue = 0
            End If
            valsDict(fldName) = valsDict(fldName) + dblValue
          Next

          ''RETRIEVE THE RECORD NUMBER
          'intDevTypeRow = -1
          'rowTemp = Nothing
          'For intRow = 1 To m_intDevTypeMax
          '    rowTemp = m_tblDevelopmentTypes.GetRow(intRow)
          '    If CStr(rowTemp.Value(m_tblDevelopmentTypes.FindField("DEVELOPMENT_TYPE"))) = strDevType Then
          '        intDevTypeRow = intRow
          '        rowTemp = Nothing
          '        Exit For
          '    End If
          'Next

          ''EXIT THE SCRIPT IF THE MATCHING DEVELOPMENT TYPE IS NOT FOUND....NOTHING TO SUBTRACT
          'If intDevTypeRow = -1 Then
          '    pFeat = pFeatureCursor.NextFeature
          '    Continue Do
          'End If

          ''REMOVE THE ACRES FROM THE SCENARIO LOOKUP TABLES
          'For intScLookUpFld = 1 To (tblScLookUpTbl.Fields.FieldCount - 1)
          '    fldSc = tblScLookUpTbl.Fields.Field(intScLookUpFld)
          '    strFieldName = fldSc.Name
          '    If fldSc.Type = esriFieldType.esriFieldTypeDouble Then
          '        If m_pEditFeatureLyr.FeatureClass.FindField(strFieldName) >= 0 Then
          '            dblValue = 0
          '            Try
          '                dblValue = CType(pFeat.Value(pFeat.Fields.FindField(strFieldName)), Double)
          '            Catch ex As Exception
          '                dblValue = 0
          '            End Try
          '        End If

          '        'ADD VALUES TO SCENARIO TABLE
          '        If intDevTypeRow >= 0 Then
          '            dblPreviousValue = 0
          '            rowScLookUp = tblScLookUpTbl.GetRow(intDevTypeRow)
          '            Try
          '                dblPreviousValue = CType(rowScLookUp.Value(intScLookUpFld), Double)
          '            Catch ex As Exception
          '                dblPreviousValue = 0
          '            End Try
          '            If dblPreviousValue >= 0 Then
          '                rowScLookUp.Value(intScLookUpFld) = dblPreviousValue - dblValue
          '            Else
          '                rowScLookUp.Value(intScLookUpFld) = dblValue
          '            End If
          '            rowScLookUp.Store()
          '        End If
          '    End If
          'Next

          pFeat = pFeatureCursor.NextFeature
          m_dockEnvisionWinForm.prgBarEnvision.Value = CType((intCount / intTotalCount) * 100, Int32)
          m_dockEnvisionWinForm.prgBarEnvision.Refresh()
        Loop

        'SUBTRACT DEVD ACRES FROM EXISTING LU PAINT TABLE
        If existingLUDict.Count <> 0 Then
          Try
            cursorEX_LU = m_tblExistingLU.Search(Nothing, False)
            rowEX_LU = cursorEX_LU.NextRow
            Do While rowEX_LU IsNot Nothing
              If rowEX_LU.Value(intEXLUFld) IsNot DBNull.Value Then
                Dim val As String = rowEX_LU.Value(intEXLUFld).ToString
                If val <> String.Empty AndAlso existingLUDict.ContainsKey(val) Then
                  dblEX_LUAcres = CType(rowEX_LU.Value(intSumDEVDACREfld), Double)
                  rowEX_LU.Value(intSumDEVDACREfld) = dblEX_LUAcres - existingLUDict(val)
                  rowEX_LU.Store()
                End If
              End If
              rowEX_LU = cursorEX_LU.NextRow
            Loop
          Catch ex As Exception
          End Try
        End If

        ' Subtract values from the scenario table
        If devTypeDict.Count <> 0 Then
          Try
            Dim intScLookupDevTypeFld As Int32 = tblScLookUpTbl.FindField("DEVELOPMENT_TYPE")
            cursorScLookup = tblScLookUpTbl.Search(Nothing, False)
            rowScLookUp = cursorScLookup.NextRow
            Do While rowScLookUp IsNot Nothing
              If rowScLookUp.Value(intScLookupDevTypeFld) IsNot DBNull.Value AndAlso rowScLookUp.Value(intScLookupDevTypeFld).ToString <> String.Empty Then
                Dim devType As String = rowScLookUp.Value(intScLookupDevTypeFld).ToString
                If devTypeDict.ContainsKey(devType) Then
                  For Each kvp As KeyValuePair(Of String, Double) In devTypeDict(devType)
                    intScLookUpFld = tblScLookUpTbl.FindField(kvp.Key)
                    Try
                      dblPreviousValue = CType(rowScLookUp.Value(intScLookUpFld), Double)
                    Catch ex As Exception
                      dblPreviousValue = 0
                    End Try
                    rowScLookUp.Value(intScLookUpFld) = dblPreviousValue - kvp.Value
                  Next
                  rowScLookUp.Store()
                End If
              End If
              rowScLookUp = cursorScLookup.NextRow
            Loop
          Catch ex As Exception
          End Try
        End If

        GoTo CleanUp

      Catch ex As Exception
        MessageBox.Show(ex.Message, "SUBTRACT DEV TYPE SUB ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
        WriteSubtractDevType = False
        GoTo CleanUp
      End Try
    End If

CleanUp:
    pFeatSelection = Nothing
    'pCursor = Nothing
    'pFeatureCursor = Nothing
    If pCursor IsNot Nothing Then Marshal.FinalReleaseComObject(pCursor)
    If cursorEX_LU IsNot Nothing Then Marshal.FinalReleaseComObject(cursorEX_LU)
    If cursorScLookup IsNot Nothing Then Marshal.FinalReleaseComObject(cursorScLookup)
    pFeat = Nothing
    pFeatClass = Nothing
    pFeatTable = Nothing
    intTotalCount = Nothing
    intCount = Nothing
    dblDevelopedAcres = Nothing
    intDevTypeFld = Nothing
    intVacantAcresFld = Nothing
    intDevelopedAcresFld = Nothing
    intDevTypeRow = Nothing
    intRow = Nothing
    'rowTemp = Nothing
    rowScLookUp = Nothing
    tblScLookUpTbl = Nothing
    'strFieldName = Nothing
    dblPreviousValue = Nothing
    intScLookUpFld = Nothing
    dblValue = Nothing
    fldSc = Nothing

    GC.WaitForPendingFinalizers()
    GC.Collect()
  End Function

  Public Sub WriteAddDevType()
    m_appEnvision.StatusBar.Message(0) = "Executing sub routine to ADD Dev Type to selected feature(s)"
    System.Windows.Forms.Application.DoEvents()
    Dim pFeatSelection As IFeatureSelection
    Dim pCursor As ICursor = Nothing
    Dim cursorEX_LU As ICursor = Nothing
    Dim pFeatureCursor As IFeatureCursor
    Dim pFeat As IFeature
    Dim pFeatClass As IFeatureClass = m_pEditFeatureLyr.FeatureClass
    Dim pFeatTable As ITable = CType(pFeatClass, ITable)
    Dim intTotalCount As Integer
    Dim intCount As Integer
    Dim dblVacantAcres As Double = 0
    Dim dblDevelopedAcres As Double = 0
    Dim dblValue As Double = 0
    Dim intDevTypeFld As Integer = -1
    Dim intVacantAcresFld As Integer = -1
    Dim intDevelopedAcresFld As Integer = -1
    Dim intEXLUFeatFld As Integer = -1
    Dim intEXLUFld As Integer = -1
    Dim intSumDEVDACREfld As Integer = -1
    Dim intRedevFld As Integer = -1
    Dim intAbandonFld As Integer = -1
    Dim rowDevType As IRow = Nothing
    Dim intDevTypeRow As Integer = -1
    Dim intRow As Integer
    Dim rowTemp As IRow
    Dim intScVacLookUpFld As Integer = -1
    Dim intScDevdLookUpFld As Integer = -1
    Dim rowScLookUp As IRow = Nothing
    Dim dblScVacantAcres As Double
    Dim dblScDevelopedAcres As Double
    Dim tblScLookUpTbl As ITable = Nothing
    Dim strFromFieldName As String = ""
    Dim strToFieldName As String = ""
    Dim dblCalcAcres As Double = 0
    Dim dblRedevRate As Double = 0
    Dim dblAbandonRate As Double = 0
    'Dim dblExistingXValue As Double = 0
    Dim dblXValue As Double = 0
    Dim intFldCount As Integer = 0
    Dim dblScExValue As Double
    Dim intScExField As Integer
    Dim intScLookUpFld As Integer = 0
    Dim strFieldName As String = ""
    Dim dblPreviousValue As Double = 0
    Dim fldSc As IField
    Dim strEX_LU As String = ""
    Dim dblEX_LUAcres As Double = 0
    Dim rowEX_LU As IRow
    Dim intEX_LU_RowCount As Integer = 0
    Dim dblNetAcre As Double = 0
    Dim intNetAcreFld As Integer
    Dim strDepVarFieldName As String = ""
    Dim dblDepVar As Double = 0
    Dim existingLUDict As Dictionary(Of String, Double) = New Dictionary(Of String, Double)

    intDevTypeFld = pFeatClass.FindField("Dev_Type")
    intVacantAcresFld = pFeatClass.FindField("VAC_ACRE")
    intDevelopedAcresFld = pFeatClass.FindField("DEVD_ACRE")
    intEXLUFeatFld = pFeatClass.FindField("EX_LU")
    If m_tblExistingLU IsNot Nothing Then
      intEXLUFld = m_tblExistingLU.FindField("EX_LU")
      intSumDEVDACREfld = m_tblExistingLU.FindField("Sum_DEVD_ACRE")
    End If

    ' Get the row count for the existing LU table
    Dim existingLUCount As Int32
    If m_tblExistingLU IsNot Nothing Then existingLUCount = m_tblExistingLU.RowCount(Nothing)

    '*********************************************************
    'MAKE THE FEATURE SELECTION
    '*********************************************************
    pFeatSelection = CType(m_pEditFeatureLyr, IFeatureSelection)
    pFeatSelection.SelectionSet.Search(Nothing, False, pCursor)
    pFeatureCursor = DirectCast(pCursor, IFeatureCursor)

    intTotalCount = pFeatSelection.SelectionSet.Count
    If intTotalCount > 0 Then
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

        intScVacLookUpFld = tblScLookUpTbl.FindField("")
        intScDevdLookUpFld = tblScLookUpTbl.FindField("DEVD_ACRE")

        'RETRIEVE THE RECORD NUMBER AND ROW FOR THE SELECTED DEVELOPMENT TYPE
        rowTemp = Nothing
        rowDevType = Nothing
        dblRedevRate = 0
        dblNetAcre = 0
        dblAbandonRate = 0
        For intRow = 1 To m_intDevTypeMax
          rowTemp = m_tblDevelopmentTypes.GetRow(intRow)
          Try
            If CStr(rowTemp.Value(m_tblDevelopmentTypes.FindField("DEVELOPMENT_TYPE"))) = m_strDWriteValue Then
              intDevTypeRow = intRow
              rowDevType = rowTemp
              'intAbandonFld = m_tblDevelopmentTypes.FindField("ABANDON_RATE")
              'If intAbandonFld <> -1 Then
              '  dblAbandonRate = CType(rowDevType.Value(intAbandonFld), Double)
              'End If
              intRedevFld = m_tblDevelopmentTypes.FindField("REDEV_RATE")
              If intRedevFld <> -1 Then
                dblRedevRate = CType(rowDevType.Value(intRedevFld), Double)
              End If
              intNetAcreFld = m_tblDevelopmentTypes.FindField("NET_ACRE")
              If intNetAcreFld >= 0 Then
                dblNetAcre = CType(rowDevType.Value(intNetAcreFld), Double)
              End If
              rowScLookUp = tblScLookUpTbl.GetRow(intDevTypeRow)
              Exit For
            End If
          Catch ex As Exception
            GoTo CleanUp
          End Try
        Next

        ' Make a list of scenario fields and create a dictionary of total field values
        Dim valsDict As Dictionary(Of String, Double) = New Dictionary(Of String, Double)
        Dim scenarioFldList As List(Of String) = New List(Of String)
        For intScLookUpFld = 1 To (tblScLookUpTbl.Fields.FieldCount - 1)
          fldSc = tblScLookUpTbl.Fields.Field(intScLookUpFld)
          If fldSc.Type = esriFieldType.esriFieldTypeDouble Then
            If m_pEditFeatureLyr.FeatureClass.FindField(fldSc.Name) <> -1 Then
              scenarioFldList.Add(fldSc.Name)
              valsDict.Add(fldSc.Name, 0)
            End If
          End If
        Next

        ' Loop through the features
        pFeat = pFeatureCursor.NextFeature
        Do While Not pFeat Is Nothing
          intCount = intCount + 1
          'RETRIEVE VACANT AND DEVELOPED ACRES
          dblDevelopedAcres = 0
          dblVacantAcres = 0
          Try
            dblVacantAcres = CType((pFeat.Value(intVacantAcresFld).ToString), Double)
            If dblVacantAcres < 0 Then
              dblVacantAcres = 0
            End If
          Catch ex As Exception
            dblVacantAcres = 0
          End Try
          Try
            dblDevelopedAcres = CType((pFeat.Value(intDevelopedAcresFld).ToString), Double)
            If dblDevelopedAcres < 0 Then
              dblDevelopedAcres = 0
            End If
          Catch ex As Exception
            dblDevelopedAcres = 0
          End Try

          ' Add to the dictionary by LU
          If m_strDWriteValue <> String.Empty AndAlso m_tblExistingLU IsNot Nothing AndAlso intEXLUFeatFld <> -1 Then
            Try
              strEX_LU = pFeat.Value(intEXLUFeatFld).ToString
              If strEX_LU.Length > 0 Then
                If existingLUDict.ContainsKey(strEX_LU) Then
                  existingLUDict(strEX_LU) = existingLUDict(strEX_LU) + Round(dblDevelopedAcres, 6)
                Else
                  existingLUDict.Add(strEX_LU, Round(dblDevelopedAcres, 6))
                End If
              End If
            Catch ex As Exception
            End Try
          End If

          ' Add the field values to the values dictionary
          If m_strDWriteValue <> String.Empty Then
            For Each fldName As String In scenarioFldList
              If pFeat.Value(pFeat.Fields.FindField(fldName)) IsNot DBNull.Value Then
                dblValue = CType(pFeat.Value(pFeat.Fields.FindField(fldName)), Double)
              Else
                dblValue = 0
              End If
              valsDict(fldName) = valsDict(fldName) + dblValue
            Next
          End If

          'APPLY THE DEVELOPMENT TYPE ATTTIBUTES IF THE CALCS ARE ON AND THE DEV TYPE RECORD WAS FOUND
          If m_blnRunCalcs Then
            If m_arrWriteDevTypeFields.Count > 0 Then
              For intFldCount = 0 To m_arrWriteDevTypeFields.Count - 1
                strFromFieldName = m_arrWriteDevTypeFields.Item(intFldCount).ToString
                Try
                  If Not m_strDWriteValue = "" Then
                    pFeat.Value(pFeatClass.FindField(strFromFieldName)) = rowDevType.Value(m_tblDevelopmentTypes.FindField(strFromFieldName))
                  Else
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
                  End If
                Catch ex As Exception

                End Try
                If m_arrWriteDevTypeTotalFields.IndexOf(strFromFieldName) <> -1 Then
                  Dim existingFldIndex As Int32 = pFeatClass.FindField("EX_" & strFromFieldName)
                  If existingFldIndex <> -1 Then
                    Dim existingVal As Double = 0
                    If pFeat.Value(existingFldIndex) IsNot DBNull.Value Then
                      existingVal = CType(pFeat.Value(existingFldIndex), Double)
                    End If
                    If m_strDWriteValue <> String.Empty Then
                      dblXValue = CType(rowDevType.Value(m_tblDevelopmentTypes.FindField(strFromFieldName)), Double)
                    Else
                      dblXValue = 0
                      dblRedevRate = 0
                    End If
                    pFeat.Value(pFeatClass.FindField(m_arrWriteDevTypeTotalFieldName(m_arrWriteDevTypeTotalFields.IndexOf(strFromFieldName)).ToString)) = (dblXValue * dblRedevRate) + (existingVal * (1 - dblRedevRate))
                  End If
                End If
              Next
            End If
            If m_arrWriteDevTypeAcresFields.Count > 0 Then
              For intFldCount = 0 To m_arrWriteDevTypeAcresFields.Count - 1
                strFromFieldName = m_arrWriteDevTypeAcresFields.Item(intFldCount).ToString
                strToFieldName = m_arrWriteDevTypeAcresFieldsAltName.Item(intFldCount).ToString
                strDepVarFieldName = m_arrWriteDevTypeAcres2ndVarFields.Item(intFldCount).ToString
                If strToFieldName = "" Then
                  strToFieldName = strFromFieldName
                End If
                Dim existingFldIndex As Int32
                Dim totFldIndex As Int32
                Try
                  If m_strDWriteValue <> String.Empty Then
                    dblXValue = CDbl(rowDevType.Value(m_tblDevelopmentTypes.FindField(strFromFieldName)))
                    'dblExistingXValue = 0
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
                    'ONLY APPLY GROSS ATTRIBUTE IF USER CALC ONLY OPTION OFF
                    If m_arrWriteDevTypeAcresFieldsOnly.IndexOf(strToFieldName) <= -1 Then
                      pFeat.Value(pFeatClass.FindField(strFromFieldName)) = dblXValue
                    End If
                    If m_arrWriteDevTypeTotalFields.IndexOf(strToFieldName) <> -1 Then
                      existingFldIndex = pFeatClass.FindField("EX_" & strToFieldName)
                      If existingFldIndex <> -1 Then
                        Dim existingVal As Double = 0
                        If pFeat.Value(existingFldIndex) IsNot DBNull.Value Then
                          existingVal = CType(pFeat.Value(existingFldIndex), Double)
                        End If
                        pFeat.Value(pFeatClass.FindField(m_arrWriteDevTypeTotalFieldName(m_arrWriteDevTypeTotalFields.IndexOf(strToFieldName)).ToString)) = dblCalcAcres + (existingVal * (1 - dblRedevRate))
                      End If
                    End If
                  Else
                    pFeat.Value(pFeatClass.FindField(strToFieldName)) = 0
                    pFeat.Value(pFeatClass.FindField(strFromFieldName)) = 0
                    If m_arrWriteDevTypeTotalFields.IndexOf(strToFieldName) <> -1 Then
                      existingFldIndex = pFeatClass.FindField("EX_" & strToFieldName)
                      If existingFldIndex <> -1 Then
                        Dim existingVal As Double = 0
                        If pFeat.Value(existingFldIndex) IsNot DBNull.Value Then
                          existingVal = CType(pFeat.Value(existingFldIndex), Double)
                        End If
                        totFldIndex = pFeatClass.FindField(m_arrWriteDevTypeTotalFieldName(m_arrWriteDevTypeTotalFields.IndexOf(strToFieldName)).ToString)
                        If totFldIndex <> -1 Then pFeat.Value(totFldIndex) = existingVal
                      End If
                    End If
                  End If
                Catch ex As Exception

                End Try
              Next
            End If
          End If

          'PUSH VALUES TO SCENARIO FEATURECLASS
          Dim isNewDevType As Boolean = pFeat.Value(intDevTypeFld).ToString <> m_strDWriteValue
          pFeat.Value(intDevTypeFld) = m_strDWriteValue
          If isNewDevType OrElse m_arrWriteDevTypeAcresFields.Count > 0 OrElse m_arrWriteDevTypeFields.Count > 0 Then
            pFeat.Store()
          End If
          pFeat = pFeatureCursor.NextFeature

          m_dockEnvisionWinForm.prgBarEnvision.Value = CType((intCount / intTotalCount) * 100, Int32)
          m_dockEnvisionWinForm.prgBarEnvision.Refresh()
        Loop

        ' Calculate DevType field more efficiently if possible
        ' Note that if there is an edit session, undoing the edit session will not undo this
        ' If Not CalculateField(pFeatClass, "Dev_Type", "'" & m_strDWriteValue & "'", "VB") Then Throw New Exception("There was an error calculating the 'Dev_Type' field.")

        'ADD DEVD ACRES FROM EXISTING LU PAINT TABLE
        If m_strDWriteValue <> String.Empty AndAlso m_tblExistingLU IsNot Nothing AndAlso intEXLUFeatFld <> -1 Then
          Try
            cursorEX_LU = m_tblExistingLU.Search(Nothing, False)
            rowEX_LU = cursorEX_LU.NextRow
            Do While rowEX_LU IsNot Nothing
              If rowEX_LU.Value(intEXLUFld) IsNot DBNull.Value Then
                Dim val As String = rowEX_LU.Value(intEXLUFld).ToString
                If val <> String.Empty AndAlso existingLUDict.ContainsKey(val) Then
                  dblEX_LUAcres = CType(rowEX_LU.Value(intSumDEVDACREfld), Double)
                  rowEX_LU.Value(intSumDEVDACREfld) = dblEX_LUAcres + existingLUDict(val)
                  rowEX_LU.Store()
                End If
              End If
              rowEX_LU = cursorEX_LU.NextRow
            Loop
          Catch ex As Exception
          End Try
        End If

        ' Save the scenario field values to the scenario table
        If m_strDWriteValue <> String.Empty Then
          Try
            If rowScLookUp IsNot Nothing Then
              For Each kvp As KeyValuePair(Of String, Double) In valsDict
                intScLookUpFld = tblScLookUpTbl.FindField(kvp.Key)
                Try
                  dblPreviousValue = CType(rowScLookUp.Value(intScLookUpFld), Double)
                Catch ex As Exception
                  dblPreviousValue = 0
                End Try
                rowScLookUp.Value(intScLookUpFld) = dblPreviousValue + kvp.Value
              Next
              rowScLookUp.Store()
            End If
          Catch ex As Exception
          End Try
        End If

        GoTo CleanUp
      Catch ex As Exception
        MessageBox.Show(ex.Message, "ADD DEV TYPE SUB ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
        GoTo CleanUp
      End Try
    End If


CleanUp:
    pFeatSelection = Nothing
    'pCursor = Nothing
    'pFeatureCursor = Nothing
    If pCursor IsNot Nothing Then Marshal.FinalReleaseComObject(pCursor)
    If cursorEX_LU IsNot Nothing Then Marshal.FinalReleaseComObject(cursorEX_LU)
    pFeat = Nothing
    pFeatClass = Nothing
    pFeatTable = Nothing
    intTotalCount = Nothing
    intCount = Nothing
    dblVacantAcres = Nothing
    dblDevelopedAcres = Nothing
    intDevTypeFld = Nothing
    intVacantAcresFld = Nothing
    intDevelopedAcresFld = Nothing
    rowDevType = Nothing
    intDevTypeRow = Nothing
    intRow = Nothing
    rowTemp = Nothing
    intScVacLookUpFld = Nothing
    intScDevdLookUpFld = Nothing
    rowScLookUp = Nothing
    dblScVacantAcres = Nothing
    dblScDevelopedAcres = Nothing
    tblScLookUpTbl = Nothing
    intRedevFld = Nothing
    strFromFieldName = Nothing
    strToFieldName = Nothing
    dblCalcAcres = Nothing
    dblRedevRate = Nothing
    dblXValue = Nothing
    intFldCount = Nothing
    dblScExValue = Nothing
    intScExField = Nothing
    strFieldName = Nothing
    dblPreviousValue = Nothing
    intScLookUpFld = Nothing
    dblValue = Nothing
    fldSc = Nothing
    strDepVarFieldName = Nothing
    dblDepVar = Nothing
    GC.WaitForPendingFinalizers()
    GC.Collect()
  End Sub

  ''' <summary>
  ''' Calculates a field using the ArcToolbox Data Management - Calculate Field tool
  ''' </summary>
  ''' <param name="featureclass"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function CalculateField(ByVal featureclass As IFeatureClass, ByVal fieldName As String, _
  ByVal expression As String, ByVal expressionType As String) As Boolean
    Try
      Dim pCalculateField As ESRI.ArcGIS.DataManagementTools.CalculateField = New ESRI.ArcGIS.DataManagementTools.CalculateField
      pCalculateField.in_table = featureclass
      pCalculateField.field = fieldName
      pCalculateField.expression = expression
      pCalculateField.expression_type = expressionType

      ' Get the folder containing the featureclass
      Dim folder As String = DirectCast(featureclass, IDataset).Workspace.PathName

      ' Initialize the geoprocessor
      Dim gp As Geoprocessor = New Geoprocessor
      gp.SetEnvironmentValue("workspace", folder)

      RunTool(gp, pCalculateField)

      '' Create an IVariantArray to hold the parameter values
      'Dim parameters As IVariantArray = New VarArray

      '' Set the parameters
      'parameters.Add(featureclass)
      'parameters.Add(fieldName)
      'parameters.Add(expression)
      'parameters.Add(expressionType)

      '' Execute the model tool by name
      'gp.Execute("CalculateField_Management", parameters, Nothing)

      ' Return true if we get this far
      Return True
    Catch ex As Exception
      Return False
    End Try
  End Function

  Public Sub LocateEnvisionEditLayer()
    'CHECKS FOR AN ENVISION EDIT LAYER UPON INITIAL OPENING OF FORM
    Dim pMxDocument As IMxDocument
    Dim mapCurrent As Map
    Dim blnLayerFound As Boolean = False
    Dim pFeatLayer As IFeatureLayer
    Dim pWksFactory As FileGDBWorkspaceFactory = New FileGDBWorkspaceFactory
    Dim pFeatWks As IFeatureWorkspace
    Dim pFeatClass As IFeatureClass
    Dim pDataset As IDataset
    Dim intLayer As Integer
    Dim pTable As ITable

    Try
      pMxDocument = DirectCast(m_appEnvision.Document, IMxDocument)
      mapCurrent = CType(pMxDocument.FocusMap, Map)

      'REVIEW CURRENT LAYERS FOR ENVISION LAYER
      If mapCurrent.LayerCount > 0 Then
        For intLayer = 0 To mapCurrent.LayerCount - 1
          If TypeOf mapCurrent.Layer(intLayer) Is IFeatureLayer Then
            pFeatLayer = CType(mapCurrent.Layer(intLayer), IFeatureLayer)
            pFeatClass = pFeatLayer.FeatureClass
            pDataset = CType(pFeatClass, IDataset)
            If pFeatLayer.Name.Contains("<ACTIVE>") Then
              m_pEditFeatureLyr = pFeatLayer
              m_strFeaturePath = pDataset.Workspace.PathName.ToString
              If m_strFeaturePath.Contains("\\") Then
                m_strFeaturePath = m_strFeaturePath.Replace("\\", "\")
              End If
              m_strFeatureName = pDataset.Name
              m_dockEnvisionWinForm.itmOpenEnvisionFGDB.ToolTipText = m_strFeaturePath & "\" & m_strFeatureName
              GC.WaitForPendingFinalizers()
              GC.Collect()
            End If
            GC.WaitForPendingFinalizers()
            GC.Collect()
          End If
        Next
      End If
    Catch ex As Exception
      If m_pEditFeatureLyr Is Nothing Then
        m_dockEnvisionWinForm.itmOpenEnvisionFGDB.ToolTipText = ""
      End If
      MessageBox.Show(ex.Message, "Data Layer Review Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
      GoTo CleanUp
    End Try

CleanUp:
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
    GC.WaitForPendingFinalizers()
    GC.Collect()
  End Sub

  Public Sub UpdateEnvisionLyrLegend()
    m_strProcessing = m_strProcessing & "Starting sub UpdateEnvisionLyrLegend: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
    If m_dockEnvisionWinForm.dgvDevelopmentTypes.RowCount <= 0 Or m_pEditFeatureLyr Is Nothing Then
      Exit Sub
    End If

    Dim mxApplication As IMxApplication
    Dim pDoc As IMxDocument
    Dim pMap As IMap
    Dim strFieldName As String = ""
    Dim pRender As IUniqueValueRenderer
    Dim symd As ISimpleFillSymbol
    Dim intCount As Integer = 0
    Dim blnDevTypFldIsString As Boolean = True
    Dim strDevType As String = ""
    Dim intDevType As Integer = 0
    Dim intRed As Integer = 0
    Dim intGreen As Integer = 0
    Dim intBlue As Integer = 0
    Dim pNewColor As IRgbColor
    Dim ValFound As Boolean
    Dim NoValFound As Boolean
    Dim uh As Integer
    Dim symx As ISimpleFillSymbol
    Dim pLyr As IGeoFeatureLayer
    Dim pOutlineSymbol As ILineSymbol
    Dim intDevTypeCount As Integer
    Dim pTable As ITable


    'SET THE LEGEND FIELD
    strFieldName = "DEV_TYPE"

    'LOOK FOR THE DEV_TYPE FIELD FIRST...ADD IF MISSING
    If m_pEditFeatureLyr.FeatureClass.FindField("DEV_TYPE") <= -1 Then
      pTable = CType(m_pEditFeatureLyr.FeatureClass, ITable)
      If Not AddEnvisionField(pTable, "DEV_TYPE", "STRING", 150, 0) Then
        GoTo CleanUp
      End If
    End If

    'DEFINE SYMBOL OUTLINE
    pOutlineSymbol = New SimpleLineSymbol
    pOutlineSymbol.Width = 0.1
    pNewColor = New RgbColor
    pNewColor.Red = 0
    pNewColor.Blue = 0
    pNewColor.Green = 0
    pOutlineSymbol.Color = pNewColor

    Try
      m_strProcessing = m_strProcessing & "Setting Legend objects" & vbNewLine
      mxApplication = CType(m_appEnvision, IMxApplication)
      pDoc = CType(m_appEnvision.Document, IMxDocument)
      pMap = pDoc.FocusMap

      m_strProcessing = m_strProcessing & "Erase Color" & vbNewLine
      pNewColor = New RgbColor
      pNewColor.Red = 225
      pNewColor.Blue = 225
      pNewColor.Green = 225
      pOutlineSymbol.Color = pNewColor
      pOutlineSymbol.Width = 0.1

      symd = New SimpleFillSymbol
      symd.Style = esriSimpleFillStyle.esriSFSHollow
      symd.Outline = pOutlineSymbol

      '** Make the renderer
      '** These properties should be set prior to adding values
      m_strProcessing = m_strProcessing & "Construct Renderer" & vbNewLine
      pRender = New UniqueValueRenderer
      pRender.FieldCount = 1
      pRender.Field(0) = strFieldName
      pRender.DefaultSymbol = DirectCast(symd, ISymbol)
      pRender.UseDefaultSymbol = True

      'RESET SYMBOL OUT LINE TO 0
      pOutlineSymbol.Width = 0.0

      pLyr = DirectCast(m_pEditFeatureLyr, IGeoFeatureLayer)
      'DETERMINE DEVELOPMENT TYPE FIELD TYPE
      'If pLyr.FeatureClass.Fields.Field(pLyr.FeatureClass.FindField(strFieldName)).Type = esriFieldType.esriFieldTypeInteger Then
      '    m_strProcessing = m_strProcessing & "Determine Dev Type field value type: Integer" & vbNewLine
      '    blnDevTypFldIsString = False
      'Else
      'm_strProcessing = m_strProcessing & "Determine Dev Type field value type: String" & vbNewLine
      blnDevTypFldIsString = True
      'End If

      intDevTypeCount = m_dockEnvisionWinForm.dgvDevelopmentTypes.RowCount
      m_strProcessing = m_strProcessing & "Number of Dev Types to Load into Legend: " & intDevTypeCount.ToString & vbNewLine

      If intDevTypeCount > 0 Then
        For intCount = 0 To intDevTypeCount - 1
          strDevType = CType(m_dockEnvisionWinForm.dgvDevelopmentTypes.Rows(intCount).Cells(1).Value, String)
          intDevType = CType(m_dockEnvisionWinForm.dgvDevelopmentTypes.Rows(intCount).Cells(2).Value, Integer)
          m_strProcessing = m_strProcessing & "Dev Type String (" & strDevType & "); Dev Type Integer (" & intDevType & ") " & m_dockEnvisionWinForm.dgvDevelopmentTypes.RowCount.ToString & vbNewLine

          If strDevType = "ERASE" And intDevType = -1 Then
            Continue For
          Else
            intRed = CType(m_dockEnvisionWinForm.dgvDevelopmentTypes.Rows(intCount).Cells(2).Value, Integer)
            intGreen = CType(m_dockEnvisionWinForm.dgvDevelopmentTypes.Rows(intCount).Cells(3).Value, Integer)
            intBlue = CType(m_dockEnvisionWinForm.dgvDevelopmentTypes.Rows(intCount).Cells(4).Value, Integer)
            pNewColor = New RgbColor
            pNewColor.Red = intRed
            pNewColor.Blue = intBlue
            pNewColor.Green = intGreen

            symx = New SimpleFillSymbol
            symx.Style = esriSimpleFillStyle.esriSFSSolid
            symx.Outline = pOutlineSymbol
            symx.Color = pNewColor
          End If

          'CHECK FOR PREEXISTING SYMBOL
          ValFound = False
          For uh = 0 To (pRender.ValueCount - 1)
            If blnDevTypFldIsString Then
              If pRender.Value(uh) = strDevType Then
                NoValFound = True
                Exit For
              End If
            Else
              If CType(pRender.Value(uh), Int32) = intDevType Then
                NoValFound = True
                Exit For
              End If
            End If
          Next uh

          If Not ValFound Then
            If blnDevTypFldIsString Then
              pRender.AddValue(strDevType, strFieldName, DirectCast(symx, ISymbol))
              pRender.Label(strDevType) = strDevType
              pRender.Symbol(strDevType) = DirectCast(symx, ISymbol)
              If strDevType = "" Then
                pRender.AddReferenceValue(" ", strDevType)
              End If
            Else
              pRender.AddValue(intDevType.ToString, strFieldName, DirectCast(symx, ISymbol))
              pRender.Label(intDevType.ToString) = strDevType
              pRender.Symbol(intDevType.ToString) = DirectCast(symx, ISymbol)
            End If
          End If
        Next
        'GoTo CleanUp
        pRender.ColorScheme = "Custom"
        pRender.FieldType(0) = True
        pLyr.Renderer = DirectCast(pRender, IFeatureRenderer)
        pLyr.DisplayField = strFieldName
        '** Refresh the TOC
        pDoc.ActiveView.ContentsChanged()
        pDoc.UpdateContents()

        '** Draw the map
        'pDoc.ActiveView.PartialRefresh(esriViewDrawPhase.esriViewGeoSelection, pLyr, pDoc.ActiveView.Extent)
        pDoc.ActiveView.Refresh()
      End If
      GoTo CleanUp

    Catch ex As Exception
      m_strProcessing = m_strProcessing & "Error in sub UpdateEnvisionLyrLegend: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
      m_strProcessing = m_strProcessing & "Error Message:  " & vbNewLine & ex.Message & vbNewLine
      MessageBox.Show(ex.Message, "Legend Renderer Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
      GoTo CleanUp
    End Try


CleanUp:
    m_strProcessing = m_strProcessing & "Ending sub UpdateEnvisionLyrLegend: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
    mxApplication = Nothing
    pDoc = Nothing
    pMap = Nothing
    strFieldName = Nothing
    pRender = Nothing
    symd = Nothing
    intCount = Nothing
    blnDevTypFldIsString = Nothing
    strDevType = Nothing
    intDevType = Nothing
    intRed = Nothing
    intGreen = Nothing
    intBlue = Nothing
    pNewColor = Nothing
    ValFound = Nothing
    NoValFound = Nothing
    uh = Nothing
    symx = Nothing
    pLyr = Nothing
    GC.WaitForPendingFinalizers()
    GC.Collect()
  End Sub

  Public Function RunTool(ByVal geoProcessor As ESRI.ArcGIS.Geoprocessor.Geoprocessor, ByVal process As IGPProcess) As Boolean
    '*******************************************************
    ' Set the overwrite output option to true
    '*******************************************************
    Try
      geoProcessor.OverwriteOutput = True
      geoProcessor.Execute(process, Nothing)
      RunTool = ReturnMessages(geoProcessor)
      GoTo CleanUp
    Catch err As Exception
      RunTool = ReturnMessages(geoProcessor)
      GoTo CleanUp
    End Try
CleanUp:
    process = Nothing
    geoProcessor = Nothing
    GC.WaitForFullGCComplete()
    GC.WaitForPendingFinalizers()
    GC.Collect()
  End Function

  Public Function ReturnMessages(ByVal GP As ESRI.ArcGIS.Geoprocessor.Geoprocessor) As Boolean
    '*******************************************************
    ' Function for returning the tool messages.
    ' Print out the messages from tool executions
    '*******************************************************
    ReturnMessages = True
    Dim strMessage As String = ""
    Dim Count As Integer
    Try
      If GP.MessageCount > 0 Then
        For Count = 0 To GP.MessageCount - 1
          Try
            strMessage = strMessage & vbNewLine & GP.GetMessage(Count)
          Catch ex As Exception
            GoTo CleanUp
          End Try
        Next
        If strMessage.Contains("ERROR") = True Or strMessage.Contains("ERR:") = True Then
          MessageBox.Show(strMessage, "")
          ReturnMessages = False
        End If
        GoTo CleanUp
      End If
    Catch ex As Exception
      GoTo CleanUp
    End Try

CleanUp:
    'm_strProcessing = m_strProcessing & strMessage
    GP = Nothing
    GC.WaitForFullGCComplete()
    GC.WaitForPendingFinalizers()
    GC.Collect()
    Return ReturnMessages
  End Function

  '    Public Function EnvisionProjectSetup_LoadData(ByVal sender As Object, ByVal e As System.EventArgs) As Boolean
  '        EnvisionProjectSetup_LoadData = True
  '        Dim mxApplication As IMxApplication = Nothing
  '        Dim pMxDocument As IMxDocument = Nothing
  '        Dim mapCurrent As Map
  '        Dim pSpatRef As ISpatialReference
  '        Dim intCount As Integer
  '        Dim pLyr As ILayer
  '        Dim pFeatLyr As IFeatureLayer
  '        Dim pRasterLyr As IRasterLayer
  '        Dim intLayer As Integer
  '        Dim intFeatCount As Integer = 0
  '        Dim pActiveView As IActiveView = Nothing
  '        Dim pPCS As IProjectedCoordinateSystem
  '        Dim intLayerCount As Integer = 0

  '        Try
  '            'MAKE ALL TOOLBARS VISIBLE
  '            m_frmEnvisionProjectSetup.ToolStrip_Constraints.Visible = True
  '            m_frmEnvisionProjectSetup.ToolStrip_LandUse.Visible = True
  '            m_frmEnvisionProjectSetup.ToolStrip_InfoTab1.Visible = True
  '            m_frmEnvisionProjectSetup.ToolStrip_InfoTab2.Visible = True
  '            m_frmEnvisionProjectSetup.ToolStrip_InfoTab3.Visible = True

  '            '********************************************************************
  '            'Populate the combo boxes with layer information
  '            '********************************************************************
  '            If Not TypeOf m_appEnvision Is IApplication Then
  '                GoTo CloseForm
  '            End If

  '            pMxDocument = DirectCast(m_appEnvision.Document, IMxDocument)
  '            mapCurrent = CType(pMxDocument.FocusMap, Map)
  '            pActiveView = CType(pMxDocument.FocusMap, IActiveView)
  '            MessageBox.Show(mapCurrent.LayerCount.ToString, "LAyer Count")


  '            Dim layerCLSID As String = "{E156D7E5-22AF-11D3-9F99-00C04F6BC78E}"

  '            Dim uid As ESRI.ArcGIS.esriSystem.IUID = New ESRI.ArcGIS.esriSystem.UIDClass
  '            uid.Value = layerCLSID ' Example: "{E156D7E5-22AF-11D3-9F99-00C04F6BC78E}" = IGeoFeatureLayer

  '            Try
  '                Dim enumLayer As ESRI.ArcGIS.Carto.IEnumLayer = mapCurrent.Layers((CType(uid, ESRI.ArcGIS.esriSystem.UID)), True) ' Explicit Cast
  '                enumLayer.Reset()
  '                Dim layer As ESRI.ArcGIS.Carto.ILayer = enumLayer.Next
  '                Do While Not (layer Is Nothing)
  '                    ' TODO - Implement your code here...


  '                    layer = enumLayer.Next()
  '                    MessageBox.Show(layer.Name.ToString)
  '                Loop
  '            Catch ex As System.Exception
  '                'System.Windows.Forms.MessageBox.Show("No layers of type: " + uid.Value.ToString);
  '            End Try

  '            If mapCurrent.LayerCount = 0 Then
  '                'DISABLE PARCEL CONTROLS
  '                m_frmEnvisionProjectSetup.rdbSourceParcels.Enabled = False
  '                m_frmEnvisionProjectSetup.rdbSourceHybrid.Enabled = False
  '                m_frmEnvisionProjectSetup.cmbParcelLayers.Enabled = False
  '                m_frmEnvisionProjectSetup.rdbExtentParcel.Enabled = False
  '                m_frmEnvisionProjectSetup.rdbExtentLayer.Enabled = False
  '                m_frmEnvisionProjectSetup.cmbExtentLayers.Enabled = False
  '                'DISABLE CONSTRAINTS CONTROLS
  '                m_frmEnvisionProjectSetup.ToolStrip_Constraints.Enabled = False
  '                m_frmEnvisionProjectSetup.dgvConstraints.Enabled = False
  '                m_frmEnvisionProjectSetup.Panel1.Enabled = False
  '                m_frmEnvisionProjectSetup.Panel2.Enabled = False
  '                'DISABLE LAND USE CONTROLS
  '                m_frmEnvisionProjectSetup.ToolStrip_LandUse.Enabled = False
  '                m_frmEnvisionProjectSetup.dgvLandUseAttributes.Enabled = False
  '                m_frmEnvisionProjectSetup.Panel1.Enabled = False
  '                m_frmEnvisionProjectSetup.Panel2.Enabled = False
  '                'DISABLE SUBAREA CONTROLS
  '                m_frmEnvisionProjectSetup.dgvSubAreas.Enabled = False
  '            End If

  '            'BY DEFUALT RETRIEVE AND LOAD THE VIEW DOCUMENT SPATIAL REFERENCE PROJECTION
  '            pSpatRef = pMxDocument.FocusMap.SpatialReference
  '            Try
  '                pPCS = pSpatRef
  '                m_frmEnvisionProjectSetup.tbxProjection.Text = pSpatRef.Name
  '                m_pSpatRefProject = pSpatRef
  '                m_frmEnvisionProjectSetup.lblProjectUnits.Text = pPCS.CoordinateUnit.Name
  '            Catch ex As Exception
  '                m_frmEnvisionProjectSetup.gpbExtentCoor.Enabled = True
  '                m_frmEnvisionProjectSetup.rdbExtentCustom.Checked = True
  '            End Try

  '            'BUILD LIST OF AVAILABLE FEATURE CLASSES
  '            m_arrFeatureLayers = New ArrayList
  '            m_arrFeatureLayers.Clear()
  '            m_arrRasterLayers = New ArrayList
  '            m_arrRasterLayers.Clear()
  '            m_frmEnvisionProjectSetup.cmbParcelLayers.Items.Clear()
  '            m_frmEnvisionProjectSetup.cmbExtentLayers.Items.Clear()
  '            m_frmEnvisionProjectSetup.dgvConstraints.Rows.Clear()
  '            m_frmEnvisionProjectSetup.cmbLandUseLayers.Items.Clear()
  '            m_frmEnvisionProjectSetup.cmbSubaraLayers.Items.Clear()
  '            m_frmEnvisionProjectSetup.dgvSubAreas.Rows.Clear()
  '            If mapCurrent.LayerCount > 0 Then
  '                For intLayer = 0 To mapCurrent.LayerCount - 1
  '                    pLyr = CType(mapCurrent.Layer(intLayer), ILayer)
  '                    m_frmEnvisionProjectSetup.cmbExtentLayers.Items.Add(pLyr.Name)
  '                    If TypeOf mapCurrent.Layer(intLayer) Is IRasterLayer Then
  '                        pRasterLyr = CType(mapCurrent.Layer(intLayer), IRasterLayer)
  '                        m_arrRasterLayers.Add(pRasterLyr)
  '                        pRasterLyr = Nothing
  '                    ElseIf TypeOf mapCurrent.Layer(intLayer) Is IFeatureLayer Then
  '                        pFeatLyr = CType(mapCurrent.Layer(intLayer), IFeatureLayer)
  '                        m_frmEnvisionProjectSetup.cmbParcelLayers.Items.Add(pFeatLyr.Name)
  '                        m_frmEnvisionProjectSetup.dgvConstraints.Rows.Add()
  '                        m_frmEnvisionProjectSetup.cmbLandUseLayers.Items.Add(pFeatLyr.Name)
  '                        m_frmEnvisionProjectSetup.cmbSubaraLayers.Items.Add(pFeatLyr.Name)
  '                        m_arrFeatureLayers.Add(pFeatLyr)
  '                        intFeatCount = intFeatCount + 1
  '                        pFeatLyr = Nothing
  '                    End If
  '                Next
  '                If m_arrRasterLayers.Count <= 0 Then
  '                    m_frmEnvisionProjectSetup.pnlSlope.Enabled = False
  '                Else
  '                    m_frmEnvisionProjectSetup.pnlSlope.Enabled = True
  '                End If
  '                GoTo CleanUp
  '            Else
  '                'MessageBox.Show("No layers were found in the current map document.  Please add layers you would like to use in building an Envision Project.", "No Layers Found", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
  '                m_arrFeatureLayers.Clear()
  '                m_arrFeatureLayers = Nothing
  '                m_arrRasterLayers.Clear()
  '                m_arrRasterLayers = Nothing
  '                'GoTo CloseForm
  '            End If
  '            GoTo CleanUp
  '        Catch ex As Exception
  '            MessageBox.Show(ex.Message, "Envision Project Setup Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
  '            GoTo CloseForm
  '        End Try
  '        GoTo CleanUp

  'CloseForm:
  '        EnvisionProjectSetup_LoadData = False
  '        GoTo CleanUp

  'CleanUp:
  '        mxApplication = Nothing
  '        pMxDocument = Nothing
  '        mapCurrent = Nothing
  '        pSpatRef = Nothing
  '        intCount = Nothing
  '        pLyr = Nothing
  '        pFeatLyr = Nothing
  '        pRasterLyr = Nothing
  '        intLayer = Nothing
  '        intFeatCount = Nothing
  '        pActiveView = Nothing
  '        GC.WaitForPendingFinalizers()
  '        GC.Collect()
  '    End Function

  Public Sub ResetEnvisionAttributeFieldTrackingTbl()
    'CLEAR OUT ALL PREVIOUS FIELD ENTERIES AND LOAD NEW 
    m_strProcessing = m_strProcessing & "Starting sub ResetEnvisionAttributeFieldTrackingTbl: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
    If m_tblAttribFields Is Nothing Then
      GoTo CleanUp
    End If

    'CLEAR PREVIOUS RECORDS
    m_tblAttribFields.DeleteSearchedRows(Nothing)

    'CHECK FOR REQUIRED DEVELOPMENT TYPE TABLE LOOKUP TABLE FIELDS
    Dim shtDevType As Microsoft.Office.Interop.Excel.Worksheet = Nothing
    Dim intCol As Integer
    Dim strCellValue As String = ""
    Dim intRow As Integer
    Dim strFieldName As String = ""
    Dim strFieldType As String = "STRING"
    Dim intFieldWidth As Integer = 1
    Dim intFieldDecimal As Integer = 0
    Dim strAlias As String = ""
    Dim intCalcByAcres As Integer = 0
    Dim strCalcFieldName As String = ""
    Dim strDepVarFieldName As String = ""
    Dim totFldName As String
    Dim rowTemp As IRow
    Dim strFldCellValue As String = ""
    Dim intFieldRow As Integer = 5
    Dim intStartRow As Integer = 6

    'SET THE ARCGIS VIEW STATUS MESSAGE
    m_appEnvision.StatusBar.Message(0) = "RETRIEVING DEV TYPE ATTRIBUTE FIELD OPTIONS"
    m_dockEnvisionWinForm.prgBarEnvision.Value = 0


    If m_xlPaintWB1 Is Nothing Or m_xlPaintApp Is Nothing Then
      GoTo CleanUp
    End If

    Try
      'RETRIEVE DEVELOPMENT TYPE TAB
      If Not TypeOf m_xlPaintWB1.Sheets("Dev Type Attributes") Is Microsoft.Office.Interop.Excel.Worksheet Then
        MessageBox.Show("The 'Devt' tab could not be found in the selected ENVISION Excel file. Please select another file.", "SCENARIO NOT FOUND", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
        GoTo CleanUp
      Else
        shtDevType = DirectCast(m_xlPaintWB1.Sheets("Dev Type Attributes"), Microsoft.Office.Interop.Excel.Worksheet)
      End If

      'FIND THE STARTING POINT
      For intRow = 1 To 10
        strFldCellValue = CStr(CType(shtDevType.Cells(intRow, 1), Microsoft.Office.Interop.Excel.Range).Value)
        If Not strFldCellValue Is Nothing Then
          If UCase(strFldCellValue) = "DEV_TYPE" Then
            intFieldRow = intRow
            Exit For
          Else
            If strFldCellValue.Contains(",") Then
              strFieldName = CStr(strFldCellValue.Split(","c)(0))
              If UCase(strFieldName) = "DEV_TYPE" Then
                intFieldRow = intRow
                Exit For
              End If
            End If
          End If
        End If
      Next

      'RETRIEVE FIELD NAMES FROM ROW 5
      intRow = m_tblAttribFields.RowCount(Nothing)
      For intCol = 1 To 400
        strCellValue = CStr(CType(shtDevType.Cells(intFieldRow, intCol), Microsoft.Office.Interop.Excel.Range).Value)
        If strCellValue Is Nothing Then
          Continue For
        End If
        If strCellValue.Contains(",") Then
          strFieldName = ""
          strFieldType = "STRING"
          intFieldWidth = 1
          intFieldDecimal = 0
          strCalcFieldName = ""
          strDepVarFieldName = ""
          totFldName = String.Empty
          Try
            strFieldName = CStr(strCellValue.Split(","c)(0))
          Catch ex As Exception
            strFieldName = ""
          End Try
          Try
            strFieldType = CStr(strCellValue.Split(","c)(1))
          Catch ex As Exception
            strFieldType = ""
          End Try
          Try
            intFieldWidth = CInt(strCellValue.Split(","c)(2))
          Catch ex As Exception
            intFieldWidth = 0
          End Try
          Try
            intFieldDecimal = CInt(strCellValue.Split(","c)(3))
          Catch ex As Exception
            intFieldDecimal = 0
          End Try
          Try
            strAlias = CStr(strCellValue.Split(","c)(4))
          Catch ex As Exception
            strAlias = ""
          End Try
          Try
            intCalcByAcres = CInt(strCellValue.Split(","c)(5))
          Catch ex As Exception
            intCalcByAcres = 0
          End Try
          Try
            strCalcFieldName = CStr(strCellValue.Split(","c)(6))
          Catch ex As Exception
            strCalcFieldName = ""
          End Try
          Try
            strDepVarFieldName = CStr(strCellValue.Split(","c)(7))
          Catch ex As Exception
            strDepVarFieldName = ""
          End Try
          If m_pEditFeatureLyr IsNot Nothing Then totFldName = GetDefaultTotalFieldName(intCalcByAcres, strCalcFieldName, strFieldName, m_pEditFeatureLyr.FeatureClass)

          'ONLY ADD THE DEV TYPE ATTRUIBUTE IF ALL THE PARAMETERS ARE PROVIDED
          If strFieldName.Length > 0 And intFieldWidth > 0 And strFieldType.Length > 0 Then
            If Not UCase(strFieldName) = "DEV_TYPE" And Not UCase(strFieldName) = "DEVELOPMENT_TYPE" Then
              'INSERT NEW RECORD
              m_tblAttribFields.CreateRow()
              intRow = intRow + 1
              Try
                rowTemp = m_tblAttribFields.GetRow(intRow)
                rowTemp.Value(m_tblAttribFields.FindField("FIELD_NAME")) = strFieldName
                rowTemp.Value(m_tblAttribFields.FindField("FIELD_TYPE")) = strFieldType
                rowTemp.Value(m_tblAttribFields.FindField("FIELD_ALIAS")) = strAlias
                rowTemp.Value(m_tblAttribFields.FindField("FIELD_WIDTH")) = intFieldWidth
                rowTemp.Value(m_tblAttribFields.FindField("FIELD_DECIMAL")) = intFieldDecimal
                rowTemp.Value(m_tblAttribFields.FindField("USE")) = 0
                rowTemp.Value(m_tblAttribFields.FindField("CALC_BY_ACRES")) = intCalcByAcres
                rowTemp.Value(m_tblAttribFields.FindField("CALC_FIELD_NAME")) = strCalcFieldName
                rowTemp.Value(m_tblAttribFields.FindField("DEP_VAR_FIELD_NAME")) = strDepVarFieldName
                rowTemp.Value(m_tblAttribFields.FindField("Calc_Total")) = 0
                rowTemp.Value(m_tblAttribFields.FindField("Total_Field_Name")) = totFldName
                rowTemp.Store()
              Catch ex As Exception
                Continue For
              End Try
            End If
          End If
        End If

        m_dockEnvisionWinForm.prgBarEnvision.Value = CType((intCol / 400) * 100, Int32)
      Next
      m_dockEnvisionWinForm.prgBarEnvision.Value = 0
    Catch ex As Exception
      MessageBox.Show(ex.Message, "ResetEnvisionAttributeFieldTrackingTbl SUB ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
      GoTo CleanUp
    End Try

CleanUp:
    GC.WaitForPendingFinalizers()
    GC.Collect()
  End Sub

  Public Sub LookUpTablesEnvisionDevTypesTblCheck(ByVal pTable As ITable)
    'CHECK FOR REQUIRED DEVELOPMENT TYPE TABLE LOOKUP TABLE FIELDS
    m_strProcessing = m_strProcessing & "Starting function LookUpTablesEnvisionDevTypesTblCheck: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
    Dim shtDevType As Microsoft.Office.Interop.Excel.Worksheet = Nothing
    Dim intCol As Integer
    Dim strCellValue As String = ""
    Dim intRow As Integer
    Dim intRowCount As Integer
    Dim intFldCount As Integer
    Dim pField As IField
    Dim intFld As Integer
    Dim strVal As String
    Dim intVal As Integer
    Dim rowTemp As IRow
    Dim blnStore As Boolean

    Dim strFieldName As String = ""
    Dim strFieldType As String = "STRING"
    Dim intFieldWidth As Integer = 1
    Dim intFieldDecimal As Integer = 0
    Dim strFieldCalcName As String = ""


    'SET THE ARCGIS VIEW STATUS MESSAGE
    If m_appEnvision Is Nothing Then
      MessageBox.Show("Error trapped and the application variable.")
    End If
    m_appEnvision.StatusBar.Message(0) = "CHECKING FOR REQUIRED DEVELOPMENT TYPE TABLE LOOKUP TABLE FIELDS"
    m_dockEnvisionWinForm.prgBarEnvision.Value = 0


    If m_xlPaintWB1 Is Nothing Or m_xlPaintApp Is Nothing Then
      GoTo CleanUp
    End If

    Try
      'RETRIEVE DEVELOPMENT TYPE TAB
      If Not TypeOf m_xlPaintWB1.Sheets("Dev Type Attributes") Is Microsoft.Office.Interop.Excel.Worksheet Then
        MessageBox.Show("The 'Devt' tab could not be found in the selected ENVISION Excel file. Please select another file.", "SCENARIO NOT FOUND", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
        GoTo CleanUp
      Else
        shtDevType = DirectCast(m_xlPaintWB1.Sheets("Dev Type Attributes"), Microsoft.Office.Interop.Excel.Worksheet)
      End If

      'SET THE 

      'RETRIEVE FIELD NAMES FROM ROW 5
      For intCol = 1 To 400
        strCellValue = CStr(CType(shtDevType.Cells(5, intCol), Microsoft.Office.Interop.Excel.Range).Value)
        If strCellValue Is Nothing Then
          Continue For
        End If
        If strCellValue.Contains(",") Then
          m_appEnvision.StatusBar.Message(0) = "CHECKING FOR REQUIRED DEVELOPMENT TYPE TABLE LOOKUP TABLE FIELDS: FIELD (" & strCellValue & ")"
          strFieldName = ""
          strFieldType = "STRING"
          intFieldWidth = 1
          intFieldDecimal = 0
          strFieldCalcName = ""
          Try
            strFieldName = CStr(strCellValue.Split(","c)(0))
          Catch ex As Exception
            strFieldName = ""
          End Try
          If UCase(strCellValue) = "DEV_TYPE" Or UCase(strCellValue) = "DEVELOPMENT TYPE" Or UCase(strCellValue) = "DEV TYPE ATTRIBUTES" Then
            Continue For
          End If
          Try
            strFieldType = CStr(strCellValue.Split(","c)(1))
          Catch ex As Exception
            strFieldType = "STRING"
          End Try
          Try
            intFieldWidth = CInt(strCellValue.Split(","c)(2))
          Catch ex As Exception
            intFieldWidth = 1
          End Try
          Try
            intFieldDecimal = CInt(strCellValue.Split(","c)(3))
          Catch ex As Exception
            intFieldDecimal = 0
          End Try
          Try
            strFieldCalcName = CStr(strCellValue.Split(","c)(4))
          Catch ex As Exception
            strFieldCalcName = ""
          End Try
          If strFieldName.Length > 0 And intFieldWidth > 0 And strFieldType.Length > 0 Then
            AddEnvisionField(pTable, strFieldName, strFieldType, intFieldWidth, intFieldDecimal)
          End If
        End If

        If pTable.FindField("DEVELOPMENT_TYPE") <= -1 Then
          If Not AddEnvisionField(pTable, "DEVELOPMENT_TYPE", "STRING", 150, 0) Then
            GoTo CleanUp
          End If
        End If
        m_dockEnvisionWinForm.prgBarEnvision.Value = CType((intCol / 400) * 100, Int32)
      Next
      m_dockEnvisionWinForm.prgBarEnvision.Value = 0
    Catch ex As Exception
      MessageBox.Show(ex.Message, "DevTypesTblFldsCheck SUB ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
      GoTo CleanUp
    End Try

    'CHECK FOR THE MINIMUM NUMBER OF RECORDS
    m_appEnvision.StatusBar.Message(0) = "CHECKING FOR REQUIRED DEVELOPMENT TYPE TABLE LOOKUP TABLE FIELDS: CHECKING FOR NULL VALUES"
    'intRowCount = pTable.RowCount(Nothing)
    'If intRowCount < m_intDevTypeMax Then
    '    For intRowCount = intRowCount To m_intDevTypeMax
    '        pTable.CreateRow()
    '    Next
    'End If
    intRowCount = pTable.RowCount(Nothing)

    'SET ALL NULL VALUES TO DEFAULTS of "" OR 0
    intRowCount = pTable.RowCount(Nothing)
    intFldCount = pTable.Fields.FieldCount
    If intRowCount > 0 Then
      For intRow = 1 To intRowCount
        blnStore = False
        rowTemp = pTable.GetRow(intRow)
        For intFld = 1 To intFldCount - 1
          pField = pTable.Fields.Field(intFld)
          m_appEnvision.StatusBar.Message(0) = "CHECKING FOR REQUIRED DEVELOPMENT TYPE TABLE LOOKUP TABLE FIELDS: CHECKING FOR NULL VALUES; FIELD (" & pField.Name & "); RECORD (" & intRow.ToString & ")"
          Try
            If pField.Type = esriFieldType.esriFieldTypeString Then
              Try
                strVal = rowTemp.Value(intFld).ToString
              Catch ex As Exception
                blnStore = True
                rowTemp.Value(intFld) = ""
              End Try
            Else
              Try
                intVal = CInt(rowTemp.Value(intFld))
              Catch ex As Exception
                blnStore = True
                rowTemp.Value(intFld) = 0
              End Try
            End If
          Catch ex As Exception
            MessageBox.Show(ex.Message, "Dev ERRR")
          End Try
        Next
        If blnStore Then
          rowTemp.Store()
        End If
      Next
    End If
    GoTo CleanUp

CleanUp:
    m_strProcessing = m_strProcessing & "Ending function LookUpTablesEnvisionDevTypesTblCheck: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine & vbNewLine
    m_appEnvision.StatusBar.Message(0) = ""
    m_dockEnvisionWinForm.prgBarEnvision.Value = 0
    m_blnLoadingDevTypes = False
    shtDevType = Nothing
    intCol = Nothing
    strCellValue = Nothing
    intRow = Nothing
    intRowCount = Nothing
    intFldCount = Nothing
    pField = Nothing
    intFld = Nothing
    strVal = Nothing
    intVal = Nothing
    rowTemp = Nothing
    blnStore = Nothing
    GC.Collect()
    GC.WaitForPendingFinalizers()
  End Sub

  Public Function LookUpTablesEnvisionCheck(ByVal strPath As String) As Boolean
    'BUILD ENVISION LOOKUP TABLES TO STORE DEVELOPMENT TYPE VARIABLES AND OUTPUTS
    m_strProcessing = m_strProcessing & "Starting function LookUpTablesEnvisionCheck: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
    LookUpTablesEnvisionCheck = True
    GC.Collect()
    Dim GP As ESRI.ArcGIS.Geoprocessor.Geoprocessor
    Dim pCreateTable As ESRI.ArcGIS.DataManagementTools.CreateTable
    Dim pWksFactory As ESRI.ArcGIS.Geodatabase.IWorkspaceFactory = Nothing
    Dim pFeatWks As ESRI.ArcGIS.Geodatabase.IFeatureWorkspace

    Try
      'RETRIEVE SELECTED FGDB PATH
      If strPath.Length > 0 Then
        pWksFactory = New FileGDBWorkspaceFactory
        If Directory.Exists(strPath) Then
          GP = New ESRI.ArcGIS.Geoprocessor.Geoprocessor
          GP.AddOutputsToMap = True
          pFeatWks = DirectCast(pWksFactory.OpenFromFile(strPath, 0), IFeatureWorkspace)
          'ATTRIBUTE TRACKING TABLE
          Try
            m_tblAttribFields = pFeatWks.OpenTable("ENVISION_DEVTYPE_FIELD_TRACKING")
          Catch ex As Exception
            m_strProcessing = m_strProcessing & "Missing table, ENVISION_DEVTYPE_FIELD_TRACKING, executing geoprocessing tool to create table: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
            pCreateTable = New ESRI.ArcGIS.DataManagementTools.CreateTable
            pCreateTable.out_name = "ENVISION_DEVTYPE_FIELD_TRACKING"
            pCreateTable.out_path = strPath
            RunTool(GP, pCreateTable)
            m_tblAttribFields = Nothing
            m_tblAttribFields = pFeatWks.OpenTable("ENVISION_DEVTYPE_FIELD_TRACKING")
          End Try
          If TypeOf m_tblAttribFields Is ITable Then
            LookUpTablesEnvisionAttributeFieldTrackingTblCheck(m_tblAttribFields)
          End If
          m_strProcessing = m_strProcessing & vbNewLine & "ATTRIBUTE TRACKING TABLE code complete" & vbNewLine
          'DEVELOPMENT TYPES TABLE
          Try
            m_tblDevelopmentTypes = pFeatWks.OpenTable("DEVELOPMENT_TYPE_ATTRIBUTES")
          Catch ex As Exception
            m_strProcessing = m_strProcessing & "Missing table, DEVELOPMENT_TYPE_ATTRIBUTES, executing geoprocessing tool to create table: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
            pCreateTable = New ESRI.ArcGIS.DataManagementTools.CreateTable
            pCreateTable.out_name = "DEVELOPMENT_TYPE_ATTRIBUTES"
            pCreateTable.out_path = strPath
            RunTool(GP, pCreateTable)
            m_tblDevelopmentTypes = pFeatWks.OpenTable("DEVELOPMENT_TYPE_ATTRIBUTES")
            blnLoadDevTypes = True
          End Try
          If TypeOf m_tblDevelopmentTypes Is ITable Then
            Try
              m_intDevTypeMax = m_tblDevelopmentTypes.RowCount(Nothing)
              If m_intDevTypeMax <= -1 Then
                m_intDevTypeMax = 1
              End If
            Catch ex As Exception
              m_intDevTypeMax = 1
            End Try

            LookUpTablesEnvisionDevTypesTblCheck(m_tblDevelopmentTypes)
          End If
          m_strProcessing = m_strProcessing & vbNewLine & "DEVELOPMENT TYPES TABLE code complete" & vbNewLine
          'SCENARIO 1 INPUT TABLE
          Try
            m_tblSc1 = pFeatWks.OpenTable("SCENARIO1")
          Catch ex As Exception
            m_strProcessing = m_strProcessing & "Missing table, SCENARIO1, executing geoprocessing tool to create table: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
            pCreateTable = New ESRI.ArcGIS.DataManagementTools.CreateTable
            pCreateTable.out_name = "SCENARIO1"
            pCreateTable.out_path = strPath
            RunTool(GP, pCreateTable)
            m_tblSc1 = pFeatWks.OpenTable("SCENARIO1")
          End Try
          If TypeOf m_tblSc1 Is ITable Then
            LookUpTablesEnvisionScenarioTblCheck(m_tblSc1)
          End If
          'SCENARIO 2 INPUT TABLE
          Try
            m_tblSc2 = pFeatWks.OpenTable("SCENARIO2")
          Catch ex As Exception
            m_strProcessing = m_strProcessing & "Missing table, SCENARIO2, executing geoprocessing tool to create table: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
            pCreateTable = New ESRI.ArcGIS.DataManagementTools.CreateTable
            pCreateTable.out_name = "SCENARIO2"
            pCreateTable.out_path = strPath
            RunTool(GP, pCreateTable)
            m_tblSc2 = pFeatWks.OpenTable("SCENARIO2")
          End Try
          If TypeOf m_tblSc2 Is ITable Then
            LookUpTablesEnvisionScenarioTblCheck(m_tblSc2)
          End If
          'SCENARIO 3 INPUT TABLE
          Try
            m_tblSc3 = pFeatWks.OpenTable("SCENARIO3")
          Catch ex As Exception
            m_strProcessing = m_strProcessing & "Missing table, SCENARIO3, executing geoprocessing tool to create table: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
            pCreateTable = New ESRI.ArcGIS.DataManagementTools.CreateTable
            pCreateTable.out_name = "SCENARIO3"
            pCreateTable.out_path = strPath
            RunTool(GP, pCreateTable)
            m_tblSc3 = pFeatWks.OpenTable("SCENARIO3")
          End Try
          If TypeOf m_tblSc3 Is ITable Then
            LookUpTablesEnvisionScenarioTblCheck(m_tblSc3)
          End If
          'SCENARIO 4 INPUT TABLE
          Try
            m_tblSc4 = pFeatWks.OpenTable("SCENARIO4")
          Catch ex As Exception
            m_strProcessing = m_strProcessing & "Missing table, SCENARIO4, executing geoprocessing tool to create table: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
            pCreateTable = New ESRI.ArcGIS.DataManagementTools.CreateTable
            pCreateTable.out_name = "SCENARIO4"
            pCreateTable.out_path = strPath
            RunTool(GP, pCreateTable)
            m_tblSc4 = pFeatWks.OpenTable("SCENARIO4")
          End Try
          If TypeOf m_tblSc4 Is ITable Then
            LookUpTablesEnvisionScenarioTblCheck(m_tblSc4)
          End If
          'SCENARIO 5 INPUT TABLE
          Try
            m_tblSc5 = pFeatWks.OpenTable("SCENARIO5")
          Catch ex As Exception
            m_strProcessing = m_strProcessing & "Missing table, SCENARIO5, executing geoprocessing tool to create table: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
            pCreateTable = New ESRI.ArcGIS.DataManagementTools.CreateTable
            pCreateTable.out_name = "SCENARIO5"
            pCreateTable.out_path = strPath
            RunTool(GP, pCreateTable)
            m_tblSc5 = pFeatWks.OpenTable("SCENARIO5")
          End Try
          If TypeOf m_tblSc5 Is ITable Then
            LookUpTablesEnvisionScenarioTblCheck(m_tblSc5)
          End If
          'SCENARIO SUMMATION TABLE
          Try
            m_tblScSummary = pFeatWks.OpenTable("SCENARIO_SUMMARY")
          Catch ex As Exception
            m_strProcessing = m_strProcessing & "Missing table, SCENARIO_SUMMARY, executing geoprocessing tool to create table: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
            pCreateTable = New ESRI.ArcGIS.DataManagementTools.CreateTable
            pCreateTable.out_name = "SCENARIO_SUMMARY"
            pCreateTable.out_path = strPath
            RunTool(GP, pCreateTable)
            m_tblScSummary = pFeatWks.OpenTable("SCENARIO_SUMMARY")
          End Try
          If TypeOf m_tblScSummary Is ITable Then
            LookUpTablesEnvisionSummaryTblCheck(m_tblScSummary)
          End If
          'TRAVEL TIMES/DISTANCES TABLE
          Try
            m_tblTravel = pFeatWks.OpenTable("TRAVEL_INPUTS")
          Catch ex As Exception
            m_strProcessing = m_strProcessing & "Missing table, TRAVEL_INPUTS, executing geoprocessing tool to create table: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
            pCreateTable = New ESRI.ArcGIS.DataManagementTools.CreateTable
            pCreateTable.out_name = "TRAVEL_INPUTS"
            pCreateTable.out_path = strPath
            RunTool(GP, pCreateTable)
            m_tblTravel = pFeatWks.OpenTable("TRAVEL_INPUTS")
          End Try
          If TypeOf m_tblTravel Is ITable Then
            LookUpTablesEnvisionTravelInputsTblCheck(m_tblTravel)
          End If
          'EXISTING LU TABLE
          Try
            m_tblExistingLU = pFeatWks.OpenTable("EXISTING_DEVELOPED_AREA")
          Catch ex As Exception
            m_tblExistingLU = Nothing
          End Try
          'EXISITNG LAND USE REFERENCE TABLE
          Try
            m_tblEXLURef = pFeatWks.OpenTable("EXISTING_LU_REFERENCE")
          Catch ex As Exception
            m_strProcessing = m_strProcessing & "Missing table, EXISTING_LU_REFERENCE, executing geoprocessing tool to create table: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
            pCreateTable = New ESRI.ArcGIS.DataManagementTools.CreateTable
            pCreateTable.out_name = "EXISTING_LU_REFERENCE"
            pCreateTable.out_path = strPath
            RunTool(GP, pCreateTable)
            m_tblEXLURef = pFeatWks.OpenTable("EXISTING_LU_REFERENCE")
            LookUpTablesEnvisionEXLUREFTblCheck(m_tblEXLURef)
          End Try
          If TypeOf m_tblTravel Is ITable Then

          End If
        End If
      End If
      GoTo CleanUp
    Catch ex As Exception
      m_strProcessing = m_strProcessing & "Error in creating envision lookup tables: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
      m_strProcessing = m_strProcessing & "Error Message:  " & vbNewLine & ex.Message & vbNewLine & "Stack Trace: " & ex.StackTrace & vbNewLine
      MessageBox.Show(ex.Message & vbNewLine & m_strProcessing, "Error in creating envision lookup tables", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
      LookUpTablesEnvisionCheck = False
      GoTo CleanUp
    End Try

CleanUp:
    m_strProcessing = m_strProcessing & "Ending function LookUpTablesEnvisionCheck: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine & vbNewLine
    GP = Nothing
    pCreateTable = Nothing
    pWksFactory = Nothing
    pFeatWks = Nothing
    GC.WaitForPendingFinalizers()
    GC.Collect()
  End Function

  Public Sub LoadExcelDevTypes()
    'LOAD THE DEVELOPMENT TYPE INFO FROM THE SELECTED ENVISION EXCEL FILE
    Try
      If OpenEnvisionExcel() Then
        'DELETE THE DEVTYPES AND SCENARIO TRACKING LOOKUP TABLES
        If DeleteEnvisionLookupTables() = False Then
          GoTo CleanUp
        End If
        'REVIEW THE LOOKUPTABLES
        LookUpTablesEnvisionCheck(m_strFeaturePath)
        'CLEAR AND RESET THE ENVISION ATTRIBUTES FIELD TABLE
        ResetEnvisionAttributeFieldTrackingTbl()
        'RETRIEVE DEVELOPMENT TYPE DATA FROM EXCEL 
        m_appEnvision.StatusBar.Message(0) = "Loading Development Types from selected Excel file.".ToUpper
        LoadDevTypesFromExcel()
        'LOAD THE TRAVEL INPUTS
        LoadTravelInputsFromExcel()
        'LOAD THE EXISTING LAND USE REFERENCE TABLE
        LoadEXLUInputsFromExcel()
        If LookUpTablesEnvisionCheck(m_strFeaturePath) Then
          m_appEnvision.StatusBar.Message(0) = "Retrieving Development Types"
          RetrieveDevTypeData()
          m_dockEnvisionWinForm.itmEnvisionExcel.Enabled = True
        Else
          GoTo CleanUp
        End If
        m_dockEnvisionWinForm.dgvDevelopmentTypes.Refresh()
        'UPDATE THE LEGEND
        If m_dockEnvisionWinForm.itmAutoUpdateLegend.Checked Then
          m_appEnvision.StatusBar.Message(0) = "Updating Envision legend"
          UpdateEnvisionLyrLegend()
        End If
        Dim intRow As Integer
        If m_dockEnvisionWinForm.dgvDevelopmentTypes.Rows.Count > 0 Then
          'CYCLE THROUGH DATA GRID VIEW ELEMENTS TO TRY AND REFRESH
          For intRow = 0 To m_dockEnvisionWinForm.dgvDevelopmentTypes.RowCount - 1
            m_dockEnvisionWinForm.dgvDevelopmentTypes.Rows(intRow).Cells(0).Selected = True
          Next
          m_dockEnvisionWinForm.dgvDevelopmentTypes.Rows(0).Cells(0).Selected = True
        End If
        GoTo CleanUp
      Else
        m_strEnvisionExcelFile = ""
        m_dockEnvisionWinForm.itmEnvisionExcel.ToolTipText = ""
        GoTo CleanUp
      End If
      GoTo CleanUp
    Catch ex As Exception
      m_xlPaintApp = Nothing
      m_xlPaintWB1 = Nothing
      MessageBox.Show(ex.Message, "Envision Excel Load Error")
    End Try

CleanUp:
    m_appEnvision.StatusBar.Message(0) = ""
    m_dockEnvisionWinForm.prgBarEnvision.Value = 0
    GC.Collect()
    GC.WaitForPendingFinalizers()
  End Sub


  Public Function DeleteEnvisionLookupTables() As Boolean
    DeleteEnvisionLookupTables = True
    'DELETE THE SUMMARY TABLES
    Dim GP As ESRI.ArcGIS.Geoprocessor.Geoprocessor
    Dim pDeleteTable As ESRI.ArcGIS.DataManagementTools.Delete
    GP = New ESRI.ArcGIS.Geoprocessor.Geoprocessor
    Dim pWksFactory As IWorkspaceFactory = Nothing
    Dim pFeatWks As IFeatureWorkspace = Nothing
    Dim pFeatClass As IFeatureClass = Nothing
    Dim featureClassName As String = String.Empty

    'REMOVE PRE-EXISTING INSTANCES OF SUMMARY TABLES
    Try
      If m_strFeaturePath.Length > 0 Then
        ' This process can destroy the featureclass of the scenario map layer and so we will repair it if needed.
        If m_pEditFeatureLyr IsNot Nothing Then featureClassName = m_pEditFeatureLyr.FeatureClass.AliasName
        GC.WaitForPendingFinalizers()
        GC.Collect()
        m_tblExistingLU = Nothing
        m_tblDevelopmentTypes = Nothing
        m_tblSc1 = Nothing
        m_tblSc2 = Nothing
        m_tblSc3 = Nothing
        m_tblSc4 = Nothing
        m_tblSc5 = Nothing
        GC.Collect()
        pDeleteTable = New ESRI.ArcGIS.DataManagementTools.Delete
        pDeleteTable.in_data = m_strFeaturePath & "\" & "DEVELOPMENT_TYPE_ATTRIBUTES"
        pDeleteTable.data_type = "Table"
        RunTool(GP, pDeleteTable)
        pDeleteTable = New ESRI.ArcGIS.DataManagementTools.Delete
        pDeleteTable.in_data = m_strFeaturePath & "\" & "SCENARIO1"
        pDeleteTable.data_type = "Table"
        RunTool(GP, pDeleteTable)
        pDeleteTable = New ESRI.ArcGIS.DataManagementTools.Delete
        pDeleteTable.in_data = m_strFeaturePath & "\" & "SCENARIO2"
        pDeleteTable.data_type = "Table"
        RunTool(GP, pDeleteTable)
        pDeleteTable = New ESRI.ArcGIS.DataManagementTools.Delete
        pDeleteTable.in_data = m_strFeaturePath & "\" & "SCENARIO3"
        pDeleteTable.data_type = "Table"
        RunTool(GP, pDeleteTable)
        pDeleteTable = New ESRI.ArcGIS.DataManagementTools.Delete
        pDeleteTable.in_data = m_strFeaturePath & "\" & "SCENARIO4"
        pDeleteTable.data_type = "Table"
        RunTool(GP, pDeleteTable)
        pDeleteTable = New ESRI.ArcGIS.DataManagementTools.Delete
        pDeleteTable.in_data = m_strFeaturePath & "\" & "SCENARIO5"
        pDeleteTable.data_type = "Table"
        RunTool(GP, pDeleteTable)
        If m_pEditFeatureLyr IsNot Nothing AndAlso m_pEditFeatureLyr.FeatureClass Is Nothing Then
          pWksFactory = New FileGDBWorkspaceFactory
          pFeatWks = DirectCast(pWksFactory.OpenFromFile(m_strFeaturePath, 0), IFeatureWorkspace)
          pFeatClass = pFeatWks.OpenFeatureClass(featureClassName)
          m_pEditFeatureLyr.FeatureClass = pFeatClass
        End If
      End If
    Catch ex As Exception
      MessageBox.Show("Error occured in attempting to delete previous versions of the Envison lookup tables. Error: " & vbNewLine & ex.Message, "Lookup Table Delete Error")
      DeleteEnvisionLookupTables = False
    End Try


    GP = Nothing
    pDeleteTable = Nothing
    pWksFactory = Nothing
    pFeatWks = Nothing
    pFeatClass = Nothing
    GC.WaitForPendingFinalizers()
    GC.Collect()
  End Function


  Public Sub LookUpTablesEnvisionGlobalVarTrackingTblCheck(ByVal pTable As ITable)
    'CHECK FOR GLOBAL VARIABLE LOOKUP TABLE FIELDS
    Dim intRowCount As Integer
    m_strProcessing = m_strProcessing & "Starting function LookUpTablesEnvisionGlobalVarTrackingTblCheck: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
    'CHECK FOR FIELD, VARIABLE_NAME, which holds global value name
    If pTable.FindField("VARIABLE_NAME") <= -1 Then
      If Not AddEnvisionField(pTable, "FIELD_NAME", "STRING", 50, 0) Then
        GoTo CleanUp
      End If
    End If
    'CHECK FOR FIELD, VARIABLE_INTEGER, which holds INTEGER global values
    If pTable.FindField("VARIABLE_INTEGER") <= -1 Then
      If Not AddEnvisionField(pTable, "VARIABLE_INTEGER", "INTEGER", 16, 0) Then
        GoTo CleanUp
      End If
    End If

    'CHECK FOR THE MINIMUM NUMBER OF RECORDS
    m_appEnvision.StatusBar.Message(0) = "CHECKING FOR REQUIRED DEVELOPMENT TYPE GRAPH TABLE LOOKUP TABLE FIELDS: CHECKING FOR NULL VALUES"
    intRowCount = pTable.RowCount(Nothing)
    If intRowCount = 0 Then
      For intRowCount = intRowCount To 1
        pTable.CreateRow()
      Next
    End If
    GoTo CleanUp

CleanUp:
    m_strProcessing = m_strProcessing & "Ending function LookUpTablesEnvisionGlobalVarTrackingTblCheck: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine & vbNewLine
    pTable = Nothing
    GC.WaitForPendingFinalizers()
    GC.Collect()
    m_appEnvision.StatusBar.Message(0) = ""
    m_dockEnvisionWinForm.prgBarEnvision.Value = 0
  End Sub


  Public Sub LookUpTablesEnvisionAttributeFieldTrackingTblCheck(ByVal pTable As ITable)
    'CHECK FOR FIELD TRACKING TABLE LOOKUP TABLE FIELDS
    'THESE FIELDS ARE DEFINED FROM THE SELECTED ENVISION EXCEL FILE AND THE DEV TYPE TAB
    m_strProcessing = m_strProcessing & "Starting function LookUpTablesEnvisionAttributeFieldTrackingTblCheck: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
    'CHECK FOR FIELD NAME FIELD
    If pTable.FindField("FIELD_NAME") <= -1 Then
      If Not AddEnvisionField(pTable, "FIELD_NAME", "STRING", 150, 0) Then
        GoTo CleanUp
      End If
    End If
    'CHECK FOR FIELD TYPE FIELD
    If pTable.FindField("FIELD_TYPE") <= -1 Then
      If Not AddEnvisionField(pTable, "FIELD_TYPE", "STRING", 16, 0) Then
        GoTo CleanUp
      End If
    End If
    'CHECK FOR FIELD ALIAS NAME FIELD
    If pTable.FindField("FIELD_ALIAS") <= -1 Then
      If Not AddEnvisionField(pTable, "FIELD_ALIAS", "STRING", 150, 0) Then
        GoTo CleanUp
      End If
    End If
    'CHECK FOR FIELD FIELD WIDTH FIELD
    If pTable.FindField("FIELD_WIDTH") <= -1 Then
      If Not AddEnvisionField(pTable, "FIELD_WIDTH", "INTEGER", 50, 0) Then
        GoTo CleanUp
      End If
    End If
    'CHECK FOR FIELD DECIMAL PLACE FIELD
    If pTable.FindField("FIELD_DECIMAL") <= -1 Then
      If Not AddEnvisionField(pTable, "FIELD_DECIMAL", "INTEGER", 50, 0) Then
        GoTo CleanUp
      End If
    End If
    'CHECK FOR USE FIELD..indicates if the field is in "use" for storying value
    If pTable.FindField("USE") <= -1 Then
      If Not AddEnvisionField(pTable, "USE", "INTEGER", 1, 0) Then
        GoTo CleanUp
      End If
    End If
    'CHECK FOR CALC BY ACRES FIELD..indicates if the field value will be calculate by acres
    If pTable.FindField("CALC_BY_ACRES") <= -1 Then
      If Not AddEnvisionField(pTable, "CALC_BY_ACRES", "INTEGER", 1, 0) Then
        GoTo CleanUp
      End If
    End If
    'CHECK FOR CALC FIELD NAME FIELD..Name of the shapefile field is different than dev type field, because is a calculated value
    If pTable.FindField("CALC_FIELD_NAME") <= -1 Then
      If Not AddEnvisionField(pTable, "CALC_FIELD_NAME", "STRING", 150, 0) Then
        GoTo CleanUp
      End If
    End If
    'CHECK FOR CALC FIELD NAME FIELD..Name of the shapefile field is different than dev type field, because is a calculated value
    If pTable.FindField("CALC_ONLY") <= -1 Then
      If Not AddEnvisionField(pTable, "CALC_ONLY", "INTEGER", 1, 0) Then
        GoTo CleanUp
      End If
    End If
    'CHECK FOR CALC FIELD NAME FIELD..Name of the shapefile field is different than dev type field, because is a calculated value
    If pTable.FindField("DEP_VAR_FIELD_NAME") <= -1 Then
      If Not AddEnvisionField(pTable, "DEP_VAR_FIELD_NAME", "STRING", 150, 0) Then
        GoTo CleanUp
      End If
    End If
    ' Total field
    If pTable.FindField("Calc_Total") = -1 Then
      If Not AddEnvisionField(pTable, "Calc_Total", "INTEGER", 1, 0) Then GoTo cleanup
    End If
    ' Total field name
    If pTable.FindField("Total_Field_Name") = -1 Then
      If Not AddEnvisionField(pTable, "Total_Field_Name", "STRING", 150, 0) Then GoTo cleanup
    End If
    ' Total field name for neighborhood aggregator
    If pTable.FindField("Total_Field_Name_Sum") = -1 Then
      If Not AddEnvisionField(pTable, "Tot_Fld_Name_Sum", "STRING", 150, 0) Then GoTo cleanup
    End If
    ' Total field name for neighborhood aggregator when calculated by acres
    If pTable.FindField("Tot_Fld_Name_Sum_By_Acres") = -1 Then
      If Not AddEnvisionField(pTable, "Tot_Fld_Name_Sum_By_Acres", "STRING", 150, 0) Then GoTo cleanup
    End If
    GoTo CleanUp

CleanUp:
    m_strProcessing = m_strProcessing & "Ending function LookUpTablesEnvisionAttributeFieldTrackingTblCheck: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine & vbNewLine
    pTable = Nothing
    GC.WaitForPendingFinalizers()
    GC.Collect()
    'm_appEnvision.StatusBar.Message(0) = ""
    'm_dockEnvisionWinForm.prgBarEnvision.Value = 0
  End Sub

  Public Sub LookUpTablesEnvisionDevTypeGraphTblCheck(ByVal pTable As ITable)
    'CHECK FOR REQUIRED DEVELOPMENT TYPE TABLE LOOKUP TABLE FIELDS
    m_strProcessing = m_strProcessing & "Starting function LookUpTablesEnvisionDevTypeGraphTblCheck: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
    Dim shtDevType As Microsoft.Office.Interop.Excel.Worksheet = Nothing
    Dim intCol As Integer
    Dim strCellValue As String = ""
    Dim intRed As Integer = 0
    Dim intGreen As Integer = 0
    Dim intBlue As Integer = 0
    Dim bmpTemp As Bitmap
    Dim intRow As Integer
    Dim intRowCount As Integer
    Dim intFldCount As Integer
    Dim pField As IField
    Dim intFld As Integer
    Dim strVal As String
    Dim intVal As Integer
    Dim rowTemp As IRow

    'CHECK FOR HOUSING UNIT LABEL FIELD
    If pTable.FindField("HU_LABEL") <= -1 Then
      If Not AddEnvisionField(pTable, "HU_LABEL", "STRING", 50, 0) Then
        GoTo CleanUp
      End If
    End If
    'CHECK FOR HOUSING_UNITS FIELD
    If pTable.FindField("HOUSING_UNITS") <= -1 Then
      If Not AddEnvisionField(pTable, "HOUSING_UNITS", "DOUBLE", 16, 4) Then
        GoTo CleanUp
      End If
    End If
    'CHECK FOR EMPLOYMENT LABEL FIELD
    If pTable.FindField("EMP_LABEL") <= -1 Then
      If Not AddEnvisionField(pTable, "EMP_LABEL", "STRING", 50, 0) Then
        GoTo CleanUp
      End If
    End If 'CHECK FOR EMPLOYMENT FIELD
    If pTable.FindField("EMPLOYMENT") <= -1 Then
      If Not AddEnvisionField(pTable, "EMPLOYMENT", "DOUBLE", 16, 4) Then
        GoTo CleanUp
      End If
    End If

    'CHECK FOR THE MINIMUM NUMBER OF RECORDS
    m_appEnvision.StatusBar.Message(0) = "CHECKING FOR REQUIRED DEVELOPMENT TYPE GRAPH TABLE LOOKUP TABLE FIELDS: CHECKING FOR NULL VALUES"
    intRowCount = pTable.RowCount(Nothing)
    If intRowCount < 3 Then
      For intRowCount = intRowCount To 2
        pTable.CreateRow()
      Next
    End If

    'SET ALL NULL VALUES TO DEFAULTS of "" OR 0
    intRowCount = pTable.RowCount(Nothing)
    intFldCount = pTable.Fields.FieldCount
    Dim blnStore As Boolean
    If intRowCount > 0 Then
      For intRow = 1 To intRowCount
        blnStore = False
        rowTemp = pTable.GetRow(intRow)
        For intFld = 1 To intFldCount - 1
          pField = pTable.Fields.Field(intFld)
          Try
            If pField.Type = esriFieldType.esriFieldTypeString Then
              Try
                strVal = rowTemp.Value(intFld).ToString
              Catch ex As Exception
                blnStore = True
                rowTemp.Value(intFld) = ""
              End Try
            Else
              Try
                intVal = CInt(rowTemp.Value(intFld))
              Catch ex As Exception
                blnStore = True
                rowTemp.Value(intFld) = 0
              End Try
            End If
          Catch ex As Exception
          End Try
        Next
        If blnStore Then
          rowTemp.Store()
        End If
      Next
    End If

    'SET THE HU AND EMP LABELS
    For intRowCount = 1 To 3
      blnStore = False
      rowTemp = pTable.GetRow(intRowCount)
      If Not rowTemp.Value(pTable.FindField("HU_LABEL")).ToString.Length > 0 Then
        If intRowCount = 1 Then
          rowTemp.Value(pTable.FindField("HU_LABEL")) = "SF %"
        ElseIf intRowCount = 2 Then
          rowTemp.Value(pTable.FindField("HU_LABEL")) = "TH %"
        ElseIf intRowCount = 3 Then
          rowTemp.Value(pTable.FindField("HU_LABEL")) = "MF %"
        End If
        blnStore = True
      End If
      If Not rowTemp.Value(pTable.FindField("EMP_LABEL")).ToString.Length > 0 Then
        If intRowCount = 1 Then
          rowTemp.Value(pTable.FindField("EMP_LABEL")) = "RET %"
        ElseIf intRowCount = 2 Then
          rowTemp.Value(pTable.FindField("EMP_LABEL")) = "OFF %"
        ElseIf intRowCount = 3 Then
          rowTemp.Value(pTable.FindField("EMP_LABEL")) = "IND %"
        End If
        blnStore = True
      End If
      If blnStore Then
        rowTemp.Store()
      End If
    Next
    GoTo CleanUp

CleanUp:
    m_strProcessing = m_strProcessing & "Ending function LookUpTablesEnvisionDevTypeGraphTblCheck: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine & vbNewLine
    m_blnLoadingDevTypes = False
    shtDevType = Nothing
    intCol = Nothing
    strCellValue = Nothing
    intRed = Nothing
    intGreen = Nothing
    intBlue = Nothing
    bmpTemp = Nothing
    pField = Nothing
    strVal = Nothing
    intVal = Nothing
    rowTemp = Nothing
    blnStore = Nothing
    GC.Collect()
    GC.WaitForPendingFinalizers()
    m_blnLoadingDevTypes = False
    m_appEnvision.StatusBar.Message(0) = ""
    m_dockEnvisionWinForm.prgBarEnvision.Value = 0
  End Sub

  Public Function LookUpTablesEnvisionScenarioTblCheck(ByVal pTable As ITable) As Boolean
    'CHECK FOR THE REQUIRED FIELDS
    m_strProcessing = m_strProcessing & "Starting function LookUpTablesEnvisionScenarioTblCheck: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
    Dim strFields As String
    Dim strFieldName As String
    Dim intFldCount As Integer
    Dim intRow As Integer
    Dim intRowCount As Integer
    Dim intFld As Integer
    Dim rowTemp As IRow
    Dim pField As IField
    Dim strVal As String
    Dim intVal As Integer
    Dim blnStore As Boolean
    Dim shtScenario As Microsoft.Office.Interop.Excel.Worksheet = Nothing
    Dim intCol As Integer = 0
    Dim strFldCellValue As String = ""
    Dim intScTabStartRow As Integer = 4
    Dim intScTabFieldRow As Integer = -1

    'SET THE ARCGIS VIEW STATUS MESSAGE
    If Not m_appEnvision Is Nothing Then
      m_appEnvision.StatusBar.Message(0) = "CHECKING FOR REQUIRED SCENARIO TABLE LOOKUP TABLE FIELDS"
    End If
    m_dockEnvisionWinForm.prgBarEnvision.Value = 5

    If pTable.FindField("DEVELOPMENT_TYPE") <= -1 Then
      If Not m_appEnvision Is Nothing Then
        m_appEnvision.StatusBar.Message(0) = "CHECKING FOR REQUIRED SCENARIO SUMMARY TABLE LOOKUP TABLE FIELDS: FIELD (DEVELOPMENT_TYPE)"
      End If
      m_dockEnvisionWinForm.prgBarEnvision.Value = 15

      If Not AddEnvisionField(pTable, "DEVELOPMENT_TYPE", "STRING", 150, 0) Then
        Return False
      End If
    End If

    'TWO REQUIRED FIEDS
    strFields = "VAC_ACRE,DEVD_ACRE"
    intFldCount = strFields.Split(","c).Length
    For Each strFieldName In strFields.Split(","c)
      If Not m_appEnvision Is Nothing Then
        m_appEnvision.StatusBar.Message(0) = "CHECKING FOR REQUIRED ENVISION SCENARIO TRACKING TABLE FIELDS AND FIELD CALCULATIONS: FIELD (" & strFieldName & ")"
      End If
      m_dockEnvisionWinForm.prgBarEnvision.Value = 25

      If pTable.FindField(strFieldName) >= 0 Then
        Continue For
      End If
      If Not AddEnvisionField(pTable, strFieldName, "DOUBLE", 16, 6) Then
        Return False
      End If
      intFld = intFld + 1
      m_dockEnvisionWinForm.prgBarEnvision.Value = CType((intFld / intFldCount) * 100, Int32)
    Next

    '---------------------------------------------------------------------------------------------------------------------------------------------
    'IF EXCEL IS OPEN, THEN REVIEW SCENARIO TAB FOR REQUIRED "EX_" FIELDS
    'RETRIEVE OR OPEN THE SELECTED ENVISION EXCEL FILE
    If m_xlPaintWB1 Is Nothing Then
      GoTo CleanUp
    End If
    Try
      Try
        If Not TypeOf m_xlPaintWB1.Sheets("SCENARIO" & m_intEditScenario.ToString) Is Microsoft.Office.Interop.Excel.Worksheet Then
        End If
      Catch ex As Exception
        CloseEnvisionExcel()
        OpenEnvisionExcel()
      End Try
      If m_xlPaintApp Is Nothing Then
        m_xlPaintApp = New Microsoft.Office.Interop.Excel.Application
        m_xlPaintApp.DisplayAlerts = False
        m_xlPaintApp.WindowState = Interop.Excel.XlWindowState.xlMinimized
        m_xlPaintApp.Visible = True
      End If
      If m_xlPaintWB1 Is Nothing Then
        m_xlPaintWB1 = CType(m_xlPaintApp.Workbooks.Open(m_strEnvisionExcelFile), Microsoft.Office.Interop.Excel.Workbook)
      End If
    Catch ex As Exception
      'MessageBox.Show(ex.Message, "Opening Envision Excel File Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
      GoTo CloseExcel
    End Try


    If Not m_xlPaintWB1 Is Nothing And Not m_xlPaintApp Is Nothing Then
      Try
        'RETRIEVE SCENARIO TAB
        Try
          If Not TypeOf m_xlPaintWB1.Sheets("SCENARIO" & m_intEditScenario.ToString) Is Microsoft.Office.Interop.Excel.Worksheet Then
            MessageBox.Show("The 'SCENARIO" & m_intEditScenario.ToString & "' tab could not be found in the selected ENVISION Excel file. Please select another file.", "SCENARIO TAB NOT FOUND", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo CloseExcel
          Else
            shtScenario = DirectCast(m_xlPaintWB1.Sheets("SCENARIO" & m_intEditScenario.ToString), Microsoft.Office.Interop.Excel.Worksheet)
          End If
        Catch ex As Exception
          MessageBox.Show("The 'SCENARIO" & m_intEditScenario.ToString & "' tab could not be found in the selected ENVISION Excel file. Please select another file.", "SCENARIO TAB NOT FOUND", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
          GoTo CloseExcel
        End Try


        'REVIEW SCENARIO TAB ROWS TO DETERMINE THE START ROW AND FIELD ROW (IF AVAILABLE)
        For intRow = 1 To 35
          Try
            strFldCellValue = CStr(CType(shtScenario.Cells(intRow, 1), Microsoft.Office.Interop.Excel.Range).Value)
            If strFldCellValue.Contains(",") Then
              strFldCellValue = CStr(strFldCellValue.Split(","c)(0))
            End If
          Catch ex As Exception
            strFldCellValue = ""
          End Try
          If UCase(strFldCellValue) = "DEV_TYPE" Then
            intScTabFieldRow = intRow
            intScTabStartRow = intRow + 1
            Exit For
          End If
        Next

        'BUILD LIST OF COLUMNS TO PROCESS
        For intCol = 2 To 400
          Try
            strFieldName = CStr(CType(shtScenario.Cells(intScTabFieldRow, intCol), Microsoft.Office.Interop.Excel.Range).Value)
            If Not strFieldName = "" Then
              If pTable.FindField(strFieldName) >= 0 Then
                Continue For
              End If
              If Not AddEnvisionField(pTable, strFieldName, "DOUBLE", 16, 6) Then
                Return False
              End If
            Else
              Continue For
            End If
          Catch ex As Exception
          End Try
        Next

      Catch ex As Exception

      End Try
    End If
    '---------------------------------------------------------------------------------------------------------------------------------------------

    'CHECK FOR THE MINIMUM NUMBER OF RECORDS BASED ON DEV TYPE MAXIMUM VALUE
    intRowCount = pTable.RowCount(Nothing)
    'If intRowCount <= m_intDevTypeMax Then
    '    For intRowCount = intRowCount To m_intDevTypeMax
    '        pTable.CreateRow()
    '    Next
    'End If

    'SET ALL NULL VALUES TO DEFAULTS of "" OR 0
    If Not m_appEnvision Is Nothing Then
      m_appEnvision.StatusBar.Message(0) = "CHECKING FOR REQUIRED SCENARIO TABLE LOOKUP TABLE FIELDS: CHECKING FOR NULL VALUES"
    End If
    m_dockEnvisionWinForm.prgBarEnvision.Value = 45

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
              strVal = rowTemp.Value(intFld).ToString
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
          If Not m_appEnvision Is Nothing Then
            m_appEnvision.StatusBar.Message(0) = "CHECKING FOR REQUIRED SCENARIO TABLE LOOKUP TABLE FIELDS: CHECKING FOR NULL VALUES; FIELD (" & pField.Name & "); RECORD (" & intRow.ToString & ")"
          End If
          m_dockEnvisionWinForm.prgBarEnvision.Value = 65

        Next
        If blnStore Then
          rowTemp.Store()
        End If
      Next
    End If
    GoTo CleanUp

CloseExcel:
    CloseEnvisionExcel()
    GoTo CleanUp

CleanUp:
    m_strProcessing = m_strProcessing & "Ending function LookUpTablesEnvisionScenarioTblCheck: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine & vbNewLine
    strFields = Nothing
    strFieldName = Nothing
    intFldCount = Nothing
    intRowCount = Nothing
    intFld = Nothing
    pField = Nothing
    strVal = Nothing
    intVal = Nothing
    rowTemp = Nothing
    GC.Collect()
    ' GC.WaitForPendingFinalizers()
    If Not m_appEnvision Is Nothing Then
      m_appEnvision.StatusBar.Message(0) = ""
    End If
    m_dockEnvisionWinForm.prgBarEnvision.Value = 0
  End Function

  Public Function LookUpTablesEnvisionSummaryTblCheck(ByVal pTable As ITable) As Boolean
    m_strProcessing = m_strProcessing & "Starting function LookUpTablesEnvisionSummaryTblCheck: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
    'CHECK FOR REQUIRED SCENARIO SUMMARY TABLE LOOKUP TABLE FIELDS
    Dim strFields As String
    Dim strFieldName As String
    Dim intFld As Integer = 0
    Dim intFldCount As Integer = 21
    Dim intRowCount As Integer
    Dim rowTemp As IRow
    Dim pField As IField
    Dim strVal As String
    Dim intVal As Integer
    Dim blnStore As Boolean

    'SET THE ARCGIS VIEW STATUS MESSAGE
    m_appEnvision.StatusBar.Message(0) = "CHECKING FOR REQUIRED SCENARIO SUMMARY TABLE LOOKUP TABLE FIELDS"
    m_dockEnvisionWinForm.prgBarEnvision.Value = 0

    If pTable.FindField("LAYER_NAME") <= -1 Then
      m_appEnvision.StatusBar.Message(0) = "CHECKING FOR REQUIRED SCENARIO SUMMARY TABLE LOOKUP TABLE FIELDS: FIELD (LAYER_NAME)"
      If Not AddEnvisionField(pTable, "LAYER_NAME", "STRING", 250, 0) Then
        Return False
      End If
    End If
    intFld = intFld + 1
    m_dockEnvisionWinForm.prgBarEnvision.Value = CType((intFld / intFldCount) * 100, Int32)

    If pTable.FindField("SCENARIO") <= -1 Then
      m_appEnvision.StatusBar.Message(0) = "CHECKING FOR REQUIRED SCENARIO SUMMARY TABLE LOOKUP TABLE FIELDS: FIELD (SCENARIO)"
      If Not AddEnvisionField(pTable, "SCENARIO", "STRING", 150, 0) Then
        Return False
      End If
    End If
    intFld = intFld + 1
    m_dockEnvisionWinForm.prgBarEnvision.Value = CType((intFld / intFldCount) * 100, Int32)

    'ACRES, HOUSING AND EMPLOYMENT VARIABLES
    strFields = "VACANT_ACRES,DEVELOPED_ACRES"
    For Each strFieldName In strFields.Split(","c)
      m_appEnvision.StatusBar.Message(0) = "CHECKING FOR REQUIRED SCENARIO SUMMARY TABLE LOOKUP TABLE FIELDS: FIELD (" & strFieldName & ")"
      If pTable.FindField(strFieldName) <= -1 Then
        If Not AddEnvisionField(pTable, strFieldName, "DOUBLE", 16, 10) Then
          Return False
        End If
      End If
      intFld = intFld + 1
      m_dockEnvisionWinForm.prgBarEnvision.Value = CType((intFld / intFldCount) * 100, Int32)
    Next

    'CHECK FOR THE SET NUMBER OF RECORDS
    intRowCount = pTable.RowCount(Nothing)
    If Not intRowCount = 5 Then
      pTable.DeleteSearchedRows(Nothing)
      intFld = pTable.FindField("SCENARIO")
      For intRowCount = 1 To 5
        rowTemp = pTable.CreateRow()
        If intRowCount = 1 Then
          rowTemp.Value(intFld) = "SCENARIO 1"
        ElseIf intRowCount = 2 Then
          rowTemp.Value(intFld) = "SCENARIO 2"
        ElseIf intRowCount = 3 Then
          rowTemp.Value(intFld) = "SCENARIO 3"
        ElseIf intRowCount = 4 Then
          rowTemp.Value(intFld) = "SCENARIO 4"
        ElseIf intRowCount = 5 Then
          rowTemp.Value(intFld) = "SCENARIO 5"
        End If
        rowTemp.Store()
      Next
    End If

    'SET ALL NULL VALUES TO DEFAULTS of "" OR 0
    m_appEnvision.StatusBar.Message(0) = "CHECKING FOR REQUIRED SCENARIO SUMMARY TABLE LOOKUP TABLE FIELDS: CHECKING FOR NULL VALUES"
    intRowCount = pTable.RowCount(Nothing)
    intFldCount = pTable.Fields.FieldCount
    If intRowCount > 0 Then
      For intRowCount = 1 To intRowCount
        blnStore = False
        rowTemp = pTable.GetRow(intRowCount)
        For intFld = 1 To intFldCount - 1
          pField = pTable.Fields.Field(intFld)
          If pField.Type = esriFieldType.esriFieldTypeString Then
            Try
              strVal = rowTemp.Value(intFld).ToString
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
    GoTo CleanUp

CleanUp:
    m_strProcessing = m_strProcessing & "Ending function LookUpTablesEnvisionSummaryTblCheck: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
    strFields = Nothing
    strFieldName = Nothing
    intFld = Nothing
    rowTemp = Nothing
    intRowCount = Nothing
    pField = Nothing
    strVal = Nothing
    intVal = Nothing
    rowTemp = Nothing
    blnStore = Nothing
    GC.Collect()
    GC.WaitForPendingFinalizers()
    m_appEnvision.StatusBar.Message(0) = ""
    m_dockEnvisionWinForm.prgBarEnvision.Value = 0
  End Function

  Public Function LookUpTablesEnvisionEXLUREFTblCheck(ByVal pTable As ITable) As Boolean
    m_strProcessing = m_strProcessing & "Starting function LookUpTablesEnvisionEXLUREFTblCheck: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
    'CHECK FOR REQUIRED SCENARIO SUMMARY TABLE LOOKUP TABLE FIELDS
    Dim strFields As String
    Dim strFieldName As String
    Dim intFld As Integer = 0
    Dim intRowCount As Integer = 0
    Dim rowTemp As IRow
    Dim pField As Field
    Dim strVal As String
    Dim intVal As Integer
    Dim blnStore As Boolean
    'Dim strLUNames As String = "Mixed-Use,Multifamily,Townhome,Single Family Small Lot,Single Family Conventional Lot,Single Family Large Lot,Mobile Home,Retail,Office,Industrial,Public / Civic,Educational,Hotel / Hospitality,Utilities / Infrastructure,Agricultural,Open Space,Vacant"
    Dim strLUNames As String = "Mixed-Use,Multifamily,Townhome,Single Family Small Lot,Single Family Conventional Lot,Single Family Large Lot,Mobile,Retail,Office,Industrial,Public / Civic,Educational,Hotel / Hospitality,Utilities / Infrastructure,Commercial Parking,Agricultural,Open Space,Vacant,Unknown"
    'Dim strEXLU As String = "MU,MF,TH,SF_SM,SF_MD,SF_LRG,MH,RET,OFF,IND,PUB,EDU,HOTEL,UTIL,AG,OS,VAC"
    Dim strEXLU As String = "MU,MF,TH,SF_SM,SF_MD,SF_LRG,MH,RET,OFF,IND,PUB,EDU,HOTEL,UTIL,PKG,AG,OS,VAC,NONE"
    Dim strEXLUMIX As String = "COM,RES,RES,RES,RES,RES,RES,COM,COM,COM,PUB,PUB,COM,,,,,,"
    Dim strValue As String

    'SET THE ARCGIS VIEW STATUS MESSAGE
    m_appEnvision.StatusBar.Message(0) = "CHECKING FOR REQUIRED EXISITNG LU REFERENCE TABLE LOOKUP TABLE FIELDS"
    m_dockEnvisionWinForm.prgBarEnvision.Value = 0

    If pTable.FindField("EX_LU_NAME") <= -1 Then
      m_appEnvision.StatusBar.Message(0) = "CHECKING FOR REQUIRED TRAVEL INPUTS TABLE LOOKUP TABLE FIELDS: FIELD (EX_LU_NAME)"
      If Not AddEnvisionField(pTable, "EX_LU_NAME", "STRING", 100, 0) Then
        Return False
      End If
    End If
    m_dockEnvisionWinForm.prgBarEnvision.Value = 25

    If pTable.FindField("EX_LU") <= -1 Then
      m_appEnvision.StatusBar.Message(0) = "CHECKING FOR REQUIRED TRAVEL INPUTS TABLE LOOKUP TABLE FIELDS: FIELD (EX_LU)"
      If Not AddEnvisionField(pTable, "EX_LU", "STRING", 10, 0) Then
        Return False
      End If
    End If
    m_dockEnvisionWinForm.prgBarEnvision.Value = 50

    If pTable.FindField("EX_LU_MIX_TYPE") <= -1 Then
      m_appEnvision.StatusBar.Message(0) = "CHECKING FOR REQUIRED TRAVEL INPUTS TABLE LOOKUP TABLE FIELDS: FIELD (EX_LU_MIX_TYPE)"
      If Not AddEnvisionField(pTable, "EX_LU_MIX_TYPE", "STRING", 10, 0) Then
        Return False
      End If
    End If
    m_dockEnvisionWinForm.prgBarEnvision.Value = 75

    Try
      'CHECK FOR THE SET NUMBER OF RECORDS
      intRowCount = pTable.RowCount(Nothing)
      If Not intRowCount = 20 Then
        pTable.DeleteSearchedRows(Nothing)
        For intRowCount = 1 To m_intDevTypeMax
          rowTemp = pTable.CreateRow()
          rowTemp.Value(1) = ""
          rowTemp.Value(2) = ""
          rowTemp.Value(3) = ""
          rowTemp.Store()
        Next
      Else
        GoTo CleanUp
      End If
      'LOAD IN DEFAULT VALUES
      m_dockEnvisionWinForm.prgBarEnvision.Value = 90
      intRowCount = 0
      For Each strValue In Split(strLUNames, ",")
        rowTemp = pTable.GetRow((intRowCount + 1))
        rowTemp.Value(1) = strValue
        rowTemp.Value(2) = Split(strEXLU, ",")(intRowCount)
        rowTemp.Value(3) = Split(strEXLUMIX, ",")(intRowCount)
        rowTemp.Store()
        intRowCount = intRowCount + 1
      Next
      m_dockEnvisionWinForm.prgBarEnvision.Value = 100
    Catch ex As Exception

    End Try

    GoTo CleanUp

CleanUp:
    m_strProcessing = m_strProcessing & "Ending function LookUpTablesEnvisionEXLUREFTblCheck: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
    strFields = Nothing
    strFieldName = Nothing
    intFld = Nothing
    rowTemp = Nothing
    intRowCount = Nothing
    pField = Nothing
    strVal = Nothing
    intVal = Nothing
    rowTemp = Nothing
    blnStore = Nothing
    GC.Collect()
    GC.WaitForPendingFinalizers()
    m_appEnvision.StatusBar.Message(0) = ""
    m_dockEnvisionWinForm.prgBarEnvision.Value = 0
  End Function

  Public Function LookUpTablesEnvisionTravelInputsTblCheck(ByVal pTable As ITable) As Boolean
    m_strProcessing = m_strProcessing & "Starting function LookUpTablesEnvisionTravelInputsTblCheck: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
    'CHECK FOR REQUIRED SCENARIO SUMMARY TABLE LOOKUP TABLE FIELDS
    Dim strFields As String
    Dim strFieldName As String
    Dim intFld As Integer = 0
    Dim intRowCount As Integer = 0
    Dim rowTemp As IRow
    Dim pField As Field
    Dim strVal As String
    Dim intVal As Integer
    Dim blnStore As Boolean

    'SET THE ARCGIS VIEW STATUS MESSAGE
    m_appEnvision.StatusBar.Message(0) = "CHECKING FOR REQUIRED TRAVEL INPUTS TABLE LOOKUP TABLE FIELDS"
    m_dockEnvisionWinForm.prgBarEnvision.Value = 0

    If pTable.FindField("DISTANCE_MILES") <= -1 Then
      m_appEnvision.StatusBar.Message(0) = "CHECKING FOR REQUIRED TRAVEL INPUTS TABLE LOOKUP TABLE FIELDS: FIELD (DISTANCE_MILES)"
      If Not AddEnvisionField(pTable, "DISTANCE_MILES", "DOUBLE", 16, 6) Then
        Return False
      End If
    End If
    m_dockEnvisionWinForm.prgBarEnvision.Value = 25

    If pTable.FindField("TIME_MINUTES") <= -1 Then
      m_appEnvision.StatusBar.Message(0) = "CHECKING FOR REQUIRED TRAVEL INPUTS TABLE LOOKUP TABLE FIELDS: FIELD (TIME_MINUTES)"
      If Not AddEnvisionField(pTable, "TIME_MINUTES", "DOUBLE", 16, 6) Then
        Return False
      End If
    End If
    m_dockEnvisionWinForm.prgBarEnvision.Value = 50

    If pTable.FindField("TRAVEL_MODE") <= -1 Then
      m_appEnvision.StatusBar.Message(0) = "CHECKING FOR REQUIRED TRAVEL INPUTS TABLE LOOKUP TABLE FIELDS: FIELD (TRAVEL_MODE)"
      If Not AddEnvisionField(pTable, "TRAVEL_MODE", "STRING", 10, 0) Then
        Return False
      End If
    End If
    m_dockEnvisionWinForm.prgBarEnvision.Value = 75

    If pTable.FindField("SPEED_MPH") <= -1 Then
      m_appEnvision.StatusBar.Message(0) = "CHECKING FOR REQUIRED TRAVEL INPUTS TABLE LOOKUP TABLE FIELDS: FIELD (SPEED_MPH)"
      If Not AddEnvisionField(pTable, "SPEED_MPH", "INTEGER", 16, 0) Then
        Return False
      End If
    End If
    m_dockEnvisionWinForm.prgBarEnvision.Value = 80


    Try
      'CHECK FOR THE SET NUMBER OF RECORDS
      intRowCount = pTable.RowCount(Nothing)
      If Not intRowCount = 10 Then
        pTable.DeleteSearchedRows(Nothing)
        For intRowCount = 1 To 10
          rowTemp = pTable.CreateRow()
        Next
      Else
        GoTo CleanUp
      End If
      'LOAD IN DEFAULT VALUES
      m_dockEnvisionWinForm.prgBarEnvision.Value = 90
      rowTemp = pTable.GetRow(1)
      rowTemp.Value(1) = 0.25
      rowTemp.Value(2) = 5
      rowTemp.Value(3) = "Walk"
      rowTemp.Value(4) = 3
      rowTemp.Store()
      rowTemp = pTable.GetRow(2)
      rowTemp.Value(1) = 0.25
      rowTemp.Value(2) = 1.5
      rowTemp.Value(3) = "Bike"
      rowTemp.Value(4) = 10
      rowTemp.Store()
      rowTemp = pTable.GetRow(3)
      rowTemp.Value(1) = 0.5
      rowTemp.Value(2) = 10
      rowTemp.Value(3) = "Walk"
      rowTemp.Value(4) = 3
      rowTemp.Store()
      rowTemp = pTable.GetRow(4)
      rowTemp.Value(1) = 0.5
      rowTemp.Value(2) = 3
      rowTemp.Value(3) = "Bike"
      rowTemp.Value(4) = 10
      rowTemp.Store()
      rowTemp = pTable.GetRow(5)
      rowTemp.Value(1) = 1
      rowTemp.Value(2) = 20
      rowTemp.Value(3) = "Walk"
      rowTemp.Value(4) = 3
      rowTemp.Store()
      rowTemp = pTable.GetRow(6)
      rowTemp.Value(1) = 1
      rowTemp.Value(2) = 6
      rowTemp.Value(3) = "Bike"
      rowTemp.Value(4) = 10
      rowTemp.Store()
      rowTemp = pTable.GetRow(7)
      rowTemp.Value(1) = 4.5
      rowTemp.Value(2) = 10
      rowTemp.Value(3) = "Auto"
      rowTemp.Value(4) = 27
      rowTemp.Store()
      rowTemp = pTable.GetRow(8)
      rowTemp.Value(1) = 9
      rowTemp.Value(2) = 20
      rowTemp.Value(3) = "Auto"
      rowTemp.Value(4) = 27
      rowTemp.Store()
      rowTemp = pTable.GetRow(9)
      rowTemp.Value(1) = 13.5
      rowTemp.Value(2) = 30
      rowTemp.Value(3) = "Auto"
      rowTemp.Value(4) = 27
      rowTemp.Store()
      rowTemp = pTable.GetRow(10)
      rowTemp.Value(1) = 10
      rowTemp.Value(2) = 30
      rowTemp.Value(3) = "Transit"
      rowTemp.Value(4) = 20
      rowTemp.Store()
      m_dockEnvisionWinForm.prgBarEnvision.Value = 100
    Catch ex As Exception

    End Try

    GoTo CleanUp

CleanUp:
    m_strProcessing = m_strProcessing & "Ending function LookUpTablesEnvisionTravelInputsTblCheck: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
    strFields = Nothing
    strFieldName = Nothing
    intFld = Nothing
    rowTemp = Nothing
    intRowCount = Nothing
    pField = Nothing
    strVal = Nothing
    intVal = Nothing
    rowTemp = Nothing
    blnStore = Nothing
    GC.Collect()
    GC.WaitForPendingFinalizers()
    m_appEnvision.StatusBar.Message(0) = ""
    m_dockEnvisionWinForm.prgBarEnvision.Value = 0
  End Function

  Public Sub LoadDevTypesFromExcel()
    'POPULATE THE ENVISION LOOKUP TABLES WITH VALUES FROM THE SELECTED ENVISION EXCEL FILE
    Dim shtDevType As Microsoft.Office.Interop.Excel.Worksheet = Nothing
    Dim strFieldName As String = ""
    Dim intField As Integer
    Dim strCellValue As String = ""
    Dim strFldCellValue As String = ""
    Dim strFieldWidth As Integer = 0
    Dim strFieldDecimals As Integer = 0

    Dim intCol As Integer
    Dim intRow As Integer
    Dim rowDevType As IRow
    Dim rowSc1 As IRow
    Dim rowSc2 As IRow
    Dim rowSc3 As IRow
    Dim rowSc4 As IRow
    Dim rowSc5 As IRow
    Dim intRowNmbr As Integer
    Dim intProgress As Integer
    Dim pField As IField
    Dim intFieldRow As Integer = 5
    Dim intStartRow As Integer = 6
    Dim intSearchMax As Integer = 6
    Dim strDevTypeName As String = ""
    Dim intRecCount As Integer = 0
    Dim intDevCount As Integer = 0
    Dim intAddDeleteRows As Integer = 0


    If m_xlPaintWB1 Is Nothing Or m_xlPaintApp Is Nothing Or m_tblDevelopmentTypes Is Nothing Then
      GoTo CleanUp
    End If
    Try
      'RETRIEVE DEVELOPMENT TYPE TAB
      If Not TypeOf m_xlPaintWB1.Sheets("Dev Type Attributes") Is Microsoft.Office.Interop.Excel.Worksheet Then
        MessageBox.Show("The 'Dev Type Attributes' tab could not be found in the selected ENVISION Excel file. Please select another file.", "SCENARIO NOT FOUND", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
        GoTo CleanUp
      Else
        shtDevType = DirectCast(m_xlPaintWB1.Sheets("Dev Type Attributes"), Microsoft.Office.Interop.Excel.Worksheet)
      End If

      'FIND THE STARTING POINT
      For intRow = 1 To 10
        strFldCellValue = CStr(CType(shtDevType.Cells(intRow, 1), Microsoft.Office.Interop.Excel.Range).Value)
        If Not strFldCellValue Is Nothing Then
          If UCase(strFldCellValue) = "DEV_TYPE" Then
            intFieldRow = intRow
            intStartRow = intRow + 1
            Exit For
          Else
            If strFldCellValue.Contains(",") Then
              strFieldName = CStr(strFldCellValue.Split(","c)(0))
              If UCase(strFieldName) = "DEV_TYPE" Then
                intFieldRow = intRow
                intStartRow = intRow + 1
                intSearchMax = intRow + 1
                Exit For
              End If
            End If
          End If
        End If
      Next

      Dim blnStop As Boolean = True

      'CYCLE THROUGHT THE LIST OF DEVELOPMENT TYPES UNTIL THE A ROW IS FOUND WITH NO DEV TYPE NAME, WHICH WOULD INDICATE THE LAST ROW HAS BEEN FOUND
      strDevTypeName = ""
      Do Until blnStop = False
        strDevTypeName = CStr(CType(shtDevType.Cells(intSearchMax, 1), Microsoft.Office.Interop.Excel.Range).Value)
        'EXIT THE LOOP IF DEV TYPE IS BLANK ELSE INCREASE THE MAX NUMBER OF DEV TYPES BY 1
        If strDevTypeName Is Nothing OrElse strDevTypeName = "" OrElse strDevTypeName = "0" Then
          Exit Do
        Else
          If strDevTypeName = "" Or strDevTypeName.Length <= 0 Then
            Exit Do
          Else
            intDevCount = intDevCount + 1
            intSearchMax = intSearchMax + 1
          End If
        End If

      Loop
      'IF THE MAX NUMBER OF DEV TYPES IS GREATER THAN 1 USE
      If intDevCount <= 0 Then
        MessageBox.Show("No Development Types could be found in the selected Envision excel file.  Please select another file.", "Dev Types No Found", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
        GoTo CleanUp
      Else
        m_intDevTypeMax = intDevCount
      End If


      '---------------------------------------------------------------------------------------------------------------------
      'ADD RECORDS FOR EACH DEVE TYPE NEEDED
      '---------------------------------------------------------------------------------------------------------------------
      intRecCount = m_tblDevelopmentTypes.RowCount(Nothing)
      If intRecCount < m_intDevTypeMax Then
        'ADD NEW ROWS
        intAddDeleteRows = m_intDevTypeMax - intRecCount
        Do Until intAddDeleteRows = 0
          m_tblDevelopmentTypes.CreateRow()
          m_tblSc1.CreateRow()
          m_tblSc2.CreateRow()
          m_tblSc3.CreateRow()
          m_tblSc4.CreateRow()
          m_tblSc5.CreateRow()
          intAddDeleteRows = intAddDeleteRows - 1
        Loop
      End If
      'If intRecCount > m_intDevTypeMax Then
      '    'DELETE NEW ROWS
      '    intAddDeleteRows = intRecCount - m_intDevTypeMax
      '    Do Until intAddDeleteRows = 0
      '        m_tblDevelopmentTypes.GetRow(m_tblDevelopmentTypes.RowCount(Nothing) - 1).Delete()
      '        m_tblSc1.GetRow(m_tblSc1.RowCount(Nothing) - 1).Delete()
      '        m_tblSc2.GetRow(m_tblSc2.RowCount(Nothing) - 1).Delete()
      '        m_tblSc3.GetRow(m_tblSc3.RowCount(Nothing) - 1).Delete()
      '        m_tblSc4.GetRow(m_tblSc4.RowCount(Nothing) - 1).Delete()
      '        m_tblSc5.GetRow(m_tblSc5.RowCount(Nothing) - 1).Delete()
      '        intAddDeleteRows = intAddDeleteRows - 1
      '    Loop
      'End If

      'CYCLE THROUGH EACH DEVELOPMENT TYPE
      intRowNmbr = 1
      Dim startMaxColumns As Int32 = 400
      Dim maxColumns As Int32 = startMaxColumns
      Dim newMaxColumns As Int32 = maxColumns
      intProgress = 0
      For intRow = intStartRow To (intStartRow + (m_intDevTypeMax - 1))
        Try
          rowDevType = m_tblDevelopmentTypes.GetRow(intRowNmbr)
          rowSc1 = m_tblSc1.GetRow(intRowNmbr)
          rowSc2 = m_tblSc2.GetRow(intRowNmbr)
          rowSc3 = m_tblSc3.GetRow(intRowNmbr)
          rowSc4 = m_tblSc4.GetRow(intRowNmbr)
          rowSc5 = m_tblSc5.GetRow(intRowNmbr)
        Catch ex As Exception
          MessageBox.Show("Record " & intRowNmbr.ToString & " could not be found in either the Dev Types or Scenario tracking lookup tables.  Please review the ObjectId field for each table.", "Record Not Found", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
          GoTo CleanUp
        End Try


        'CYCLE THROUGH EACH ENVISION EXCEL FILE COLUMN
        'm_dockEnvisionWinForm.prgBarEnvision.Value = 0
        If intRow = intStartRow + 1 Then newMaxColumns = maxColumns
        For intCol = 1 To newMaxColumns
          strCellValue = CStr(CType(shtDevType.Cells(intRow, intCol), Microsoft.Office.Interop.Excel.Range).Value)
          strFldCellValue = CStr(CType(shtDevType.Cells(intFieldRow, intCol), Microsoft.Office.Interop.Excel.Range).Value)

          'ALWAYS DEFAULT TO A VALUE OF 0 IF CELL VALUE IS NULL/EMPTY
          If strCellValue Is Nothing Then
            strCellValue = "0"
          End If
          If strFldCellValue Is Nothing Then
            Continue For
          ElseIf intRow = intStartRow Then
            maxColumns = intCol
          End If

          'RETRIEVE FIELD NAME
          If strFldCellValue.Length <= 0 Then
            Continue For
          Else
            If intCol = startMaxColumns Then
              Throw New Exception("There are too many columns in the 'Dev Type Attributes' worksheet.  A maximum of " & intCol.ToString & " columns are allowed.")
            End If
            If Not strFldCellValue.Contains(",") Then
              Continue For
            Else
              Try
                If strFldCellValue.Contains(",") Then
                  strFieldName = CStr(strFldCellValue.Split(","c)(0))
                Else
                  strFieldName = strFldCellValue
                End If
              Catch ex As Exception
              End Try
            End If
          End If

          If UCase(strFieldName) = "DEV_TYPE" Then
            strFieldName = "DEVELOPMENT_TYPE"
          End If

          'RETRIEVE FIELD AND SET VALUE ACCORDINGLY
          intField = m_tblDevelopmentTypes.FindField(strFieldName)
          If intField >= 0 Then
            pField = m_tblDevelopmentTypes.Fields.Field(intField)
            If pField.Type = esriFieldType.esriFieldTypeDouble Then
              Try
                rowDevType.Value(intField) = CDbl(strCellValue)
              Catch ex As Exception
                rowDevType.Value(intField) = 0
              End Try
            ElseIf pField.Type = esriFieldType.esriFieldTypeInteger Then
              Try
                rowDevType.Value(intField) = CInt(strCellValue)
              Catch ex As Exception
                rowDevType.Value(intField) = 0
              End Try
            ElseIf pField.Type = esriFieldType.esriFieldTypeString Then
              Try
                rowDevType.Value(intField) = CStr(strCellValue)
              Catch ex As Exception
                rowDevType.Value(intField) = 0
              End Try
            End If

            If UCase(strFieldName) = "DEVELOPMENT_TYPE" Or UCase(strFieldName) = "DEV_TYPE" Or UCase(strFieldName) = "DEVELOPMENT TYPE" Or UCase(strFieldName) = "DEV TYPE" Then
              rowDevType.Value(intField) = strCellValue
              rowSc1.Value(m_tblSc1.FindField("DEVELOPMENT_TYPE")) = strCellValue
              rowSc2.Value(m_tblSc2.FindField("DEVELOPMENT_TYPE")) = strCellValue
              rowSc3.Value(m_tblSc3.FindField("DEVELOPMENT_TYPE")) = strCellValue
              rowSc4.Value(m_tblSc4.FindField("DEVELOPMENT_TYPE")) = strCellValue
              rowSc5.Value(m_tblSc5.FindField("DEVELOPMENT_TYPE")) = strCellValue
            Else
            End If
          End If
        Next
        rowDevType.Store()
        rowSc1.Store()
        rowSc2.Store()
        rowSc3.Store()
        rowSc4.Store()
        rowSc5.Store()
        intRowNmbr = intRowNmbr + 1
        'm_dockEnvisionWinForm.prgBarEnvision.Value = 100
        intProgress = intProgress + 1
        m_dockEnvisionWinForm.prgBarEnvision.Value = CType((intProgress / m_intDevTypeMax) * 100, Int32)
      Next
      m_dockEnvisionWinForm.prgBarEnvision.Value = 0

      GoTo CleanUp
    Catch ex As Exception
      MessageBox.Show(ex.Message, "LoadDevTypesFromExcel SUB ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
      GoTo CleanUp
    End Try

CleanUp:
    shtDevType = Nothing
    strFieldName = Nothing
    intField = Nothing
    strCellValue = Nothing
    intCol = Nothing
    intRow = Nothing
    rowDevType = Nothing
    rowSc1 = Nothing
    rowSc2 = Nothing
    rowSc3 = Nothing
    rowSc4 = Nothing
    rowSc5 = Nothing
    intRowNmbr = Nothing
    intProgress = Nothing
    GC.Collect()
    GC.WaitForPendingFinalizers()
  End Sub

  Public Sub LoadTravelInputsFromExcel()
    'POPULATE THE ENVISION TRAVEL INPUTS TABLE WITH VALUES FROM THE SELECTED ENVISION EXCEL FILE
    Dim shtDevType As Microsoft.Office.Interop.Excel.Worksheet = Nothing
    Dim strFieldName As String = ""
    Dim intField As Integer
    Dim intCellValue1 As Integer = 0
    Dim intCellValue2 As Integer = 0
    Dim intCellValue3 As Integer = 0
    Dim strCellValue1 As String = ""

    Dim intCol As Integer
    Dim intRow As Integer
    Dim intRec As Integer
    Dim rowTemp As IRow
    Dim intRowNmbr As Integer
    Dim intProgress As Integer
    Dim intStartRow As Integer = 1

    If m_xlPaintWB1 Is Nothing Or m_xlPaintApp Is Nothing Or m_tblTravel Is Nothing Then
      GoTo CleanUp
    End If
    Try
      'RETRIEVE DEVELOPMENT TYPE TAB
      If Not TypeOf m_xlPaintWB1.Sheets("Travel Times") Is Microsoft.Office.Interop.Excel.Worksheet Then
        'MessageBox.Show("The 'Travel Times' tab could not be found in the selected ENVISION Excel file. Please select another file.", "Travel Times tab NOT FOUND", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
        GoTo CleanUp
      Else
        shtDevType = DirectCast(m_xlPaintWB1.Sheets("Travel Times"), Microsoft.Office.Interop.Excel.Worksheet)
      End If

      'FIND THE STARTING POINT
      For intRow = 1 To 10
        strCellValue1 = CStr(CType(shtDevType.Cells(intRow, 1), Microsoft.Office.Interop.Excel.Range).Value)
        If Not strCellValue1 Is Nothing Then
          If strCellValue1 = "Distance (MI)" Then
            intStartRow = intRow + 1
            Exit For
          End If
        End If
      Next

      'CYCLE THROUGH EACH TRAVEL INPUT
      intProgress = 0
      intRec = 0
      For intRow = intStartRow To (intStartRow + 9)
        intRec = intRec + 1
        rowTemp = m_tblTravel.GetRow(intRec)
        m_dockEnvisionWinForm.prgBarEnvision.Value = 0
        'RETRIEVE TRAVEL INPUTS
        intCellValue1 = 0
        intCellValue2 = 0
        intCellValue3 = 0
        strCellValue1 = ""
        Try
          intCellValue1 = CInt(CType(shtDevType.Cells(intRow, 1), Microsoft.Office.Interop.Excel.Range).Value)
          intCellValue2 = CInt(CType(shtDevType.Cells(intRow, 2), Microsoft.Office.Interop.Excel.Range).Value)
          strCellValue1 = CStr(CType(shtDevType.Cells(intRow, 3), Microsoft.Office.Interop.Excel.Range).Value)
          intCellValue3 = CInt(CType(shtDevType.Cells(intRow, 4), Microsoft.Office.Interop.Excel.Range).Value)
        Catch ex As Exception

        End Try

        'WRITE VALUES TO TRAVEL INPUTS LOOKUP TABLE
        Try
          rowTemp.Value(1) = intCellValue1
          rowTemp.Value(2) = intCellValue2
          rowTemp.Value(3) = strCellValue1
          rowTemp.Value(4) = intCellValue3
          rowTemp.Store()
        Catch ex As Exception

        End Try
        rowTemp.Store()
        intProgress = intProgress + 1
        m_dockEnvisionWinForm.prgBarEnvision.Value = CType((intProgress / 10) * 100, Int32)
      Next
      m_dockEnvisionWinForm.prgBarEnvision.Value = 0

      GoTo CleanUp
    Catch ex As Exception
      'MessageBox.Show(ex.Message, "LoadTravelInputsFromExcel SUB ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
      GoTo CleanUp
    End Try

CleanUp:
    shtDevType = Nothing
    strFieldName = Nothing
    intField = Nothing
    intCellValue1 = Nothing
    intCellValue2 = Nothing
    intCellValue3 = Nothing
    strCellValue1 = Nothing
    intCol = Nothing
    intRow = Nothing
    rowTemp = Nothing
    intRowNmbr = Nothing
    intProgress = Nothing
    GC.Collect()
    GC.WaitForPendingFinalizers()
  End Sub

  Public Sub LoadEXLUInputsFromExcel()
    'POPULATE THE ENVISION EXISTING LAND USE REFENCE INPUTS TABLE WITH VALUES FROM THE SELECTED ENVISION EXCEL FILE
    Dim shtDevType As Microsoft.Office.Interop.Excel.Worksheet = Nothing
    Dim shtExistingLandUse As Microsoft.Office.Interop.Excel.Worksheet = Nothing
    Dim strFieldName As String = ""
    Dim intField As Integer
    Dim strCellValue1 As String = ""
    Dim strCellValue2 As String = ""
    Dim strCellValue3 As String = ""
    Dim intEXLUcolumn As Integer = -1
    Dim intLUMixcolumn As Integer = -1
    Dim intCol As Integer
    Dim intRow As Integer
    Dim intRec As Integer
    Dim rowTemp As IRow
    Dim intRowNmbr As Integer
    Dim intProgress As Integer
    Dim intStartRow As Integer = 1
    Dim intEndRow As Int32
    Dim intCount As Integer = 0
    Dim pTable As ITable

    If m_xlPaintWB1 Is Nothing Or m_xlPaintApp Is Nothing Or m_tblTravel Is Nothing Then
      GoTo CleanUp
    End If
    Try
      'RETRIEVE DEVELOPMENT TYPE TAB
      If Not TypeOf m_xlPaintWB1.Sheets("Existing Developed Area") Is Microsoft.Office.Interop.Excel.Worksheet Then
        'MessageBox.Show("The 'Existing Developed Area' tab could not be found in the selected ENVISION Excel file. Please select another file.", "Existing Developed Area tab NOT FOUND", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
        GoTo CleanUp
      Else
        shtDevType = DirectCast(m_xlPaintWB1.Sheets("Existing Developed Area"), Microsoft.Office.Interop.Excel.Worksheet)

        If shtDevType Is Nothing Then
          GoTo CleanUp
        End If
      End If

      'FIND THE STARTING POINT
      For intRow = 1 To 10
        strCellValue1 = CStr(CType(shtDevType.Cells(intRow, 1), Microsoft.Office.Interop.Excel.Range).Value)
        If Not strCellValue1 Is Nothing Then
          If strCellValue1 = "EXISTING" Then
            intStartRow = intRow
            Exit For
          End If
        End If
      Next

      'FIND THE ENDING POINT
      For intRow = intStartRow To 200
        strCellValue1 = CStr(CType(shtDevType.Cells(intRow, 1), Microsoft.Office.Interop.Excel.Range).Value)
        If strCellValue1 Is Nothing OrElse strCellValue1 = String.Empty Then
          intEndRow = intRow - 1
          Exit For
        End If
      Next

      'FIND THE INPUT COLUMNS
      For intCol = 1 To 30
        strCellValue1 = CStr(CType(shtDevType.Cells((intStartRow), intCol), Microsoft.Office.Interop.Excel.Range).Value)
        If Not strCellValue1 Is Nothing Then
          If strCellValue1 = "EX_LU" Then
            intEXLUcolumn = intCol
          End If
          If strCellValue1 = "EX_LU_TYPES" Then
            intLUMixcolumn = intCol
          End If
        End If
      Next

      ' Create rows if they don't exist yet
      If m_tblEXLURef.RowCount(Nothing) = 0 Then
        For intRec = 1 To intEndRow - intStartRow
          rowTemp = m_tblEXLURef.CreateRow()
          rowTemp.Store()
        Next
      End If

      'CLEAR PREVIOUS VALUES
      For intRec = 1 To m_intDevTypeMax
        Try
          rowTemp = m_tblEXLURef.GetRow(intRec)

          strCellValue1 = CStr(CType(shtDevType.Cells((intStartRow + intRec), 1), Microsoft.Office.Interop.Excel.Range).Value)
          rowTemp.Value(1) = strCellValue1
          If intEXLUcolumn > 1 Then
            strCellValue2 = CStr(CType(shtDevType.Cells((intStartRow + intRec), intEXLUcolumn), Microsoft.Office.Interop.Excel.Range).Value)
            rowTemp.Value(2) = strCellValue2
          Else
            rowTemp.Value(2) = ""
          End If
          If intLUMixcolumn > 1 Then
            strCellValue3 = CStr(CType(shtDevType.Cells((intStartRow + intRec), intLUMixcolumn), Microsoft.Office.Interop.Excel.Range).Value)
            rowTemp.Value(3) = strCellValue3
          Else
            rowTemp.Value(3) = ""
          End If
          rowTemp.Store()
        Catch ex As Exception
        End Try
      Next

      'RETRIEVE THE EXISITNG LAND USE TAB TO FIELD REQUIRED PARCELS, CALC NUMERIC FIELDS TO 0 WHEN ADDED
      Try
        If m_pEditFeatureLyr Is Nothing Then
          GoTo CleanUp
        Else
          pTable = CType(m_pEditFeatureLyr.FeatureClass, ITable)
        End If
        m_appEnvision.StatusBar.Message(0) = "Retrieving 'Existing Land Use' tab"
        shtExistingLandUse = DirectCast(m_xlPaintWB1.Sheets("Existing Land Use"), Microsoft.Office.Interop.Excel.Worksheet)

        'FIND THE STARTING POINT
        For intRow = 1 To 10
          strCellValue1 = CStr(CType(shtExistingLandUse.Cells(intRow, 1), Microsoft.Office.Interop.Excel.Range).Value)
          If Not strCellValue1 Is Nothing Then
            If UCase(strCellValue1) = "EXISTING" Then
              intStartRow = intRow
              Exit For
            End If
          End If
        Next

        'RETRIEVE EXISITNG LAND USE ABBREVIATION COLUMN
        m_appEnvision.StatusBar.Message(0) = "'Existing Land Use' tab - building existing conditions field list"
        intCount = 0
        For intCol = 2 To m_intDevTypeMax
          strCellValue1 = CStr(CType(shtExistingLandUse.Cells(intStartRow, intCol), Microsoft.Office.Interop.Excel.Range).Value)
          If Not strCellValue1 Is Nothing Then
            If strCellValue1.Contains("EX_") Then
              'CHECK TO SEE IF EXISITNG LAND USE FIELD EXITST, ADD IF MISSING AND CALC TO 0
              If pTable.FindField(strCellValue1) = -1 Then
                If AddEnvisionField(pTable, strCellValue1, "Double", 16, 6) Then
                  CalcFldValues(pTable, strCellValue1, "0", "Numeric")
                End If
              End If
            End If
          End If
        Next


      Catch ex As Exception

      End Try

      GoTo CleanUp
    Catch ex As Exception
      'MessageBox.Show(ex.Message, "LoadTravelInputsFromExcel SUB ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
      GoTo CleanUp
    End Try

CleanUp:
    shtDevType = Nothing
    shtExistingLandUse = Nothing
    strFieldName = Nothing
    intField = Nothing
    strCellValue1 = Nothing
    strCellValue2 = Nothing
    strCellValue3 = Nothing
    strCellValue1 = Nothing
    intCol = Nothing
    intRow = Nothing
    rowTemp = Nothing
    intRowNmbr = Nothing
    intProgress = Nothing
    intCount = Nothing
    pTable = Nothing
    GC.Collect()
  End Sub

  Public Sub RetrieveDevTypeData()
    'RETRIEVE THE DEVELOPMENT TYPE NAME, GRID CODE, 
    'RGB COLOR VALUES, AND ACRE SUMS FROM SELECTED ENVISION EXCEL FILE
    m_strProcessing = m_strProcessing & "Starting sub RetrieveDevTypeData: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
    Dim intRow As Integer
    Dim strCellValue As String = ""
    Dim intRed As Integer = 0
    Dim intGreen As Integer = 0
    Dim intBlue As Integer = 0
    Dim bmpTemp As Bitmap
    Dim rowTemp As IRow

    Try
      'EXIT IF EMPTY TABLE
      If m_tblDevelopmentTypes.RowCount(Nothing) <= 0 Then
        GoTo CleanUp
      End If
      'CLEAR PREVIOUSLY LOADED DEVELOPMENT TYPES
      m_blnLoadingDevTypes = True
      m_dockEnvisionWinForm.dgvDevelopmentTypes.Rows.Clear()
      'INSERT THE ERASE OPTION
      m_dockEnvisionWinForm.dgvDevelopmentTypes.Rows.Add()
      bmpTemp = CreateColorBMP(255, 255, 255)
      m_dockEnvisionWinForm.dgvDevelopmentTypes.Rows(m_dockEnvisionWinForm.dgvDevelopmentTypes.RowCount - 1).Cells(0).Value = bmpTemp
      m_dockEnvisionWinForm.dgvDevelopmentTypes.Rows(m_dockEnvisionWinForm.dgvDevelopmentTypes.RowCount - 1).Cells(1).Value = "ERASE"
      m_dockEnvisionWinForm.dgvDevelopmentTypes.Rows(m_dockEnvisionWinForm.dgvDevelopmentTypes.RowCount - 1).Cells(2).Value = 255
      m_dockEnvisionWinForm.dgvDevelopmentTypes.Rows(m_dockEnvisionWinForm.dgvDevelopmentTypes.RowCount - 1).Cells(3).Value = 255
      m_dockEnvisionWinForm.dgvDevelopmentTypes.Rows(m_dockEnvisionWinForm.dgvDevelopmentTypes.RowCount - 1).Cells(4).Value = 255

      m_strDWriteValue = ""
      m_dockEnvisionWinForm.prgBarEnvision.Value = 0
      If m_tblDevelopmentTypes Is Nothing Then
        GoTo CleanUp
      End If
      If m_tblDevelopmentTypes.RowCount(Nothing) < m_intDevTypeMax Then
        GoTo CleanUp
      End If
      For intRow = 1 To m_intDevTypeMax
        m_dockEnvisionWinForm.prgBarEnvision.Value = CType(((intRow / m_intDevTypeMax) * 100), Int32)
        rowTemp = m_tblDevelopmentTypes.GetRow(intRow)
        Try
          strCellValue = CStr(rowTemp.Value(m_tblDevelopmentTypes.FindField("DEVELOPMENT_TYPE")))
        Catch ex As Exception
          Exit For
        End Try
        'MAKE SURE THERE IS A DEVELOPMENT TYPE
        If strCellValue.Length > 0 And (Not strCellValue = "" And Not strCellValue = "0") Then
          m_appEnvision.StatusBar.Message(0) = "LOADING DEVELOPMENT TYPE: " & strCellValue
          m_dockEnvisionWinForm.dgvDevelopmentTypes.Rows.Add()
          'COLUMN A = DEVELOPMENT TYPE NAME
          m_dockEnvisionWinForm.dgvDevelopmentTypes.Rows(m_dockEnvisionWinForm.dgvDevelopmentTypes.RowCount - 1).Cells(1).Value = strCellValue
          'SYMBOL RED COLOR 
          Try
            If IsNumeric(CStr(rowTemp.Value(m_tblDevelopmentTypes.FindField("RED")))) Then
              intRed = CInt(rowTemp.Value(m_tblDevelopmentTypes.FindField("RED")))
            Else
              intRed = 0
            End If
          Catch ex As Exception
            intRed = 0
          End Try
          m_dockEnvisionWinForm.dgvDevelopmentTypes.Rows(m_dockEnvisionWinForm.dgvDevelopmentTypes.RowCount - 1).Cells(2).Value = intRed
          'SYMBOL GREEN COLOR
          Try
            If IsNumeric(CStr(rowTemp.Value(m_tblDevelopmentTypes.FindField("GREEN")))) Then
              intGreen = CInt(rowTemp.Value(m_tblDevelopmentTypes.FindField("GREEN")))
            Else
              intGreen = 0
            End If
          Catch ex As Exception
            intGreen = 0
          End Try
          m_dockEnvisionWinForm.dgvDevelopmentTypes.Rows(m_dockEnvisionWinForm.dgvDevelopmentTypes.RowCount - 1).Cells(3).Value = intGreen
          'SYMBOL BLUE COLOR
          Try
            If IsNumeric(CStr(rowTemp.Value(m_tblDevelopmentTypes.FindField("BLUE")))) Then
              intBlue = CInt(rowTemp.Value(m_tblDevelopmentTypes.FindField("BLUE")))
            Else
              intBlue = 0
            End If
          Catch ex As Exception
            intBlue = 0
          End Try
          m_dockEnvisionWinForm.dgvDevelopmentTypes.Rows(m_dockEnvisionWinForm.dgvDevelopmentTypes.RowCount - 1).Cells(4).Value = intBlue
          'COLOR BITMAP SYMBOL
          bmpTemp = CreateColorBMP(intRed, intGreen, intBlue)
          m_dockEnvisionWinForm.dgvDevelopmentTypes.Rows(m_dockEnvisionWinForm.dgvDevelopmentTypes.RowCount - 1).Cells(0).Value = bmpTemp

          m_dockEnvisionWinForm.dgvDevelopmentTypes.Rows(m_dockEnvisionWinForm.dgvDevelopmentTypes.RowCount - 1).Cells(1).Selected = True
          'Else
          '    Exit For
        End If
      Next

      If m_dockEnvisionWinForm.dgvDevelopmentTypes.RowCount = 1 Then
        m_dockEnvisionWinForm.dgvDevelopmentTypes.Rows.Clear()
        GoTo CleanUp
      End If
      m_dockEnvisionWinForm.prgBarEnvision.Value = 100
      m_dockEnvisionWinForm.dgvDevelopmentTypes.Rows(0).Cells(0).Selected = True
      m_strDWriteValue = ""
      GoTo CleanUp

    Catch ex As Exception
      MessageBox.Show(ex.Message & vbNewLine & "Please review the development type data pull from the selected Envision excel file as attribute field may be missing.", "DEVELOPMENT TYPE LOAD ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
      GoTo CleanUp
    End Try

CleanUp:
    m_strProcessing = m_strProcessing & "Ending sub RetrieveDevTypeData: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine & vbNewLine
    m_blnLoadingDevTypes = False
    intRow = Nothing
    strCellValue = Nothing
    intRed = Nothing
    intGreen = Nothing
    intBlue = Nothing
    bmpTemp = Nothing
    GC.Collect()
    GC.WaitForPendingFinalizers()
    m_blnLoadingDevTypes = False
    m_appEnvision.StatusBar.Message(0) = ""
    m_dockEnvisionWinForm.prgBarEnvision.Value = 0
  End Sub

  Public Function CreateColorBMP(ByVal intRed As Integer, ByVal intGreen As Integer, ByVal intBlue As Integer) As Bitmap
    'Create a bitmap
    Dim bmpNew As Bitmap = New Drawing.Bitmap(48, 48)
    CreateColorBMP = bmpNew

    'Create a Graphics object and Device Contect
    Dim memoryGraphic As Drawing.Graphics = Drawing.Graphics.FromImage(bmpNew)
    Dim bmpDC As System.IntPtr = memoryGraphic.GetHdc
    Dim geom As ESRI.ArcGIS.Geometry.IGeometry
    Dim env As ESRI.ArcGIS.Geometry.Envelope
    Dim symbNew As ISimpleFillSymbol
    Dim pFillSym As IFillSymbol
    Dim newColor As RgbColor
    Dim pSymbol As ESRI.ArcGIS.Display.ISymbol

    Try
      env = New ESRI.ArcGIS.Geometry.Envelope
      env.PutCoords(0, 0, 48, 48)
      geom = CType(env, ESRI.ArcGIS.Geometry.IGeometry)

      'Draw the symbol to the device context
      newColor = New RgbColor
      newColor.Red = intRed
      newColor.Blue = intBlue
      newColor.Green = intGreen

      symbNew = New SimpleFillSymbol
      pFillSym = (New SimpleFillSymbol)
      With pFillSym
        .Color = newColor
      End With
      CreateColorBMP = bmpNew

      pSymbol = CType(pFillSym, ESRI.ArcGIS.Display.ISymbol)
      pSymbol.SetupDC(bmpDC.ToInt32, Nothing)
      pSymbol.ROP2 = esriRasterOpCode.esriROPCopyPen
      If Not geom Is Nothing Then
        pSymbol.Draw(geom)
      End If
      pSymbol.ResetDC()
      GoTo CleanUp
    Catch ex As Exception
      MessageBox.Show(ex.Message, "CreateColorBMP")
      GoTo CleanUp
    End Try

CleanUp:
    bmpNew = Nothing
    memoryGraphic = Nothing
    bmpDC = Nothing
    geom = Nothing
    env = Nothing
    symbNew = Nothing
    pFillSym = Nothing
    newColor = Nothing
    pSymbol = Nothing
    GC.Collect()
    GC.WaitForPendingFinalizers()
  End Function

  Public Function ReturnAcres(ByVal dblValue As Double, ByVal UnitofMeasure As String) As Double
    'CONVERT PROVIDED VALUE INTO ACRES
    ReturnAcres = dblValue
    UnitofMeasure = UCase(UnitofMeasure)
    If (UnitofMeasure = "ESRIFEET") Or UnitofMeasure = "3" Or UnitofMeasure = "FOOT" Or UnitofMeasure = "FOOT_US" Then
      ReturnAcres = (dblValue * 0.000022956)
    ElseIf (UnitofMeasure = "ESRIMETERS") Or UnitofMeasure = "9" Or UnitofMeasure = "METER" Then
      ReturnAcres = (dblValue * 0.0002471054)
    ElseIf (UnitofMeasure = "ESRIYARDS") Or UnitofMeasure = "4" Or UnitofMeasure = "YARD" Then
      ReturnAcres = (dblValue * 0.000206611)
    ElseIf (UnitofMeasure = "ESRIMILES") Or UnitofMeasure = "5" Or UnitofMeasure = "MILE_US" Then
      ReturnAcres = (dblValue * 640.0)
    ElseIf (UnitofMeasure = "ESRIKILOMETERS") Or UnitofMeasure = "10" Or UnitofMeasure = "KILOMETER" Then
      ReturnAcres = (dblValue * 247.1054)
    ElseIf (UnitofMeasure = "ESRIMILLIMETERS") Or UnitofMeasure = "7" Or UnitofMeasure = "MILLIMETER" Then
      ReturnAcres = (dblValue * 0.0000000002471054)
    ElseIf (UnitofMeasure = "ESRICENTIMETERS") Or UnitofMeasure = "8" Or UnitofMeasure = "CENTIMETER" Then
      ReturnAcres = (dblValue * 0.00000002471054)
    ElseIf (UnitofMeasure = "ESRIINCHES") Or UnitofMeasure = "1" Or UnitofMeasure = "INCH" Or UnitofMeasure = "INCH_US" Then
      ReturnAcres = (dblValue * 0.00000015942250790735638)
    End If
  End Function

  Public Function ReturnMiles(ByVal dblValue As Double, ByVal UnitofMeasure As String) As Double
    'CONVERT PROVIDED VALUE INTO MILES
    ReturnMiles = dblValue
    UnitofMeasure = UCase(UnitofMeasure)
    If (UnitofMeasure = "ESRIFEET") Or UnitofMeasure = "3" Or UnitofMeasure = "FOOT" Or UnitofMeasure = "FOOT_US" Then
      ReturnMiles = (dblValue * 0.000189393939)
    ElseIf (UnitofMeasure = "ESRIMETERS") Or UnitofMeasure = "9" Or UnitofMeasure = "METER" Then
      ReturnMiles = (dblValue * 0.000621371192)
    ElseIf (UnitofMeasure = "ESRIYARDS") Or UnitofMeasure = "4" Or UnitofMeasure = "YARD" Then
      ReturnMiles = (dblValue * 0.000568181818)
    ElseIf (UnitofMeasure = "ESRIMILES") Or UnitofMeasure = "5" Or UnitofMeasure = "MILE_US" Then
      ReturnMiles = (dblValue * 1.0)
    ElseIf (UnitofMeasure = "ESRIKILOMETERS") Or UnitofMeasure = "10" Or UnitofMeasure = "KILOMETER" Then
      ReturnMiles = (dblValue * 0.621371192)
    ElseIf (UnitofMeasure = "ESRIMILLIMETERS") Or UnitofMeasure = "7" Or UnitofMeasure = "MILLIMETER" Then
      ReturnMiles = (dblValue * 0.000000621371192)
    ElseIf (UnitofMeasure = "ESRICENTIMETERS") Or UnitofMeasure = "8" Or UnitofMeasure = "CENTIMETER" Then
      ReturnMiles = (dblValue * 0.00000621371192)
    ElseIf (UnitofMeasure = "ESRIINCHES") Or UnitofMeasure = "1" Or UnitofMeasure = "INCH" Or UnitofMeasure = "INCH_US" Then
      ReturnMiles = (dblValue * 0.0000157828283)
    End If
  End Function

  Public Function ReturnFeet(ByVal dblValue As Double, ByVal UnitofMeasure As String) As Double
    'CONVERT PROVIDED VALUE INTO FEET
    ReturnFeet = dblValue
    UnitofMeasure = UCase(UnitofMeasure)
    If (UnitofMeasure = "ESRIFEET") Or UnitofMeasure = "3" Or UnitofMeasure = "FOOT" Or UnitofMeasure = "FOOT_US" Then
      ReturnFeet = (dblValue * 1.0)
    ElseIf (UnitofMeasure = "ESRIMETERS") Or UnitofMeasure = "9" Or UnitofMeasure = "METER" Then
      ReturnFeet = (dblValue * 3.2808399)
    ElseIf (UnitofMeasure = "ESRIYARDS") Or UnitofMeasure = "4" Or UnitofMeasure = "YARD" Then
      ReturnFeet = (dblValue * 3.0)
    ElseIf (UnitofMeasure = "ESRIMILES") Or UnitofMeasure = "5" Or UnitofMeasure = "MILE_US" Then
      ReturnFeet = (dblValue * 5280.0)
    ElseIf (UnitofMeasure = "ESRIKILOMETERS") Or UnitofMeasure = "10" Or UnitofMeasure = "KILOMETER" Then
      ReturnFeet = (dblValue * 3280.0)
    ElseIf (UnitofMeasure = "ESRIMILLIMETERS") Or UnitofMeasure = "7" Or UnitofMeasure = "MILLIMETER" Then
      ReturnFeet = (dblValue * 0.0032808399)
    ElseIf (UnitofMeasure = "ESRICENTIMETERS") Or UnitofMeasure = "8" Or UnitofMeasure = "CENTIMETER" Then
      ReturnFeet = (dblValue * 0.032808399)
    ElseIf (UnitofMeasure = "ESRIINCHES") Or UnitofMeasure = "1" Or UnitofMeasure = "INCH" Or UnitofMeasure = "INCH_US" Then
      ReturnFeet = (dblValue * 0.0833333333)
    End If
  End Function

  '  Public Sub TrackIndicatorGraphs()
  '    'CREATE AND TRACK SCENARIO GRAPHS WITHIN ARCGIS AND AS GRAPHIC JPEG FILES

  '    'RETRIEVE THE SELECTED FGDB FILE
  '    If m_strFeaturePath.Length <= 0 Then
  '      GoTo cleanup
  '      If Not Directory.Exists(m_strFeaturePath) Then
  '        GoTo cleanup
  '      End If
  '    End If

  '    Dim mxApplication As IMxApplication
  '    Dim pMxDoc As IMxDocument
  '    Dim pSTableColl As IStandaloneTableCollection
  '    Dim pSTable As IStandaloneTable = Nothing
  '    Dim pTable As ITable = Nothing
  '    Dim pWksFactory As IWorkspaceFactory = Nothing
  '    Dim pFeatWks As IFeatureWorkspace
  '    Dim blnTableFound As Boolean = False
  '    Dim intCount As Integer
  '    Dim pDataset As IDataset
  '    Dim pDataGraphs As IDataGraphCollection
  '    Dim pDataGraph As IDataGraphT
  '    Dim pDataGraphBase As IDataGraphBase
  '    Dim imgGraph As Image
  '    Dim blnGraph As Boolean = False
  '    Dim pSP As ISeriesProperties

  '    'CHECK GRAPHS DIRECTORY TO STORE THE IMAGE FILES
  '    If Not Directory.Exists(m_strFeaturePath + "\ENVISION\INDICATORS\GRAPH_IMAGES") Then
  '      Try
  '        Directory.CreateDirectory((m_strFeaturePath + "\ENVISION\INDICATORS\GRAPH_IMAGES"))
  '      Catch ex As Exception
  '        MessageBox.Show(ex.Message, "CREATE DIRECTORY ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
  '        GoTo cleanup
  '      End Try
  '    End If

  '    'RETRIEVE THE CURRENT VIEW DOCUMENT
  '    mxApplication = CType(m_appEnvision, IMxApplication)
  '    pMxDoc = CType(m_appEnvision.Document, IMxDocument)

  '    'LOAD THE SCENARIO SUMMARY TABLE INTO THE CURRENT VIEW DOCUMENT
  '    pSTableColl = DirectCast(pMxDoc.FocusMap, IStandaloneTableCollection)

  '    'REVIEW CURRENT STANDALONE COLLECTION FOR THE SUMMARY TABLE
  '    For intCount = 0 To pSTableColl.StandaloneTableCount - 1
  '      Try
  '        pSTable = pSTableColl.StandaloneTable(intCount)
  '        If pSTable.Name = "SCENARIO_SUMMARY" Then
  '          pTable = pSTable.Table
  '          pDataset = DirectCast(pTable, IDataset)
  '          If pDataset.Workspace.PathName = m_strFeaturePath Then
  '            blnTableFound = True
  '            Exit For
  '          End If
  '        End If
  '      Catch ex As Exception

  '      End Try
  '    Next
  '    If Not blnTableFound Then
  '      GoTo CleanUp
  '    End If

  '    Try
  '      'REVIEW GRAPH LIST TO SEE IF GRAPH ALREADY EXISTS AND UPDATE THE TABLE IF NEEDED
  '      pDataGraphs = DirectCast(pMxDoc, IDataGraphCollection)

  '      '----------------------------------------------------------------------------------
  '      'CHECK FOR THE ACRES INDICATOR GRAPH
  '      Dim arrIList As ArrayList = New ArrayList
  '      Dim strIndicator As String = "ACRES"
  '      Dim strVal As String = ""
  '      arrIList.Add("ACRES")
  '      arrIList.Add("HOUSING MIX")
  '      arrIList.Add("EMPLOYMENT MIX")
  '      arrIList.Add("JOBS HOUSING BALANCE")
  '      arrIList.Add("VO per HU")
  '      arrIList.Add("VMT per HU")
  '      arrIList.Add("MODE SPLIT")
  '      arrIList.Add("GREENHOUSE GASES")
  '      For Each strIndicator In arrIList
  '        blnGraph = False
  '        If pDataGraphs.DataGraphCount > 0 Then
  '          For intCount = 0 To pDataGraphs.DataGraphCount - 1
  '            pDataGraph = DirectCast(pDataGraphs.DataGraph(intCount), IDataGraphT)
  '            strVal = pDataGraph.Name
  '            If strVal = "ENVISION - " & strIndicator & " INDICATOR" Then
  '              Try
  '                pSP = pDataGraph.SeriesProperties(0)
  '                pTable = DirectCast(pSP.SourceData, ITable)
  '                pDataset = DirectCast(pTable, IDataset)
  '                If m_strFeaturePath.Contains("\\") Then
  '                  m_strFeaturePath = m_strFeaturePath.Replace("\\", "\")
  '                End If
  '                If pDataset.Workspace.PathName = m_strFeaturePath Then
  '                  blnGraph = True
  '                  pDataGraph.Reload()
  '                  Try
  '                    pDataGraph.ExportToFileEx(m_strFeaturePath & "\ENVISION\INDICATORS\GRAPH_IMAGES\" & strIndicator & ".jpg", 462, 324)
  '                  Catch ex As Exception
  '                    'MessageBox.Show(ex.ToString, "")
  '                  End Try
  '                  pDataGraph = Nothing
  '                  GC.Collect()
  '                  GC.WaitForPendingFinalizers()
  '                  Exit For
  '                End If
  '              Catch ex As Exception
  '              End Try
  '            End If
  '          Next
  '        End If
  '      Next
  '      GoTo CleanUp
  '    Catch ex As Exception
  '      MessageBox.Show(ex.Message, "TrackIndicatorGraphs Sub Error")
  '      GoTo CleanUp
  '    End Try

  'CleanUp:
  '    mxApplication = Nothing
  '    pMxDoc = Nothing
  '    pSTableColl = Nothing
  '    pSTable = Nothing
  '    pTable = Nothing
  '    pWksFactory = Nothing
  '    pFeatWks = Nothing
  '    blnTableFound = Nothing
  '    intCount = Nothing
  '    pDataset = Nothing
  '    pDataGraphs = Nothing
  '    pDataGraph = Nothing
  '    pDataGraphBase = Nothing
  '    imgGraph = Nothing
  '    blnGraph = Nothing
  '    pSP = Nothing
  '    m_appEnvision.StatusBar.Message(0) = ""
  '    m_dockEnvisionWinForm.prgBarEnvision.Value = 0
  '    GC.Collect()
  '    GC.WaitForPendingFinalizers()

  '  End Sub

  Public Function OpenEnivisionGRF(ByVal strGRFFile As String) As Boolean
    OpenEnivisionGRF = True
    'OPENS THE INDICATOR GRAPH FILE, IF EXISTS, FOR THE SELECTED GRAPH
    Dim mxApplication As IMxApplication
    Dim pMxDoc As IMxDocument
    Dim pDataGraphs As IDataGraphCollection
    Dim pDataGraph As IDataGraphT
    Dim pDataGraphBase As IDataGraphBase
    Dim pGraphWindow As IDataGraphWindow2
    'CHECK FOR EXISTING FILE
    If Not File.Exists(m_strFeaturePath & "\ENVISION\INDICATORS\" & strGRFFile & ".grf") Then
      OpenEnivisionGRF = False
      GoTo CleanUp
    End If

    'RETRIEVE THE CURRENT VIEW DOCUMENT
    mxApplication = CType(m_appEnvision, IMxApplication)
    pMxDoc = CType(m_appEnvision.Document, IMxDocument)

    Try
      pDataGraphBase = New DataGraphT
      pDataGraph = DirectCast(pDataGraphBase, IDataGraphT)
      With pDataGraph
        .LoadFromFile((m_strFeaturePath & "\ENVISION\INDICATORS\" & strGRFFile & ".grf"))
        .UseSelectedSet = True
        .Reload()
      End With

      pDataGraphs = DirectCast(pMxDoc, IDataGraphCollection)
      pDataGraphs.AddDataGraph(pDataGraph)
      pDataGraph.ExportToFile((m_strFeaturePath & "\ENVISION\INDICATORS\GRAPH_IMAGES\" & strGRFFile & ".jpg"))
      pGraphWindow = New DataGraphWindow
      pGraphWindow.DataGraphBase = pDataGraphBase
      pGraphWindow.Application = mxApplication
      'pGraphWindow.Show(True)
      GC.Collect()
      GC.WaitForPendingFinalizers()
    Catch ex As Exception
      MessageBox.Show(ex.Message, "OpenEnivisionGRF Sub Error")
    End Try
CleanUp:
    mxApplication = Nothing
    pMxDoc = Nothing
    pDataGraphs = Nothing
    pDataGraph = Nothing
    pGraphWindow = Nothing
    GC.Collect()
    GC.WaitForPendingFinalizers()
  End Function

  Public Function SelectEnvisionExcelFile() As Boolean
    'WRITE VALUES TO ENVISION EXCEL FILE
    'RETRIEVE OR OPEN THE SELECTED ENVISION EXCEL FILE
    SelectEnvisionExcelFile = True
    Dim FileDialog1 As New OpenFileDialog
    Dim blnEditSession As Boolean = False
    Try
      'END ANY EDIT SESSIONS
      If Not m_dockEnvisionWinForm.btnEditing.Text = "Start Edit" Then
        m_strProcessing = m_strProcessing & "An edit session is currently started and needs to be stopped before proceeding: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine & vbNewLine
        blnEditSession = True
        EditSession()
      End If

      FileDialog1.Title = "Select an ENVISION Excel File"
      FileDialog1.Filter = "Excel Files (*.xls,*.xlsx,*.xlsm)|*.xls;*.xlsx;*.xlsm"
      FileDialog1.CheckPathExists = True
      If FileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
        'CLOSE ANY EXCEL, WHICH MAY BE OPEN
        CloseEnvisionExcel()
        'APPLY TO MODULE VARIABLE
        m_strEnvisionExcelFile = FileDialog1.FileName
        'SET OTHER CONTROL VARIABLES
        m_dockEnvisionWinForm.itmSelectEnvisionFile.ToolTipText = FileDialog1.FileName
        m_dockEnvisionWinForm.itmLoadEnvisionExcelFile.Enabled = True
        m_dockEnvisionWinForm.itmCloseEnvisionExcel.Enabled = True
        ' m_dockEnvisionWinForm.itmAutoUpdateStatus.Enabled = True
        'OPEN THE EXCEL 
        OpenEnvisionExcel()

        'REVIEW THE LOOKUPTABLES
        LookUpTablesEnvisionCheck(m_strFeaturePath)
      Else
        SelectEnvisionExcelFile = False
        GoTo cleanup
      End If
      GoTo CleanUp
    Catch ex As Exception
      SelectEnvisionExcelFile = False
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
    GC.WaitForPendingFinalizers()
  End Function

  Public Sub EditSession()
    'CONTROLS THE EDIT SESSION ON THE ENVISION EDIT LAYER
    Dim pEditor As IEditor
    Dim pID As New UID
    Dim pDataset As IDataset
    m_strProcessing = m_strProcessing & "Starting Function EditSession: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine

    Try

      'RETRIEVE THE EDITOR EXTENSION
      pID.Value = "esriEditor.Editor"
      pEditor = DirectCast(m_appEnvision.FindExtensionByCLSID(pID), IEditor)
      m_strProcessing = m_strProcessing & "Retrieving Editor object: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine

      If m_dockEnvisionWinForm.btnEditing.Text = "Start Edit" Then
        m_strProcessing = m_strProcessing & "Envision edit button indicates no active edit session: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
        m_strProcessing = m_strProcessing & "Starting a New Edit Session"
        'CHECK FOR EDIT LAYER
        If m_pEditFeatureLyr Is Nothing Then
          m_strProcessing = m_strProcessing & "Exiting Function, No Edit Feature Layer Found: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
          GoTo CleanUp
        Else
          m_strProcessing = m_strProcessing & "Edit Layer: " & m_pEditFeatureLyr.Name
          Try
            'RETRIEVE DATASET FROM SELECTED ENVISION EDIT FEATURE CLASS
            m_strProcessing = m_strProcessing & "Retrieving Edit Layer Dataset object" & Date.Now.ToString("hh:mm:ss tt")
            pDataset = CType(m_pEditFeatureLyr.FeatureClass, IDataset)
            'START EDITING THE DATASET WORKSPACE
            m_strProcessing = m_strProcessing & "Starting Edit Session" & Date.Now.ToString("hh:mm:ss tt")
            pEditor.StartEditing(pDataset.Workspace)
            m_strProcessing = m_strProcessing & "Edit session activated" & Date.Now.ToString("hh:mm:ss tt")
            m_dockEnvisionWinForm.btnEditing.Text = "End Edit" & vbNewLine & " - Save"
            m_dockEnvisionWinForm.btnSaveEdits.Visible = False
            m_dockEnvisionWinForm.tlsEndEdit.Visible = False
            m_dockEnvisionWinForm.sep2.Visible = False
          Catch ex As Exception
            MessageBox.Show(ex.Message, "Start Edit Error")
            GoTo CleanUp
          End Try
          m_dockEnvisionWinForm.btnSaveEdits.Visible = True
          m_dockEnvisionWinForm.tlsEndEdit.Visible = True
          m_dockEnvisionWinForm.sep2.Visible = True
          GoTo CleanUp
        End If
      Else
        m_strProcessing = m_strProcessing & "Envision edit button indicates an edit session is active: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
        m_strProcessing = m_strProcessing & "Ending Edit Session with Saving Edits"
        Try
          pEditor.StopEditing(True)
          m_strProcessing = m_strProcessing & "Edit Halted: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
          m_dockEnvisionWinForm.btnEditing.Text = "Start Edit"
          m_dockEnvisionWinForm.btnSaveEdits.Visible = False
          m_dockEnvisionWinForm.tlsEndEdit.Visible = False
          m_dockEnvisionWinForm.sep2.Visible = False
        Catch ex As Exception
          m_strProcessing = m_strProcessing & "Error in halting the active edit session: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
          m_strProcessing = m_strProcessing & ex.Message
          MessageBox.Show(ex.Message, "End Edit Error")
          GoTo CleanUp
        End Try
      End If
    Catch ex As Exception
      m_strProcessing = m_strProcessing & "Error in Function EditSession: " & Date.Now.ToString("hh:mm:ss tt") & vbNewLine
      m_strProcessing = m_strProcessing & ex.Message
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


  Public Sub UpdateSummaryTableSelection()
    'QUERY THE INDICATOR SUMMARY TABLE TO MATCH THE SELECTED MENU ITEMS
    If m_blnScrMenuCheck Then
      GoTo CLeanUp
    End If

    Dim mxApplication As IMxApplication
    Dim pMxDoc As IMxDocument
    Dim pSTableColl As IStandaloneTableCollection
    Dim pSTable As IStandaloneTable = Nothing
    Dim pTable As ITable
    Dim pWksFactory As IWorkspaceFactory = Nothing
    Dim pFeatWks As IFeatureWorkspace
    Dim blnTableFound As Boolean = False
    Dim intCount As Integer
    Dim pDataset As IDataset
    Dim pTableSelection As ITableSelection

    'RETRIEVE THE CURRENT VIEW DOCUMENT
    mxApplication = CType(m_appEnvision, IMxApplication)
    pMxDoc = CType(m_appEnvision.Document, IMxDocument)

    'LOAD THE SCENARIO SUMMARY TABLE INTO THE CURRENT VIEW DOCUMENT
    pSTableColl = DirectCast(pMxDoc.FocusMap, IStandaloneTableCollection)
    If m_strFeaturePath.Length <= 0 Then
      GoTo CleanUp
    End If
    If pSTableColl.StandaloneTableCount = 0 Then
      Try
        If TypeOf m_tblScSummary Is ITable Then
          pSTable = New StandaloneTable
          pSTable.Table = m_tblScSummary
          pSTableColl.AddStandaloneTable(pSTable)
          blnTableFound = True
        End If
      Catch ex As Exception
        MessageBox.Show(ex.Message, "Error")
      End Try
    Else
      'REVIEW CURRENT STANDALONE COLLECTION FOR THE SUMMARY TABLE
      For intCount = 0 To pSTableColl.StandaloneTableCount - 1
        Try
          pSTable = pSTableColl.StandaloneTable(intCount)
          If pSTable.Name = "SCENARIO_SUMMARY" Then
            pTable = pSTable.Table
            pDataset = DirectCast(pTable, IDataset)
            If pDataset.Workspace.PathName = m_strFeaturePath Then
              blnTableFound = True
              Exit For
            End If
          End If
        Catch ex As Exception
          'MessageBox.Show(ex.Message, "THe problem")
        End Try
      Next
    End If

    'LOAD THE SUMMARY TABLE IF MISSING
    If Not blnTableFound Then
      If m_tblScSummary Is Nothing Then
        Try
          pWksFactory = New FileGDBWorkspaceFactory
          pFeatWks = DirectCast(pWksFactory.OpenFromFile(m_strFeaturePath, 0), IFeatureWorkspace)
          pTable = pFeatWks.OpenTable("SCENARIO_SUMMARY")
          pSTable = CType(pTable, IStandaloneTable)
          pSTableColl.AddStandaloneTable(pSTable)
        Catch ex As Exception
          MessageBox.Show(ex.Message, "Update Scenario Summary Table Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
          GoTo CleanUp
        End Try
      Else
        pSTable = New StandaloneTable
        pSTable.Table = m_tblScSummary
        pSTableColl.AddStandaloneTable(pSTable)
      End If
    End If



CleanUp:
    mxApplication = Nothing
    pMxDoc = Nothing
    pSTableColl = Nothing
    pSTable = Nothing
    pTable = Nothing
    pWksFactory = Nothing
    pFeatWks = Nothing
    blnTableFound = Nothing
    intCount = Nothing
    pDataset = Nothing
    pTableSelection = Nothing
    GC.WaitForPendingFinalizers()
    GC.Collect()

  End Sub

  Public Sub RetrieveEvisionFields()
    'RETRIEVE ENVISION FIELDS FOR THE DEFINITION QUERY 
    Dim pField As IField
    Dim intFldCount As Integer
    If m_pEditFeatureLyr Is Nothing Then
      GoTo cleanup
    Else
      'CLEAR OUT PREVIOUS 
      m_dockEnvisionWinForm.itmcmbSubareaFields.Items.Clear()
      m_dockEnvisionWinForm.itmcmbSubareaFields.Text = ""

      For intFldCount = 1 To m_pEditFeatureLyr.FeatureClass.Fields.FieldCount - 1
        pField = m_pEditFeatureLyr.FeatureClass.Fields.Field(intFldCount)
        If pField.Type = esriFieldType.esriFieldTypeInteger Or pField.Type = esriFieldType.esriFieldTypeString Or pField.Type = esriFieldType.esriFieldTypeDouble Then
          m_dockEnvisionWinForm.itmcmbSubareaFields.Items.Add(pField.Name)
        End If
      Next
    End If

CleanUp:
    pField = Nothing
    intFldCount = Nothing
    GC.WaitForPendingFinalizers()
    GC.Collect()
  End Sub

  Public Sub LoadUniqueSubareaValues()
    Dim pDataStatistics As IDataStatistics
    Dim pCursor As ICursor
    Dim pEnumVar As System.Collections.IEnumerator
    Dim intUniqueCount As Integer
    Dim pQueryFilter As IQueryFilter
    Dim pFC As IFeatureClass = m_pEditFeatureLyr.FeatureClass
    Dim pTable As ITable = CType(pFC, ITable)
    Dim pRow As Row
    Dim dic As New SortedList(Of Object, String)
    Dim intField As Integer
    Dim strValue As String
    Dim intCount As Integer
    pQueryFilter = New QueryFilter


    'LOAD UNIQUE SUBAREA VALUES FROM THE ENVISION LAYER AND SELECTED FIELD
    If m_dockEnvisionWinForm.itmcmbSubareaFields.Text = "" Then
      GoTo CleanUp
    End If
    m_appEnvision.StatusBar.Message(0) = "Checking for Subarea unique values from selected field " & m_dockEnvisionWinForm.itmcmbSubareaFields.Text

    If m_pEditFeatureLyr Is Nothing Then
      GoTo cleanup
    End If

    Try
      'CLEAR OUT PREVIOUS VALUES
      m_dockEnvisionWinForm.itmSelectSubarea.Items.Clear()
      'CHECK FOR NUMBER OF UNIQUE VALUES
      pDataStatistics = New DataStatistics
      pCursor = CType(pTable.Search(Nothing, False), ICursor)
      pDataStatistics.Field = m_dockEnvisionWinForm.itmcmbSubareaFields.Text
      pDataStatistics.Cursor = pCursor
      'pEnumVar = CType(pDataStatistics.UniqueValues, System.Collections.IEnumerator)
      intUniqueCount = pDataStatistics.UniqueValueCount()
      m_appEnvision.StatusBar.Message(0) = "Field, " & m_dockEnvisionWinForm.itmcmbSubareaFields.Text & ", unique value count: " & intUniqueCount.ToString
      If intUniqueCount > 50 Then
        If MessageBox.Show("There are " + intUniqueCount.ToString + " unique values in the selected field.  Would you like to continue to load them all?", "Significant Number of Values", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2, MessageBoxOptions.DefaultDesktopOnly, False) = Windows.Forms.DialogResult.No Then
          Exit Try
        End If
      End If

      'pDataStatistics = Nothing
      'If intUniqueCount > 0 Then
      '    Dim blnMoved As Boolean
      '    Dim objValue As Object
      '    m_dockEnvisionWinForm.itmSelectSubarea.Text = ""
      '    m_dockEnvisionWinForm.itmSelectSubarea.Items.Clear()
      '    blnMoved = pEnumVar.MoveNext
      '    Do Until blnMoved = False
      '        objValue = pEnumVar.Current
      '        m_dockEnvisionWinForm.itmSelectSubarea.Items.Add(objValue.ToString)
      '        blnMoved = pEnumVar.MoveNext
      '    Loop
      'End If


      '***************************************************************************
      'WORKAROUND
      '***************************************************************************
      Try
        If Not pCursor Is Nothing Then
          pRow = CType(pCursor.NextRow(), Row)
          intField = pRow.Fields.FindField(m_dockEnvisionWinForm.itmcmbSubareaFields.Text)

          'loop through the data and get the unique values
          Do Until pRow Is Nothing
            strValue = pRow.Value(intField).ToString
            If dic.ContainsKey(strValue) = False Then
              dic.Add(strValue, "")
              If dic.Count = intUniqueCount Then
                Exit Do
              End If
            End If
            pRow = CType(pCursor.NextRow(), Row)
          Loop

          If dic.Count > 0 Then
            For intCount = 0 To dic.Count - 1
              strValue = dic.Keys.Item(intCount).ToString
              m_dockEnvisionWinForm.itmSelectSubarea.Items.Add(strValue.ToString)
            Next
          End If
        End If
      Catch ex As Exception
        MessageBox.Show("Error is building a list of unique land use types." & vbNewLine & ex.Message, "Land Use Value Error")
      End Try

      m_dockEnvisionWinForm.itmcmbSubareaFields.Tag = m_dockEnvisionWinForm.itmcmbSubareaFields.Text
      GoTo cleanup
    Catch ex As Exception
      MessageBox.Show(ex.Message, "Unique Subarea Value Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
      GoTo cleanup
    End Try

CleanUp:
    pTable = Nothing
    pDataStatistics = Nothing
    pCursor = Nothing
    pEnumVar = Nothing
    intUniqueCount = Nothing
    pQueryFilter = Nothing
    pFC = Nothing
    pRow = Nothing
    dic = Nothing
    intField = Nothing
    strValue = Nothing
    intCount = Nothing
    GC.Collect()
  End Sub

  'Public Function RoundNumber(ByVal dblNumber As Double, ByVal intDecimals As Integer) As Double
  '    'ROUND NUMBER TO A SET NUMBER OF DECIMAL PLACES
  '    Dim strNumber As String = ""
  '    RoundNumber = 0
  '    If intDecimals = 1 Then
  '        strNumber = dblNumber.ToString("#.#")
  '    ElseIf intDecimals = 2 Then
  '        strNumber = dblNumber.ToString("#.##")
  '    ElseIf intDecimals = 3 Then
  '        strNumber = dblNumber.ToString("#.###")
  '    ElseIf intDecimals = 4 Then
  '        strNumber = dblNumber.ToString("#.####")
  '    ElseIf intDecimals = 5 Then
  '        strNumber = dblNumber.ToString("#.#####")
  '    ElseIf intDecimals = 6 Then
  '        strNumber = dblNumber.ToString("#.######")
  '    ElseIf intDecimals = 7 Then
  '        strNumber = dblNumber.ToString("#.#######")
  '    ElseIf intDecimals = 8 Then
  '        strNumber = dblNumber.ToString("#.########")
  '    ElseIf intDecimals = 9 Then
  '        strNumber = dblNumber.ToString("#.#########")
  '    ElseIf intDecimals = 10 Then
  '        strNumber = dblNumber.ToString("#.##########")
  '    End If

  '    If strNumber = "" Then
  '        RoundNumber = 0
  '    Else
  '        RoundNumber = CDbl(strNumber)
  '    End If
  'End Function

  '  Public Sub EnvisionFieldTrack()
  '    'OPEN THE ENVISION FIELD TRACKING TABLE AND LOAD OPTIONS
  '    'FIELD TRACKING TABLE
  '    Dim GP As ESRI.ArcGIS.Geoprocessor.Geoprocessor
  '    Dim pCreateTable As ESRI.ArcGIS.DataManagementTools.CreateTable
  '    Dim pWksFactory As ESRI.ArcGIS.Geodatabase.IWorkspaceFactory = Nothing
  '    Dim pFeatWks As ESRI.ArcGIS.Geodatabase.IFeatureWorkspace = Nothing
  '    Dim rowTemp As IRow
  '    Dim strFieldName As String = ""
  '    Dim intUse As Integer
  '    Dim intCalcByAcres As Integer
  '    Dim strDepVarFieldName As String = ""
  '    Dim strCalcFieldName As String = ""
  '    Dim pCursor As ICursor
  '    Dim iRow As Integer
  '    Dim pTable As ITable

  '    'DEV TYPE ATTRIBUTE TABLE ATTRIUBUTES
  '    If m_tblAttribFields Is Nothing Then
  '      Try
  '        pWksFactory = New FileGDBWorkspaceFactory
  '        pFeatWks = DirectCast(pWksFactory.OpenFromFile(m_strFeaturePath, 0), IFeatureWorkspace)
  '        m_tblAttribFields = pFeatWks.OpenTable("ENVISION_DEVTYPE_FIELD_TRACKING")
  '      Catch ex As Exception
  '        GoTo CleanUp
  '      End Try
  '    End If

  '    'CLEAR THE EXISTING FIELD ARRAY LIST 
  '    m_arrWriteCalcFields.Clear()
  '    m_arrWriteDevTypeFields.Clear()
  '    m_arrWriteDevTypeAcresFields.Clear()
  '    m_arrWriteDevTypeAcresFieldsAltName.Clear()

  '    pCursor = m_tblAttribFields.Search(Nothing, False)
  '    rowTemp = pCursor.NextRow
  '    Do Until rowTemp Is Nothing
  '      strFieldName = ""
  '      intUse = 0
  '      intCalcByAcres = 0
  '      strCalcFieldName = ""
  '      Try
  '        strFieldName = CStr(rowTemp.Value(rowTemp.Fields.FindField("FIELD_NAME")))
  '      Catch ex As Exception
  '      End Try
  '      Try
  '        intUse = CInt(rowTemp.Value(rowTemp.Fields.FindField("USE")))
  '      Catch ex As Exception
  '      End Try
  '      Try
  '        intCalcByAcres = CInt(rowTemp.Value(rowTemp.Fields.FindField("CALC_BY_ACRES")))
  '      Catch ex As Exception
  '      End Try
  '      Try
  '        strDepVarFieldName = CStr(rowTemp.Value(rowTemp.Fields.FindField("DEP_VAR_FIELD_NAME")))
  '      Catch ex As Exception
  '      End Try
  '      Try
  '        strCalcFieldName = CStr(rowTemp.Value(rowTemp.Fields.FindField("CALC_FIELD_NAME")))
  '      Catch ex As Exception
  '      End Try
  '      If CType(intUse, Boolean) Then
  '        If Not strFieldName = "DEV_TYPE" And Not strFieldName = "VAC_ACRES" And Not strFieldName = "DEVD_ACRES" And Not strFieldName = "CONSTRAINED_ACRE" Then
  '          If intCalcByAcres = 0 Then
  '            m_arrWriteDevTypeFields.Add(strFieldName)
  '          End If
  '          If intCalcByAcres = 1 Or intCalcByAcres = 2 Then
  '            m_arrWriteDevTypeAcresFields.Add(strFieldName)
  '            m_arrWriteDevTypeAcresFieldsAltName.Add(strCalcFieldName)
  '            If intCalcByAcres = 2 Then
  '              m_arrWriteDevTypeAcres2ndVarFields.Add(strDepVarFieldName)
  '            Else
  '              m_arrWriteDevTypeAcres2ndVarFields.Add("")
  '            End If
  '          End If
  '        End If
  '      End If
  '      rowTemp = pCursor.NextRow
  '    Loop

  'CleanUp:
  '    GP = Nothing
  '    pCreateTable = Nothing
  '    pWksFactory = Nothing
  '    pFeatWks = Nothing
  '    rowTemp = Nothing
  '    strFieldName = Nothing
  '    intUse = Nothing
  '    intCalcByAcres = Nothing
  '    strDepVarFieldName = Nothing
  '    pCursor = Nothing
  '    iRow = Nothing
  '    pTable = Nothing
  '    GC.Collect()
  '    GC.WaitForPendingFinalizers()
  '  End Sub

  '  Public Sub EnvisionDevTypeFieldTrack()
  '    'OPEN THE ENVISION DEV TYPES TRACKING TABLE AND LOAD OPTIONS
  '    Dim GP As ESRI.ArcGIS.Geoprocessor.Geoprocessor
  '    Dim pCreateTable As ESRI.ArcGIS.DataManagementTools.CreateTable
  '    Dim pWksFactory As ESRI.ArcGIS.Geodatabase.IWorkspaceFactory = Nothing
  '    Dim pFeatWks As ESRI.ArcGIS.Geodatabase.IFeatureWorkspace = Nothing
  '    Dim pFieldTable As ITable
  '    Dim rowTemp As IRow
  '    Dim strFieldName As String = ""
  '    Dim intUse As Integer = 0
  '    Dim intCalcByAcres As Integer = 0
  '    Dim strDepVarFieldName As String = ""
  '    Dim pCursor As ICursor
  '    Dim iRow As Integer
  '    Dim pTable As ITable

  '    Try
  '      pWksFactory = New FileGDBWorkspaceFactory
  '      pFeatWks = DirectCast(pWksFactory.OpenFromFile(m_strFeaturePath, 0), IFeatureWorkspace)
  '      pFieldTable = pFeatWks.OpenTable("ENVISION_ATTRIBUTE_FIELD_TRACKING")
  '    Catch ex As Exception
  '      GoTo CleanUp
  '    End Try

  '    'CLEAR THE EXITING FIELD ARRAY LIST
  '    m_arrWriteDevTypeFields.Clear()
  '    m_arrWriteDevTypeAcresFields.Clear()
  '    m_arrWriteDevTypeAcres2ndVarFields.Clear()

  '    pCursor = pFieldTable.Search(Nothing, False)
  '    rowTemp = pCursor.NextRow
  '    Do Until rowTemp Is Nothing
  '      strFieldName = ""
  '      intUse = 0
  '      intCalcByAcres = 0
  '      Try
  '        strFieldName = CStr(rowTemp.Value(rowTemp.Fields.FindField("FIELD_NAME")))
  '      Catch ex As Exception
  '      End Try
  '      Try
  '        intUse = CInt(rowTemp.Value(rowTemp.Fields.FindField("USE")))
  '      Catch ex As Exception
  '      End Try
  '      Try
  '        intCalcByAcres = CInt(rowTemp.Value(rowTemp.Fields.FindField("CALC_BY_ACRES")))
  '      Catch ex As Exception
  '      End Try
  '      Try
  '        strDepVarFieldName = CStr(rowTemp.Value(rowTemp.Fields.FindField("DEP_VAR_FIELD_NAME")))
  '      Catch ex As Exception
  '      End Try
  '      If intUse = 1 Then
  '        If intCalcByAcres = 0 Then
  '          m_arrWriteDevTypeFields.Add(strFieldName)
  '          If intCalcByAcres >= 1 Then
  '            m_arrWriteDevTypeAcresFields.Add(strFieldName)
  '            If intCalcByAcres = 2 Then
  '              m_arrWriteDevTypeAcres2ndVarFields.Add(strDepVarFieldName)
  '            End If
  '          End If
  '        End If
  '      End If
  '      rowTemp = pCursor.NextRow
  '    Loop

  'CleanUp:
  '    GP = Nothing
  '    pCreateTable = Nothing
  '    pWksFactory = Nothing
  '    pFeatWks = Nothing
  '    pFieldTable = Nothing
  '    rowTemp = Nothing
  '    strFieldName = Nothing
  '    intUse = Nothing
  '    intCalcByAcres = Nothing
  '    strDepVarFieldName = Nothing
  '    pCursor = Nothing
  '    iRow = Nothing
  '    pTable = Nothing
  '    GC.WaitForPendingFinalizers()
  '    GC.Collect()
  '  End Sub

  Public Sub WriteExistingLUtoExcel()
    'WRITE THE EXISITNG LU ACRE VALUES TO THE SELECTED ENVISION EXCEL FILE
    Dim intRow As Integer
    Dim intTblRow As Integer
    Dim intFieldRow As Integer = -1
    Dim intStartRow As Integer = -1
    Dim strFldCellValue As String = ""
    Dim intCol As Integer = 0
    Dim shtExisting As Microsoft.Office.Interop.Excel.Worksheet = Nothing
    Dim strDevdAcresAllFld As String = ""
    Dim strDevdAcresPaintedFld As String = ""
    Dim intExLUCol As Integer = 0
    Dim intDevdAcresAllCol As Integer = -1
    Dim intDevdAcresPaintedCol As Integer = -1
    Dim intExcelFormulaCalc As Integer = 0

    'EXIT IF EXISTING LU TABLE IS MISSING
    If m_tblExistingLU Is Nothing Then
      GoTo CleanUp
    End If

    'RETRIEVE OR OPEN THE SELECTED ENVISION EXCEL FILE
    Try
      If m_xlPaintWB1 Is Nothing Then
        GoTo CleanUp
      End If
      Try
        shtExisting = DirectCast(m_xlPaintWB1.Sheets("Existing Developed Area"), Microsoft.Office.Interop.Excel.Worksheet)
      Catch ex As Exception
      End Try
    Catch ex As Exception
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
      GoTo CleanUp
    End If

    'CYCLE THROUGH THE FIRST 25 COULMNS FOR THE INPUT FIELDS, FIRST FOUND IS EXISTING LAND USE FIELD, SECOND IS DEVELOPED ACREAS FOR ALL FEATURES, THIRD IS DEVELOPED ACREAS FOR PAINTED FEATURES
    For intCol = 1 To 55
      strFldCellValue = CStr(CType(shtExisting.Cells(intFieldRow, intCol), Microsoft.Office.Interop.Excel.Range).Value)
      If Not strFldCellValue Is Nothing Then
        If UCase(strFldCellValue) = ("PAINTED" & m_intEditScenario.ToString) And strFldCellValue.Length > 0 Then
          intDevdAcresPaintedCol = intCol
          'Exit For
        End If
        If UCase(strFldCellValue) = ("EX_LU") Then
          intExLUCol = intCol
          'Exit For
        End If
      End If
    Next
    If intDevdAcresPaintedCol = -1 Then
      GoTo CleanUp
    End If

    'CLEAR VALUES
    For intRow = intStartRow To (intStartRow + 20)
      CType(shtExisting.Cells(intRow, intDevdAcresPaintedCol), Microsoft.Office.Interop.Excel.Range).Value = ""
    Next

    'CODE TO WRITE THE SUMMARY VALUES TO EXCEL FILE
    Try
      For intRow = intStartRow To (intStartRow + 20)
        strFldCellValue = CStr(CType(shtExisting.Cells(intRow, intExLUCol), Microsoft.Office.Interop.Excel.Range).Value)
        If Not strFldCellValue Is Nothing Then
          'PAINTED FEATURES
          If Not m_tblExistingLU Is Nothing Then
            For intTblRow = 1 To m_tblExistingLU.RowCount(Nothing)
              Try
                If strFldCellValue = CStr(m_tblExistingLU.GetRow(intTblRow).Value(1)) Then
                  CType(shtExisting.Cells(intRow, intDevdAcresPaintedCol), Microsoft.Office.Interop.Excel.Range).Value = m_tblExistingLU.GetRow(intTblRow).Value(2)
                End If
              Catch ex As Exception

              End Try
            Next
          End If
        End If
      Next
    Catch ex As Exception
      GoTo CleanUp
    End Try
    GoTo CleanUp


CleanUp:
    'RESET FORMULA CALC SETTING
    If Not m_xlPaintApp Is Nothing Then
      If intExcelFormulaCalc = 1 Then
        m_xlPaintApp.Calculation = Interop.Excel.XlCalculation.xlCalculationAutomatic
      ElseIf intExcelFormulaCalc = 2 Then
        m_xlPaintApp.Calculation = Interop.Excel.XlCalculation.xlCalculationManual
      ElseIf intExcelFormulaCalc = 3 Then
        m_xlPaintApp.Calculation = Interop.Excel.XlCalculation.xlCalculationSemiautomatic
      End If
    End If
    intRow = Nothing
    intTblRow = Nothing
    intFieldRow = Nothing
    intStartRow = Nothing
    strFldCellValue = Nothing
    intCol = Nothing
    shtExisting = Nothing
    strDevdAcresAllFld = Nothing
    strDevdAcresPaintedFld = Nothing
    intExLUCol = Nothing
    intDevdAcresAllCol = Nothing
    intDevdAcresPaintedCol = Nothing
    GC.WaitForPendingFinalizers()
    GC.Collect()
  End Sub

  Public Sub EnvisionSyncEX_LU_CreateSummaryTablesPaint()
    'POPULATES THE EXISTING DEVELOPED AREA TAB WITH VALEUS FROM 2 SUMMARY TABLES 

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
    Dim pData As IDataStatistics
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
    Dim pSumAllTable As ITable = Nothing
    Dim pWkSpName As IName
    Dim pOutTabName As ITableName
    Dim pOutDatasetName As IDatasetName
    Dim pDataSet As IDataset
    Dim pWkSpDS As IDataset
    Dim pFeatSelection As IFeatureSelection
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
          m_xlPaintApp.DisplayAlerts = False
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
        If Not TypeOf m_xlPaintWB1.Sheets("Existing Developed Area") Is Microsoft.Office.Interop.Excel.Worksheet Then
          'MessageBox.Show("The 'Existing Developed Area' tab could not be found in the selected ENVISION Excel file. Please select another file.", "TAB NOT FOUND", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
          GoTo CloseExcel
        Else
          shtExisting = DirectCast(m_xlPaintWB1.Sheets("Existing Developed Area"), Microsoft.Office.Interop.Excel.Worksheet)
        End If
      Catch ex As Exception
        ' MessageBox.Show("The 'Existing Developed Area' tab could not be found in the selected ENVISION Excel file. Please select another file.", "TAB NOT FOUND", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
        GoTo CloseExcel
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
      MessageBox.Show("The 'Existing Developed Area' tab does not appear to be formatted correctly.", "Value Entry 1 not Found", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
      GoTo CleanUp
    End If

    'SELECT ALL FEATURES / APPLY QUERY DEFINITION IF AVAILABLE
    pFeatClass = pLayer.FeatureClass
    pDef = DirectCast(pLayer, IFeatureLayerDefinition2)
    strDefExpression = pDef.DefinitionExpression
    pQFilter = New QueryFilter
    If strDefExpression.Length > 0 Then
      pQFilter.WhereClause = pDef.DefinitionExpression
      pCursor = DirectCast(pFeatClass.Search(pQFilter, False), ICursor)
    Else
      pCursor = DirectCast(pFeatClass.Search(Nothing, False), ICursor)
    End If

    'CYCLE THROUGH THE FIRST 50 COULMNS FOR THE INPUT FIELDS, FIRST FOUND IS EXISTING LAND USE FIELD, SECOND IS DEVELOPED ACREAS FOR ALL FEATURES, THIRD IS DEVELOPED ACREAS FOR PAINTED FEATURES
    intCount = 0
    intTotalCount = 50
    For intCol = 1 To 50
      intCount = intCount + 1
      strFldCellValue = CStr(CType(shtExisting.Cells(intFieldRow, intCol), Microsoft.Office.Interop.Excel.Range).Value)
      If Not strFldCellValue Is Nothing Then
        'If pFeatClass.FindField(strFldCellValue) >= 0 Then
        If UCase(strFldCellValue) = ("PAINTED" & m_intEditScenario.ToString) And strFldCellValue.Length > 0 Then
          strDevdAcresPaintedFld = "DEVD_ACRE"
          intDevdAcresPaintedCol = intCol
          Exit For
        End If
        'End If
        If strExLUFld.Length > 0 And strDevdAcresAllFld.Length > 0 And strDevdAcresPaintedFld.Length > 0 Then
          Exit For
        End If
      End If
    Next
    If intDevdAcresPaintedCol = -1 Then
      GoTo CleanUp
    End If

    pTable = CType(pFeatClass, ITable)
    pDataSet = DirectCast(pTable, IDataset)
    pWkSpDS = DirectCast(pDataSet.Workspace, IDataset)
    pWkSpName = pWkSpDS.FullName

    'FIRST PROCESS THE DEV TYPE PAINTED FEATURES ONLY
    Try
      pFeatClass = pLayer.FeatureClass
      pDef = DirectCast(pLayer, IFeatureLayerDefinition2)
      strDefExpression = pDef.DefinitionExpression
      pQFilter = New QueryFilter
      strQString = "NOT DEV_TYPE = ''"
      If strDefExpression.Length > 0 Then
        pQFilter.WhereClause = "(" & strQString & ") AND (" & pDef.DefinitionExpression & ")"
        pFeatSelection = CType(pLayer, IFeatureSelection)
        pFeatSelection.SelectFeatures(pQFilter, esriSelectionResultEnum.esriSelectionResultNew, False)
      Else
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
      pOutDatasetName = DirectCast(pOutTabName, IDatasetName)
      pOutDatasetName.Name = "EXISTING_DEVELOPED_AREA_PAINTED"
      pOutDatasetName.WorkspaceName = DirectCast(pWkSpName, IWorkspaceName)

      pGeoProc = New BasicGeoprocessor
      If strExLUFld.Length > 0 And strDevdAcresPaintedFld.Length > 0 Then
        m_appEnvision.StatusBar.Message(0) = "Creating summary table for all PAINTED features."
        m_tblExistingLU = pGeoProc.Dissolve(DirectCast(pLayer, ITable), True, strExLUFld, ("Minimum." & strExLUFld & ",Sum." & strDevdAcresPaintedFld), pOutDatasetName)
      End If
    Catch ex As Exception

    End Try

    'CLEAR VALUES
    m_appEnvision.StatusBar.Message(0) = "Clearing previous developed acre values from excel table."
    For intRow = intStartRow To (intStartRow + 50)
      strFldCellValue = CStr(CType(shtExisting.Cells(intRow, intExLUCol), Microsoft.Office.Interop.Excel.Range).Value)
      If Not strFldCellValue Is Nothing Then
        'ALL FEATURES
        CType(shtExisting.Cells(intRow, intDevdAcresAllCol), Microsoft.Office.Interop.Excel.Range).Value = ""
        'PAINTED FEATURES
        CType(shtExisting.Cells(intRow, intDevdAcresPaintedCol), Microsoft.Office.Interop.Excel.Range).Value = ""
      End If
    Next

    'CODE TO WRITE THE SUMMARY VALUES TO EXCEL FILE
    For intRow = intStartRow To (intStartRow + 50)
      strFldCellValue = CStr(CType(shtExisting.Cells(intRow, intExLUCol), Microsoft.Office.Interop.Excel.Range).Value)
      If Not strFldCellValue Is Nothing Then
        m_appEnvision.StatusBar.Message(0) = "Pulling SUM values for LU, " & strFldCellValue & "."
        'ALL FEATURES
        If Not pSumAllTable Is Nothing Then
          For intTblRow = 1 To pSumAllTable.RowCount(Nothing)
            Try
              If strFldCellValue = CStr(pSumAllTable.GetRow(intTblRow).Value(1)) Then
                CType(shtExisting.Cells(intRow, intDevdAcresAllCol), Microsoft.Office.Interop.Excel.Range).Value = pSumAllTable.GetRow(intTblRow).Value(2)
              End If
            Catch ex As Exception

            End Try
          Next
        End If
        'PAINTED FEATURES
        If Not m_tblExistingLU Is Nothing Then
          For intTblRow = 1 To m_tblExistingLU.RowCount(Nothing)
            Try
              If strFldCellValue = CStr(m_tblExistingLU.GetRow(intTblRow).Value(1)) Then
                CType(shtExisting.Cells(intRow, intDevdAcresPaintedCol), Microsoft.Office.Interop.Excel.Range).Value = m_tblExistingLU.GetRow(intTblRow).Value(2)
              End If
            Catch ex As Exception

            End Try
          Next
        End If
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

    m_appEnvision.StatusBar.Message(0) = String.Empty

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
    pFeatureCursor = Nothing
    intRow = Nothing
    intTblRow = Nothing
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
    pGeoProc = Nothing
    strExLUFld = Nothing
    strDevdAcresAllFld = Nothing
    strDevdAcresPaintedFld = Nothing
    intExLUCol = Nothing
    intDevdAcresAllCol = Nothing
    intDevdAcresPaintedCol = Nothing
    pSumAllTable = Nothing
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

  ''' -----------------------------------------------------------------------------
  ''' <summary>
  ''' Makes the form float on top of it's parent.
  ''' </summary>
  ''' <param name="FormHwnd"></param>
  ''' <param name="appHwnd"></param>
  ''' <remarks>
  ''' Uses SetWindowLong() API call to make the form minimize with ArcMap.
  ''' </remarks>
  ''' <history>
  ''' </history>
  ''' -----------------------------------------------------------------------------
  Public Sub OnTop(ByVal FormHwnd As Int32, ByVal appHwnd As Int32)
    SetWindowLong(FormHwnd, GWL_HWNDPARENT, appHwnd)
  End Sub

  ''' <summary>
  ''' Returns the default name for the total field in the scenario featureclass
  ''' </summary>
  ''' <param name="calcByAcres"></param>
  ''' <param name="calcFieldName"></param>
  ''' <param name="fieldName"></param>
  ''' <param name="featureClass"></param>
  ''' <returns></returns>
  ''' <remarks>
  ''' If we are calculating by acres, then the field name is based on the specified calculated field name.  Otherwise it is based on the field name of the attribute.
  ''' </remarks>
  Private Function GetDefaultTotalFieldName(calcByAcres As Int32, calcFieldName As String, fieldName As String, featureClass As IFeatureClass) As String
    Dim totFldName As String = String.Empty
    Dim exfldname As String
    If calcByAcres <> 0 Then
      exfldname = calcFieldName
    Else
      exfldname = fieldName
    End If
    If featureClass.Fields.FindField("EX_" & exfldname) <> -1 Then
      totFldName = "TOT_" & exfldname
    End If
    Return totFldName
  End Function

End Module

