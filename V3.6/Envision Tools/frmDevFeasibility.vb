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

Public Class frmDevFeasibility
  Public lstSlopeValues As ArrayList = New ArrayList
  Public lstYInterceptValues As ArrayList = New ArrayList
  Public lstClassType As ArrayList = New ArrayList

  Private Sub frmRedevFeasibility_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
    'CLEAR LISTS VARIABLES
    lstSlopeValues.Clear()
    lstSlopeValues = Nothing
    lstYInterceptValues.Clear()
    lstYInterceptValues = Nothing

    Try
      m_frmRedevTiming = Nothing
      GC.Collect()
      GC.WaitForPendingFinalizers()
      GoTo CleanUp
    Catch ex As Exception
      MessageBox.Show(Me, ex.Message, "Closing Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
      GoTo CleanUp
    End Try

CleanUp:
    m_frmDevFeasibility = Nothing
    GC.Collect()
    GC.WaitForPendingFinalizers()
  End Sub

  Private Sub frmRedevFeasibility_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    'LOAD DEV TYPES AND DEV TYPE ATTRIBUTE VALUES FROM THE LOOK UP TABLE
    'REQUIRED DEV TYPE ATTRIBUTES ARE DEV TYPE NAme," SLOPE(m) and Y-INTERCEPT (b)
    'EQUAITONS:
    'STEP 1: X = Land Value (1+r) + Improvement Value) / Parcel Square Footage
    'STEP 2: Y = mX + b
    Dim mxApplication As IMxApplication = Nothing
    Dim pMxDocument As IMxDocument = Nothing
    Dim mapCurrent As Map
    Dim pActiveView As IActiveView = Nothing
    Dim pFeatSelection As IFeatureSelection
    Dim pCursor As ICursor = Nothing
    Dim pField As IField
    Dim intFldCount As Integer
    Dim intRow As Integer
    Dim rowTemp As IRow
    Dim intDevTypeFld As Integer
    Dim intSlopeFld As Integer
    Dim intYInterceptFld As Integer
    Dim strDevTypeName As String = ""
    Dim dblSlope As Double
    Dim dblYIntercept As Double
    Dim strType As String = ""
    Dim intTypeField As String

    'RETRIEVE CURRENT VIEW DOCUMENT TO OBTAIN LIST OF CURRENT LAYER(S)
    If Not TypeOf m_appEnvision Is IApplication Then
      GoTo CleanUp
    Else
      pMxDocument = DirectCast(m_appEnvision.Document, IMxDocument)
      mapCurrent = CType(pMxDocument.FocusMap, Map)
      pActiveView = CType(pMxDocument.FocusMap, IActiveView)
    End If


    'EXIT IF THE USER HAS NOT DEFINED A PARCEL LAYER
    If m_pEditFeatureLyr Is Nothing Then
      MessageBox.Show(Me, "Please select a Parcel layer to use this tool.", "Parcel Layer Not Defined", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1)
      GoTo CloseForm
    Else
      'RETRIEVE SELECTED NEIGHBORHOOD LAYER
      pFeatSelection = CType(m_pEditFeatureLyr, IFeatureSelection)
      pFeatSelection.SelectionSet.Search(Nothing, False, pCursor)
      'Default to ALL features
      If pFeatSelection.SelectionSet.Count > 0 Then
        Me.chkParcelSelected.Enabled = True
        Me.chkParcelSelected.Text = "Use " & pFeatSelection.SelectionSet.Count.ToString & " Selected Features"
      Else
        Me.chkParcelSelected.Enabled = False
      End If
    End If

    'RETRIEVE LIST OF VALUE FIELDS 
    For intFldCount = 1 To m_pEditFeatureLyr.FeatureClass.Fields.FieldCount - 1
      pField = m_pEditFeatureLyr.FeatureClass.Fields.Field(intFldCount)
      If pField.Type = esriFieldType.esriFieldTypeInteger Or pField.Type = esriFieldType.esriFieldTypeDouble Then
        Me.cmbLandValueFld.Items.Add(pField.Name)
        Me.cmbImprovementValueFld.Items.Add(pField.Name)
        Me.cmbSqFtFld.Items.Add(pField.Name)
      End If
    Next

    'RETRIEVE THE DEV TYPES TABLE
    If m_tblDevelopmentTypes Is Nothing Then
      MessageBox.Show(Me, "The Development Types table has not been set.  Please Load Development Type Attributes.", "Table Not Found", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1)
    Else
      'RETRIEVE THE REQUIRED DEVTYPE TABLE FIELDS
      intDevTypeFld = m_tblDevelopmentTypes.FindField("DEVELOPMENT_TYPE")
      If intDevTypeFld <= -1 Then
        MessageBox.Show(Me, "The required field, DEVTYPE, could not be found within the current Development Types lookup table.", "Required Field Not Found", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1)
        GoTo CloseForm
      End If
      intSlopeFld = m_tblDevelopmentTypes.FindField("Slope")
      If intSlopeFld <= -1 Then
        MessageBox.Show(Me, "The required field, Slope, could not be found within the current Development Types lookup table.", "Required Field Not Found", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1)
        GoTo CloseForm
      End If
      intYInterceptFld = m_tblDevelopmentTypes.FindField("Y_Intercept")
      If intYInterceptFld <= -1 Then
        MessageBox.Show(Me, "The required field, Y_Intercept, could not be found within the current Development Types lookup table.", "Required Field Not Found", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1)
        GoTo CloseForm
      End If
      intTypeField = m_tblDevelopmentTypes.FindField("Product_Type")
      If intTypeField <= -1 Then
        MessageBox.Show(Me, "The required field, Product_Type, could not be found within the current Development Types lookup table.", "Required Field Not Found", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1)
        GoTo CloseForm
      End If
      'EXIT IF EMPTY TABLE
      If m_tblDevelopmentTypes.RowCount(Nothing) <= 0 Then
        MessageBox.Show(Me, "The Development Types table is empty.  Please load development types.", "Empty Table", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1)
        GoTo CloseForm
      End If

      'CLEAR ANY PREVIOUS VALUES FROM THE DEVELOPMENT TYPES LIST
      Me.cmbDevTypes.Items.Clear()
      lstSlopeValues.Clear()
      lstYInterceptValues.Clear()
      lstClassType.Clear()


      'BUILD LIST OF DEV TYPES WHERE THE SLOPE AND YINTERCEPT ARE BOTH GREATER THAN 0
      For intRow = 1 To m_intDevTypeMax
        Try
          rowTemp = m_tblDevelopmentTypes.GetRow(intRow)
          strDevTypeName = ""
          dblSlope = 0
          dblYIntercept = 0
          Try
            strDevTypeName = CStr(rowTemp.Value(intDevTypeFld))
          Catch ex As Exception
            Continue For
          End Try
          Try
            dblSlope = CDbl(rowTemp.Value(intSlopeFld))
          Catch ex As Exception
            Continue For
          End Try
          Try
            dblYIntercept = CDbl(rowTemp.Value(intYInterceptFld))
          Catch ex As Exception
            Continue For
          End Try
          Try
            strType = CStr(rowTemp.Value(intTypeField))
          Catch ex As Exception
            Continue For
          End Try
          If dblSlope > 0 And dblYIntercept > 0 Then
            Me.cmbDevTypes.Items.Add(strDevTypeName)
            lstSlopeValues.Add(dblSlope)
            lstYInterceptValues.Add(dblYIntercept)
            lstClassType.Add(strType)
          End If
        Catch ex As Exception

        End Try

      Next
    End If
    GoTo CleanUp

CloseForm:
    Me.Close()
    GoTo CleanUp

CleanUp:
    mxApplication = Nothing
    pMxDocument = Nothing
    mapCurrent = Nothing
    pActiveView = Nothing
    pFeatSelection = Nothing
    pCursor = Nothing
    pField = Nothing
    intFldCount = Nothing
    intRow = Nothing
    rowTemp = Nothing
    intDevTypeFld = Nothing
    intSlopeFld = Nothing
    intYInterceptFld = Nothing
    strDevTypeName = Nothing
    dblSlope = Nothing
    dblYIntercept = Nothing
    GC.WaitForPendingFinalizers()
    GC.Collect()
  End Sub

  Private Sub tbxLandValueAppreciation_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbxLandValueAppreciation.TextChanged
    'MAKE SURE TEXT IS NUMERICAL
    If IsNumeric(Me.tbxLandValueAppreciation.Text) Then
      Me.tbxLandValueAppreciation.ForeColor = Color.Black
    Else
      Me.tbxLandValueAppreciation.ForeColor = Color.Red
    End If
  End Sub

  Private Sub tbxTargetRent_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbxTargetRent.TextChanged
    'MAKE SURE TEXT IS NUMERICAL
    If IsNumeric(Me.tbxTargetRent.Text) Then
      Me.tbxTargetRent.ForeColor = Color.Black
    Else
      Me.tbxTargetRent.ForeColor = Color.Red
    End If
  End Sub

  Private Sub btnRun_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRun.Click
    Dim dblAppreciation As Double = 0
    Dim dblTargetRent As Double = 0
    Dim intFldLandValue As Integer = 0
    Dim intFldImprovmentValue As Integer = 0
    Dim intFldSqFtParcel As Integer = 0
    Dim dblLandValue As Double = 0
    Dim dblImprovmentValue As Double = 0
    Dim dblSqFtParcel As Double = 0
    Dim pFeatSelection As IFeatureSelection
    Dim pFeatClass As IFeatureClass
    Dim pCursor As ICursor = Nothing
    Dim dblSlope As Double
    Dim dblYIntercept As Double
    Dim pQueryFilter As IQueryFilter
    Dim strQString As String = ""
    Dim pFeatureCursor As IFeatureCursor = Nothing
    Dim pFeat As IFeature
    Dim dblYValue As Double = 0
    Dim intWriteFld As Integer = 0
    Dim intTargetFld As Integer = -1
    Dim pTable As ITable
    Dim mxApplication As IMxApplication = Nothing
    Dim pMxDocument As IMxDocument = Nothing
    Dim mapCurrent As Map
    Dim pActiveView As IActiveView = Nothing
    Dim intLayer As Integer
    Dim pFeatLayer As IFeatureLayer = Nothing
    Dim pDataset As IDataset
    Dim blnLayerFound As Boolean = False
    Dim strPaintName As String = ""
    Dim intTotalCount As Integer
    Dim intCount As Integer = 0
    Dim dblRentTarget As Double

    Me.lblStatus.Text = "Please wait while Envision calculates Development Feasibility"
    Me.Refresh()
    'REVIEW ALL INPUTS
    '----------------------------------------------------------------------------------
    If Me.cmbDevTypes.FindString(Me.cmbDevTypes.Text) = -1 Then
      MessageBox.Show(Me, "Please select a Development Type.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Stop)
      GoTo CleanUp
    Else
      Try
        dblSlope = lstSlopeValues.Item(Me.cmbDevTypes.FindString(Me.cmbDevTypes.Text))
        dblYIntercept = lstYInterceptValues.Item(Me.cmbDevTypes.FindString(Me.cmbDevTypes.Text))
      Catch ex As Exception
        MessageBox.Show(Me, "Failure to retrieve Slope and Y-Intercept values for the selected Development Type.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Stop)
        GoTo CleanUp
      End Try
    End If

    If IsNumeric(Me.tbxLandValueAppreciation.Text) Then
      dblAppreciation = (CDbl(Me.tbxLandValueAppreciation.Text) / 100)
    Else
      MessageBox.Show(Me, "Please provide a numeric value for Land Value Appreciation.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Stop)
      GoTo CleanUp
    End If

    If IsNumeric(Me.tbxTargetRent.Text) Then
      dblTargetRent = CDbl(Me.tbxTargetRent.Text)
    Else
      MessageBox.Show(Me, "Please provide a numeric value for Target Archievable Rent.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Stop)
      GoTo CleanUp
    End If

    If Me.cmbLandValueFld.Text.Length > 0 Then
      intFldLandValue = m_pEditFeatureLyr.FeatureClass.Fields.FindField(Me.cmbLandValueFld.Text)
      If intFldLandValue <= -1 Then
        MessageBox.Show(Me, "Please select a field for Land Value.  The field input, " & Me.cmbLandValueFld.Text & ", could not be found.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Stop)
        GoTo CleanUp
      End If
    Else
      MessageBox.Show(Me, "Please select a field for Land Value.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Stop)
      GoTo CleanUp
    End If

    If Me.cmbImprovementValueFld.Text.Length > 0 Then
      intFldImprovmentValue = m_pEditFeatureLyr.FeatureClass.Fields.FindField(Me.cmbImprovementValueFld.Text)
      If intFldImprovmentValue <= -1 Then
        MessageBox.Show(Me, "Please select a field for Improvement Value.  The field input, " & Me.cmbImprovementValueFld.Text & ", could not be found.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Stop)
        GoTo CleanUp
      End If
    Else
      MessageBox.Show(Me, "Please select a field for Improvement Value.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Stop)
      GoTo CleanUp
    End If

    If Me.cmbSqFtFld.Text.Length > 0 Then
      intFldSqFtParcel = m_pEditFeatureLyr.FeatureClass.Fields.FindField(Me.cmbSqFtFld.Text)
      If intFldSqFtParcel <= -1 Then
        MessageBox.Show(Me, "Please select a field for Parcel Square Footage.  The field input, " & Me.cmbSqFtFld.Text & ", could not be found.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Stop)
        GoTo CleanUp
      End If
    Else
      MessageBox.Show(Me, "Please select a field for Parcel Square Footage.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Stop)
      GoTo CleanUp
    End If

    'CHECK FOR THE OUTPUT FIELD, REDEV_FEASIBILITY.
    pTable = CType(m_pEditFeatureLyr.FeatureClass, ITable)
    If pTable.FindField("REDEV_FEASIBILITY") <= -1 Then
      If Not AddEnvisionField(pTable, "REDEV_FEASIBILITY", "DOUBLE", 16, 10) Then
        GoTo CleanUp
      End If
    End If
    intWriteFld = m_pEditFeatureLyr.FeatureClass.FindField("REDEV_FEASIBILITY")
    'CHECK FOR THE OUTPUT FIELD, REDEV_TARGET.
    If pTable.FindField("REDEV_TARGET") <= -1 And CDbl(Me.tbxTargetRent.Text) > 0 Then
      If Not AddEnvisionField(pTable, "REDEV_TARGET", "DOUBLE", 16, 10) Then
        GoTo CleanUp
      End If
    End If
    intTargetFld = m_pEditFeatureLyr.FeatureClass.FindField("REDEV_TARGET")

    '----------------------------------------------------------------------------------
    'EXECUTE THE CALCULATIONS

    'RETREIEVE THE SELECTION SET TO CYCLE THROUGH 
    'EITHER THE USER DEFINED OR SELECT THE DEV TYPE THE USER DEFINED
    pFeatSelection = CType(m_pEditFeatureLyr, IFeatureSelection)
    If Me.chkParcelSelected.Checked Then
      pFeatSelection.SelectionSet.Search(Nothing, False, pCursor)
      pFeatureCursor = DirectCast(pCursor, IFeatureCursor)
      intTotalCount = pFeatSelection.SelectionSet.Count
    Else
      Try
        pFeatClass = m_pEditFeatureLyr.FeatureClass
        pFeatureCursor = pFeatClass.Search(Nothing, False)
        intTotalCount = pFeatClass.FeatureCount(Nothing)
      Catch ex As Exception
        'MessageBox.Show(me,ex.Message)
      End Try
    End If

    pFeat = pFeatureCursor.NextFeature
    Do While Not pFeat Is Nothing
      Try
        dblLandValue = pFeat.Value(intFldLandValue)
      Catch ex As Exception
        dblLandValue = 0
      End Try
      Try
        dblImprovmentValue = pFeat.Value(intFldImprovmentValue)
      Catch ex As Exception
        dblImprovmentValue = 0
      End Try
      Try
        dblSqFtParcel = pFeat.Value(intFldSqFtParcel)
      Catch ex As Exception
        dblSqFtParcel = 0
      End Try
      Try
        If dblSqFtParcel <= 0 Then
          dblYValue = 0
        Else
          dblYValue = (dblSlope * (((dblLandValue + dblImprovmentValue) / dblSqFtParcel) * (1 + dblAppreciation))) + dblYIntercept
        End If

        pFeat.Value(intWriteFld) = dblYValue
        If dblTargetRent > 0 And intTargetFld >= 0 Then
          dblRentTarget = (1 - (dblYValue / dblTargetRent))
          If dblRentTarget < -1 Then
            dblRentTarget = -1
          End If
          If dblRentTarget > 1 Then
            dblRentTarget = 1
          End If
          pFeat.Value(intTargetFld) = dblRentTarget
        End If
        pFeat.Store()
      Catch ex As Exception
        MessageBox.Show(Me, ex.Message, "Calc Failure")
      End Try
      pFeat = pFeatureCursor.NextFeature
      'm_appEnvision.StatusBar.Message(0) = (((intCount / intTotalCount) * 100)).ToString & "% Complete"
    Loop
    Me.lblStatus.Text = "Calculating Development Feasibility has finished"
    Me.Refresh()

    '--------------------------------------------------------------------------------------
    'ADD THE PARCEL LAYER FOR SYMBOLZING TO THE DEVELOPMENT FEASIBILITY CALCS
    'RETRIEVE CURRENT VIEW DOCUMENT TO OBTAIN LIST OF CURRENT LAYER(S)
    If Not TypeOf m_appEnvision Is IApplication Then
      GoTo CleanUp
    Else
      pMxDocument = DirectCast(m_appEnvision.Document, IMxDocument)
      mapCurrent = CType(pMxDocument.FocusMap, Map)
      pActiveView = CType(pMxDocument.FocusMap, IActiveView)
    End If

    'ADD A COPY OF THE PAINT LAYER FOR SYMBOLIZING OF THE DEVELOPMENT FEASIBILITY VALUES
    pFeatClass = m_pEditFeatureLyr.FeatureClass
    pDataset = CType(pFeatClass, IDataset)
    strPaintName = pDataset.Name & " <REDEV FEASIBILITY>"
    If mapCurrent.LayerCount > 0 Then
      For intLayer = 0 To mapCurrent.LayerCount - 1
        If TypeOf mapCurrent.Layer(intLayer) Is IFeatureLayer Then
          pFeatLayer = CType(mapCurrent.Layer(intLayer), IFeatureLayer)
          If pFeatLayer.Name = strPaintName Then
            blnLayerFound = True
            Exit For
          End If
        End If
      Next
    End If
    If Not blnLayerFound Then
      pFeatLayer = New FeatureLayer
      pFeatLayer.FeatureClass = pFeatClass
      pFeatLayer.Name = pDataset.Name & " <REDEV FEASIBILITY>"
      mapCurrent.AddLayer(pFeatLayer)
    End If
    RedevFeasibilityLegend(pFeatLayer, dblTargetRent)

CleanUp:
    dblAppreciation = Nothing
    dblTargetRent = Nothing
    intFldLandValue = Nothing
    intFldImprovmentValue = Nothing
    intFldSqFtParcel = Nothing
    dblLandValue = Nothing
    dblImprovmentValue = Nothing
    dblSqFtParcel = Nothing
    pFeatSelection = Nothing
    pCursor = Nothing
    dblSlope = Nothing
    dblYIntercept = Nothing
    pQueryFilter = Nothing
    strQString = Nothing
    pFeatureCursor = Nothing
    pFeat = Nothing
    dblYValue = Nothing
    intWriteFld = Nothing
    pTable = Nothing
    mxApplication = Nothing
    pMxDocument = Nothing
    mapCurrent = Nothing
    pActiveView = Nothing
    intLayer = Nothing
    pFeatLayer = Nothing
    pDataset = Nothing
    blnLayerFound = Nothing
    strPaintName = Nothing
    intTotalCount = Nothing
    intCount = Nothing
    dblRentTarget = Nothing
    GC.Collect()
    GC.WaitForPendingFinalizers()
    'MessageBox.Show(me,"Development Feasibility calculations have completed.", "Calcualtions Finalized", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
  End Sub

  Private Sub RedevFeasibilityLegend(ByVal pRedevFeasLyr As IFeatureLayer, ByVal dblTargetRent As Double)
    'CREATE A CLASS BREAKS LEGEND TO THE DEVELOPMENT FEASIBILITY COPY OF THE DEV TYPE PAINT LAYER
    Dim mxApplication As IMxApplication
    Dim pDoc As IMxDocument
    Dim pMap As IMap
    Dim strFieldName As String = ""
    Dim pRender As IClassBreaksRenderer
    Dim symd As ISimpleFillSymbol
    Dim intCount As Integer = 0
    Dim intRed As Integer = 0
    Dim intGreen As Integer = 0
    Dim intBlue As Integer = 0
    Dim pNewColor As IRgbColor
    Dim symx As ISimpleFillSymbol
    Dim pLyr As IGeoFeatureLayer
    Dim pOutlineSymbol As ILineSymbol
    Dim pLayerAffects As ILayerEffects
    Dim intCurrentYr As Integer = CInt(Date.Now.Year)

    Dim aoiClassifyGen As IClassifyGEN
    Dim aoiTableHistogram As ITableHistogram
    Dim aoiHistogram As IHistogram
    Dim vValues As Object = Nothing
    Dim vFrequenzy As Object = Nothing
    Dim aoiTable As ITable
    Dim intClass As Integer
    Dim intTopClass As Integer = 0


    'SET THE LEGEND FIELD
    strFieldName = "REDEV_FEASIBILITY"

    'LOOK FOR THE DEV_TYPE FIELD FIRST...ADD IF MISSING
    If pRedevFeasLyr.FeatureClass.FindField(strFieldName) <= -1 Then
      GoTo CleanUp
    Else
      aoiTable = CType(pRedevFeasLyr.FeatureClass, ITable)
    End If

    aoiTableHistogram = New TableHistogram
    With aoiTableHistogram
      .Table = aoiTable
      .Field = strFieldName
    End With

    aoiHistogram = aoiTableHistogram
    aoiHistogram.GetHistogram(vValues, vFrequenzy)
    aoiClassifyGen = New Quantile  'EqualInterval  'NaturalBreaks  'DefinedInterval
    aoiClassifyGen.Classify(vValues, vFrequenzy, 10)

    'DEFINE SYMBOL OUTLINE
    pOutlineSymbol = New SimpleLineSymbol
    pOutlineSymbol.Width = 0.0
    pNewColor = New RgbColor
    pNewColor.Red = 0
    pNewColor.Blue = 0
    pNewColor.Green = 0
    pOutlineSymbol.Color = pNewColor

    Try
      mxApplication = CType(m_appEnvision, IMxApplication)
      pDoc = CType(m_appEnvision.Document, IMxDocument)
      pMap = pDoc.FocusMap

      ' -------------------------------------------------------------------------- 
      ' Renderer Properties: 
      ' make new class breaks renderer with 5 classes 
      pRender = New ClassBreaksRenderer
      pRender.Field = strFieldName
      pRender.BreakCount = 10

      'DETERMINE THE TOP END OF THE CLASSICIAFICATION IN CASE THERE ARE NOT 10 CLASSES
      pRender.MinimumBreak = aoiClassifyGen.ClassBreaks(0)
      For intClass = 1 To 10
        Try
          pRender.Break((intClass - 1)) = aoiClassifyGen.ClassBreaks(intClass)
        Catch ex As Exception
          Exit For
        End Try
      Next

      intTopClass = intClass - 1
      If intTopClass < 10 Then
        pRender.BreakCount = intTopClass
      End If

      ' set class breaks 
      'pRender.Break(1) = aoiClassifyGen.ClassBreaks(2)
      'pRender.Break(2) = aoiClassifyGen.ClassBreaks(3)
      'pRender.Break(3) = aoiClassifyGen.ClassBreaks(4)
      'pRender.Break(4) = aoiClassifyGen.ClassBreaks(5)
      'pRender.Break(5) = aoiClassifyGen.ClassBreaks(6)
      'pRender.Break(6) = aoiClassifyGen.ClassBreaks(7)
      'pRender.Break(7) = aoiClassifyGen.ClassBreaks(8)
      'pRender.Break(8) = aoiClassifyGen.ClassBreaks(9)
      'pRender.Break(9) = aoiClassifyGen.ClassBreaks(10)

      'EXIT IF ONLY 1 CLASS
      If intTopClass <= 0 Then
        MessageBox.Show(Me, "No legend classes could be created with the values in the field, " & strFieldName & ".", "Insuffiecent Legend Classes", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1)
        '** Refresh the TOC
        pDoc.ActiveView.ContentsChanged()
        pDoc.UpdateContents()
        '** Draw the map
        pDoc.ActiveView.Refresh()
        GoTo CleanUp
      End If

      'SET CLASS LABELS
      For intClass = 0 To intTopClass - 1
        pRender.Label(intClass) = (aoiClassifyGen.ClassBreaks(intClass).ToString & " - " & aoiClassifyGen.ClassBreaks(intClass + 1).ToString)
      Next
      'pRender.Label(0) = (aoiClassifyGen.ClassBreaks(0).ToString & " - " & aoiClassifyGen.ClassBreaks(1).ToString)
      'pRender.Label(1) = (aoiClassifyGen.ClassBreaks(1).ToString & " - " & aoiClassifyGen.ClassBreaks(2).ToString)
      'pRender.Label(2) = (aoiClassifyGen.ClassBreaks(2).ToString & " - " & aoiClassifyGen.ClassBreaks(3).ToString)
      'pRender.Label(3) = (aoiClassifyGen.ClassBreaks(3).ToString & " - " & aoiClassifyGen.ClassBreaks(4).ToString)
      'pRender.Label(4) = (aoiClassifyGen.ClassBreaks(4).ToString & " - " & aoiClassifyGen.ClassBreaks(5).ToString)
      'pRender.Label(5) = (aoiClassifyGen.ClassBreaks(5).ToString & " - " & aoiClassifyGen.ClassBreaks(6).ToString)
      'pRender.Label(6) = (aoiClassifyGen.ClassBreaks(6).ToString & " - " & aoiClassifyGen.ClassBreaks(7).ToString)
      'pRender.Label(7) = (aoiClassifyGen.ClassBreaks(7).ToString & " - " & aoiClassifyGen.ClassBreaks(8).ToString)
      'pRender.Label(8) = (aoiClassifyGen.ClassBreaks(8).ToString & " - " & aoiClassifyGen.ClassBreaks(9).ToString)
      'pRender.Label(9) = (aoiClassifyGen.ClassBreaks(9).ToString & " - " & aoiClassifyGen.ClassBreaks(10).ToString)

      m_strProcessing = m_strProcessing & "Erase Color" & vbNewLine
      pNewColor = New RgbColor
      pNewColor.Red = 255
      pNewColor.Green = 255
      pNewColor.Blue = 255
      pOutlineSymbol.Color = pNewColor
      pOutlineSymbol.Width = 0.0
      'BACKGROUND HOLLOW SYMBOL
      symd = New SimpleFillSymbol
      symd.Style = esriSimpleFillStyle.esriSFSHollow
      symd.Outline = pOutlineSymbol

      '** Make the renderer
      '** These properties should be set prior to adding values
      '------------------------------------------------------------------------------------------
      'SYMBOL FOR BREAK 1
      If intTopClass >= 1 Then
        pNewColor = New RgbColor
        pNewColor.Red = 0
        pNewColor.Green = 97
        pNewColor.Blue = 0

        symx = New SimpleFillSymbol
        symx.Style = esriSimpleFillStyle.esriSFSSolid
        symx.Outline = pOutlineSymbol
        symx.Color = pNewColor
        pRender.Symbol(0) = symx
      End If
      '------------------------------------------------------------------------------------------
      'SYMBOL FOR BREAK 2 
      If intTopClass >= 2 Then
        pNewColor = New RgbColor
        pNewColor.Red = 60
        pNewColor.Green = 128
        pNewColor.Blue = 0

        symx = New SimpleFillSymbol
        symx.Style = esriSimpleFillStyle.esriSFSSolid
        symx.Outline = pOutlineSymbol
        symx.Color = pNewColor
        pRender.Symbol(1) = symx
      End If
      '------------------------------------------------------------------------------------------
      'SYMBOL FOR BREAK 3 
      If intTopClass >= 3 Then
        pNewColor = New RgbColor
        pNewColor.Red = 107
        pNewColor.Green = 161
        pNewColor.Blue = 0

        symx = New SimpleFillSymbol
        symx.Style = esriSimpleFillStyle.esriSFSSolid
        symx.Outline = pOutlineSymbol
        symx.Color = pNewColor
        pRender.Symbol(2) = symx
      End If
      '------------------------------------------------------------------------------------------
      'SYMBOL FOR BREAK 4 
      If intTopClass >= 4 Then
        pNewColor = New RgbColor
        pNewColor.Red = 164
        pNewColor.Green = 196
        pNewColor.Blue = 0

        symx = New SimpleFillSymbol
        symx.Style = esriSimpleFillStyle.esriSFSSolid
        symx.Outline = pOutlineSymbol
        symx.Color = pNewColor
        pRender.Symbol(3) = symx
      End If
      '------------------------------------------------------------------------------------------
      'SYMBOL FOR BREAK 5 
      If intTopClass >= 5 Then
        pNewColor = New RgbColor
        pNewColor.Red = 223
        pNewColor.Green = 235
        pNewColor.Blue = 0

        symx = New SimpleFillSymbol
        symx.Style = esriSimpleFillStyle.esriSFSSolid
        symx.Outline = pOutlineSymbol
        symx.Color = pNewColor
        pRender.Symbol(4) = symx
      End If
      '------------------------------------------------------------------------------------------
      'SYMBOL FOR BREAK 6
      If intTopClass >= 6 Then
        pNewColor = New RgbColor
        pNewColor.Red = 255
        pNewColor.Green = 234
        pNewColor.Blue = 0

        symx = New SimpleFillSymbol
        symx.Style = esriSimpleFillStyle.esriSFSSolid
        symx.Outline = pOutlineSymbol
        symx.Color = pNewColor
        pRender.Symbol(5) = symx
      End If
      '------------------------------------------------------------------------------------------
      'SYMBOL FOR BREAK 7 
      If intTopClass >= 7 Then
        pNewColor = New RgbColor
        pNewColor.Red = 255
        pNewColor.Green = 187
        pNewColor.Blue = 0

        symx = New SimpleFillSymbol
        symx.Style = esriSimpleFillStyle.esriSFSSolid
        symx.Outline = pOutlineSymbol
        symx.Color = pNewColor
        pRender.Symbol(6) = symx
      End If
      '------------------------------------------------------------------------------------------
      'SYMBOL FOR BREAK 8 
      If intTopClass >= 8 Then
        pNewColor = New RgbColor
        pNewColor.Red = 255
        pNewColor.Green = 145
        pNewColor.Blue = 0

        symx = New SimpleFillSymbol
        symx.Style = esriSimpleFillStyle.esriSFSSolid
        symx.Outline = pOutlineSymbol
        symx.Color = pNewColor
        pRender.Symbol(7) = symx
      End If
      '------------------------------------------------------------------------------------------
      'SYMBOL FOR BREAK 9 
      If intTopClass >= 9 Then
        pNewColor = New RgbColor
        pNewColor.Red = 255
        pNewColor.Green = 98
        pNewColor.Blue = 0

        symx = New SimpleFillSymbol
        symx.Style = esriSimpleFillStyle.esriSFSSolid
        symx.Outline = pOutlineSymbol
        symx.Color = pNewColor
        pRender.Symbol(8) = symx
      End If
      '------------------------------------------------------------------------------------------
      'SYMBOL FOR BREAK 10 
      If intTopClass >= 10 Then
        pNewColor = New RgbColor
        pNewColor.Red = 255
        pNewColor.Green = 34
        pNewColor.Blue = 0

        symx = New SimpleFillSymbol
        symx.Style = esriSimpleFillStyle.esriSFSSolid
        symx.Outline = pOutlineSymbol
        symx.Color = pNewColor
        pRender.Symbol(9) = symx
      End If
      '------------------------------------------------------------------------------------------

      pLyr = pRedevFeasLyr
      pLyr.Renderer = pRender
      pLyr.DisplayField = strFieldName

      '** Set Layer Transparency
      pLayerAffects = pRedevFeasLyr
      pLayerAffects.Transparency = 35

      '** Refresh the TOC
      pDoc.ActiveView.ContentsChanged()
      pDoc.UpdateContents()
      '** Draw the map
      pDoc.ActiveView.Refresh()

      GoTo CleanUp

    Catch ex As Exception
      MessageBox.Show(Me, ex.Message, "Legend Renderer Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
      GoTo CleanUp
    End Try


CleanUp:
    mxApplication = Nothing
    pDoc = Nothing
    pMap = Nothing
    strFieldName = Nothing
    pRender = Nothing
    symd = Nothing
    intCount = Nothing
    intRed = Nothing
    intGreen = Nothing
    intBlue = Nothing
    pNewColor = Nothing
    symx = Nothing
    pLyr = Nothing
    pOutlineSymbol = Nothing
    pLayerAffects = Nothing
    GC.WaitForPendingFinalizers()
    GC.Collect()
  End Sub

  Private Sub cmbDevTypes_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbDevTypes.SelectedIndexChanged
    'CHANGE THE TARGET TEXT TO REFLECT THE VALUE IS EITHER A SALE OR RENT PRICE
    Dim strProductType As String = ""

    Try
      strProductType = lstClassType.Item(m_frmDevFeasibility.cmbDevTypes.SelectedIndex)
      'DEFAULT TEXT FOR ALL OTHER PRODUCT TYPES
      m_frmDevFeasibility.lblTarget.Text = "Target NNN Lease Rate : ($/sf)"

      If strProductType = "Residential Owner" Then

        m_frmDevFeasibility.lblTarget.Text = "Target Achievable Sales Price: ($/sf)"

      ElseIf strProductType = "Residential Rental" Then

        m_frmDevFeasibility.lblTarget.Text = "Target Achievable Rental Rate : ($/sf)"

      ElseIf strProductType = "Hotel" Then

        m_frmDevFeasibility.lblTarget.Text = "Target Achievable Nightly Rate : ($/sf)"

      End If
    Catch ex As Exception

    End Try
  End Sub

End Class