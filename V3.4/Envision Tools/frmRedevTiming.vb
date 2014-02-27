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

Public Class frmRedevTiming

    Private m_PolygonLyrs As ArrayList
    Private m_ParcelLayer As IFeatureLayer
    Dim blnOpenForm As Boolean = True

    Private Sub frmRedevTiming_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim pUID As New UID
        Try
            m_frmRedevTiming = Nothing
            m_PolygonLyrs = Nothing
            m_ParcelLayer = Nothing
            blnOpenForm = False
            GC.WaitForPendingFinalizers()
            GC.Collect()

            Me.WindowState = FormWindowState.Minimized

            pUID.Value = "esriCore.SelectTool"
            If Not m_appEnvision Is Nothing Then
                m_appEnvision.CurrentTool = m_appEnvision.Document.CommandBars.Find(pUID)
            End If
            GoTo CleanUp
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Closing Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo CleanUp
        End Try

CleanUp:
        pUID = Nothing
        m_frmEnvisionProjectSetup = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub

    Private Sub frmRedevTiming_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim mxApplication As IMxApplication = Nothing
        Dim pMxDocument As IMxDocument = Nothing
        Dim mapCurrent As Map
        Dim intCount As Integer
        Dim pLyr As ILayer
        Dim pFeatLyr As IFeatureLayer
        Dim intLayer As Integer
        Dim intFeatCount As Integer = 0
        Dim pActiveView As IActiveView = Nothing

        Try
            '********************************************************************
            'Populate the combo boxes with layer information
            '********************************************************************
            If Not TypeOf m_appEnvision Is IApplication Then
                GoTo CleanUp
            End If

            pMxDocument = DirectCast(m_appEnvision.Document, IMxDocument)
            mapCurrent = CType(pMxDocument.FocusMap, Map)
            pActiveView = CType(pMxDocument.FocusMap, IActiveView)
            If mapCurrent.LayerCount = 0 Then
                MessageBox.Show("No polygon feature layers could be found in the current view document.  Please added polygon layers to view before utilizing this tool.", "No Polygon Layers Found", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
                GoTo CleanUp
            End If

            'LOAD CURRENT YEAR
            Me.tbxCurrentYear.Text = Date.Now.ToString("yyyy")

            'BUILD LIST OF AVAILABLE FEATURE CLASSES
            m_PolygonLyrs = New ArrayList
            m_PolygonLyrs.Clear()
            Me.cmbParcelLayers.Items.Clear()
            Me.cmbYearBuilt.Items.Clear()
            Me.cmbImproveValue.Items.Clear()
            Me.cmbLandValue.Items.Clear()
            Me.cmbParcelLayers.Text = ""
            Me.cmbYearBuilt.Text = ""
            Me.cmbImproveValue.Text = ""
            Me.cmbLandValue.Text = ""

            For intLayer = 0 To mapCurrent.LayerCount - 1
                pLyr = CType(mapCurrent.Layer(intLayer), ILayer)
                If TypeOf mapCurrent.Layer(intLayer) Is IFeatureLayer Then
                    pFeatLyr = CType(mapCurrent.Layer(intLayer), IFeatureLayer)
                    If pFeatLyr.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                        Me.cmbParcelLayers.Items.Add(pFeatLyr.Name)
                        m_PolygonLyrs.Add(pFeatLyr)
                        intFeatCount = intFeatCount + 1
                    End If
                    pFeatLyr = Nothing
                End If
            Next
            If intFeatCount <= 0 Then
                MessageBox.Show("No polygon feature layers could be found in the current view document.  Please added polygon layers to view before utilizing this tool.", "No Polygon Layers Found", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
                Me.Close()
                GoTo CleanUp
            End If

            GoTo CleanUp
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Redevelopment Timing Calculator Form Opening Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            Me.Close()
            GoTo CleanUp
        End Try
        GoTo CleanUp

CleanUp:
        mxApplication = Nothing
        pMxDocument = Nothing
        mapCurrent = Nothing
        intCount = Nothing
        pLyr = Nothing
        pFeatLyr = Nothing
        intLayer = Nothing
        intFeatCount = Nothing
        pActiveView = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub

    Private Sub cmbParcelLayers_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbParcelLayers.SelectedIndexChanged
        Dim intCount As Integer
        Dim blnLayerFound As Boolean = False
        Dim pFeatureClass As IFeatureClass
        Dim pField As IField
        Dim intFld As Integer

        If Not blnOpenForm Then
            GoTo CleanUp
        End If
        'LOAD THE FIELDS FROM THE SELECTED PARCEL LAYER
        If Me.cmbParcelLayers.Text.Length <= 0 Then
            GoTo CleanUp
        Else
            'MAKE SURE THE SELECTED LAYER IS VALID
            If Me.cmbParcelLayers.Items.Contains(Me.cmbParcelLayers.Text) Then
                blnLayerFound = True
            End If
        End If

        If m_PolygonLyrs.Count > 0 And blnLayerFound Then
            Try
                m_ParcelLayer = CType(m_PolygonLyrs.Item(Me.cmbParcelLayers.SelectedIndex), IFeatureLayer)
                pFeatureClass = m_ParcelLayer.FeatureClass
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Return Parcel Layer Fields Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
                GoTo CleanUp
            End Try

            Me.cmbYearBuilt.Items.Clear()
            Me.cmbImproveValue.Items.Clear()
            Me.cmbLandValue.Items.Clear()
            Me.cmbYearBuilt.Text = ""
            Me.cmbImproveValue.Text = ""
            Me.cmbLandValue.Text = ""
            For intFld = 0 To pFeatureClass.Fields.FieldCount - 1
                pField = pFeatureClass.Fields.Field(intFld)
                If pField.Type = esriFieldType.esriFieldTypeInteger Or pField.Type = esriFieldType.esriFieldTypeDouble Or pField.Type = esriFieldType.esriFieldTypeSmallInteger Or pField.Type = esriFieldType.esriFieldTypeSingle Then
                    If Not UCase(pField.Name) = "OBJECTID" And Not UCase(pField.Name) = "SHAPE_LENGTH" And Not UCase(pField.Name) = "SHAPE_AREA" Then
                        Me.cmbYearBuilt.Items.Add(pField.Name)
                        Me.cmbImproveValue.Items.Add(pField.Name)
                        Me.cmbLandValue.Items.Add(pField.Name)
                        Me.cmbSqFt.Items.Add(pField.Name)
                    End If
                End If
            Next
        End If

CleanUp:
        intCount = Nothing
        blnLayerFound = Nothing
        pFeatureClass = Nothing
        pField = Nothing
        intFld = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub

    Private Sub cmbYearBuilt_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbYearBuilt.SelectedIndexChanged
        If Not blnOpenForm Then
            GoTo CleanUp
        End If
CleanUp:
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub

    Private Sub cmbImproveValue_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbImproveValue.SelectedIndexChanged
        If Not blnOpenForm Then
            GoTo CleanUp
        End If
CleanUp:
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub

    Private Sub cmbLandValue_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbLandValue.SelectedIndexChanged
        If Not blnOpenForm Then
            GoTo CleanUp
        End If

CleanUp:
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub

    Private Sub tbxYearBuilt_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbxCurrentYear.TextChanged
        'MAKE SURE WHAT USER ENTER IS A 4 DIGIT NUMERIC YEAR VALUE
        If Me.tbxCurrentYear.Text.Length = 4 Then
            If IsNumeric(Me.tbxCurrentYear.Text) Then
                Me.tbxCurrentYear.ForeColor = Color.Black
            Else
                Me.tbxCurrentYear.ForeColor = Color.Red
            End If
        Else
            Me.tbxCurrentYear.ForeColor = Color.Red
        End If
    End Sub

    Private Sub tbxLifespan_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbxLifespan.TextChanged
        'MAKE SURE WHAT USER ENTER IS NUMERIC YEAR VALUE
        If Me.tbxLifespan.Text.Length > 0 Then
            If IsNumeric(Me.tbxLifespan.Text) Then
                Me.tbxLifespan.ForeColor = Color.Black
            Else
                Me.tbxLifespan.ForeColor = Color.Red
            End If
        Else
            Me.tbxLifespan.ForeColor = Color.Red
        End If
    End Sub

    Private Sub tbxAppreciation_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbxAppreciation.TextChanged
        'MAKE SURE WHAT USER ENTER IS NUMERIC YEAR VALUE
        If Me.tbxAppreciation.Text.Length > 0 Then
            If IsNumeric(Me.tbxLifespan.Text) Then
                Me.tbxAppreciation.ForeColor = Color.Black
            Else
                Me.tbxAppreciation.ForeColor = Color.Red
            End If
        Else
            Me.tbxAppreciation.ForeColor = Color.Red
        End If
    End Sub

    Private Sub btnRun_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRun.Click
        'RUN THE BOTTOM QUARTILE OR DEPRECIATION SUB
        If Me.rdbDepreciation.Checked Then
            RunDepreciation()
        Else
            RunBottomQuartile()
        End If
    End Sub

    Private Sub RunDepreciation()
        'RETRIEVE VALUES AND CALCULATE THE REDEVE_YR FIELD
        Dim pFeatClass As IFeatureClass
        Dim intYearBuiltFld As Integer = -1
        Dim intImprovementValueFld As Integer = -1
        Dim intLandValueFld As Integer = -1
        Dim intOutputFld As Integer = -1
        Dim pTable As ITable
        Dim strEquation As String = ""

        'EQUATION VARIABLES
        Dim intEffectiveYearBuilt As Integer = 0
        Dim dblImprovmentValue As Double = 0
        Dim dblLandValue As Double = 0
        Dim intCurrentYear As Integer = 0
        Dim intLifespan As Integer = 0
        Dim dblLandAppreciation As Double = 0
        Dim intPlanningHorizon As Integer = 0
        Dim dlbEquationResults As Double = 0

        Dim pQFilter As IQueryFilter
        Dim pFeatureCursor As IFeatureCursor
        Dim intCount As Integer = 0
        Dim intTotalCount As Integer = 0
        Dim pFeat As IFeature
        Dim intDateYear As Integer = CInt(Date.Now.ToString("yyyy"))
        Dim intFailureCount As Integer = 0
        Dim strFailureMsg As String = ""

        If m_ParcelLayer Is Nothing Then
            MessageBox.Show("Please select a parcel layer.", "No Parcel Layer Selected", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo CleanUp
        Else
            pFeatClass = m_ParcelLayer.FeatureClass
            pTable = CType(pFeatClass, Table)
        End If

        'CHECK FOR INPUT FIELDS
        If Me.cmbYearBuilt.Text.Length <= 0 Then
            MessageBox.Show("Please select a Year Built field.", "Field Required", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo CleanUp
        Else
            intYearBuiltFld = pFeatClass.FindField(Me.cmbYearBuilt.Text)
        End If
        If Me.cmbImproveValue.Text.Length <= 0 Then
            MessageBox.Show("Please select a Improvement or Building Value field.", "Field Required", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo CleanUp
        Else
            intImprovementValueFld = pFeatClass.FindField(Me.cmbImproveValue.Text)
        End If
        If Me.cmbLandValue.Text.Length <= 0 Then
            MessageBox.Show("Please select a Land Value field.", "Field Required", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo CleanUp
        Else
            intLandValueFld = pFeatClass.FindField(Me.cmbLandValue.Text)
        End If

        'RETREIVE USER INPUTS
        If Me.tbxCurrentYear.Text.Length = 4 Then
            If IsNumeric(Me.tbxCurrentYear.Text) Then
                intCurrentYear = CInt(Me.tbxCurrentYear.Text)
            Else
                MessageBox.Show("Please enter a valid Current Year value.", "Invalid User Input", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
                GoTo CleanUp
            End If
        Else
            MessageBox.Show("Please enter a valid Current Year value.", "Invalid User Input", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo CleanUp
        End If

        If IsNumeric(Me.tbxLifespan.Text) Then
            intLifespan = CInt(Me.tbxLifespan.Text)
        Else
            MessageBox.Show("Please enter a valid Building Lifespan value.", "Invalid User Input", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo CleanUp
        End If

        If IsNumeric(Me.tbxAppreciation.Text) Then
            dblLandAppreciation = CDbl(Me.tbxAppreciation.Text)
        Else
            MessageBox.Show("Please enter a valid Annual Land Appreciation value.", "Invalid User Input", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo CleanUp
        End If

        'CHECK FOR THE FIELD TO BE CALCULATED
        If pTable.FindField("REDEV_YR") <= -1 Then
            If Not AddEnvisionField(pTable, "REDEV_YR", "DOUBLE", 16, 6) Then
                GoTo CleanUp
            End If
        End If
        intOutputFld = pTable.FindField("REDEV_YR")

        'CYCLE THROUGH EACH RECORD TO RETRIEVE VARIABLES AND EVALUATE FOR EQUATION CALC
        Me.barProgress.Visible = True
        pFeatClass = m_ParcelLayer.FeatureClass
        pQFilter = New QueryFilter
        pFeatureCursor = pFeatClass.Search(Nothing, False)
        intTotalCount = pFeatClass.FeatureCount(Nothing)
        pFeat = pFeatureCursor.NextFeature
        Do While Not pFeat Is Nothing
            intCount = intCount + 1
            Me.barProgress.Value = ((intCount / intTotalCount) * 100)
            intEffectiveYearBuilt = 0
            Try
                intEffectiveYearBuilt = pFeat.Value(intYearBuiltFld)
            Catch ex As Exception
            End Try
            dblImprovmentValue = 0
            Try
                dblImprovmentValue = pFeat.Value(intImprovementValueFld)
            Catch ex As Exception
            End Try
            dblLandValue = 0
            Try
                dblLandValue = pFeat.Value(intLandValueFld)
            Catch ex As Exception
            End Try

            'CALCULATE THE EQUATION
            Try
                dlbEquationResults = 0
                If dblLandValue > dblImprovmentValue Then
                    dlbEquationResults = intCurrentYear
                Else
                    If (intCurrentYear - intEffectiveYearBuilt) > intLifespan Then
                        dlbEquationResults = 0
                    Else
                        If (intCurrentYear - intEffectiveYearBuilt) = intLifespan Then
                            dlbEquationResults = intCurrentYear
                        Else
                            'dlbEquationResults = ((dblImprovmentValue - dblLandValue) / ((dblImprovmentValue / intLifespan) - (intCurrentYear - intEffectiveYearBuilt) + (dblLandValue * (dblLandAppreciation / 100)))) + intCurrentYear
                            dlbEquationResults = ((dblImprovmentValue - dblLandValue) / ((dblImprovmentValue / (intLifespan - (intCurrentYear - intEffectiveYearBuilt))) + (dblLandValue * (dblLandAppreciation / 100)))) + intCurrentYear
                        End If
                    End If
                End If
            Catch ex As Exception
            End Try
            'WRITE EQUATION OUTPUT TO CURRENT FEATURE
            Try
                pFeat.Value(intOutputFld) = dlbEquationResults
                pFeat.Store()
            Catch ex As Exception
                intFailureCount = intFailureCount + 1
                strFailureMsg = ex.Message
            End Try
            pFeat = pFeatureCursor.NextFeature
        Loop

        'APPLY THE APPROPRIATE LEGEND
        If Me.rdbDepreciation.Checked Then
            DepreciationLegend()
        Else
            BottomQuartileLegend()
        End If
        GoTo CleanUp

CleanUp:
        If intFailureCount >= 1 Then
            MessageBox.Show(CStr(intFailureCount) & " features failed to have a value writen to a file.  Last error message was: " & vbNewLine & strFailureMsg, "Process Complete: Error(s) Found", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
        End If
        Me.barProgress.Visible = False
        Me.barProgress.Value = 0
        pFeatClass = Nothing
        intYearBuiltFld = Nothing
        intImprovementValueFld = Nothing
        intLandValueFld = Nothing
        intOutputFld = Nothing
        pTable = Nothing
        strEquation = Nothing
        intEffectiveYearBuilt = Nothing
        dblImprovmentValue = Nothing
        dblLandValue = Nothing
        intCurrentYear = Nothing
        intLifespan = Nothing
        dblLandAppreciation = Nothing
        intPlanningHorizon = Nothing
        dlbEquationResults = Nothing
        pQFilter = Nothing
        pFeatureCursor = Nothing
        intCount = Nothing
        intTotalCount = Nothing
        pFeat = Nothing
        intDateYear = Nothing
        'Me.Close()
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub

    Private Sub RunBottomQuartile()
        'RETRIEVE VALUES AND CALCULATE THE REDEVE_YR FIELD
        Dim pFeatClass As IFeatureClass
        Dim intImprovementValueFld As Integer = -1
        Dim intLandValueFld As Integer = -1
        Dim intSqFtValueFld As Integer = -1
        Dim intOutputFld As Integer = -1
        Dim pTable As ITable
        Dim strEquation As String = ""

        'EQUATION VARIABLES
        Dim dblSqFt As Double = 0
        Dim dblImprovmentValue As Double = 0
        Dim dblLandValue As Double = 0
        Dim dlbEquationResults As Double = 0

        Dim pQFilter As IQueryFilter
        Dim pFeatureCursor As IFeatureCursor
        Dim intCount As Integer = 0
        Dim intTotalCount As Integer = 0
        Dim pFeat As IFeature
        Dim intDateYear As Integer = CInt(Date.Now.ToString("yyyy"))
        Dim intFailureCount As Integer = 0
        Dim strFailureMsg As String = ""

        If m_ParcelLayer Is Nothing Then
            MessageBox.Show("Please select a parcel layer.", "No Parcel Layer Selected", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo CleanUp
        Else
            pFeatClass = m_ParcelLayer.FeatureClass
            pTable = CType(pFeatClass, Table)
        End If

        'CHECK FOR INPUT FIELDS
        If Me.cmbImproveValue.Text.Length <= 0 Then
            MessageBox.Show("Please select a Improvement or Building Value field.", "Field Required", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo CleanUp
        Else
            intImprovementValueFld = pFeatClass.FindField(Me.cmbImproveValue.Text)
        End If
        If Me.cmbLandValue.Text.Length <= 0 Then
            MessageBox.Show("Please select a Land Value field.", "Field Required", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo CleanUp
        Else
            intLandValueFld = pFeatClass.FindField(Me.cmbLandValue.Text)
        End If
        If Me.cmbSqFt.Text.Length <= 0 Then
            MessageBox.Show("Please select a Square Footage Value field.", "Field Required", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo CleanUp
        Else
            intSqFtValueFld = pFeatClass.FindField(Me.cmbSqFt.Text)
        End If

        'CHECK FOR THE FIELD TO BE CALCULATED
        If pTable.FindField("Val_per_sqft") <= -1 Then
            If Not AddEnvisionField(pTable, "Val_per_sqft", "DOUBLE", 16, 6) Then
                GoTo CleanUp
            End If
        End If
        intOutputFld = pTable.FindField("Val_per_sqft")

        'CYCLE THROUGH EACH RECORD TO RETRIEVE VARIABLES AND EVALUATE FOR EQUATION CALC
        Me.barProgress.Visible = True
        pFeatClass = m_ParcelLayer.FeatureClass
        pQFilter = New QueryFilter
        pFeatureCursor = pFeatClass.Search(Nothing, False)
        intTotalCount = pFeatClass.FeatureCount(Nothing)
        pFeat = pFeatureCursor.NextFeature
        Do While Not pFeat Is Nothing
            intCount = intCount + 1
            Me.barProgress.Value = ((intCount / intTotalCount) * 100)
            dblImprovmentValue = 0
            Try
                dblImprovmentValue = CDbl(pFeat.Value(intImprovementValueFld))
            Catch ex As Exception
            End Try
            dblLandValue = 0
            Try
                dblLandValue = CDbl(pFeat.Value(intLandValueFld))
            Catch ex As Exception
            End Try
            dblSqFt = 1
            Try
                dblSqFt = CDbl(pFeat.Value(intSqFtValueFld))
            Catch ex As Exception
            End Try

            'CALCULATE THE EQUATION
            Try
                dlbEquationResults = 0
                dlbEquationResults = (dblImprovmentValue + dblLandValue) / dblSqFt
            Catch ex As Exception
            End Try
            'WRITE EQUATION OUTPUT TO CURRENT FEATURE
            Try
                pFeat.Value(intOutputFld) = dlbEquationResults
                pFeat.Store()
            Catch ex As Exception
                intFailureCount = intFailureCount + 1
                strFailureMsg = ex.Message
            End Try
            pFeat = pFeatureCursor.NextFeature
        Loop

        'APPLY THE APPROPRIATE LEGEND
        If Me.rdbDepreciation.Checked Then
            DepreciationLegend()
        Else
            BottomQuartileLegend()
        End If
        GoTo CleanUp

CleanUp:
        If intFailureCount >= 1 Then
            MessageBox.Show(CStr(intFailureCount) & " features failed to have a value writen to a file.  Last error message was: " & vbNewLine & strFailureMsg, "Process Complete: Error(s) Found", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
        End If
        Me.barProgress.Visible = False
        Me.barProgress.Value = 0
        pFeatClass = Nothing
        intImprovementValueFld = Nothing
        intLandValueFld = Nothing
        intOutputFld = Nothing
        pTable = Nothing
        strEquation = Nothing
        dblImprovmentValue = Nothing
        dblLandValue = Nothing
        dlbEquationResults = Nothing
        pQFilter = Nothing
        pFeatureCursor = Nothing
        intCount = Nothing
        intTotalCount = Nothing
        pFeat = Nothing
        intDateYear = Nothing
        'Me.Close()
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub

    Private Sub btnApplyLegend_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnApplyLegend.Click
        'APPLY THE APPROPRIATE LEGEND
        If Me.rdbDepreciation.Checked Then
            DepreciationLegend()
        Else
            BottomQuartileLegend()
        End If
    End Sub

    Private Sub DepreciationLegend()
        'APPLY THE REDEVELOPMENT CANDIDATE TO THE SELECTED PARCEL LAYER
        If m_ParcelLayer Is Nothing Then
            Exit Sub
        End If

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
        Dim pTable As ITable
        Dim intCurrentYr As Integer = CInt(Date.Now.Year)


        'SET THE LEGEND FIELD
        strFieldName = "REDEV_YR"

        'LOOK FOR THE DEV_TYPE FIELD FIRST...ADD IF MISSING
        If m_ParcelLayer.FeatureClass.FindField("REDEV_YR") <= -1 Then
            pTable = CType(m_ParcelLayer.FeatureClass, ITable)
            If Not AddEnvisionField(pTable, "REDEV_YR", "DOUBLE", 16, 6) Then
                GoTo CleanUp
            End If
        End If

        'EXIT IF FIELD NOT FOUND
        If m_ParcelLayer.FeatureClass.FindField(strFieldName) <= -1 Then
            GoTo CleanUp
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

            ' -------------------------------------------------------------------------- 
            ' Renderer Properties: 
            ' make new class breaks renderer with 5 classes 
            pRender = New ClassBreaksRenderer
            pRender.Field = strFieldName
            pRender.BreakCount = 5

            ' set class breaks 
            pRender.MinimumBreak = 0 ' min value in dataset 
            pRender.Break(0) = 0
            pRender.Break(1) = intCurrentYr '2012
            pRender.Break(2) = intCurrentYr + 5 '2017
            pRender.Break(3) = intCurrentYr + 10 '2022
            pRender.Break(4) = intCurrentYr + 50 '2060

            ' set class labels (not necessary, but these look better than the default). 
            pRender.Label(0) = "0"
            pRender.Label(1) = "Today"
            pRender.Label(2) = "0-5yrs"
            pRender.Label(3) = "5-10yrs"
            pRender.Label(4) = "10+ yrs"

            m_strProcessing = m_strProcessing & "Erase Color" & vbNewLine
            pNewColor = New RgbColor
            pNewColor.Red = 255
            pNewColor.Blue = 255
            pNewColor.Green = 255
            pOutlineSymbol.Color = pNewColor
            pOutlineSymbol.Width = 0.0
            'BACKGROUND HOLLOW SYMBOL
            symd = New SimpleFillSymbol
            symd.Style = esriSimpleFillStyle.esriSFSHollow
            symd.Outline = pOutlineSymbol

            '** Make the renderer
            '** These properties should be set prior to adding values
            m_strProcessing = m_strProcessing & "Construct Renderer" & vbNewLine
            'pRender.BackgroundSymbol = symd
            'DEFINE MINIMUM BREAK VALUES
            'DEFINE LABELS
            'DEFINE SYMBOLS
            '------------------------------------------------------------------------------------------
            'SYMBOL FOR BREAK 1 - 0
            pNewColor = New RgbColor
            pNewColor.Red = 255
            pNewColor.Blue = 255
            pNewColor.Green = 255

            symx = New SimpleFillSymbol
            symx.Style = esriSimpleFillStyle.esriSFSHollow
            symx.Outline = pOutlineSymbol
            symx.Color = pNewColor
            pRender.Symbol(0) = symx
            '------------------------------------------------------------------------------------------
            'SYMBOL FOR BREAK 2 - Today
            pNewColor = New RgbColor
            pNewColor.Red = 168
            pNewColor.Green = 0
            pNewColor.Blue = 0

            symx = New SimpleFillSymbol
            symx.Style = esriSimpleFillStyle.esriSFSSolid
            symx.Outline = pOutlineSymbol
            symx.Color = pNewColor
            pRender.Symbol(1) = symx
            '------------------------------------------------------------------------------------------
            'SYMBOL FOR BREAK 3 - 0-5years
            pNewColor = New RgbColor
            pNewColor.Red = 230
            pNewColor.Green = 76
            pNewColor.Blue = 0

            symx = New SimpleFillSymbol
            symx.Style = esriSimpleFillStyle.esriSFSSolid
            symx.Outline = pOutlineSymbol
            symx.Color = pNewColor
            pRender.Symbol(2) = symx
            '------------------------------------------------------------------------------------------
            'SYMBOL FOR BREAK 4 - 5-10years
            pNewColor = New RgbColor
            pNewColor.Red = 255
            pNewColor.Green = 170
            pNewColor.Blue = 0

            symx = New SimpleFillSymbol
            symx.Style = esriSimpleFillStyle.esriSFSSolid
            symx.Outline = pOutlineSymbol
            symx.Color = pNewColor
            pRender.Symbol(3) = symx
            '------------------------------------------------------------------------------------------
            'SYMBOL FOR BREAK 5 10+years
            pNewColor = New RgbColor
            pNewColor.Red = 255
            pNewColor.Green = 235
            pNewColor.Blue = 175

            symx = New SimpleFillSymbol
            symx.Style = esriSimpleFillStyle.esriSFSSolid
            symx.Outline = pOutlineSymbol
            symx.Color = pNewColor
            pRender.Symbol(4) = symx
            '------------------------------------------------------------------------------------------

            pLyr = m_ParcelLayer
            pLyr.Renderer = pRender
            pLyr.DisplayField = strFieldName

            '** Set Layer Transparency
            pLayerAffects = m_ParcelLayer
            pLayerAffects.Transparency = 35

            '** Refresh the TOC
            pDoc.ActiveView.ContentsChanged()
            pDoc.UpdateContents()
            '** Draw the map
            pDoc.ActiveView.Refresh()

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

    Private Sub BottomQuartileLegend()
        'APPLY THE REDEVELOPMENT CANDIDATE TO THE SELECTED PARCEL LAYER
        If m_ParcelLayer Is Nothing Then
            Exit Sub
        End If

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
        Dim pTable As ITable

        Dim arcTable As ITable
        Dim arcTableHistogram As ITableHistogram
        Dim arcHistogram As IHistogram
        Dim varDataValues As Object = Nothing
        Dim varDataFrequencies As Object = Nothing
        Dim arcClassify As IClassify
        Dim Classes() As Double = Nothing
        Dim arcExclussion As IDataExclusion
        Dim intNumBreak As Integer = 4
        Dim pRamp As IColorRamp
        Dim l_oFromColor As IRgbColor
        Dim l_oToColor As IRgbColor
        Dim l_oAlgoRamp As ESRI.ArcGIS.Display.IAlgorithmicColorRamp
        Dim intItem As Integer

        'SET THE LEGEND FIELD
        strFieldName = "Val_per_sqft"

        'SET CLASS BREAKS, DEFAULT TO 4
        If Me.tbxBreaks.Text.Length > 0 Then
            If IsNumeric(Me.tbxBreaks.Text) Then
                intNumBreak = CInt(Me.tbxBreaks.Text)
            End If
        End If

        'LOOK FOR THE DEV_TYPE FIELD FIRST...ADD IF MISSING
        If m_ParcelLayer.FeatureClass.FindField(strFieldName) <= -1 Then
            pTable = CType(m_ParcelLayer.FeatureClass, ITable)
            If Not AddEnvisionField(pTable, strFieldName, "DOUBLE", 16, 6) Then
                GoTo CleanUp
            End If
        End If

        'EXIT IF FIELD NOT FOUND
        If m_ParcelLayer.FeatureClass.FindField(strFieldName) <= -1 Then
            GoTo CleanUp
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

            ' -------------------------------------------------------------------------- 
            ' Renderer Properties: 
            ' make new class breaks renderer with 5 classes 
            pRender = New ClassBreaksRenderer
            pRender.Field = strFieldName
            pRender.BreakCount = intNumBreak

            Try
                arcTable = CType(m_ParcelLayer.FeatureClass, ITable)
                ' Define the table histogram.
                arcTableHistogram = New TableHistogram
                arcTableHistogram.Field = strFieldName
                arcTableHistogram.Table = arcTable
                arcExclussion = CType(pRender, IDataExclusion)
                arcExclussion.ExclusionClause = strFieldName & " = 0"
                arcTableHistogram.Exclusion = arcExclussion
                ' Derive the data values and frequencies from the histogram.
                arcHistogram = arcTableHistogram
                arcHistogram.GetHistogram(varDataValues, varDataFrequencies)
                arcClassify = New Quantile
                ' Prepare a classify object.
                arcClassify.SetHistogramData(varDataValues, varDataFrequencies)
                arcClassify.Classify(intNumBreak)
                ' Create an array of class breaks.
                Classes = arcClassify.ClassBreaks
                'MessageBox.Show(Classes.Length.ToString)
                'For i = 0 To 3
                '    MessageBox.Show(Classes(i).ToString)
                'Next i
            Catch ex As Exception

            End Try

            'Color ramp generation
            l_oFromColor = New RgbColor
            l_oFromColor.Red = 255
            l_oFromColor.Green = 0
            l_oFromColor.Blue = 0
            l_oToColor = New RgbColor
            l_oToColor.Red = 255
            l_oToColor.Green = 255
            l_oToColor.Blue = 190
            l_oAlgoRamp = New AlgorithmicColorRamp
            l_oAlgoRamp.Name = "Temperature"
            l_oAlgoRamp.FromColor = l_oFromColor
            l_oAlgoRamp.ToColor = l_oToColor
            l_oAlgoRamp.Size = intNumBreak
            l_oAlgoRamp.Algorithm = esriColorRampAlgorithm.esriHSVAlgorithm
            l_oAlgoRamp.CreateRamp(True)

            For intItem = 0 To intNumBreak - 1
                'FIRST SET THE CLASS BREAK
                pRender.Break(intItem) = Classes(intItem + 1)
                'SECOND SET THE BREAK LABEL
                If intItem = 0 Then
                    pRender.Label(intItem) = "<= " & CInt(Classes(intItem + 1)).ToString
                ElseIf Not intItem = 0 And Not intItem = (intNumBreak - 1) Then
                    pRender.Label(intItem) = CInt(Classes(intItem + 1)).ToString & " to " & CInt(Classes(intItem + 2)).ToString
                ElseIf intItem = (intNumBreak - 1) Then
                    pRender.Label(intItem) = ">= " & CInt(Classes(intItem + 1)).ToString
                End If
                'THIRD CREATE AND APPLY SYMBOL
                symx = New SimpleFillSymbol
                symx.Style = esriSimpleFillStyle.esriSFSSolid
                symx.Outline = pOutlineSymbol
                symx.Color = l_oAlgoRamp.Color(intItem)
                pRender.Symbol(intItem) = symx
            Next

            pLyr = m_ParcelLayer
            pLyr.Renderer = pRender
            pLyr.DisplayField = strFieldName

            '** Set Layer Transparency
            pLayerAffects = m_ParcelLayer
            pLayerAffects.Transparency = 35

            '** Refresh the TOC
            pDoc.ActiveView.ContentsChanged()
            pDoc.UpdateContents()
            '** Draw the map
            pDoc.ActiveView.Refresh()

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
        intRed = Nothing
        intGreen = Nothing
        intBlue = Nothing
        pNewColor = Nothing
        symx = Nothing
        pLyr = Nothing
        pOutlineSymbol = Nothing
        pLayerAffects = Nothing
        intNumBreak = Nothing
        pRamp = Nothing
        l_oFromColor = Nothing
        l_oToColor = Nothing
        l_oAlgoRamp = Nothing
        intItem = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()

    End Sub

    Private Function GetColorRamp(ByVal sCRName As String) As IColorRamp  'gets a colorramp object
        GetColorRamp = Nothing
        Dim pEnum As IEnumStyleGalleryItem
        Dim pItem As IStyleGalleryItem
        Dim pStyleGallery As IStyleGallery

        pStyleGallery = New StyleGallery
        pEnum = pStyleGallery.Items("Color Ramps", "ESRI.style", "Default Ramps")
        pEnum.Reset()
        pItem = pEnum.Next

        Do While Not pItem Is Nothing
            If pItem.Name = sCRName Then
                GetColorRamp = pItem.Item
                Exit Do
            End If
            pItem = pEnum.Next
        Loop
    End Function

    Private Sub EnabledControls(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdbDepreciation.CheckedChanged, rdbBottomQurtile.CheckedChanged
        'TOGGLE THE ENABLED STATUS OF CONTROLS TO MATCH INPUTS OF SELECTED METHOD
        Me.lblYearBuilt.Enabled = Me.rdbDepreciation.Checked
        Me.cmbYearBuilt.Enabled = Me.rdbDepreciation.Checked
        Me.lblSqFt.Enabled = Me.rdbBottomQurtile.Checked
        Me.cmbSqFt.Enabled = Me.rdbBottomQurtile.Checked
        Me.LblBreaks.Enabled = Me.rdbBottomQurtile.Checked
        Me.tbxBreaks.Enabled = Me.rdbBottomQurtile.Checked
        Me.pnlInputs.Enabled = Me.rdbDepreciation.Checked
    End Sub
End Class