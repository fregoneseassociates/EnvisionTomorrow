Imports System
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Runtime.InteropServices
Imports System.Collections
Imports System.IO
Imports System.Drawing.Printing
Imports System.Data
Imports System.Math

Imports ESRI.ArcGIS.DataManagementTools
Imports ESRI.ArcGIS.Geoprocessor
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.Geodatabase

Imports ESRI.ArcGIS.DataSourcesRaster
'Imports ESRI.ArcGIS.SpatialAnalyst
'Imports ESRI.ArcGIS.SpatialAnalystUI


Public Class frmEnvisionSlope

    Private Sub frmEnvisionSlope_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'CLOSE THE SLOPE/HILLSHADE FORM
        m_frmEnvisionSlope = Nothing
    End Sub

    Private Sub frmEnvisionSlope_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim intCount As Integer
        Dim pLayer As IRasterLayer

        Try
            'CLOSE IF EXTENT IS NOT DEFINED
            If m_pETSetupExtentEnv Is Nothing Then
                m_frmEnvisionSlope.Close()
                GoTo CleanUp
            End If
            'WORKSPACE DIRECTORY
            If m_frmEnvisionProjectSetup.tbxWorkspace.Text = "" Then
                MessageBox.Show("A Envision Project Workspace Directory is required.  Please select the a directory in the tab labeled, *Step 1 - Project Settings.", "Workspace Path Reequired", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
                m_frmEnvisionSlope.Close()
                GoTo CleanUp
            End If

            'LOAD GRID LAYERS AND DEFINED GRID CELL SIZE INTO FORM
            Me.cmbGridLayers.Items.Clear()
            For intCount = 0 To m_arrRasterLayers.Count - 1
                pLayer = CType(m_arrRasterLayers.Item(intCount), IRasterLayer)
                Me.cmbGridLayers.Items.Add(pLayer.Name)
            Next

            If m_intSlopeGridCellSize <= -1 And IsNumeric(m_frmEnvisionProjectSetup.tbxConstraintsCellSize.Text) Then
                Me.tbxCellSize.Text = m_frmEnvisionProjectSetup.tbxConstraintsCellSize.Text
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Slope/Hillshade Form Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo CleanUp
        End Try
CleanUp:
        intCount = Nothing
        pLayer = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub

    Private Sub btnReturnGridCellSize_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReturnGridCellSize.Click
        'RETRIEVE THE GRID CELL SIZE FROM THE SELECTED GRID LAYER
        Dim pRFeatLyr As IRasterLayer
        Dim pRaster2 As IRaster2
        Dim pRasterProps As IRasterProps

        Try
            If Me.cmbGridLayers.SelectedIndex = -1 Then
                GoTo CleanUp
            End If

            pRFeatLyr = CType(m_arrRasterLayers.Item(Me.cmbGridLayers.SelectedIndex), IRasterLayer)
            pRaster2 = CType(pRFeatLyr.Raster, IRaster2)

            pRasterProps = CType(pRaster2, IRasterProps)
            Me.tbxCellSize.Text = CType(pRasterProps.MeanCellSize.X, String)
            GoTo CleanUp
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString, "Grid Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo CleanUp
        End Try

CleanUp:
        pRFeatLyr = Nothing
        pRaster2 = Nothing
        pRasterProps = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub

    Private Sub btnSlopeRun_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSlopeRun.Click
        'CREATE THE SLOPE AND HILLSHADE GRID, AND THE CONSTRAINTS SLOPE FEATURE CLASS
        If Me.cmbGridLayers.SelectedIndex = -1 Then
            MessageBox.Show("Please select a DEM grid to derieve slope and hillshade.", "Select a grid", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            Me.BringToFront()
            GoTo CleanUp
        End If
        If Not IsNumeric(Me.tbxCellSize.Text) Then
            MessageBox.Show("Please enter a numeric output grid cell size.", "Invalid Grid Cell Value", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            Me.BringToFront()
            GoTo CleanUp
        End If
        If m_frmEnvisionProjectSetup.tbxWorkspace.Text = "" Then
            MessageBox.Show("A Envision Project Workspace Directory is required.  Please select the a directory in the tab labeled, *Step 1 - Project Settings, of the Envision Project Setup form.", "Workspace Path Reequired", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            Me.BringToFront()
            GoTo CleanUp
        End If

        Dim mxApplication As IMxApplication
        Dim mxDoc As IMxDocument = Nothing
        Dim pActiveView As IActiveView = Nothing
        Dim pViewEnv As IEnvelope
        Dim pPoint As ESRI.ArcGIS.Geometry.IPoint
        Dim pRLyr As IRasterLayer
        Dim pRaster2 As IRaster2 = Nothing
        Dim pRasterDataset As IRasterDataset
        Dim pGeoDataset As IGeoDataset
        Dim pRasterEnv As IEnvelope
        Dim pixelWidth As Double
        Dim pixelHeight As Double
        Dim intCellSize As Double
        Dim lCol As Double
        Dim lRow As Double
        Dim intNumCol As Double
        Dim intNumRow As Double
        Dim pEnv As Envelope = New Envelope
        Dim objExtent As Object
        Dim objOutputSpatref As Object
        Dim objWorkspace As Object


        Dim pTempRLyr As IRasterLayer
        Dim GP As ESRI.ArcGIS.Geoprocessor.Geoprocessor = New ESRI.ArcGIS.Geoprocessor.Geoprocessor
        Dim pCreateSlope As ESRI.ArcGIS.SpatialAnalystTools.Slope
        Dim pCreateHillshade As ESRI.ArcGIS.SpatialAnalystTools.HillShade
        Dim pCreate15Slope As ESRI.ArcGIS.SpatialAnalystTools.Reclassify
        Dim pCreate15Lyr As ESRI.ArcGIS.ConversionTools.RasterToPolygon
        Dim pCreate25Slope As ESRI.ArcGIS.SpatialAnalystTools.Reclassify
        Dim pCreate25Lyr As ESRI.ArcGIS.ConversionTools.RasterToPolygon
        Me.lblStatus.Visible = True
        Me.lblStatus.Text = "Please Wait, Processing....."
        Me.Refresh()

        Try
            If m_arrRasterLayers.Count > 0 Then
                'CURRENT VIEW DOCUMENT
                mxApplication = CType(m_appEnvision, IMxApplication)
                mxDoc = CType(m_appEnvision.Document, IMxDocument)
                pActiveView = CType(mxDoc.FocusMap, IActiveView)
                pViewEnv = pActiveView.Extent

                'RASTER LAYER VARIABLES
                pRLyr = CType(m_arrRasterLayers.Item(Me.cmbGridLayers.SelectedIndex), IRasterLayer)
                pRaster2 = CType(pRLyr.Raster, IRaster2)
                pRasterDataset = pRaster2.RasterDataset
                pGeoDataset = CType(pRasterDataset, IGeoDataset)
                pRasterEnv = pGeoDataset.Extent
                intCellSize = CDbl(Me.tbxCellSize.Text)
                pPoint = New ESRI.ArcGIS.Geometry.Point
                pPoint.X = m_pETSetupExtentEnv.XMin
                pPoint.Y = m_pETSetupExtentEnv.YMin

                'CALCULATE RASTER ORIGIN / EXTENT
                pixelWidth = pRasterEnv.Width / pRLyr.ColumnCount
                If pixelWidth <= 0 Then
                    pixelWidth = 1
                End If
                pixelHeight = pRasterEnv.Height / pRLyr.RowCount
                If pixelHeight <= 0 Then
                    pixelHeight = 1
                End If
                lCol = CType(Round(Abs(pPoint.X - pRasterEnv.XMin + (0.5 * pixelWidth)) / pixelWidth), Long)
                lRow = CType(Round(Abs(pPoint.Y - pRasterEnv.YMax - (0.5 * pixelHeight)) / pixelHeight), Long)
                intNumRow = CInt(m_pETSetupExtentEnv.Width / intCellSize)
                intNumCol = CInt(m_pETSetupExtentEnv.Height / intCellSize)

                pPoint = New ESRI.ArcGIS.Geometry.Point
                pPoint.X = (pRasterEnv.XMin + (lCol * pixelWidth)) - pixelWidth
                pPoint.Y = (pRasterEnv.YMax - (lRow * pixelHeight))

                pEnv.XMax = CDbl(pPoint.X + (intNumRow * intCellSize))
                pEnv.XMin = CDbl(pPoint.X)
                pEnv.YMax = CDbl(pPoint.Y + (intNumCol * intCellSize))
                pEnv.YMin = CDbl(pPoint.Y)
            End If
            objExtent = pEnv.XMax.ToString + " " + pEnv.YMax.ToString + " " + pEnv.XMin.ToString + " " + pEnv.YMin.ToString
            objOutputSpatref = m_pSpatRefProject.FactoryCode
            objWorkspace = m_frmEnvisionProjectSetup.tbxWorkspace.Text

            Try
                GP.SetEnvironmentValue("cellSize", intCellSize)
                GP.SetEnvironmentValue("extent", objExtent)
                GP.SetEnvironmentValue("workspace", objWorkspace)
                GP.SetEnvironmentValue("outputCoordinateSystem", objOutputSpatref)
                GP.OverwriteOutput = True
                GP.AddOutputsToMap = True
                GP.TemporaryMapLayers = False
            Catch ex As Exception
                MessageBox.Show(ex.Message, "GeoProcessor Settings Error")
                GoTo CleanUp
            End Try

            'HILLSHADE
            Me.lblStatus.Text = "Please Wait, Creating Hillshade Grid"
            Me.Refresh()
            pCreateHillshade = New ESRI.ArcGIS.SpatialAnalystTools.HillShade
            pCreateHillshade.in_raster = pRaster2
            pCreateHillshade.altitude = "45"
            pCreateHillshade.azimuth = "315"
            pCreateHillshade.out_raster = (m_frmEnvisionProjectSetup.tbxWorkspace.Text + "\Hillshade")
            pCreateHillshade.z_factor = "3.2808"
            RunTool(GP, pCreateHillshade)
            pCreateHillshade = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()
            If Not mxDoc.SelectedLayer.Name = "Hillshade" Then
                MessageBox.Show("The envision tools were unable to create the Hillshade raster from the selected elevation raster.", "Hillshade Process Failed", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            End If

            'Create Slope Grid from selected raster
            Me.lblStatus.Text = "Please Wait, Creating Slope Grid"
            Me.Refresh()
            pCreateSlope = New ESRI.ArcGIS.SpatialAnalystTools.Slope
            pCreateSlope.in_raster = pRaster2
            pCreateSlope.output_measurement = "DEGREE"
            pCreateSlope.out_raster = (m_frmEnvisionProjectSetup.tbxWorkspace.Text + "\Slope")
            pCreateSlope.z_factor = "3.2808"
            RunTool(GP, pCreateSlope)
            pCreateSlope = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()

            If mxDoc.SelectedLayer.Name = "Slope" Then
                m_pSlopeLyr = CType(mxDoc.SelectedLayer, IRasterLayer)
                pRaster2 = CType(m_pSlopeLyr.Raster, IRaster2)
            Else
                MessageBox.Show("The envision tools were unable to create the Slope raster from the selected elevation raster.", "Slope Process Failed", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
                GoTo CleanUp
            End If

            '15-25% SLOPES
            Me.lblStatus.Text = "Please Wait, Reclassing 15-25% Slope Grid"
            Me.Refresh()
            pCreate15Slope = New ESRI.ArcGIS.SpatialAnalystTools.Reclassify
            pCreate15Slope.in_raster = pRaster2
            pCreate15Slope.out_raster = (m_frmEnvisionProjectSetup.tbxWorkspace.Text + "\Slp15to25")
            pCreate15Slope.remap = "15 25 1"
            pCreate15Slope.missing_values = "NODATA"
            pCreate15Slope.reclass_field = "Value"
            RunTool(GP, pCreate15Slope)
            pCreate15Slope = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()
            If mxDoc.SelectedLayer.Name = "Slp15to25" Then
                pRaster2 = CType(CType(mxDoc.SelectedLayer, IRasterLayer).Raster, IRaster2)
            Else
                MessageBox.Show("The envision tools were unable to create the 15-25% Slope raster.", "Slope Process Failed", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            End If

            'RASRTER TO POLYGONS 15-25% SLOPES
            Me.lblStatus.Text = "Please Wait, Coverting Raster to Polygons"
            Me.Refresh()
            pCreate15Lyr = New ESRI.ArcGIS.ConversionTools.RasterToPolygon
            pCreate15Lyr.in_raster = pRaster2
            pCreate15Lyr.out_polygon_features = (m_frmEnvisionProjectSetup.tbxWorkspace.Text + "\Slp15to25.shp")
            pCreate15Lyr.simplify = "NO_SIMPLIFY"
            pCreate15Lyr.raster_field = "Value"
            RunTool(GP, pCreate15Lyr)
            pCreate15Lyr = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()

            'GREATER THAN 25% SLOPES
            Me.lblStatus.Text = "Please Wait, Reclassing 25+% Slope Grid"
            Me.Refresh()
            pCreate25Slope = New ESRI.ArcGIS.SpatialAnalystTools.Reclassify
            pRaster2 = CType(m_pSlopeLyr.Raster, IRaster2)
            pCreate25Slope.in_raster = pRaster2
            pCreate25Slope.out_raster = (m_frmEnvisionProjectSetup.tbxWorkspace.Text + "\" + "Slp25Plus")
            pCreate25Slope.remap = "25 100 1"
            pCreate25Slope.missing_values = "NODATA"
            pCreate25Slope.reclass_field = "Value"
            RunTool(GP, pCreate25Slope)
            pCreate25Slope = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()
            If mxDoc.SelectedLayer.Name = "Slp25Plus" Then
                pRaster2 = CType(CType(mxDoc.SelectedLayer, IRasterLayer).Raster, IRaster2)
            Else
                MessageBox.Show("The envision tools were unable to create the 25Plus% Slope raster.", "Slope Process Failed", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            End If

            'RASRTER TO POLYGONS 15-25% SLOPES
            Me.lblStatus.Text = "Please Wait, Converting Raster to Polygons"
            Me.Refresh()
            pCreate25Lyr = New ESRI.ArcGIS.ConversionTools.RasterToPolygon
            pCreate25Lyr.in_raster = pRaster2
            pCreate25Lyr.out_polygon_features = (m_frmEnvisionProjectSetup.tbxWorkspace.Text + "\Slp25Plus.shp")
            pCreate25Lyr.simplify = "NO_SIMPLIFY"
            pCreate25Lyr.raster_field = "Value"
            RunTool(GP, pCreate25Lyr)
            pCreate25Lyr = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()

            ' m_blnSlope = False
            m_frmEnvisionProjectSetup.cmbParcelLayers.Items.Clear()
            m_frmEnvisionProjectSetup.cmbExtentLayers.Items.Clear()
            m_frmEnvisionProjectSetup.dgvConstraints.Rows.Clear()
            m_frmEnvisionProjectSetup.cmbLandUseLayers.Items.Clear()
            m_frmEnvisionProjectSetup.cmbLandUseField.Items.Clear()
            m_frmEnvisionProjectSetup.dgvLandUseAttributes.Rows.Clear()
            m_frmEnvisionProjectSetup.dgvSubAreas.Rows.Clear()
            'EnvisionProjectSetup_LoadData(sender, e)
            Me.Close()

            GoTo CleanUp
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString, "Envision Slope/Hillshade Data Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo CleanUp
        End Try


CleanUp:
        Me.lblStatus.Text = ""
        Me.lblStatus.Visible = False
        mxApplication = Nothing
        mxDoc = Nothing
        GP = Nothing
        pCreateSlope = Nothing
        pRLyr = Nothing
        pRaster2 = Nothing
        ' pRasterProps = Nothing
        pCreateSlope = Nothing
        pCreateHillshade = Nothing
        pTempRLyr = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()

    End Sub

End Class