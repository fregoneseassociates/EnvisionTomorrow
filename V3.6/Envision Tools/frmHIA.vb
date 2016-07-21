Option Explicit On
Option Strict On

Imports System
Imports System.Windows.Forms
Imports System.Runtime.InteropServices
Imports System.IO
Imports Microsoft.Office

Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.DataSourcesGDB
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Geoprocessing
Imports ESRI.ArcGIS.GeoprocessingUI

''' <summary>
''' Class for Health Impact Assessment form
''' </summary>
''' <remarks>
''' Calls a geoprocessing tool which calls EPAP API
''' </remarks>
Public Class frmHIA

  ' Declare constants
  Private Const baseURL As String = "http://api.ud4htools.com/hmapi_post_custom_inputs/EPAP/"
  Private Const summaryBaseName As String = "HIA_Summary_"
  Private Const detailsBaseName As String = "HIA_Details_"
  Private Const existingConditionsName As String = "Existing Conditions"
  Private Const scenarioBaseName As String = "Scenario" & " "
  Private Const HIAToolBoxName As String = "ToolboxHIAA"
  Private Const HIAToolName As String = "EPAPPost"
  Private Const ExcelWorksheetName As String = "EPAP Outputs"
  Private Const ExcelVariableNameColumnName As String = "Variable Name"

  ''' <summary>
  ''' Runs the EPAP geoprocessing tool and processes the results
  ''' </summary>
  ''' <param name="sender"></param>
  ''' <param name="e"></param>
  ''' <remarks></remarks>
  Private Sub btnRun_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRun.Click
    ' Declare COM objects that require cleanup
    Dim xlApp As Microsoft.Office.Interop.Excel.Application = Nothing
    Dim xlWB As Microsoft.Office.Interop.Excel.Workbook = Nothing
    Dim xlWS As Microsoft.Office.Interop.Excel.Worksheet = Nothing

    Try
      ' Check for a selected census block group layer
      m_appEnvision.StatusBar.Message(0) = "Validating inputs..."
      lblStatus.Text = "Validating inputs..."
      System.Windows.Forms.Application.DoEvents()
      If Me.cmbLayers.Text = String.Empty OrElse Me.cmbLayers.Items.IndexOf(Me.cmbLayers.Text) = -1 Then
        MessageBox.Show("Please select a census block group layer from the list.", "Input Layer Not Selected", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
        Exit Sub
      End If

      ' Get the workspace of the selected layer
      ' This can only be a file geodatabase at this point
      Dim workspacePath As String = String.Empty
      Dim mxDocument As IMxDocument = DirectCast(m_appEnvision.Document, IMxDocument)
      Dim map As IMap = mxDocument.FocusMap
      For n As Int32 = 0 To map.LayerCount - 1
        If map.Layer(n).Name = cmbLayers.Text Then
          Dim workspace As IWorkspace = DirectCast(DirectCast(map.Layer(n), IFeatureLayer).FeatureClass, IDataset).Workspace
          workspacePath = workspace.PathName
          If Not workspacePath.ToLower.EndsWith(".gdb") Then
            MessageBox.Show("Only file geodatabase feature classes are currently supported.", "Input Layer Not Selected", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            Exit Sub
          End If
          Exit For
        End If
      Next

      ' Determine the Excel scenario column from the selected scenario or existing conditions
      ' Set the summary and details output paths
      Dim summaryName As String = String.Empty
      Dim detailsName As String = String.Empty
      Dim scenarioColumn As Int32 = 0
      Dim scenarioName As String = String.Empty
      If Me.rdbExisting.Checked Then
        scenarioColumn = 4
        summaryName = summaryBaseName & "EX"
        detailsName = detailsBaseName & "EX"
        scenarioName = existingConditionsName
      Else
        If Me.cmbScenarios.Text = scenarioBaseName & "1" Then
          scenarioColumn = 5
        ElseIf Me.cmbScenarios.Text = scenarioBaseName & "2" Then
          scenarioColumn = 6
        ElseIf Me.cmbScenarios.Text = scenarioBaseName & "3" Then
          scenarioColumn = 7
        ElseIf Me.cmbScenarios.Text = scenarioBaseName & "4" Then
          scenarioColumn = 8
        ElseIf Me.cmbScenarios.Text = scenarioBaseName & "5" Then
          scenarioColumn = 9
        End If
        Dim scenarioNum As String = cmbScenarios.Text.Replace(scenarioBaseName, "")
        summaryName = summaryBaseName & scenarioNum
        detailsName = detailsBaseName & scenarioNum
        scenarioName = cmbScenarios.Text
      End If
      If scenarioColumn = 0 Then
        MessageBox.Show("Please select a scenario from the list, or the existing conditions.", "Scenario Not Defined", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
        Exit Sub
      End If
      Dim outputSummaryFeatureclassPath As String = System.IO.Path.Combine(workspacePath, summaryName)
      Dim outputDetailsFeatureclassPath As String = System.IO.Path.Combine(workspacePath, detailsName)

      ' Check for Excel file specified
      Dim editExcel As Boolean = Me.tbxExcel.Text <> String.Empty
      If editExcel AndAlso Not File.Exists(Me.tbxExcel.Text) Then
        MessageBox.Show("The specified Health Assessment Model Excel file does not exist.", "Excel File Not Found", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
        Exit Sub
      End If

      ' Check for specified email
      If txtEmail.Text = String.Empty Then
        MessageBox.Show("Please enter an email address.", "Email Not Specified", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
        Exit Sub
      End If

      ' Initialize the geoprocessor
      m_appEnvision.StatusBar.Message(0) = "Loading Health Assessment Model tool..."
      lblStatus.Text = "Loading tool..."
      System.Windows.Forms.Application.DoEvents()
      Dim GP As ESRI.ArcGIS.Geoprocessing.GeoProcessor = New ESRI.ArcGIS.Geoprocessing.GeoProcessor()

      ' Check if the toolbox is already loaded
      Dim GPEnumListToolboxes As IGpEnumList = GP.ListToolboxes(HIAToolBoxName)
      Dim loadedToolbox As String = GPEnumListToolboxes.Next
      If loadedToolbox = String.Empty Then
        ' Check for the file in the install path
        Dim toolboxFilename As String = System.IO.Path.Combine(System.IO.Path.Combine(My.Application.Info.DirectoryPath, "UD4H"), "ToolboxHIAA.pyt")
        If Not File.Exists(toolboxFilename) Then
          ' Check for the config file
          Dim configPath As String = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "EnvisionTomorrow")
          If Not Directory.Exists(configPath) Then Directory.CreateDirectory(configPath)
          Dim configFilePath As String = System.IO.Path.Combine(configPath, "ET.txt")
          If Not File.Exists(configFilePath) Then
            ' Ask the user to located the toolbox file
            Dim FileDialog1 As New OpenFileDialog
            FileDialog1.Title = "Select the UD4H HIAA EPAP toolbox file"
            FileDialog1.Filter = "ArcGIS Toolbox File (*.pyt)|*.pyt"
            FileDialog1.CheckPathExists = True
            If FileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
              ' Set the toolbox filename
              toolboxFilename = FileDialog1.FileName

              ' Save the location in the config file
              Using streamWriter As StreamWriter = New StreamWriter(configFilePath)
                streamWriter.WriteLine(FileDialog1.FileName)
              End Using
            Else
              Exit Sub
            End If
          Else
            ' Read the config file for the path to the tool.
            Dim txt As String
            Using streamReader As StreamReader = New StreamReader(configFilePath)
              txt = streamReader.ReadLine
            End Using
            If txt = String.Empty OrElse Not File.Exists(txt) Then
              File.Delete(configFilePath)
              MessageBox.Show("The configuration file was invalid.  Please try running this tool again.  You will be asked to locate the toolbox file.", "Invalid Configuration File", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
              Exit Sub
            Else
              ' Set the toolbox filename
              toolboxFilename = txt
            End If
          End If
        End If

        ' Add the toolbox to ArcToolbox
        Try
          GP.AddToolbox(toolboxFilename)
        Catch ex As Exception
          MessageBox.Show("The following error occurred while attempting to access the tool: " & ex.Message, "Error Adding Toolbox", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
          Exit Sub
        End Try
      End If

      ' Get the tool by name
      Dim uID As UID = New UIDClass
      uID.Value = "esriGeoprocessingUI.ArcToolboxExtension"
      Dim arcToolboxExtension As IArcToolboxExtension = DirectCast(m_appEnvision.FindExtensionByCLSID(uID), IArcToolboxExtension)
      Dim arcToolbox As IArcToolbox = arcToolboxExtension.ArcToolbox
      Dim gpTool As IGPTool = Nothing
      Try
        gpTool = arcToolbox.GetToolbyNameString(HIAToolName)
      Catch ex As Exception
      End Try
      If gpTool Is Nothing Then
        MessageBox.Show("Unable to find UD4H HIAA EPAP tool.", "Tool Not Found", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
        Exit Sub
      End If

      ' Track the state of the Add Outputs to Map option and temporarily change it to off
      Dim existingAddOutputsToMap As Boolean = GP.AddOutputsToMap
      GP.AddOutputsToMap = False

      ' Get the parameters
      Dim baseOnly As String = Me.rdbExisting.Checked.ToString()

      ' Run the tool
      m_appEnvision.StatusBar.Message(0) = "Running Health Assessment Model tool, please wait..."
      lblStatus.Text = "Running tool, please wait..."
      System.Windows.Forms.Application.DoEvents()
      Dim parameters As VarArray = New VarArray
      parameters.Add(Me.cmbLayers.Text)
      parameters.Add(outputDetailsFeatureclassPath)
      parameters.Add(outputSummaryFeatureclassPath)
      parameters.Add(baseURL)
      parameters.Add(txtEmail.Text)
      parameters.Add(baseOnly)
      Try
        GP.Execute(gpTool.Name, parameters, Nothing)
      Catch ex As Exception
        MessageBox.Show("The following error occurred running the Health Assessment Model tool: " & ex.Message, "Health Assessment Model", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
        Exit Sub
      End Try
      For n As Int32 = 0 To GP.MessageCount - 1
        If GP.GetMessage(n).Contains("ERROR") OrElse GP.GetMessage(n).Contains("ERR:") Then
          MessageBox.Show("The following error occurred running the Health Assessment Model tool: " & GP.GetMessage(n), "Health Assessment Model Tool Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
          Exit Sub
        End If
      Next
      GP.AddOutputsToMap = existingAddOutputsToMap

      ' Create the details featureclass
      System.Windows.Forms.Application.DoEvents()
      Dim fileGDBWorkspaceFactory As IWorkspaceFactory2 = New FileGDBWorkspaceFactory
      Dim featureWorkspace As IFeatureWorkspace = DirectCast(fileGDBWorkspaceFactory.OpenFromFile(System.IO.Path.GetDirectoryName(outputDetailsFeatureclassPath), 0), IFeatureWorkspace)
      Dim detailsFeatureClass As IFeatureClass = featureWorkspace.OpenFeatureClass(detailsName)

      ' Display message
      m_appEnvision.StatusBar.Message(0) = "Adding details layer to the map..."
      lblStatus.Text = "Adding layer to map..."

      ' Delay map events.
      map.DelayDrawing(True)
      map.DelayEvents(True)

      ' Create the layer and add it to the map
      Dim featureLayer As IFeatureLayer
      featureLayer = New FeatureLayer
      featureLayer.FeatureClass = detailsFeatureClass
      Dim layer As ILayer = DirectCast(featureLayer, ILayer)
      layer.Name = featureLayer.FeatureClass.AliasName & " - " & scenarioName
      map.AddLayer(featureLayer)

      ' Refresh the TOC and map
      map.DelayDrawing(False)
      map.DelayEvents(False)
      Dim contentsView As IContentsView = mxDocument.CurrentContentsView
      contentsView.Refresh(Nothing)
      Dim activeView As IActiveView = DirectCast(map, IActiveView)
      activeView.Refresh()

      ' Open the Excel workbook
      If editExcel Then
        m_appEnvision.StatusBar.Message(0) = "Editing Excel file..."
        lblStatus.Text = "Editing Excel file..."
        System.Windows.Forms.Application.DoEvents()
        Try
          xlApp = New Microsoft.Office.Interop.Excel.Application
          xlApp.DisplayAlerts = False
          xlApp.WindowState = Interop.Excel.XlWindowState.xlMinimized
          xlApp.Visible = True
          xlWB = DirectCast(xlApp.Workbooks.Open(Me.tbxExcel.Text), Microsoft.Office.Interop.Excel.Workbook)
        Catch ex As Exception
          MessageBox.Show(ex.Message, "Opening Health Assessment Model Excel File Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
        End Try

        ' Open the worksheet
        Try
          xlWS = DirectCast(xlWB.Sheets(ExcelWorksheetName), Interop.Excel.Worksheet)
        Catch ex As Exception
          MessageBox.Show("Unable to open the worksheet 'EPAP Outputs'" & vbNewLine & "Error: " & ex.Message, "Excel Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
          Exit Sub
        End Try

        ' Determine the current formula calc setting to reset after function executes
        Dim intExcelFormulaCalc As Int32
        If xlApp.Calculation = Interop.Excel.XlCalculation.xlCalculationAutomatic Then
          intExcelFormulaCalc = 1
        ElseIf xlApp.Calculation = Interop.Excel.XlCalculation.xlCalculationManual Then
          intExcelFormulaCalc = 2
        ElseIf xlApp.Calculation = Interop.Excel.XlCalculation.xlCalculationSemiautomatic Then
          intExcelFormulaCalc = 3
        End If

        ' Set Excel formula calc to manual
        xlApp.Calculation = Interop.Excel.XlCalculation.xlCalculationManual

        ' Get the summary featureclass
        Dim summaryFeatureClass As IFeatureClass = featureWorkspace.OpenFeatureClass(summaryName)

        ' Edit Excel values
        ' Get the single feature for the summary featureclass
        Dim summaryFeature As IFeature = summaryFeatureClass.GetFeature(1)

        ' Find the 'Variable Name' column
        Dim varNameCol As Int32
        For col As Int32 = 1 To 100
          If DirectCast(xlWS.Cells(1, col), Microsoft.Office.Interop.Excel.Range).Value Is Nothing Then
            MessageBox.Show("Unable to find 'Variable Name' column in Excel", "Missing Variable Name Column", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            Exit Sub
          End If
          If DirectCast(xlWS.Cells(1, col), Microsoft.Office.Interop.Excel.Range).Value.ToString = ExcelVariableNameColumnName Then
            varNameCol = col
            Exit For
          End If
        Next

        ' Find the start and end rows for the variables in the Excel
        Dim startRow As Int32 = 2
        Dim endRow As Int32
        For row As Int32 = startRow To 100
          If DirectCast(xlWS.Cells(row, varNameCol), Microsoft.Office.Interop.Excel.Range).Value Is Nothing Then
            endRow = row - 1
            Exit For
          End If
        Next

        ' Make sure all current values for this scenario are cleared
        For row As Int32 = startRow To endRow
          DirectCast(xlWS.Cells(row, scenarioColumn), Microsoft.Office.Interop.Excel.Range).Value = String.Empty
        Next

        ' Loop through the feature fields, skipping the first field
        For n As Int32 = 2 To summaryFeature.Fields.FieldCount - 1
          ' Get the field name
          Dim fldName As String = summaryFeature.Fields.Field(n).AliasName

          ' Loop through the worksheet rows, getting the value of the variable name column
          For row As Int32 = startRow To endRow
            ' Get the variable name and compare to the current field
            Dim varName As String = DirectCast(xlWS.Cells(row, varNameCol), Microsoft.Office.Interop.Excel.Range).Value.ToString
            If varName.ToLower = fldName.ToLower Then
              ' Set the the cell value
              If summaryFeature.Value(n) IsNot DBNull.Value Then
                DirectCast(xlWS.Cells(row, scenarioColumn), Microsoft.Office.Interop.Excel.Range).Value = summaryFeature.Value(n)
              End If
            End If
          Next
        Next

        ' Reset formula calc setting
        If intExcelFormulaCalc = 1 Then
          xlApp.Calculation = Interop.Excel.XlCalculation.xlCalculationAutomatic
        ElseIf intExcelFormulaCalc = 2 Then
          xlApp.Calculation = Interop.Excel.XlCalculation.xlCalculationManual
        ElseIf intExcelFormulaCalc = 3 Then
          xlApp.Calculation = Interop.Excel.XlCalculation.xlCalculationSemiautomatic
        End If

        ' Save Excel changes
        xlWB.Save()
      End If

      ' Close the form and display message
      Me.WindowState = FormWindowState.Minimized
      MessageBox.Show(Me, "The tool has completed successfully!", "Health Assessment Model", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
      Me.Close()
    Catch ex As Exception
      MessageBox.Show("The following error occurred: " & ex.Message, "Health Assessment Model", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
    Finally
      ' Cleanup and close Excel
      Try
        ' Call garbage collection to release minor objects that don't have variable references
        GC.Collect()
        GC.WaitForPendingFinalizers()

        ' Cleanup Excel worksheet COM object
        If xlWS IsNot Nothing Then
          Marshal.FinalReleaseComObject(xlWS)
        End If

        ' Close the workbook and Excel application
        If xlWB IsNot Nothing Then
          xlWB.Close(False)
          Marshal.FinalReleaseComObject(xlWB)
        End If
        If xlApp IsNot Nothing Then
          xlApp.Quit()
          Marshal.FinalReleaseComObject(xlApp)
        End If

        ' Clear the status bar
        m_appEnvision.StatusBar.Message(0) = ""
        lblStatus.Text = ""
      Catch ex As Exception
        MessageBox.Show("Error in closing the Envision Excel file.  You may need to terminate the application using the Task Manager." & vbNewLine & ex.Message, "Excel Closing Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
      End Try
    End Try
  End Sub

  ''' <summary>
  ''' Allow the user to locate the Excel file on disk
  ''' </summary>
  ''' <param name="sender"></param>
  ''' <param name="e"></param>
  ''' <remarks></remarks>
  Private Sub btnExcelFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExcelFile.Click
    Try
      ' Let the user select the Excel file
      Dim FileDialog1 As New OpenFileDialog
      FileDialog1.Title = "Select Health Assessment Model Excel File"
      FileDialog1.Filter = "Excel Files (*.xls,*.xlsx,*.xlsm)|*.xls;*.xlsx;*.xlsm"
      If FileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
        Me.tbxExcel.Text = FileDialog1.FileName.ToString
      End If
    Catch ex As Exception
      MessageBox.Show(ex.Message, "Excel File Error")
    End Try
  End Sub

  ''' <summary>
  ''' Loads the HIA form
  ''' </summary>
  ''' <param name="sender"></param>
  ''' <param name="e"></param>
  ''' <remarks></remarks>
  Private Sub frmHIA_Load(sender As Object, e As EventArgs) Handles Me.Load
    Try
      ' Clear the status
      lblStatus.Text = ""

      ' Populate the list of polygon layers
      Me.cmbLayers.Items.Clear()
      Dim mxDocument As IMxDocument = DirectCast(m_appEnvision.Document, IMxDocument)
      Dim map As IMap = mxDocument.FocusMap
      Dim activeView As IActiveView = DirectCast(map, IActiveView)
      Dim layerEnum As IEnumLayer = map.Layers()
      Dim layer As ILayer = map.Layers.Next
      Do While layer IsNot Nothing
        If TypeOf layer Is IFeatureLayer Then
          Dim featureLayer As IFeatureLayer = DirectCast(layer, IFeatureLayer)
          If featureLayer.FeatureClass.ShapeType = ESRI.ArcGIS.Geometry.esriGeometryType.esriGeometryPolygon Then
            If featureLayer.FeatureClass.Fields.FindField("GEOID10") <> -1 Then
              Me.cmbLayers.Items.Add(layer.Name)
            End If
          End If
        End If
        layer = layerEnum.Next
      Loop
      If cmbLayers.Items.Count = 0 Then
        MessageBox.Show("Please add an input polygon layer with a 'GEOID10' attribute to the current view document.", "No Layers Found", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
      Else
        ' Select the first item
        cmbLayers.SelectedIndex = 0
      End If
    Catch ex As Exception
      MessageBox.Show(ex.Message, "Error Loading Health Assessment Model Form")
    End Try
  End Sub

  ''' <summary>
  ''' Responds to the user selecting a census block group layer from the list
  ''' </summary>
  ''' <param name="sender"></param>
  ''' <param name="e"></param>
  ''' <remarks>
  ''' Displays a warning icon button if there are more than 500 features or selected features
  ''' </remarks>
  Private Sub cmbLayers_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbLayers.SelectedIndexChanged
    ' Verify the layer has less than or equal to 500 features selected
    Dim mxDocument As IMxDocument = DirectCast(m_appEnvision.Document, IMxDocument)
    Dim map As IMap = mxDocument.FocusMap
    For n As Int32 = 0 To map.LayerCount - 1
      If map.Layer(n).Name = cmbLayers.Text Then
        Dim featurelayer As IFeatureLayer = DirectCast(map.Layer(n), IFeatureLayer)
        Dim selectionSet As ISelectionSet = DirectCast(featurelayer, IFeatureSelection).SelectionSet
        Dim featureClass As IFeatureClass = featurelayer.FeatureClass
        btnLayerWarning.Visible = False
        If selectionSet.Count <> 0 Then
          If selectionSet.Count > 500 Then
            btnLayerWarning.Visible = True
          End If
        Else
          If featureClass.FeatureCount(Nothing) > 500 Then
            btnLayerWarning.Visible = True
          End If
        End If
        Exit For
      End If
    Next
  End Sub

  ''' <summary>
  ''' Responds to the user clicking the warning icon
  ''' </summary>
  ''' <param name="sender"></param>
  ''' <param name="e"></param>
  ''' <remarks></remarks>
  Private Sub btnLayerWarning_Click(sender As Object, e As EventArgs) Handles btnLayerWarning.Click
    MessageBox.Show("The selected layer has more than 500 features or selected features.  The output layer will contain points representing polygon centroids instead of polygons.  The scenario summary will not be affected.", "Layer Features", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
  End Sub
End Class