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

Public Class frmOnetoManyExport

    Public blnIsOpening As Boolean = True

    Private Sub frmOnetoManyExport_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Me.WindowState = Windows.Forms.FormWindowState.Minimized
        m_frm1toManyExport = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub

    Private Sub frmOnetoManyExport_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'POPULATE THE FIELDS FROM THE SELECTED ENVISION PARCEL LAYER
        blnIsOpening = True

        'CHECK FOR ENVISION EDIT LAYER
        If m_pEditFeatureLyr Is Nothing Then
            m_frm1toManyExport.Close()
            GoTo CleanUp
        End If


        Dim pFeatClass As IFeatureClass = m_pEditFeatureLyr.FeatureClass
        Dim pTable As ITable = CType(pFeatClass, ITable)
        Dim intFldNum As Integer
        Dim pField As IField

        'CHECK FOR AVAILABLE TABLE
        pFeatClass = Nothing
        If pTable Is Nothing Then
            MessageBox.Show("The Envision Edit Layer could not be found in the current view document.  Please reselect the Envision layer.", "Layer Not Found", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            m_frm1toManyExport.Close()
            GoTo CleanUp
        End If

        'LOAD ALL THE FIELD NAME(S) INTO THE FORM CONTROLS
        m_frm1toManyExport.cmb1toManyFields.Items.Clear()
        m_frm1toManyExport.lstFields.Items.Clear()
        For intFldNum = 0 To pTable.Fields.FieldCount - 1
            pField = pTable.Fields.Field(intFldNum)
            If pField.Type = esriFieldType.esriFieldTypeInteger Or pField.Type = esriFieldType.esriFieldTypeDouble Then
                m_frm1toManyExport.cmb1toManyFields.Items.Add(pField.AliasName)
            End If
            If Not pField.Name = "Shape" And Not pField.Name = "OBJECTID" Then
                m_frm1toManyExport.lstFields.Items.Add(pField.AliasName)
            End If
        Next

        If m_frm1toManyExport.cmb1toManyFields.Items.Count > 0 Then
            m_frm1toManyExport.cmb1toManyFields.Enabled = True
            m_frm1toManyExport.lbl1toManyFields.Enabled = True

        End If
CleanUp:
        blnIsOpening = False
        pFeatClass = Nothing
        pTable = Nothing
        intFldNum = Nothing
        pField = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()

    End Sub

    Private Sub btnCheckAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCheckAll.Click
        FieldCheckStatus(True)
    End Sub

    Private Sub btnUncheckAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUncheckAll.Click
        FieldCheckStatus(False)
    End Sub

    Public Sub FieldCheckStatus(ByVal blnCheckStatus As Boolean)
        'SET CHECK STATUS TO ALL FIELD(S) AS DEFINED BY INUT VARIABLE
        Dim intCount As Integer
        For intCount = 0 To m_frm1toManyExport.lstFields.Items.Count - 1
            m_frm1toManyExport.lstFields.SetItemChecked(intCount, blnCheckStatus)
        Next
    End Sub

    Private Sub btnExport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExport.Click
        'EXECUTE THE FUNCTIONS FOR ONE TO MANY EXPORT
        m_frm1toManyExport.ProgressBar1.Visible = True
        m_frm1toManyExport.ProgressBar1.Value = 0
        GC.WaitForPendingFinalizers()
        GC.Collect()

        'CHECK FOR ENVISION EDIT LAYER
        If m_pEditFeatureLyr Is Nothing Then
            GoTo CleanUp
        End If

        Dim pCursor As ICursor = Nothing
        Dim pFeat As IFeature
        Dim pFeatureCursor As IFeatureCursor
        Dim pFeatClass As IFeatureClass = m_pEditFeatureLyr.FeatureClass
        Dim pTable As ITable = CType(pFeatClass, ITable)
        Dim pExportTable As ITable = Nothing
        Dim intTotalCount As Integer
        Dim intCount As Integer
        Dim pField As IField = Nothing
        Dim strFieldName As String = ""
        Dim intFldCount As Integer
        Dim intFldNum As Integer = 0
        Dim strValue As String = ""
        Dim pFieldEdit As IFieldEdit
        Dim GP As Geoprocessor
        Dim pCreateTable As ESRI.ArcGIS.DataManagementTools.CreateTable
        Dim pExportDBFTable As ESRI.ArcGIS.ConversionTools.TableToTable

        Dim intManyCount As Integer = 0
        Dim i As Integer = 0
        Dim rowTemp As IRow = Nothing
        Dim intRowCount As Integer = 0
        Dim pRowBuffer As IRowBuffer = Nothing
        Dim recCursor As ICursor

        Dim pWksFactory As ESRI.ArcGIS.Geodatabase.IWorkspaceFactory = Nothing
        Dim pFeatWks As ESRI.ArcGIS.Geodatabase.IFeatureWorkspace = Nothing
        Dim pDataSet As IDataset
        Dim ptrans As ITransactions

        'CHECK FOR A DEFINITION QUERY
        Dim pDef As IFeatureLayerDefinition2
        Dim pQFilter As IQueryFilter
        Dim strDefExpression As String = ""
        Dim strQString As String = Nothing

        'CHECK FOR REQUIRED FIELDS
        If pTable Is Nothing Then
            MessageBox.Show("The Envision Edit Layer could not be found in the current view document.  Please reselect the Envision layer.", "Layer Not Found", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo CleanUp
        End If

        'CHECK FOR COUNT FIELD
        If m_frm1toManyExport.cmb1toManyFields.FindString(m_frm1toManyExport.cmb1toManyFields.Text) < 0 Then
            MessageBox.Show("Please select a valid Many Count field.", "Invalid Field Selection", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo CleanUp
        End If
        If Not pTable.FindField(m_frm1toManyExport.cmb1toManyFields.Text) >= 0 Then
            MessageBox.Show("The Many Count field, " & m_frm1toManyExport.cmb1toManyFields.Text & ", could not be found.  Please select a valid Many Count field.", "Invalid Field Selection", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo CleanUp
        End If

        'CHECK FOR COUNT FIELD
        If m_frm1toManyExport.lstFields.CheckedItems.Count <= 0 Then
            If MessageBox.Show("No atribute fields have been selected.  Would you like to continue?", "Attribute Fields", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2, MessageBoxOptions.DefaultDesktopOnly, False) = Windows.Forms.DialogResult.No Then
                GoTo CleanUp
            End If
        End If

        pFeatClass = m_pEditFeatureLyr.FeatureClass
        pDef = m_pEditFeatureLyr
        strDefExpression = pDef.DefinitionExpression
        pQFilter = New QueryFilter
        'If Me.rdbPartial.Checked Then
        strQString = "NOT DEV_TYPE = ''"
        If strDefExpression.Length > 0 Then
            pQFilter.WhereClause = "(" & strQString & ") AND (" & pDef.DefinitionExpression & ")"
            pFeatureCursor = pFeatClass.Search(pQFilter, False)
        Else
            pQFilter.WhereClause = strQString
            pFeatureCursor = pFeatClass.Search(pQFilter, False)
        End If
        intTotalCount = pFeatClass.FeatureCount(pQFilter)
        'Else
        '    If strDefExpression.Length > 0 Then
        '        pQFilter.WhereClause = pDef.DefinitionExpression
        '        pFeatureCursor = pFeatClass.Search(pQFilter, False)
        '        intTotalCount = pFeatClass.FeatureCount(pQFilter)
        '    Else
        '        pFeatureCursor = pFeatClass.Search(Nothing, False)
        '        intTotalCount = pFeatClass.FeatureCount(Nothing)
        '    End If
        'End If

        '*********************************************************
        'WRITE VALUE(S) TO EXPORT FILE
        '*********************************************************
        'CREATE CSV FILE IF NEEDED
        If m_frm1toManyExport.rdbCSV.Checked Then
            Try
                If m_frm1toManyExport.tbxCSVFile.Text.Length <= 0 Then
                    MessageBox.Show("Please selected a CSV file for exporting of data.", "CSV filename Missing", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
                    GoTo CleanUp
                End If
            Catch CantCreate As Exception
                MessageBox.Show("Can't create or write the file" & vbNewLine & CantCreate.Message, "Write Error", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
                GoTo CleanUp
            End Try
        End If

        'CREATE DBF FILE IF NEEDED
        If m_frm1toManyExport.rdbDBF.Checked Then
            Try
                If m_frm1toManyExport.tbxDBFFile.Text.Length <= 0 Then
                    MessageBox.Show("Please selected a DBF file for exporting of data.", "DBF filename Missing", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
                    GoTo CleanUp
                End If
            Catch CantCreate As Exception
                MessageBox.Show("Can't create or write the file" & vbNewLine & CantCreate.Message, "Write Error", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
                GoTo CleanUp
            End Try
        End If

        GP = New Geoprocessor
        GP.AddOutputsToMap = False
        pCreateTable = New ESRI.ArcGIS.DataManagementTools.CreateTable
        pCreateTable.out_path = m_strFeaturePath

        'CREATE TEMPORARY GEO TABLE FOR THE STORAGE OF THE RECORDS BEFORE EXPORT TO DBASE FILE
        If m_frm1toManyExport.rdbCSV.Checked Or m_frm1toManyExport.rdbDBF.Checked Then
            Try
                If m_strFeaturePath.Length > 0 Then
                    pWksFactory = New FileGDBWorkspaceFactory
                    If Directory.Exists(m_strFeaturePath) Then
                        pFeatWks = pWksFactory.OpenFromFile(m_strFeaturePath, 0)
                        pExportTable = pFeatWks.OpenTable("TEMP_TABLE_ENVISION_EXPORT_1TOMANY")
                        pDataSet = pExportTable
                        If pDataSet.CanDelete Then
                            pExportTable = Nothing
                            ptrans = pFeatWks
                            ptrans.StartTransaction()
                            pDataSet.Delete()
                            ptrans.CommitTransaction()
                            pCreateTable.out_name = "TEMP_TABLE_ENVISION_EXPORT_1TOMANY"
                            RunTool(GP, pCreateTable)
                            pExportTable = pFeatWks.OpenTable("TEMP_TABLE_ENVISION_EXPORT_1TOMANY")
                        End If
                    End If
                End If
            Catch ex As Exception
                pCreateTable.out_name = "TEMP_TABLE_ENVISION_EXPORT_1TOMANY"
                RunTool(GP, pCreateTable)
                pExportTable = pFeatWks.OpenTable("TEMP_TABLE_ENVISION_EXPORT_1TOMANY")
            End Try
        End If
        pCreateTable = Nothing

        'CREATE GEO TABLE FOR THE STORAGE OF THE RECORDS BEFORE EXPORT TO DBASE FILE
        If m_frm1toManyExport.rdbTable.Checked Then
            Try
                If m_strFeaturePath.Length > 0 Then
                    pWksFactory = New FileGDBWorkspaceFactory
                    If Directory.Exists(m_strFeaturePath) Then
                        pFeatWks = pWksFactory.OpenFromFile(m_strFeaturePath, 0)
                        pExportTable = pFeatWks.OpenTable(m_frm1toManyExport.tbxExportTable.Text)
                        pDataSet = pExportTable
                        If pDataSet.CanDelete Then
                            pExportTable = Nothing
                            ptrans = pFeatWks
                            ptrans.StartTransaction()
                            pDataSet.Delete()
                            ptrans.CommitTransaction()
                            pCreateTable.out_name = m_frm1toManyExport.tbxExportTable.Text
                            RunTool(GP, pCreateTable)
                            pExportTable = pFeatWks.OpenTable(m_frm1toManyExport.tbxExportTable.Text)
                        End If
                    End If
                End If
            Catch ex As Exception
                pCreateTable.out_name = m_frm1toManyExport.tbxExportTable.Text
                RunTool(GP, pCreateTable)
                pExportTable = pFeatWks.OpenTable(m_frm1toManyExport.tbxExportTable.Text)
            End Try
        End If

        'ADD THE SELECTED FIELD(S)
        If m_frm1toManyExport.lstFields.CheckedItems.Count > 0 Then
            For Each strFieldName In m_frm1toManyExport.lstFields.CheckedItems
                pField = pTable.Fields.Field(pTable.Fields.FindField(strFieldName))
                pFieldEdit = CType(pField, IFieldEdit)
                pExportTable.AddField(pFieldEdit)
            Next
        End If

        'CYCLE THROUGH FOR THE CREATION OF THE MANY ROWS
        Try
            recCursor = pExportTable.Insert(True) 'ITable object
            intCount = 0
            intFldCount = 0
            intFldNum = pFeatClass.Fields.FindField(m_frm1toManyExport.cmb1toManyFields.Text)
            pFeat = pFeatureCursor.NextFeature
            Do While Not pFeat Is Nothing
                intCount = intCount + 1
                intManyCount = 0
                Try
                    intManyCount = pFeat.Value(intFldNum)
                Catch ex As Exception
                    intManyCount = 0
                End Try

                If intManyCount > 0 Then
                    pRowBuffer = pExportTable.CreateRowBuffer
                    For Each strFieldName In m_frm1toManyExport.lstFields.CheckedItems
                        pRowBuffer.Value(pRowBuffer.Fields.FindField(strFieldName)) = pFeat.Value(pFeat.Fields.FindField(strFieldName))
                    Next
                    For i = 1 To intManyCount
                        recCursor.InsertRow(pRowBuffer)
                    Next
                End If
                m_frm1toManyExport.ProgressBar1.Value = (intCount / intTotalCount) * 100
                m_frm1toManyExport.ProgressBar1.Refresh()
                pFeat = pFeatureCursor.NextFeature
            Loop
        Catch ex As Exception
            MessageBox.Show(ex.Message, "ROW DUPLICATION ERROR")
            GoTo CleanUp
        End Try

        'EXPORT TEMP TABLE IF DBF OPTION SELECTED
        If m_frm1toManyExport.rdbDBF.Checked Then
            Dim strExportName As String = ""
            Dim strExportPath As String = ""
            For i = 0 To m_frm1toManyExport.tbxDBFFile.Text.Split("\").Length - 1
                If i = 0 Then
                    strExportPath = m_frm1toManyExport.tbxDBFFile.Text.Split("\")(i)
                    Continue For
                ElseIf i = m_frm1toManyExport.tbxDBFFile.Text.Split("\").Length - 1 Then
                    strExportName = m_frm1toManyExport.tbxDBFFile.Text.Split("\")(i)
                    Continue For
                End If
                strExportPath = strExportPath & "\" & m_frm1toManyExport.tbxDBFFile.Text.Split("\")(i)
            Next
            Try
                pExportDBFTable = New ESRI.ArcGIS.ConversionTools.TableToTable
                pExportDBFTable.in_rows = pExportTable
                pExportDBFTable.out_name = strExportName
                pExportDBFTable.out_path = strExportPath
                RunTool(GP, pExportDBFTable)
            Catch ex As Exception
                MessageBox.Show(ex.Message, "DBASE FILE EXPORT ERROR")
                GoTo CleanUp
            End Try
            'DELETE THE TEMPORARY GEODATABASE TABLE
            pDataSet = pExportTable
            If pDataSet.CanDelete Then
                pExportTable = Nothing
                ptrans = pFeatWks
                ptrans.StartTransaction()
                pDataSet.Delete()
                ptrans.CommitTransaction()
            End If
        End If


CleanUp:
        m_frm1toManyExport.ProgressBar1.Value = 0
        m_frm1toManyExport.ProgressBar1.Visible = False
        m_frm1toManyExport.Close()
        pCursor = Nothing
        pFeat = Nothing
        pFeatureCursor = Nothing
        pFeatClass = Nothing
        pTable = Nothing
        pExportTable = Nothing
        intTotalCount = Nothing
        intCount = Nothing
        pField = Nothing
        strFieldName = Nothing
        intFldCount = Nothing
        intFldNum = Nothing
        strValue = Nothing
        pFieldEdit = Nothing
        GP = Nothing
        pCreateTable = Nothing
        pExportDBFTable = Nothing
        intManyCount = Nothing
        i = Nothing
        rowTemp = Nothing
        intRowCount = Nothing
        pRowBuffer = Nothing
        recCursor = Nothing
        pWksFactory = Nothing
        pFeatWks = Nothing
        pDataSet = Nothing
        ptrans = Nothing
        pDef = Nothing
        pQFilter = Nothing
        strDefExpression = Nothing
        strQString = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()

    End Sub

    Private Sub rdbCSV_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdbCSV.CheckedChanged
        If Not blnIsOpening Then
            m_frm1toManyExport.grpCSV.Visible = True
            m_frm1toManyExport.grpDBF.Visible = False
            m_frm1toManyExport.grpEnvisionTable.Visible = False
        End If
    End Sub

    Private Sub rdbDBF_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdbDBF.CheckedChanged
        If Not blnIsOpening Then
            m_frm1toManyExport.grpCSV.Visible = False
            m_frm1toManyExport.grpDBF.Visible = True
            m_frm1toManyExport.grpEnvisionTable.Visible = False
        End If
    End Sub

    Private Sub rdbTable_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdbTable.CheckedChanged
        If Not blnIsOpening Then
            m_frm1toManyExport.grpCSV.Visible = False
            m_frm1toManyExport.grpDBF.Visible = False
            m_frm1toManyExport.grpEnvisionTable.Visible = True
        End If
    End Sub

    Private Sub btnSaveFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSaveCSVFile.Click
        'PROVIDE SAVE FILE DIALOG FOR USER TO DESIGNATE THE OUTPUT CSV FILE
        SaveFileDialog1.AddExtension = True
        SaveFileDialog1.Filter = "Comma Delimited Text File|*.csv"
        SaveFileDialog1.Title = "Save File - One to Many Export"
        If SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            m_frm1toManyExport.tbxCSVFile.Text = SaveFileDialog1.FileName
        End If
    End Sub

    Private Sub btnDBFSaveFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDBFSaveFile.Click
        'PROVIDE SAVE FILE DIALOG FOR USER TO DESIGNATE THE OUTPUT DBF FILE
        SaveFileDialog1.AddExtension = True
        SaveFileDialog1.Filter = "dBASE File|*.dbf"
        SaveFileDialog1.Title = "Save File - One to Many Export"
        If SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            m_frm1toManyExport.tbxDBFFile.Text = SaveFileDialog1.FileName
        End If
    End Sub

End Class