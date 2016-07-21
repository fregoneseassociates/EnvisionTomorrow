Imports System
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Runtime.InteropServices
Imports System.IO

Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Display
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.GeodatabaseUI
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.SystemUI
'Imports ESRI.ArcGIS.Utility.CATIDs
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.ADF.BaseClasses


<Guid("4CFE6E02-A63E-453C-A5D9-1155A278EA2F")> _
    Public Class CustomToolBar
    Inherits BaseToolbar
    ' Implements IToolBarDef

#Region "COM Registration Function(s)"
    <ComRegisterFunction(), ComVisibleAttribute(False)> _
    Public Shared Sub RegisterFunction(ByVal registerType As Type)
        ' Required for ArcGIS Component Category Registrar support
        ArcGISCategoryRegistration(registerType)

        'Add any COM registration code after the ArcGISCategoryRegistration() call

    End Sub

    <ComUnregisterFunction(), ComVisibleAttribute(False)> _
    Public Shared Sub UnregisterFunction(ByVal registerType As Type)
        ' Required for ArcGIS Component Category Registrar support
        ArcGISCategoryUnregistration(registerType)

        'Add any COM unregistration code after the ArcGISCategoryUnregistration() call

    End Sub

#Region "ArcGIS Component Category Registrar generated code"
    ''' <summary>
    ''' Required method for ArcGIS Component Category registration -
    ''' Do not modify the contents of this method with the code editor.
    ''' </summary>
    Private Shared Sub ArcGISCategoryRegistration(ByVal registerType As Type)
        Dim regKey As String = String.Format("HKEY_CLASSES_ROOT\CLSID\{{{0}}}", registerType.GUID)
        MxCommandBars.Register(regKey)

    End Sub
    ''' <summary>
    ''' Required method for ArcGIS Component Category unregistration -
    ''' Do not modify the contents of this method with the code editor.
    ''' </summary>
    Private Shared Sub ArcGISCategoryUnregistration(ByVal registerType As Type)
        Dim regKey As String = String.Format("HKEY_CLASSES_ROOT\CLSID\{{{0}}}", registerType.GUID)
        MxCommandBars.Unregister(regKey)

    End Sub

#End Region
#End Region
    '#Region "Component Category Registration"
    '    <ComRegisterFunction()> _
    '    Public Shared Sub Reg(ByVal regKey As String)
    '        ESRI.ArcGIS.ADF.CATIDs.MxCommandBars.Register(regKey)
    '    End Sub
    '
    '    <ComUnregisterFunction()> _
    '    Public Shared Sub Unreg(ByVal regKey As String)
    '        ESRI.ArcGIS.ADF.CATIDs.MxCommandBars.Unregister(regKey)
    '    End Sub
    '#End Region


    Public Overrides ReadOnly Property Caption() As String
        Get
            ' Set the string that appears as the toolbar's title
            Return "Envision Tomorrow"
        End Get
    End Property

    'Public Overrides Sub GetItemInfo(ByVal pos As Integer, ByVal itemDef As ESRI.ArcGIS.SystemUI.IItemDef)
    '    ' Define the commands that will be on the toolbar. The 1st command
    '    ' will be the Envision Project Setup and the 2nd Envision Attribute Edit Tool
    '    Select Case pos
    '        Case 0
    '            '    itemDef.ID = "Envision_Tools.EnvisionProjectSetup"
    '            '    itemDef.Group = False
    '            'Case 1
    '            itemDef.ID = "Envision_Tools.EnvisionAttributeEditor"
    '            itemDef.Group = True
    '    End Select

    'End Sub

    'Public Overrides ReadOnly Property ItemCount() As Integer
    '    Get
    '        'Set how many commands will be on the toolbar
    '        Return 1
    '    End Get
    'End Property

    Public Overrides ReadOnly Property Name() As String
        Get
            ' Set the internal name of the toolbar.
            Return "EnvisionToolbar2"
        End Get
    End Property

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Public Sub New()
        AddItem("{3C6A7F33-FCB2-426C-B919-E9FEDE429BF6}") ' EnvisionAttributeEditor
    End Sub
End Class

<Guid("3C6A7F33-FCB2-426C-B919-E9FEDE429BF6")> _
Public Class EnvisionAttributeEditor

    Implements ESRI.ArcGIS.SystemUI.ICommand
    Implements ESRI.ArcGIS.SystemUI.ITool

    Private m_bitmap As Bitmap
    Private m_hBitmap As IntPtr
    Private m_cursor As System.Windows.Forms.Cursor
    Private m_hCursor As IntPtr

    Private m_pArc As IPolyline
    Private m_pPoly As IPolygon
    Private m_pCircArc As ICircularArc

    Dim m_pMap As IMap
    Dim m_pMxDoc As IMxDocument
    Dim m_pMxApp As IMxApplication
    Dim m_pILayer As ILayer
    Dim m_pFLayer As IFeatureLayer
    Dim m_BMousedown As Boolean

    Private m_pDoc As IMxDocument
    Private m_pNewPolyFeedback As ESRI.ArcGIS.Display.INewPolygonFeedback = Nothing
    Private m_pNewLineFeedback As ESRI.ArcGIS.Display.INewLineFeedback = Nothing
    Private m_pNewCircleFeedback As ESRI.ArcGIS.Display.INewCircleFeedback = Nothing
    Private m_pNewEnvFeedback As ESRI.ArcGIS.Display.INewEnvelopeFeedback2 = Nothing

    Private m_pCircle As ICircularArc
    Private m_pAV As IActiveView
    Private m_pScrD As IScreenDisplay
    Private m_pRad As ILine
    Private m_pPolyline As IPolyline
    Private m_pntCollection As IPointCollection
    Private m_dblFullLength As Double = 0
    Private m_pntStart As IPoint
    Private m_strUnit As String = ""
    Public m_strStatus As String = ""


    ' Needed to clear up the Hbitmap unmanaged resource
    <DllImport("gdi32.dll")> _
    Private Shared Function DeleteObject(ByVal hObject As IntPtr) As Boolean

    End Function

#Region "Component Category Registration"
    <ComRegisterFunction()> _
    Public Shared Sub Reg(ByVal regKey As String)
        ESRI.ArcGIS.ADF.CATIDs.MxCommands.Register(regKey)
    End Sub

    <ComUnregisterFunction()> _
    Public Shared Sub Unreg(ByVal regKey As String)
        ESRI.ArcGIS.ADF.CATIDs.MxCommands.Unregister(regKey)
    End Sub
#End Region

    Public Sub New()
        'm_bitmap = New System.Drawing.Bitmap((Me.GetType.Assembly.GetManifestResourceStream("Envision_Tools.ET_LOGO.bmp")))
        'If Not (m_bitmap Is Nothing) Then
        '    m_bitmap.MakeTransparent(m_bitmap.GetPixel(1, 1))
        '    m_hBitmap = m_bitmap.GetHbitmap()
        'End If
    End Sub

    Protected Overrides Sub Finalize()

        If m_hBitmap.ToInt32 <> 0 Then
            DeleteObject(m_hBitmap)
        End If

        If m_hCursor.ToInt32() <> 0 Then
            DeleteObject(m_hCursor)
        End If

        If Not m_appEnvision Is Nothing Then
            Marshal.ReleaseComObject(m_appEnvision)
        End If

    End Sub

    Public ReadOnly Property Bitmap() As Integer Implements ESRI.ArcGIS.SystemUI.ICommand.Bitmap
        Get
            ' Set the bitmap of the command. The m_hBitmap variable is set
            ' in constructor (New) of this class.
            Return m_hBitmap.ToInt32()
        End Get
    End Property

    Public ReadOnly Property Caption() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Caption
        Get
            ' Set the string that appears when the command is used as a
            ' menu item.
            Return "ET Paint Tool"
        End Get
    End Property

    Public ReadOnly Property Category() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Category
        Get
            ' Set the category of this command. This determines where the
            ' command appears in the Commands panel of the Customize dialog.
            Return "EnvisionTools3"
        End Get
    End Property

    Public ReadOnly Property Checked() As Boolean Implements ESRI.ArcGIS.SystemUI.ICommand.Checked
        Get
            Try
                m_blnEnvisionEditingFormIsOpening = True

                m_blnEnvisionEditingFormIsOpening = False
            Catch ex As Exception
                MessageBox.Show(ex.Message, "ET Attribute Editing Tool 'Checked' Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            End Try

CleanUp:
            GC.WaitForPendingFinalizers()
            GC.Collect()
            Return False

        End Get
    End Property

    Public ReadOnly Property Enabled() As Boolean Implements ESRI.ArcGIS.SystemUI.ICommand.Enabled
        Get
            ' Add some logic here to specify when the command should
            ' be enabled. In this example, the command is always enabled.
            Return True
        End Get
    End Property

    Public ReadOnly Property HelpContextID() As Integer Implements ESRI.ArcGIS.SystemUI.ICommand.HelpContextID
        Get
            Return 0
        End Get
    End Property

    Public ReadOnly Property HelpFile() As String Implements ESRI.ArcGIS.SystemUI.ICommand.HelpFile
        Get
            Return ""
        End Get
    End Property

    Public ReadOnly Property Message() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Message
        Get

            ' Set the message string that appears in the statusbar of the
            ' application when the mouse passes over the command.
            Return "ET Attribute Editing Tool"
        End Get
    End Property

    Public ReadOnly Property Name() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Name
        Get
            ' Set the internal name of this command. By convention, this
            ' name string contains the category and caption of the command.
            Return "EnvisionTools_EnvisionEditingTool2"

        End Get
    End Property

    Public Sub OnClick() Implements ESRI.ArcGIS.SystemUI.ICommand.OnClick
        m_blnEnvisionEditingFormIsOpening = True
        Try
            If m_dockEnvisionWinForm Is Nothing Then
                m_dockEnvisionWinForm = New ctlEnvisionAttributesEditor
            End If
            m_pEnvisionDockWin.Show(True)

            GoTo CleanUp
        Catch ex As Exception
            MessageBox.Show(ex.Message, "ET Paint Tool 'OnClick' Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo cleanup
        End Try

CleanUp:
        m_blnEnvisionEditingFormIsOpening = False

    End Sub

    Public Sub OnCreate(ByVal hook As Object) Implements ESRI.ArcGIS.SystemUI.ICommand.OnCreate
        m_appEnvision = CType(hook, IApplication)


        If TypeOf hook Is IMxApplication Then
            m_pEnvisionDockWinMgr = CType(hook, ESRI.ArcGIS.Framework.IDockableWindowManager)
            m_pDevTypeDockWinMgr = CType(hook, ESRI.ArcGIS.Framework.IDockableWindowManager)
            ' m_frmRedevTimingDockWinMgr = CType(hook, ESRI.ArcGIS.Framework.IDockableWindowManager)

            ' Get a reference to the selection count dockable window
            Dim u1 As New UID
            Dim u2 As New UID
            Dim u3 As New UID

            ' Get a reference to the selection count dockable window
            Try
                u1.Value = "{E37A794C-F661-41C1-883B-A05648131221}"
                u2.Value = "{E37A794C-F661-41C1-883B-A05648131223}"
                u3.Value = "{E37A794C-F661-41C1-883B-A05648131224}"
                m_pEnvisionDockWin = m_pEnvisionDockWinMgr.GetDockableWindow(u1)
                m_pDevTypeDockWin = m_pDevTypeDockWinMgr.GetDockableWindow(u2)
                'm_frmRedevTiming = m_frmRedevTimingDockWinMgr.GetDockableWindow(u3)
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End If
    End Sub

    Public ReadOnly Property Tooltip() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Tooltip
        Get
            ' Set the string that appears in the screen tip.
            Return "ET Paint Tools"
        End Get
    End Property

    Public ReadOnly Property Cursor() As Integer Implements ESRI.ArcGIS.SystemUI.ITool.Cursor
        Get
            ' Set the cursor of the command. The m_hCursor variable is set
            ' in the constructor of the class.
            Return m_hCursor.ToInt32()
        End Get
    End Property

    Public Function Deactivate() As Boolean Implements ESRI.ArcGIS.SystemUI.ITool.Deactivate
        m_dockEnvisionWinForm.ddbBrushes.Checked = False
        Return True
    End Function

    Public Function OnContextMenu(ByVal x As Integer, ByVal y As Integer) As Boolean Implements ESRI.ArcGIS.SystemUI.ITool.OnContextMenu

    End Function

    Public Sub OnDblClick() Implements ESRI.ArcGIS.SystemUI.ITool.OnDblClick
        'STOP THE FEEDBACKS FOR POLYLINES AND POLYGONS, THEN EXECUTE WRITE FUNCTIONS
        Dim mxApplication As IMxApplication
        Dim mxDoc As IMxDocument
        Dim pGeom As IGeometry
        Dim pArc As IPolyline
        Dim pPoly As IPolygon
        Dim pArea As IArea
        Dim pAcres As Double = 0
        Dim strMessage As String = ""
        Try
            'EXIT IF NOT FEEDBACK IS AVAILABLE
            mxApplication = CType(m_appEnvision, IMxApplication)
            mxDoc = CType(m_appEnvision.Document, IMxDocument)
            If UCase(m_strEnvisionEditOption) = UCase("Polyline") Then
                pGeom = m_pNewLineFeedback.Stop
                If pGeom Is Nothing Then
                    GoTo CleanUp
                Else
                    m_pNewLineFeedback = Nothing
                    pArc = pGeom
                    m_appEnvision.StatusBar.Message(0) = "Envision Polyline Full Length: " & pArc.Length.ToString("#.####") & " " & m_strUnit
                    WriteValues(Nothing, Nothing, pArc)
                    strMessage = m_appEnvision.StatusBar.Message(0)
                    m_appEnvision.StatusBar.Message(0) = "Envision Polyline Full Length: " & pArc.Length.ToString("#.####") & " " & m_strUnit & "----" & strMessage
                End If
            ElseIf UCase(m_strEnvisionEditOption) = UCase("Polygon") Then
                pGeom = m_pNewPolyFeedback.Stop
                If Not pGeom Is Nothing Then
                    pPoly = pGeom
                    pArea = pPoly
                    pAcres = ReturnAcres(pArea.Area, mxDoc.FocusMap.MapUnits.ToString)
                    WriteValues(Nothing, pPoly, Nothing)
                    strMessage = m_appEnvision.StatusBar.Message(0)
                    strMessage = "Envision Polygon Area: " & pAcres.ToString("#.####") & " Acres" & "; " & strMessage
                    strMessage = strMessage.Replace("-", "")
                    m_appEnvision.StatusBar.Message(0) = strMessage
                End If
                m_pNewPolyFeedback = Nothing
            End If
        Catch ex As Exception

        End Try

CleanUp:
        mxApplication = Nothing
        mxDoc = Nothing
        pGeom = Nothing
        pArc = Nothing
        pPoly = Nothing
        pArea = Nothing
        pAcres = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub

    Public Sub OnKeyDown(ByVal keyCode As Integer, ByVal shift As Integer) Implements ESRI.ArcGIS.SystemUI.ITool.OnKeyDown
        Dim mxApplication As IMxApplication
        Dim pMxDoc As IMxDocument
        If keyCode = 27 Then
            mxApplication = CType(m_appEnvision, IMxApplication)
            pMxDoc = CType(m_appEnvision.Document, IMxDocument)
            m_pNewPolyFeedback = Nothing
            m_pNewLineFeedback = Nothing
            m_pNewCircleFeedback = Nothing
            m_pNewEnvFeedback = Nothing
            If Not m_pEditFeatureLyr Is Nothing Then
                pMxDoc.ActiveView.PartialRefresh(esriViewDrawPhase.esriViewGeoSelection, m_pEditFeatureLyr, pMxDoc.ActiveView.Extent)
            End If
        End If
    End Sub

    Public Sub OnKeyUp(ByVal keyCode As Integer, ByVal shift As Integer) Implements ESRI.ArcGIS.SystemUI.ITool.OnKeyUp

    End Sub

    Public Sub OnMouseDown(ByVal button As Integer, ByVal shift As Integer, ByVal x As Integer, ByVal y As Integer) Implements ESRI.ArcGIS.SystemUI.ITool.OnMouseDown
        'Track the user defined point, polyline, circle or polygon the user defined,
        'but only if the LEFT mouse button is selected
        Dim mxApplication As IMxApplication
        Dim pMxDoc As IMxDocument
        Dim pDisp As IDisplay
        Dim pPnt As IPoint
        Dim pLine As ILine
        Dim pPoly As IPolygon
        Dim pArea As IArea
        Dim pRubberBand As IRubberBand
        Dim pSegmentCollection As ISegmentCollection
        Dim dblAcres As Double
        Dim dblLength As Double

        Try
            m_cursor = Cursors.WaitCursor

            'Display the mouse click point in the status bar
            mxApplication = CType(m_appEnvision, IMxApplication)
            pMxDoc = CType(m_appEnvision.Document, IMxDocument)
            pDisp = pMxDoc.ActivatedView.ScreenDisplay
            m_strUnit = ""
            If pMxDoc.FocusMap.MapUnits = esriUnits.esriDecimalDegrees Then
                m_strUnit = "Decimal Degrees"
            ElseIf pMxDoc.FocusMap.MapUnits = esriUnits.esriFeet Then
                m_strUnit = "Feet"
            ElseIf pMxDoc.FocusMap.MapUnits = esriUnits.esriMiles Then
                m_strUnit = "Miles"
            ElseIf pMxDoc.FocusMap.MapUnits = esriUnits.esriMeters Then
                m_strUnit = "Meters"
            ElseIf pMxDoc.FocusMap.MapUnits = esriUnits.esriKilometers Then
                m_strUnit = "Kilometers"
            End If
            If button = 1 Then
                pPnt = mxApplication.Display.DisplayTransformation.ToMapPoint(x, y)
                If UCase(m_strEnvisionEditOption) = UCase("Point") Then
                    m_appEnvision.StatusBar.Message(0) = "Envision Point Coordinates: " & pPnt.X.ToString() & " (X)," & pPnt.Y.ToString() & " (Y)"
                    WriteValues(pPnt, Nothing, Nothing)
                ElseIf UCase(m_strEnvisionEditOption) = UCase("Polyline") Then
                    If m_pNewLineFeedback Is Nothing Then
                        m_pNewLineFeedback = New NewLineFeedback
                        m_dblFullLength = 0
                        m_pNewLineFeedback.Display = pDisp
                        m_pNewLineFeedback.Start(pPnt)
                        m_pntCollection = New Polyline
                        m_pntCollection.AddPoint(pPnt)
                        m_pntStart = pPnt
                    Else
                        m_pNewLineFeedback.AddPoint(pPnt)
                        m_pntCollection.AddPoint(pPnt)
                        pLine = New Line
                        pLine.FromPoint = m_pntStart
                        pLine.ToPoint = pPnt
                        dblLength = pLine.Length
                        If m_dblFullLength > 0 Then
                            m_appEnvision.StatusBar.Message(0) = "Envision Polyline Full Length: " & " " & m_strUnit & m_dblFullLength.ToString("#.####") & "; Segment Length: " & dblLength.ToString("#.####") & " " & m_strUnit
                        Else
                            m_appEnvision.StatusBar.Message(0) = "Envision Polyline Full Length: " & " " & m_strUnit & dblLength.ToString("#.####") & "; Segment Length: " & dblLength.ToString("#.####") & " " & m_strUnit
                        End If
                        m_dblFullLength = m_dblFullLength + dblLength
                        m_appEnvision.StatusBar.Message(0) = "Envision Polyline Full Length:" & m_dblFullLength.ToString("#.####") & " " & m_strUnit & "; Segment Length: " & dblLength.ToString("#.####") & " " & m_strUnit
                        m_pntStart = pPnt
                    End If
                ElseIf UCase(m_strEnvisionEditOption) = UCase("Rect") Then
                    If m_pNewEnvFeedback Is Nothing Then
                        m_pNewEnvFeedback = New NewEnvelopeFeedback
                        m_pNewEnvFeedback.Display = pDisp
                        m_pNewEnvFeedback.Start(pPnt)
                        m_pntStart = pPnt
                    End If
                ElseIf UCase(m_strEnvisionEditOption) = UCase("Circle") Then
                    If m_pNewCircleFeedback Is Nothing Then
                        m_pNewCircleFeedback = New NewCircleFeedback
                        m_pNewCircleFeedback.Display = pDisp
                        m_pNewCircleFeedback.Start(pPnt)
                        m_pntStart = pPnt
                    End If
                ElseIf UCase(m_strEnvisionEditOption) = UCase("Polygon") Then
                    If m_pNewPolyFeedback Is Nothing Then
                        m_pNewPolyFeedback = New NewPolygonFeedback
                        m_pNewPolyFeedback.Display = pDisp
                        m_pNewPolyFeedback.Start(pPnt)
                        m_pntCollection = New Polygon
                        m_pntCollection.AddPoint(pPnt)
                    Else
                        m_pNewPolyFeedback.AddPoint(pPnt)
                        m_pntCollection.AddPoint(pPnt)
                        pPoly = New Polygon
                        pPoly = m_pntCollection
                        pArea = pPoly
                        dblAcres = ReturnAcres(pArea.Area, pMxDoc.FocusMap.MapUnits)
                        Dim strMessage As String
                        If m_strUnit = "" Then
                            strMessage = "Envision Area:  " & Format(dblAcres, "0.00")
                            strMessage = strMessage.Replace("-", "")
                            m_appEnvision.StatusBar.Message(0) = strMessage
                        Else
                            strMessage = "Envision Area:  " & Format(dblAcres, "0.00") & " Acres"
                            strMessage = strMessage.Replace("-", "")
                            m_appEnvision.StatusBar.Message(0) = strMessage
                        End If
                    End If
                End If
            End If

            GoTo CleanUp
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Envision Mouse Down Event Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo CleanUp
        End Try

CleanUp:
        m_cursor = Cursors.Default
        mxApplication = Nothing
        pMxDoc = Nothing
        pPnt = Nothing
        pPoly = Nothing
        pRubberBand = Nothing
        pSegmentCollection = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()

    End Sub

    Public Sub OnMouseUp(ByVal button As Integer, ByVal shift As Integer, ByVal x As Integer, ByVal y As Integer) Implements ESRI.ArcGIS.SystemUI.ITool.OnMouseUp
        Dim mxApplication As IMxApplication
        Dim pMxDoc As IMxDocument
        Dim pCircArc As IGeometry
        Dim pSegs As ISegmentCollection
        Dim pGeomPoly As IGeometry
        Dim pArea As IArea
        Dim dblAcres As Double
        Dim pPolygon As IPolygon = New Polygon
        Dim pPnt As IPoint
        Dim strMessage As String = ""

        mxApplication = CType(m_appEnvision, IMxApplication)
        pMxDoc = CType(m_appEnvision.Document, IMxDocument)
        pPnt = mxApplication.Display.DisplayTransformation.ToMapPoint(x, y)

        If UCase(m_strEnvisionEditOption) = UCase("Circle") Then
            If Not m_pNewCircleFeedback Is Nothing Then
                pCircArc = m_pNewCircleFeedback.Stop
                pSegs = pPolygon
                pSegs.AddSegment(pCircArc)
                pPolygon.Close()
                pPolygon.SpatialReference = pCircArc.SpatialReference
                pArea = pPolygon
                dblAcres = ReturnAcres(pArea.Area, pMxDoc.FocusMap.MapUnits)
                m_pRad.ToPoint = pPnt
                If m_strUnit = "" Then
                    m_appEnvision.StatusBar.Message(0) = "Envision Radius:  " & Format(m_pRad.Length, "0.00") & "; Area " & m_strUnit & " " & dblAcres
                    WriteValues(Nothing, pPolygon, Nothing)
                    strMessage = m_appEnvision.StatusBar.Message(0)
                    m_appEnvision.StatusBar.Message(0) = "Envision Radius:  " & Format(m_pRad.Length, "0.00") & "; Area " & m_strUnit & " " & dblAcres & "----" & strMessage
                Else
                    m_appEnvision.StatusBar.Message(0) = "Envision Radius:  " & Format(m_pRad.Length, "0.00") & "; Area " & m_strUnit & " " & dblAcres & " Acres"
                    WriteValues(Nothing, pPolygon, Nothing)
                    strMessage = m_appEnvision.StatusBar.Message(0)
                    m_appEnvision.StatusBar.Message(0) = "Envision Radius:  " & Format(m_pRad.Length, "0.00") & "; Area " & m_strUnit & " " & dblAcres & " Acres" & "----" & strMessage
                End If
                m_pNewCircleFeedback = Nothing
            End If
        ElseIf UCase(m_strEnvisionEditOption) = UCase("rect") Then
            ' Check if the user is currently using the feedback
            If Not m_pNewEnvFeedback Is Nothing Then
                pGeomPoly = m_pNewEnvFeedback.Stop
                m_pNewEnvFeedback = Nothing
                pArea = pGeomPoly
                pPolygon = pGeomPoly
                dblAcres = ReturnAcres(pArea.Area, pMxDoc.FocusMap.MapUnits)
                If m_strUnit = "" Then
                    m_appEnvision.StatusBar.Message(0) = "Envision Area:  " & Format(dblAcres, "0.00")
                    WriteValues(Nothing, pPolygon, Nothing)
                    strMessage = m_appEnvision.StatusBar.Message(0)
                    m_appEnvision.StatusBar.Message(0) = "Envision Area:  " & Format(dblAcres, "0.00") & "----" & strMessage
                Else
                    m_appEnvision.StatusBar.Message(0) = "Envision Area:  " & Format(dblAcres, "0.00") & " Acres"
                    WriteValues(Nothing, pPolygon, Nothing)
                    strMessage = m_appEnvision.StatusBar.Message(0)
                    m_appEnvision.StatusBar.Message(0) = "Envision Area:  " & Format(dblAcres, "0.00") & " Acres" & "----" & strMessage
                End If
            End If

        End If

    End Sub

    Public Sub OnMouseMove(ByVal button As Integer, ByVal shift As Integer, ByVal x As Integer, ByVal y As Integer) Implements ESRI.ArcGIS.SystemUI.ITool.OnMouseMove
        ' Check if the user is currently using the feedback
        Dim mxApplication As IMxApplication
        Dim pMxDoc As IMxDocument
        Dim pPnt As IPoint
        Dim pLine As ILine
        Dim dblLength As Double
        Dim pArea As IArea
        Dim dblAcres As Double = 0
        Dim pPoly As IPolygon
        Dim pRect As IEnvelope
        Dim pSegs As ISegmentCollection

        mxApplication = CType(m_appEnvision, IMxApplication)
        pMxDoc = CType(m_appEnvision.Document, IMxDocument)
        pPnt = mxApplication.Display.DisplayTransformation.ToMapPoint(x, y)
        If UCase(m_strEnvisionEditOption) = UCase("Polyline") Then
            If Not m_pNewLineFeedback Is Nothing Then
                m_pNewLineFeedback.MoveTo(pPnt)
                pLine = New Line
                pLine.FromPoint = m_pntStart
                pLine.ToPoint = pPnt
                dblLength = pLine.Length
                If m_dblFullLength > 0 Then
                    m_appEnvision.StatusBar.Message(0) = "Envision Polyline Full Length: " & m_dblFullLength.ToString("#.####") & " " & m_strUnit & "; Segment Length: " & dblLength.ToString("#.####") & " " & m_strUnit
                Else
                    m_appEnvision.StatusBar.Message(0) = "Envision Polyline Full Length: " & dblLength.ToString("#.####") & " " & m_strUnit & "; Segment Length: " & dblLength.ToString("#.####") & " " & m_strUnit
                End If
            End If
        ElseIf UCase(m_strEnvisionEditOption) = UCase("Rect") Then
            If Not m_pNewEnvFeedback Is Nothing Then
                m_pNewEnvFeedback.MoveTo(pPnt)
                pRect = New Envelope
                pRect.XMin = m_pntStart.X
                pRect.YMin = m_pntStart.Y
                pRect.XMax = pPnt.X
                pRect.YMax = pPnt.Y
                pArea = pRect
                dblAcres = ReturnAcres(pArea.Area, pMxDoc.FocusMap.MapUnits)
                If m_strUnit = "" Then
                    m_appEnvision.StatusBar.Message(0) = "Envision Area:  " & Format(dblAcres, "0.00")
                Else
                    m_appEnvision.StatusBar.Message(0) = "Envision Area:  " & Format(dblAcres, "0.00") & " Acres"
                End If
            End If
        ElseIf UCase(m_strEnvisionEditOption) = UCase("Circle") Then
            If Not m_pNewCircleFeedback Is Nothing Then
                m_pNewCircleFeedback.MoveTo(pPnt)
                m_pRad = New Line
                m_pRad.FromPoint = m_pntStart
                m_pRad.ToPoint = pPnt

                m_pRad.ToPoint = pPnt
                pPoly = New Polygon
                pSegs = pPoly
                pSegs.SetCircle(m_pntStart, m_pRad.Length)
                pPoly.Close()
                pArea = pPoly
                dblAcres = ReturnAcres(pArea.Area, pMxDoc.FocusMap.MapUnits)

                If m_strUnit = "" Then
                    m_appEnvision.StatusBar.Message(0) = "Envision Radius:  " & Format(m_pRad.Length, "0.00") & "; Area " & m_strUnit & " " & dblAcres
                Else
                    m_appEnvision.StatusBar.Message(0) = "Envision Radius:  " & Format(m_pRad.Length, "0.00") & "; Area " & m_strUnit & " " & dblAcres & " Acres"
                End If
            End If
        ElseIf UCase(m_strEnvisionEditOption) = UCase("Polygon") Then
            If Not m_pNewPolyFeedback Is Nothing Then
                m_pNewPolyFeedback.MoveTo(pPnt)
            End If
        End If

    End Sub

    Public Sub Refresh(ByVal hdc As Integer) Implements ESRI.ArcGIS.SystemUI.ITool.Refresh
        m_pDoc = m_appEnvision.Document
        m_pAV = m_pDoc.ActiveView
        m_pScrD = m_pAV.ScreenDisplay
    End Sub
End Class

<Guid("D425C102-74A7-4D99-A4A4-C7D38833FEDA")> _
Public Class EnvisionProjectSetup

    Implements ICommand
    Private m_pApp As IApplication
    Private m_application As IApplication
    Private m_bitmap As Bitmap
    Private m_hBitmap As IntPtr
    Private m_cursor As System.Windows.Forms.Cursor
    Private m_hCursor As IntPtr

    ' Needed to clear up the Hbitmap unmanaged resource
    <DllImport("gdi32.dll")> _
  Private Shared Function DeleteObject(ByVal hObject As IntPtr) As Boolean
    End Function

#Region "Component Category Registration"
    <ComRegisterFunction()> _
    Public Shared Sub Reg(ByVal regKey As String)
        ESRI.ArcGIS.ADF.CATIDs.MxCommands.Register(regKey)
    End Sub

    <ComUnregisterFunction()> _
    Public Shared Sub Unreg(ByVal regKey As String)
        ESRI.ArcGIS.ADF.CATIDs.MxCommands.Unregister(regKey)
    End Sub
#End Region

    Public Sub New()

        'm_bitmap = New System.Drawing.Bitmap((Me.GetType.Assembly.GetManifestResourceStream("EnvisionToolbar2.Pin_Red.bmp")))
        'If Not (m_bitmap Is Nothing) Then
        '    m_bitmap.MakeTransparent(m_bitmap.GetPixel(1, 1))
        '    m_hBitmap = m_bitmap.GetHbitmap()
        'End If

    End Sub

    Protected Overrides Sub Finalize()

        If m_hBitmap.ToInt32 <> 0 Then
            DeleteObject(m_hBitmap)
        End If

        If m_hCursor.ToInt32() <> 0 Then
            DeleteObject(m_hCursor)
        End If

        If Not m_application Is Nothing Then
            Marshal.ReleaseComObject(m_application)
        End If

    End Sub

    Public ReadOnly Property Bitmap() As Integer Implements ESRI.ArcGIS.SystemUI.ICommand.Bitmap
        Get
            ' Set the bitmap of the command. The m_hBitmap variable is set
            ' in constructor (New) of this class.

            Return m_hBitmap.ToInt32()
        End Get
    End Property

    Public ReadOnly Property Caption() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Caption
        Get
            ' Set the string that appears when the command is used as a
            ' menu item.
            Return "ET Project Setup"
        End Get
    End Property

    Public ReadOnly Property Category() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Category
        Get
            ' Set the category of this command. This determines where the
            ' command appears in the Commands panel of the Customize dialog.
            Return "EnvisionTools3"
        End Get
    End Property

    Public ReadOnly Property Checked() As Boolean Implements ESRI.ArcGIS.SystemUI.ICommand.Checked
        Get

CleanUp:
            GC.WaitForPendingFinalizers()
            GC.Collect()
            Return False

        End Get
    End Property

    Public ReadOnly Property Enabled() As Boolean Implements ESRI.ArcGIS.SystemUI.ICommand.Enabled
        Get
            ' Add some logic here to specify when the command should
            ' be enabled. In this example, the command is always enabled.
            Return True
        End Get
    End Property

    Public ReadOnly Property HelpContextID() As Integer Implements ESRI.ArcGIS.SystemUI.ICommand.HelpContextID
        Get
            Return 0
        End Get
    End Property

    Public ReadOnly Property HelpFile() As String Implements ESRI.ArcGIS.SystemUI.ICommand.HelpFile
        Get
            Return ""
        End Get
    End Property

    Public ReadOnly Property Message() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Message
        Get

            ' Set the message string that appears in the statusbar of the
            ' application when the mouse passes over the command.
            Return "Creates a custom dialog to create Envision Project"
        End Get
    End Property

    Public ReadOnly Property Name() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Name
        Get
            ' Set the internal name of this command. By convention, this
            ' name string contains the category and caption of the command.
            Return "EnvisionTools_EnvisionSetup3"

        End Get
    End Property

    Public Sub OnClick() Implements ESRI.ArcGIS.SystemUI.ICommand.OnClick
        m_blnETSetupFormIsOpening = True
        'CHECK FOR LICENSE FILE AND OPEN PROJECT SETUP FORM
        Try
            If m_frmEnvisionProjectSetup Is Nothing Then
                m_frmEnvisionProjectSetup = New frmEnvisionProjectSetup
            End If
            Try
                m_frmEnvisionProjectSetup.Show()
            Catch ex As Exception
                GoTo CleanUp
            End Try


            GoTo CleanUp

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Envision Setup Tool 'OnClick' Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            GoTo cleanup
        End Try
CleanUp:
        m_blnETSetupFormIsOpening = False
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub

    Public Sub OnCreate(ByVal hook As Object) Implements ESRI.ArcGIS.SystemUI.ICommand.OnCreate
        ' The hook argument is a pointer to Application object.
        ' Establish a hook to the application
        'm_appEnvision = CType(hook, IApplication)

    End Sub

    Public ReadOnly Property Tooltip() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Tooltip
        Get
            ' Set the string that appears in the screen tip.
            Return "ET Project Setup Tool"
        End Get
    End Property

End Class
