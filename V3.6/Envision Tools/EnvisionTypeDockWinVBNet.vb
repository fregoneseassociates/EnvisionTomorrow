' Copyright 1995-2004 ESRI

' All rights reserved under the copyright laws of the United States.

' You may freely redistribute and use this sample code, with or without modification.

' Disclaimer: THE SAMPLE CODE IS PROVIDED "AS IS" AND ANY EXPRESS OR IMPLIED 
' WARRANTIES, INCLUDING THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS 
' FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL ESRI OR 
' CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, 
' OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF 
' SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS 
' INTERRUPTION) SUSTAINED BY YOU OR A THIRD PARTY, HOWEVER CAUSED AND ON ANY 
' THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT ARISING IN ANY 
' WAY OUT OF THE USE OF THIS SAMPLE CODE, EVEN IF ADVISED OF THE POSSIBILITY OF 
' SUCH DAMAGE.

' For additional information contact: Environmental Systems Research Institute, Inc.

' Attn: Contracts Dept.

' 380 New York Street

' Redlands, California, U.S.A. 92373 

' Email: contracts@esri.com
Imports System.Runtime.InteropServices
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.Geodatabase
'Imports ESRI.ArcGIS.Utility.CATIDs

<ComClass(EnvisionDockWinVBNet.ClassId, EnvisionDockWinVBNet.InterfaceId, EnvisionDockWinVBNet.EventsId)> _
Public Class EnvisionDockWinVBNet
    Implements IDockableWindowDef

#Region "Register Unregister Component"
    <ComRegisterFunction(), ComVisibleAttribute(False)> _
    Public Shared Sub OnRegister(ByVal regKey As String)
        ESRI.ArcGIS.ADF.CATIDs.MxDockableWindows.Register(regKey)
    End Sub

    <ComUnregisterFunction(), ComVisibleAttribute(False)> _
    Public Shared Sub OnUnRegister(ByVal regKey As String)
        ESRI.ArcGIS.ADF.CATIDs.MxDockableWindows.Unregister(regKey)
    End Sub
#End Region

#Region "COM GUIDs"
    Public Const ClassId As String = "E37A794C-F661-41C1-883B-A05648131221"
    Public Const InterfaceId As String = "55B154B5-4B5D-4A41-BEAE-8D3B558EFD6E"
    Public Const EventsId As String = "FDEEC565-02F5-4F3A-B0E1-50BFAB430FF8"
#End Region

    Private m_app As IApplication
    Private m_mxDoc As IMxDocument
    Private m_mxDocument As MxDocument
    Private m_map As Map
    Private m_pageLayout As PageLayout
    'Private m_dockWinForm As frmEnvisionAttributeEditor

    Private DNewDocE As IDocumentEvents_NewDocumentEventHandler
    Private DOpenDocE As IDocumentEvents_OpenDocumentEventHandler
    Private DMapsChangedE As IDocumentEvents_MapsChangedEventHandler
    Private DSelChangedE As IActiveViewEvents_SelectionChangedEventHandler
    Private DFocusMapChangedE As IActiveViewEvents_FocusMapChangedEventHandler

    Public Sub New()
        MyBase.New()
    End Sub

    Public ReadOnly Property Caption() As String Implements ESRI.ArcGIS.Framework.IDockableWindowDef.Caption
        Get
            Caption = "ENVISION PAINT TOOLS"
        End Get
    End Property

    Public ReadOnly Property ChildHWND() As Integer Implements ESRI.ArcGIS.Framework.IDockableWindowDef.ChildHWND
        Get
            ' The dockable window will consist of a list box, so pass back the hWnd of
            ' the listbox on a form
            ChildHWND = m_dockEnvisionWinForm.Handle.ToInt32
        End Get
    End Property

    Public Sub OnCreate(ByVal hook As Object) Implements ESRI.ArcGIS.Framework.IDockableWindowDef.OnCreate
        m_dockEnvisionWinForm = New ctlEnvisionAttributesEditor
        ' The hook argument is a pointer to Application object.
        ' Establish a hook to the application
        If TypeOf hook Is IMxApplication Then
            m_app = hook

            m_mxDoc = m_app.Document
            m_mxDocument = m_app.Document
            m_map = m_mxDoc.FocusMap
            m_pageLayout = m_mxDoc.PageLayout

            If m_appEnvision Is Nothing Then m_appEnvision = CType(hook, IApplication)

            'SetupEvents()
            'OnSelectionChanged()
        End If
    End Sub

    Public ReadOnly Property Name() As String Implements ESRI.ArcGIS.Framework.IDockableWindowDef.Name
        Get
            Name = "ENVISION PAINT TOOLS 2.0"
        End Get
    End Property

    Public Sub OnDestroy() Implements ESRI.ArcGIS.Framework.IDockableWindowDef.OnDestroy
        m_map = Nothing
        m_pageLayout = Nothing
        m_mxDocument = Nothing
        m_mxDoc = Nothing
        m_app = Nothing
    End Sub

    Public ReadOnly Property UserData() As Object Implements ESRI.ArcGIS.Framework.IDockableWindowDef.UserData
        Get
            UserData = ""
        End Get
    End Property

    Private Sub SetupEvents()
        ' Add delegates to events

        ' Cast to the non-default event interface
        DNewDocE = New IDocumentEvents_NewDocumentEventHandler(AddressOf OnNewOpenDoc)
        AddHandler CType(m_mxDocument, IDocumentEvents_Event).NewDocument, DNewDocE

        DOpenDocE = New IDocumentEvents_OpenDocumentEventHandler(AddressOf OnNewOpenDoc)
        AddHandler CType(m_mxDocument, IDocumentEvents_Event).OpenDocument, DOpenDocE

        DMapsChangedE = New IDocumentEvents_MapsChangedEventHandler(AddressOf OnMapsChanged)
        AddHandler CType(m_mxDocument, IDocumentEvents_Event).MapsChanged, DMapsChangedE

        DFocusMapChangedE = New IActiveViewEvents_FocusMapChangedEventHandler(AddressOf OnMapsChanged)
        AddHandler m_pageLayout.FocusMapChanged, DFocusMapChangedE

        DSelChangedE = New IActiveViewEvents_SelectionChangedEventHandler(AddressOf OnSelectionChanged)
        AddHandler m_map.SelectionChanged, DSelChangedE
    End Sub

    Private Sub OnNewOpenDoc()
        If Not m_mxDocument Is m_app.Document Then
            m_mxDocument = m_app.Document
            AddHandler CType(m_mxDocument, IDocumentEvents_Event).NewDocument, DNewDocE
            AddHandler CType(m_mxDocument, IDocumentEvents_Event).OpenDocument, DOpenDocE
            AddHandler CType(m_mxDocument, IDocumentEvents_Event).MapsChanged, DMapsChangedE
        End If

        If Not m_map Is m_mxDoc.FocusMap Then
            m_map = m_mxDoc.FocusMap
            AddHandler m_map.SelectionChanged, DSelChangedE
        End If

        If Not m_pageLayout Is m_mxDoc.PageLayout Then
            m_pageLayout = m_mxDoc.PageLayout
            AddHandler m_pageLayout.FocusMapChanged, DFocusMapChangedE
        End If

        'OnSelectionChanged()
    End Sub

    Private Sub OnMapsChanged()
        If Not m_map Is m_mxDoc.FocusMap Then
            m_map = m_mxDoc.FocusMap
            AddHandler m_map.SelectionChanged, DSelChangedE
        End If

        'OnSelectionChanged()
    End Sub

    Private Sub OnSelectionChanged()
        'Dim focusMap As IMap
        'Dim i As Integer
        'Dim featLayer As IFeatureLayer
        'Dim featSel As IFeatureSelection

        'focusMap = m_map
        'm_dockWinForm.ListSelCount.Items.Clear()

        ''Loop through the layers in the map and add the layer's name and
        ''selection count to the list box
        'For i = 0 To focusMap.LayerCount - 1
        '    If TypeOf focusMap.Layer(i) Is IFeatureSelection Then
        '        featLayer = focusMap.Layer(i)
        '        featSel = featLayer
        '        m_dockWinForm.ListSelCount.Items.Add(featLayer.Name & ": " & featSel.SelectionSet.Count)
        '    End If
        'Next
    End Sub
End Class