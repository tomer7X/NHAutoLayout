Imports System.Windows.Forms
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.PlottingServices
Namespace MyNamespace
    Public Class LayoutsClass
        Public Shared Function GetLayoutList() As List(Of String)
            Dim dwgLayouts As New List(Of String)
            Try
                Dim doc As Document = Core.Application.DocumentManager.MdiActiveDocument
                Dim db As Database = doc.Database

                Dim layAndTab As SortedDictionary(Of Integer, String) = New SortedDictionary(Of Integer, String)

                Using tr As Transaction = db.TransactionManager.StartTransaction()
                    Dim layoutDic As DBDictionary = TryCast(tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead, False), DBDictionary)

                    For Each entry As DBDictionaryEntry In layoutDic
                        Dim layoutid As ObjectId = entry.Value
                        Dim layout As Layout = TryCast(tr.GetObject(layoutid, OpenMode.ForRead), Layout)
                        layAndTab.Add(layout.TabOrder, layout.LayoutName)
                    Next
                    tr.Commit()
                End Using

                For Each v In layAndTab.Values
                    dwgLayouts.Add(v)
                Next

                Return dwgLayouts
            Catch ex As Exception
                dwgLayouts.Add("")
                Return dwgLayouts
            End Try

        End Function

        Public Shared Function GetLayouts() As List(Of Layout)
            Dim dwgLayouts As New List(Of Layout)
            Try
                Dim doc As Document = Core.Application.DocumentManager.MdiActiveDocument
                Dim db As Database = doc.Database

                Dim layAndTab As SortedDictionary(Of Integer, Layout) = New SortedDictionary(Of Integer, Layout)

                Using tr As Transaction = db.TransactionManager.StartTransaction()
                    Dim layoutDic As DBDictionary = TryCast(tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead, False), DBDictionary)

                    For Each entry As DBDictionaryEntry In layoutDic
                        Dim layoutid As ObjectId = entry.Value
                        Dim layout As Layout = TryCast(tr.GetObject(layoutid, OpenMode.ForRead), Layout)
                        layAndTab.Add(layout.TabOrder, layout)
                    Next
                    tr.Commit()
                End Using

                For Each v In layAndTab.Values
                    If v.LayoutName <> "Model" Then
                        dwgLayouts.Add(v)
                    End If
                Next

                Return dwgLayouts
            Catch ex As Exception
                dwgLayouts.Add(Nothing)
                Return dwgLayouts
            End Try

        End Function

        Public Shared Sub SetCurrentLayoutTab(tab As String)
            LayoutManager.Current.CurrentLayout = tab
        End Sub

        Public Shared Function GetCurrentLayoutTab()
            Dim tab As String = LayoutManager.Current.CurrentLayout
            Return tab
        End Function

        Public Shared Sub AddLayout(LayoutName As String)
            Dim doc = Core.Application.DocumentManager.MdiActiveDocument
            Dim LayoutList As List(Of String) = GetLayoutList()
            Using trans As Transaction = doc.TransactionManager.StartTransaction()
                If LayoutList.Contains(LayoutName) = False Then
                    Using doc.LockDocument()
                        LayoutManager.Current.CreateLayout(LayoutName)

                    End Using
                Else
                    MessageBox.Show("Layout " & LayoutName & " already exists.")
                End If
                trans.Commit()
            End Using

        End Sub
        Public Shared Sub DeleteLayout(LayoutName As String)
            Dim doc = Core.Application.DocumentManager.MdiActiveDocument
            Using trans As Transaction = doc.TransactionManager.StartTransaction()
                If LayoutName = "Model" Then
                    doc.Editor.WriteMessage("Modelspace could not be deleted.")
                Else
                    LayoutManager.Current.DeleteLayout(LayoutName)
                End If
                trans.Commit()
            End Using
        End Sub

        Public Shared Sub DeleteLayouts(LayoutNames As List(Of String))
            Dim doc = Core.Application.DocumentManager.MdiActiveDocument
            Using trans As Transaction = doc.TransactionManager.StartTransaction()
                For Each LayoutName In LayoutNames
                    If LayoutName = "Model" Then
                        doc.Editor.WriteMessage("Modelspace could not be deleted.")
                    Else
                        LayoutManager.Current.DeleteLayout(LayoutName)
                    End If
                Next
                trans.Commit()
            End Using
        End Sub

        Public Shared Sub DeleteCurrentLayoutEntities(LayoutName As String)
            Dim doc As Document = Core.Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Dim ed As Editor = doc.Editor
            Using trans As Transaction = doc.TransactionManager.StartTransaction()
                Try
                    Dim tv As TypedValue()
                    tv = New TypedValue(0) {}
                    tv.SetValue(New TypedValue(CInt(DxfCode.LayoutName), LayoutName), 0)

                    Dim filter As New SelectionFilter(tv)
                    Dim ssPrompt As PromptSelectionResult = ed.SelectAll(filter)
                    Dim ss As SelectionSet = ssPrompt.Value

                    If ssPrompt.Status = PromptStatus.OK Then

                        Dim blks = ss.GetObjectIds()
                        For Each sObj As ObjectId In blks

                            Dim ent As Entity = TryCast(trans.GetObject(sObj, OpenMode.ForWrite), Entity)
                            ent.Erase()

                        Next
                        trans.Commit()

                    Else
                        ed.WriteMessage("No object selected.")
                        trans.Abort()
                    End If

                Catch ex As Exception
                    ed.WriteMessage("Error encountered : " & ex.Message)
                    trans.Abort()
                End Try
            End Using
        End Sub

        Public Shared Sub CreateViewport(CenterPoint As Point3d, Width As Double, Height As Double, ViewCenter As Point2d, Scale As Double, Optional LayerName As String = "")
            Dim doc As Document = Core.Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Dim ed As Editor = doc.Editor

            Dim idCol As ObjectIdCollection = New ObjectIdCollection

            If LayerName = "" Then
                LayerName = "Defpoints"
            End If

            Using trans As Transaction = db.TransactionManager.StartTransaction
                Dim idLay As ObjectId = db.LayerTableId
                Dim lt As LayerTable = trans.GetObject(idLay, OpenMode.ForRead)


                If lt.Has("P_Names") Then
                    idCol.Add(lt.Item("P_Names"))
                End If
                If lt.Has("Quantity") Then
                    idCol.Add(lt.Item("Quantity"))
                End If
                If lt.Has("Defpoints") Then
                    idCol.Add(lt.Item("Defpoints"))
                End If

                Dim bt As BlockTable = trans.GetObject(db.BlockTableId, OpenMode.ForRead)
                Dim btr As BlockTableRecord = trans.GetObject(bt(BlockTableRecord.PaperSpace), OpenMode.ForWrite)
                ed.SwitchToPaperSpace()

                Dim vp As Viewport = New Viewport()
                vp.Layer = LayerName
                vp.SetDatabaseDefaults()
                vp.CenterPoint = CenterPoint
                vp.Width = Width
                vp.Height = Height
                vp.ViewCenter = ViewCenter
                vp.ViewHeight = Height / Scale
                vp.FreezeLayersInViewport(idCol.GetEnumerator)

                btr.AppendEntity(vp)
                trans.AddNewlyCreatedDBObject(vp, True)

                vp.On = True

                trans.Commit()

            End Using

        End Sub
        Public Shared Sub CreateOrEditPageSetup()
            Dim doc As Document = Core.Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Dim ed As Editor = doc.Editor

            Using trans As Transaction = db.TransactionManager.StartTransaction()

                Dim plSets As DBDictionary =
                trans.GetObject(db.PlotSettingsDictionaryId, OpenMode.ForRead)
                Dim vStyles As DBDictionary =
                trans.GetObject(db.VisualStyleDictionaryId, OpenMode.ForRead)

                Dim plSet As PlotSettings
                Dim createNew As Boolean = False

                Dim acLayoutMgr As LayoutManager = LayoutManager.Current

                Dim acLayout As Layout =
                trans.GetObject(acLayoutMgr.GetLayoutId(acLayoutMgr.CurrentLayout),
                                  OpenMode.ForRead)

                If plSets.Contains("TH_A3 PDF") = False Then
                    createNew = True

                    plSet = New PlotSettings(acLayout.ModelType)
                    plSet.CopyFrom(acLayout)

                    plSet.PlotSettingsName = "TH_A3 PDF"
                    plSet.AddToPlotSettingsDictionary(db)
                    trans.AddNewlyCreatedDBObject(plSet, True)
                Else
                    plSet = plSets.GetAt("TH_A3 PDF").GetObject(OpenMode.ForWrite)
                End If

                Try
                    Dim acPlSetVdr As PlotSettingsValidator = PlotSettingsValidator.Current

                    acPlSetVdr.SetPlotConfigurationName(plSet,
                                                        "DWG To PDF.pc3",
                                                        "ISO_full_bleed_A3_(420.00_x_297.00_MM)")

                    acPlSetVdr.SetPlotType(plSet,
                                           Autodesk.AutoCAD.DatabaseServices.PlotType.Extents)

                    acPlSetVdr.SetPlotCentered(plSet, True)
                    acPlSetVdr.SetPlotOrigin(plSet, New Point2d(0, 0))

                    acPlSetVdr.SetUseStandardScale(plSet, True)
                    acPlSetVdr.SetPlotPaperUnits(plSet, PlotPaperUnit.Millimeters)
                    acPlSetVdr.SetStdScaleType(plSet, StdScaleType.ScaleToFit)
                    plSet.ScaleLineweights = False

                    plSet.ShowPlotStyles = True

                    acPlSetVdr.RefreshLists(plSet)

                    plSet.ShadePlot = PlotSettingsShadePlotType.AsDisplayed
                    plSet.ShadePlotResLevel = ShadePlotResLevel.Normal

                    plSet.PrintLineweights = True
                    plSet.PlotTransparency = False
                    plSet.PlotPlotStyles = True
                    plSet.DrawViewportsFirst = True

                    acPlSetVdr.SetPlotRotation(plSet, PlotRotation.Degrees000)

                    If db.PlotStyleMode = True Then
                        acPlSetVdr.SetCurrentStyleSheet(plSet, "monochrome.ctb")
                    Else
                        acPlSetVdr.SetCurrentStyleSheet(plSet, "monochrome.stb")
                    End If

                    acPlSetVdr.SetZoomToPaperOnUpdate(plSet, True)

                Catch es As Autodesk.AutoCAD.Runtime.Exception
                    ed.WriteMessage("Error encountered : " & es.Message)
                End Try

                trans.Commit()
                If createNew = True Then
                    plSet.Dispose()
                End If
            End Using
        End Sub

        Public Shared Sub AssignPageSetupToLayout()
            Dim doc As Document = Core.Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database

            Using trans As Transaction = db.TransactionManager.StartTransaction()
                Dim lm As LayoutManager = LayoutManager.Current

                Dim acLayout As Layout =
                        trans.GetObject(lm.GetLayoutId(lm.CurrentLayout),
                                           OpenMode.ForRead)

                Dim acPlSet As DBDictionary =
                        trans.GetObject(db.PlotSettingsDictionaryId, OpenMode.ForRead)

                If acPlSet.Contains("TH_A3 PDF") = True Then
                    Dim plSet As PlotSettings =
                            acPlSet.GetAt("TH_A3 PDF").GetObject(OpenMode.ForRead)

                    trans.GetObject(lm.GetLayoutId(lm.CurrentLayout), OpenMode.ForWrite)
                    acLayout.CopyFrom(plSet)

                    trans.Commit()
                Else
                    trans.Abort()
                End If
            End Using

            doc.Editor.Regen()
        End Sub

        Public Shared Sub SetPageSetupToLayout()
            Dim doc As Document = Core.Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database

            Using trans As Transaction = db.TransactionManager.StartTransaction()
                Dim lm As LayoutManager = LayoutManager.Current

                Dim acLayout As Layout =
                    trans.GetObject(lm.GetLayoutId(lm.CurrentLayout),
                                      OpenMode.ForRead)

                Dim acPlSet As DBDictionary =
                    trans.GetObject(db.PlotSettingsDictionaryId, OpenMode.ForRead)

                If acPlSet.Contains("TH_A3 PDF") = False Then CreateOrEditPageSetup()

                Dim plSet As PlotSettings =
                        acPlSet.GetAt("TH_A3 PDF").GetObject(OpenMode.ForRead)

                trans.GetObject(lm.GetLayoutId(lm.CurrentLayout), OpenMode.ForWrite)
                acLayout.CopyFrom(plSet)

                trans.Commit()
            End Using

            doc.Editor.Regen()
        End Sub

        Public Shared Sub SwapLayoutTabOrder(layoutName1 As String, layoutName2 As String)
            Dim doc As Document = Core.Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database

            Using tr As Transaction = db.TransactionManager.StartTransaction()
                Dim layoutDic As DBDictionary = TryCast(tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead, False), DBDictionary)

                Dim layout1 As Layout = Nothing
                Dim layout2 As Layout = Nothing

                For Each entry As DBDictionaryEntry In layoutDic
                    Dim lo As Layout = TryCast(tr.GetObject(entry.Value, OpenMode.ForRead), Layout)
                    If lo.LayoutName = layoutName1 Then layout1 = lo
                    If lo.LayoutName = layoutName2 Then layout2 = lo
                Next

                If layout1 IsNot Nothing AndAlso layout2 IsNot Nothing Then
                    Dim tab1 As Integer = layout1.TabOrder
                    Dim tab2 As Integer = layout2.TabOrder

                    layout1.UpgradeOpen()
                    layout2.UpgradeOpen()
                    layout1.TabOrder = tab2
                    layout2.TabOrder = tab1
                End If

                tr.Commit()
            End Using
        End Sub

    End Class
End Namespace