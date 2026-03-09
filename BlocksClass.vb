Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime
Imports System.IO

Namespace MyNamespace
    Public Class BlocksClass
        Public Shared Function GetBlockNames() As List(Of String)
            Dim BlocksList As New List(Of String)
            Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database

            Using trans As Transaction = db.TransactionManager.StartTransaction()
                Dim bt As BlockTable = TryCast(trans.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)

                For Each b In bt
                    Dim btr As BlockTableRecord = TryCast(trans.GetObject(b, OpenMode.ForRead), BlockTableRecord)
                    BlocksList.Add(btr.Name)
                Next
                trans.Commit()
            End Using

            Return BlocksList
        End Function

        Public Shared Function InsertBlock(BlockName As String, ptPosition As Point3d, Optional inPaperSpace As Boolean = False, Optional LayoutName As String = "", Optional lname As String = "0")
            Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            Dim db As Autodesk.AutoCAD.DatabaseServices.Database
            Dim BlockID As New ObjectId

            db = Application.DocumentManager.MdiActiveDocument.Database

            Using acLckDocCur As DocumentLock = doc.LockDocument
                Using trans As Transaction = db.TransactionManager.StartTransaction()
                    ' Open the Block table for read
                    Dim acBlkTbl As BlockTable
                    acBlkTbl = trans.GetObject(db.BlockTableId, OpenMode.ForRead)

                    Dim blkRecId As ObjectId = ObjectId.Null

                    If acBlkTbl.Has(BlockName) Then
                        blkRecId = acBlkTbl(BlockName)
                    Else
                        Return Nothing
                        'Exit Function
                    End If

                    ' Insert the block reference into the current space
                    If blkRecId <> ObjectId.Null Then
                        Using acBlkRef As New BlockReference(New Point3d(0, 0, 0), blkRecId)
                            With acBlkRef
                                .SetDatabaseDefaults()
                                .Visible = True
                                .Position = ptPosition
                                .ScaleFactors = New Scale3d(1, 1, 1)
                                .Rotation = 0
                                .Layer = lname
                            End With

                            Dim acBlkTblRec As BlockTableRecord

                            'Attempt #1
                            'acBlkTblRec = acTrans.GetObject(acCurDb.CurrentSpaceId, OpenMode.ForWrite)

                            'Attempt #2
                            If inPaperSpace Then
                                'acBlkTblRec = acTrans.GetObject(acBlkTbl(BlockTableRecord.PaperSpace), OpenMode.ForWrite)

                                'Attempt #3
                                If LayoutName.Length = 0 Then
                                    Return Nothing
                                    'Exit Function
                                End If

                                Dim loMgr As LayoutManager = LayoutManager.Current
                                Dim loID As ObjectId = loMgr.GetLayoutId(LayoutName)
                                Dim lo As Layout = trans.GetObject(loID, OpenMode.ForRead)
                                acBlkTblRec = trans.GetObject(lo.BlockTableRecordId, OpenMode.ForWrite)
                            Else
                                acBlkTblRec = trans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                            End If

                            acBlkTblRec.AppendEntity(acBlkRef)
                            trans.AddNewlyCreatedDBObject(acBlkRef, True)

                            'add attribute definitions
                            Dim btr As BlockTableRecord = blkRecId.GetObject(OpenMode.ForRead)
                            For Each objID As ObjectId In btr
                                Dim obj As DBObject = objID.GetObject(OpenMode.ForRead)
                                If TypeOf obj Is AttributeDefinition Then
                                    Dim ad As AttributeDefinition = objID.GetObject(OpenMode.ForRead)
                                    Dim ar As AttributeReference = New AttributeReference()
                                    ar.SetAttributeFromBlock(ad, acBlkRef.BlockTransform)
                                    ar.Position = ad.Position.TransformBy(acBlkRef.BlockTransform)
                                    acBlkRef.AttributeCollection.AppendAttribute(ar)
                                    trans.AddNewlyCreatedDBObject(ar, True)
                                End If
                            Next

                            BlockID = acBlkRef.ObjectId
                        End Using
                    End If

                    ' Save the new object to the database
                    trans.Commit()

                    ' Dispose of the transaction
                    'InsertBlock = True
                End Using
            End Using
            Return BlockID
        End Function

        Public Shared Sub SetBlockAttributes(BlockID As ObjectId, aTag As List(Of String), aValue As List(Of String), Optional tabName As String = "", Optional SetGlobal As Boolean = False)
            Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Dim ed As Editor = doc.Editor
            If tabName = "" Then
                tabName = LayoutManager.Current.CurrentLayout
            End If

            Using trans As Transaction = doc.TransactionManager.StartTransaction()
                Try
                    Dim blk As BlockReference = TryCast(trans.GetObject(BlockID, OpenMode.ForRead), BlockReference)

                    If blk.ObjectId = BlockID Then
                        For Each attRefID As ObjectId In blk.AttributeCollection
                            Dim obj As DBObject = trans.GetObject(attRefID, OpenMode.ForWrite)
                            Dim attRef As AttributeReference = TryCast(obj, AttributeReference)

                            For i = 0 To aTag.Count - 1
                                If attRef.Tag = aTag(i) Then
                                    attRef.TextString = aValue(i)
                                End If
                            Next
                        Next
                    End If

                    trans.Commit()

                Catch ex As Exception
                    ed.WriteMessage("Error encountered : " & ex.Message)
                    trans.Abort()
                End Try
            End Using
        End Sub
        Public Shared Function CheckBlockAttributes(BlockID As ObjectId, aTag As List(Of String), Optional tabName As String = "", Optional SetGlobal As Boolean = False) As List(Of String)
            Dim TagResult As List(Of String) = New List(Of String)
            Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Dim ed As Editor = doc.Editor
            If tabName = "" Then
                tabName = LayoutManager.Current.CurrentLayout
            End If

            Using trans As Transaction = doc.TransactionManager.StartTransaction()
                Try
                    Dim blk As BlockReference = TryCast(trans.GetObject(BlockID, OpenMode.ForRead), BlockReference)

                    If blk.ObjectId = BlockID Then
                        For Each attRefID As ObjectId In blk.AttributeCollection
                            Dim obj As DBObject = trans.GetObject(attRefID, OpenMode.ForRead)
                            Dim attRef As AttributeReference = TryCast(obj, AttributeReference)

                            For i = 0 To aTag.Count - 1

                                If attRef.Tag = aTag(i) Then
                                    TagResult.Add(attRef.TextString)
                                End If
                            Next
                        Next
                    End If

                    trans.Commit()

                Catch ex As Exception
                    ed.WriteMessage("Error encountered : " & ex.Message)
                    trans.Abort()
                End Try
            End Using
            Return TagResult
        End Function

        Public Shared Sub InsertBlockFromFile(FileName As String, BlockName As String)
            Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            Using acLckDocCur As DocumentLock = doc.LockDocument
                Using OpenDb As Database = New Database(False, True)
                    OpenDb.ReadDwgFile(FileName, FileShare.ReadWrite, True, "")
                    Dim ids As ObjectIdCollection = New ObjectIdCollection()
                    Using tr As Transaction = OpenDb.TransactionManager.StartTransaction()
                        Dim bt As BlockTable
                        bt = CType(tr.GetObject(OpenDb.BlockTableId, OpenMode.ForRead), BlockTable)
                        If bt.Has(BlockName) Then
                            ids.Add(bt(BlockName))
                        End If
                        tr.Commit()
                    End Using

                    If ids.Count > 0 Then
                        Dim destdb As Database = doc.Database
                        Dim iMap As IdMapping = New IdMapping()
                        destdb.WblockCloneObjects(ids, destdb.BlockTableId, iMap, DuplicateRecordCloning.Ignore, False)
                    End If
                End Using
            End Using

        End Sub

        Public Shared Function GetPolylineInBlock(BlockName As String, lname As String) As Polyline
            Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            Dim db As Autodesk.AutoCAD.DatabaseServices.Database
            Dim returnPolyline As Polyline = Nothing

            db = Application.DocumentManager.MdiActiveDocument.Database

            'InsertBlock = False

            Using acLckDocCur As DocumentLock = doc.LockDocument
                Try

                    Using trans As Transaction = db.TransactionManager.StartTransaction()
                        ' Open the Block table for read
                        Dim acBlkTbl As BlockTable
                        acBlkTbl = trans.GetObject(db.BlockTableId, OpenMode.ForRead)

                        Dim blkRecId As ObjectId = ObjectId.Null

                        If acBlkTbl.Has(BlockName) Then
                            blkRecId = acBlkTbl(BlockName)
                        Else
                            Return Nothing
                            'Exit Function
                        End If

                        ' Insert the block reference into the current space
                        If blkRecId <> ObjectId.Null Then
                            Dim objs As DBObjectCollection = New DBObjectCollection()
                            Using acBlkRef As New BlockReference(New Point3d(0, 0, 0), blkRecId)
                                acBlkRef.Explode(objs)
                                For Each obj As Object In objs
                                    Dim pl As Polyline = Nothing
                                    pl = TryCast(obj, Polyline)
                                    If pl IsNot Nothing AndAlso pl.Layer = lname Then
                                        returnPolyline = pl
                                        Exit For
                                    End If
                                Next
                                acBlkRef.Dispose()
                            End Using
                        End If

                        trans.Commit()

                    End Using
                Catch ex As Exception

                End Try
            End Using
            Return returnPolyline
        End Function

    End Class
End Namespace
