Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.ViewModel.PointCloudManager

Public Class NetDimensions

    ''' <summary>
    ''' Returns the Measurement values of all Dimension entities on the given layer
    ''' whose definition points are fully contained within the given boundary polygon.
    ''' </summary>
    ''' <param name="boundary">Closed polygon as Point3dCollection (WCS, Z ignored)</param>
    ''' <param name="layerName">Layer to filter by</param>
    ''' <param name="tr">Active transaction</param>
    ''' <returns>List of measurement values</returns>
    Public Shared Function GetSumOfDimensionValuesInBoundaryByLayer(boundary As Point3dCollection, layerName As String, tr As Transaction) As Integer
        Dim values As List(Of Integer) = GetDimensionValuesInBoundary(boundary, layerName, tr)
        Return values.Sum()
    End Function

    Public Shared Function GetDimensionValuesInBoundary(
            boundary As Point3dCollection,
            layerName As String,
            tr As Transaction) As List(Of Integer)

        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim results As New List(Of Integer)

        ' Build a 2D polygon from the boundary for point-in-polygon tests
        Dim poly2D As New List(Of Point2d)
        For Each pt As Point3d In boundary
            poly2D.Add(New Point2d(pt.X, pt.Y))
        Next

        ' Open model space and iterate all entities
        Dim modelSpace As BlockTableRecord =
            DirectCast(tr.GetObject(db.CurrentSpaceId, OpenMode.ForRead), BlockTableRecord)

        For Each objId As ObjectId In modelSpace

            Dim ent As Entity = TryCast(tr.GetObject(objId, OpenMode.ForRead), Entity)
            If ent Is Nothing Then Continue For

            ' Layer filter
            If String.Compare(ent.Layer, layerName, StringComparison.OrdinalIgnoreCase) <> 0 Then
                Continue For
            End If

            ' Dimension filter
            Dim dimension As Dimension = TryCast(ent, Dimension)
            If dimension Is Nothing Then Continue For

            ' Get all definition points of this dimension
            Dim defPoints As List(Of Point2d) = GetDimensionDefPoints(dimension)
            If defPoints Is Nothing OrElse defPoints.Count = 0 Then Continue For

            ' Check that EVERY definition point is inside the boundary
            Dim fullyContained As Boolean = True
            For Each pt As Point2d In defPoints
                If Not IsPointInPolygon(pt, poly2D) Then
                    fullyContained = False
                    Exit For
                End If
            Next

            If fullyContained Then
                results.Add(Math.Round(dimension.Measurement))
            End If

        Next

        Return results

    End Function

    ''' <summary>
    ''' Extracts all definition points from any Dimension subtype.
    ''' These are the geometry-defining points (not the text position).
    ''' </summary>
    Public Shared Function GetDimensionDefPoints(dimension As Dimension) As List(Of Point2d)

        Dim pts As New List(Of Point2d)

        Select Case True

            Case TypeOf dimension Is RotatedDimension
                Dim d As RotatedDimension = DirectCast(dimension, RotatedDimension)
                pts.Add(To2D(d.XLine1Point))
                pts.Add(To2D(d.XLine2Point))
                pts.Add(To2D(d.DimLinePoint))

            Case TypeOf dimension Is AlignedDimension
                Dim d As AlignedDimension = DirectCast(dimension, AlignedDimension)
                pts.Add(To2D(d.XLine1Point))
                pts.Add(To2D(d.XLine2Point))
                pts.Add(To2D(d.DimLinePoint))

            Case TypeOf dimension Is RadialDimension
                Dim d As RadialDimension = DirectCast(dimension, RadialDimension)
                pts.Add(To2D(d.Center))
                pts.Add(To2D(d.ChordPoint))

            Case TypeOf dimension Is RadialDimensionLarge
                Dim d As RadialDimensionLarge = DirectCast(dimension, RadialDimensionLarge)
                pts.Add(To2D(d.Center))
                pts.Add(To2D(d.ChordPoint))

            Case TypeOf dimension Is DiametricDimension
                Dim d As DiametricDimension = DirectCast(dimension, DiametricDimension)
                pts.Add(To2D(d.ChordPoint))
                pts.Add(To2D(d.FarChordPoint))

            Case TypeOf dimension Is ArcDimension
                Dim d As ArcDimension = DirectCast(dimension, ArcDimension)
                pts.Add(To2D(d.CenterPoint))
                pts.Add(To2D(d.XLine1Point))
                pts.Add(To2D(d.XLine2Point))
                pts.Add(To2D(d.ArcPoint))

            Case TypeOf dimension Is OrdinateDimension
                Dim d As OrdinateDimension = DirectCast(dimension, OrdinateDimension)
                pts.Add(To2D(d.Origin))
                pts.Add(To2D(d.DefiningPoint))
                pts.Add(To2D(d.LeaderEndPoint))

            Case Else
                ' Unknown dimension type — fall back to GeometricExtents corners
                Try
                    Dim ext As Extents3d = dimension.GeometricExtents
                    pts.Add(To2D(ext.MinPoint))
                    pts.Add(To2D(ext.MaxPoint))
                    pts.Add(New Point2d(ext.MinPoint.X, ext.MaxPoint.Y))
                    pts.Add(New Point2d(ext.MaxPoint.X, ext.MinPoint.Y))
                Catch
                    ' No extents available
                End Try

        End Select

        Return pts

    End Function

    ''' <summary>
    ''' Ray-casting point-in-polygon test (2D, ignores Z).
    ''' Works for convex and concave polygons.
    ''' </summary>
    Private Shared Function IsPointInPolygon(pt As Point2d, polygon As List(Of Point2d)) As Boolean

        Dim n As Integer = polygon.Count
        Dim inside As Boolean = False
        Dim j As Integer = n - 1

        For i As Integer = 0 To n - 1
            Dim xi As Double = polygon(i).X
            Dim yi As Double = polygon(i).Y
            Dim xj As Double = polygon(j).X
            Dim yj As Double = polygon(j).Y

            Dim intersects As Boolean =
                ((yi > pt.Y) <> (yj > pt.Y)) AndAlso
                (pt.X < (xj - xi) * (pt.Y - yi) / (yj - yi) + xi)

            If intersects Then inside = Not inside

            j = i
        Next

        Return inside

    End Function

    ''' <summary>Converts Point3d to Point2d by dropping Z.</summary>
    Private Shared Function To2D(pt As Point3d) As Point2d
        Return New Point2d(pt.X, pt.Y)
    End Function

End Class
