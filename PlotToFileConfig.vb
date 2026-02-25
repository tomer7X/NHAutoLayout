Imports System.Collections.Generic
Imports System.IO
Imports System.Text
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.PlottingServices
Imports Autodesk.AutoCAD.Publishing

Namespace MyNamespace
    Public MustInherit Class PlotToFileConfig
        Private dsdFile, dwgFile, outputDir, outputFile, plotType As String
        Private sheetNum As Integer
        Private layouts As IEnumerable(Of Layout)
        Private Const LOG As String = "publish.log"

        Public Sub New(ByVal outputDir As String, ByVal layouts As IEnumerable(Of Layout), ByVal plotType As String)
            Dim db As Database = HostApplicationServices.WorkingDatabase
            Me.dwgFile = db.Filename
            Me.outputDir = outputDir
            Me.dsdFile = Path.ChangeExtension(Me.dwgFile, "dsd")
            Me.layouts = layouts
            Me.plotType = plotType
            Dim ext As String = If(plotType = "0" OrElse plotType = "1", "dwf", "pdf")
            Me.outputFile = Path.Combine(Me.outputDir, Path.ChangeExtension(Path.GetFileName(Me.dwgFile), ext))
        End Sub

        Public Sub Publish()
            If TryCreateDSD() Then
                Dim bgp As Object = Application.GetSystemVariable("BACKGROUNDPLOT")
                Dim ctab As Object = Application.GetSystemVariable("CTAB")

                Try
                    Application.SetSystemVariable("BACKGROUNDPLOT", 0)
                    Dim publisher As Publisher = Application.Publisher
                    Dim plotDlg As PlotProgressDialog = New PlotProgressDialog(False, Me.sheetNum, True)
                    publisher.PublishDsd(Me.dsdFile, plotDlg)
                    plotDlg.Destroy()
                    File.Delete(Me.dsdFile)
                Catch exn As System.Exception
                    Dim ed As Editor = Application.DocumentManager.MdiActiveDocument.Editor
                    ed.WriteMessage(vbLf & "Error: {0}" & vbLf & "{1}", exn.Message, exn.StackTrace)
                    Throw
                Finally
                    Application.SetSystemVariable("BACKGROUNDPLOT", bgp)
                    Application.SetSystemVariable("CTAB", ctab)
                End Try
            End If
        End Sub

        Private Function TryCreateDSD() As Boolean
            Using dsd As DsdData = New DsdData()

                Using dsdEntries As DsdEntryCollection = CreateDsdEntryCollection(Me.layouts)
                    If dsdEntries Is Nothing OrElse dsdEntries.Count <= 0 Then Return False
                    If Not Directory.Exists(Me.outputDir) Then Directory.CreateDirectory(Me.outputDir)
                    Me.sheetNum = dsdEntries.Count
                    dsd.SetDsdEntryCollection(dsdEntries)
                    dsd.SetUnrecognizedData("PwdProtectPublishedDWF", "FALSE")
                    dsd.SetUnrecognizedData("PromptForPwd", "FALSE")
                    dsd.NoOfCopies = 1
                    dsd.DestinationName = Me.outputFile
                    dsd.IsHomogeneous = False
                    dsd.LogFilePath = Path.Combine(Me.outputDir, LOG)
                    PostProcessDSD(dsd)
                    Return True
                End Using
            End Using
        End Function

        Private Function CreateDsdEntryCollection(ByVal layouts As IEnumerable(Of Layout)) As DsdEntryCollection
            Dim entries As DsdEntryCollection = New DsdEntryCollection()

            For Each layout As Layout In layouts
                Dim dsdEntry As DsdEntry = New DsdEntry()
                dsdEntry.DwgName = Me.dwgFile
                dsdEntry.Layout = layout.LayoutName
                dsdEntry.Title = Path.GetFileNameWithoutExtension(Me.dwgFile) & "-" & layout.LayoutName
                dsdEntry.Nps = layout.TabOrder.ToString()
                entries.Add(dsdEntry)
            Next

            Return entries
        End Function

        Private Sub PostProcessDSD(ByVal dsd As DsdData)
            Dim str, newStr As String
            Dim tmpFile As String = Path.Combine(Me.outputDir, "temp.dsd")
            dsd.WriteDsd(tmpFile)

            Using reader As StreamReader = New StreamReader(tmpFile, Encoding.[Default])

                Using writer As StreamWriter = New StreamWriter(Me.dsdFile, False, Encoding.[Default])

                    While Not reader.EndOfStream
                        str = reader.ReadLine()

                        If str.Contains("Has3DDWF") Then
                            newStr = "Has3DDWF=0"
                        ElseIf str.Contains("OriginalSheetPath") Then
                            newStr = "OriginalSheetPath=" & Me.dwgFile
                        ElseIf str.Contains("Type") Then
                            newStr = "Type=" & Me.plotType
                        ElseIf str.Contains("OUT") Then
                            newStr = "OUT=" & Me.outputDir
                        ElseIf str.Contains("IncludeLayer") Then
                            newStr = "IncludeLayer=TRUE"
                        ElseIf str.Contains("PromptForDwfName") Then
                            newStr = "PromptForDwfName=FALSE"
                        ElseIf str.Contains("LogFilePath") Then
                            newStr = "LogFilePath=" & Path.Combine(Me.outputDir, LOG)
                        Else
                            newStr = str
                        End If

                        writer.WriteLine(newStr)
                    End While
                End Using
            End Using

            File.Delete(tmpFile)
        End Sub
    End Class

    Public Class SingleSheetDwf
        Inherits PlotToFileConfig

        Public Sub New(ByVal outputDir As String, ByVal layouts As IEnumerable(Of Layout))
            MyBase.New(outputDir, layouts, "0")
        End Sub
    End Class

    Public Class MultiSheetsDwf
        Inherits PlotToFileConfig

        Public Sub New(ByVal outputDir As String, ByVal layouts As IEnumerable(Of Layout))
            MyBase.New(outputDir, layouts, "1")
        End Sub
    End Class

    Public Class SingleSheetPdf
        Inherits PlotToFileConfig

        Public Sub New(ByVal outputDir As String, ByVal layouts As IEnumerable(Of Layout))
            MyBase.New(outputDir, layouts, "5")
        End Sub
    End Class

    Public Class MultiSheetsPdf
        Inherits PlotToFileConfig

        Public Sub New(ByVal outputDir As String, ByVal layouts As IEnumerable(Of Layout))
            MyBase.New(outputDir, layouts, "6")
        End Sub
    End Class
End Namespace
