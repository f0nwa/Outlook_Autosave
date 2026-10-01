Option Strict Off

Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.ComTypes
Imports System.Windows.Forms
Imports OutlookAutosave.Interop

Namespace Core

    ''' <summary>
    ''' Следит за исходными папками правил через событие Outlook Items.ItemAdd и сохраняет вложения
    ''' сразу, как только письмо оказалось в папке: пришло, докачалось с сервера или перенесено вручную.
    ''' Дата получения письма при этом не важна.
    ''' Outlook может не прислать событие при массовом добавлении (больше 16 писем за раз),
    ''' поэтому это дополнение к проверке по таймеру, а не замена.
    ''' </summary>
    Friend Module FolderWatcher

        ''' <summary>Пауза после последнего события перед обработкой: письма приходят пачками.</summary>
        Private Const DebounceMilliseconds As Integer = 3000
        ''' <summary>Сколько писем держать в очереди; остальное подхватит проверка по таймеру.</summary>
        Private Const MaxQueuedItems As Integer = 1000

        Friend NotInheritable Class AddedItem
            Public Property EntryId As String
            Public Property StoreId As String
            Public Property SourcePath As String
        End Class

        Private NotInheritable Class Watch
            Public Property SourcePath As String
            Public Property StoreId As String
            Public Property Items As Object
            Public Property ConnectionPoint As IConnectionPoint
            Public Property Cookie As Integer
            Public Property Sink As ItemsSink
        End Class

        Private ReadOnly _watches As New List(Of Watch)()
        Private ReadOnly _queue As New List(Of AddedItem)()
        Private _queueTimer As Timer

        Friend ReadOnly Property WatchedFolderCount As Integer
            Get
                Return _watches.Count
            End Get
        End Property

        ''' <summary>Подписывается на исходные папки активных правил (при включенном автосохранении).</summary>
        Friend Sub Refresh()
            StopAll()

            If Not AutoSaveEngine.IsAutoSaveEnabled() Then
                Return
            End If

            Dim sourcePaths = SortRuleStore.LoadActiveRules().
                Select(Function(r) TextUtil.NormalizeOutlookFolderPath(r.SourceFolder)).
                Where(Function(p) p.Length > 0).
                Distinct(StringComparer.CurrentCultureIgnoreCase).
                ToList()

            For Each sourcePath In sourcePaths
                Try
                    Dim folder = OutlookHost.FindMailFolderByPath(sourcePath)

                    If folder Is Nothing Then
                        Trace.WriteLine("OutlookAutosave: watcher skipped missing folder: " & sourcePath)
                        Continue For
                    End If

                    Dim storeId As String = String.Empty

                    Try
                        storeId = CStr(folder.StoreID)
                    Catch
                    End Try

                    Dim items As Object = folder.Items
                    Dim container = DirectCast(items, IConnectionPointContainer)
                    Dim iid = GetType(IItemsEvents).GUID
                    Dim cp As IConnectionPoint = Nothing
                    container.FindConnectionPoint(iid, cp)

                    Dim sink As New ItemsSink(sourcePath, storeId)
                    Dim cookie As Integer
                    cp.Advise(sink, cookie)

                    ' Ссылку на Items нужно держать, иначе Outlook перестанет присылать события.
                    _watches.Add(New Watch With {
                        .SourcePath = sourcePath, .StoreId = storeId, .Items = items,
                        .ConnectionPoint = cp, .Cookie = cookie, .Sink = sink})

                    Trace.WriteLine("OutlookAutosave: watching folder " & sourcePath)
                Catch ex As Exception
                    Trace.WriteLine("OutlookAutosave: watcher failed for " & sourcePath & ": " & ex.Message)
                End Try
            Next
        End Sub

        Friend Sub StopAll()
            For Each w In _watches
                Try
                    w.ConnectionPoint.Unadvise(w.Cookie)
                Catch
                End Try

                Dim items = w.Items
                OutlookHost.Release(items)
            Next

            _watches.Clear()
            _queue.Clear()

            If _queueTimer IsNot Nothing Then
                _queueTimer.Stop()
            End If
        End Sub

        Friend Sub Shutdown()
            StopAll()

            If _queueTimer IsNot Nothing Then
                RemoveHandler _queueTimer.Tick, AddressOf OnQueueTimerTick
                _queueTimer.Dispose()
                _queueTimer = Nothing
            End If
        End Sub

        Friend Sub OnItemAdded(item As Object, sourcePath As String, storeId As String)
            Try
                If Not AutoSaveEngine.IsAutoSaveEnabled() OrElse Not OutlookHost.IsMailItem(item) Then
                    Return
                End If

                If _queue.Count >= MaxQueuedItems Then
                    ' Массовая загрузка: остальное обработает проверка по таймеру.
                    Return
                End If

                _queue.Add(New AddedItem With {.EntryId = CStr(item.EntryID), .StoreId = storeId, .SourcePath = sourcePath})
                RestartQueueTimer()
            Catch ex As Exception
                Trace.WriteLine("OutlookAutosave: ItemAdd failed: " & ex.Message)
            End Try
        End Sub

        Private Sub RestartQueueTimer()
            If _queueTimer Is Nothing Then
                _queueTimer = New Timer With {.Interval = DebounceMilliseconds}
                AddHandler _queueTimer.Tick, AddressOf OnQueueTimerTick
            End If

            _queueTimer.Stop()
            _queueTimer.Start()
        End Sub

        Private Sub OnQueueTimerTick(sender As Object, e As EventArgs)
            _queueTimer.Stop()

            If _queue.Count = 0 Then
                Return
            End If

            If AutoSaveEngine.IsBusy Then
                ' Идет проверка по таймеру — попробуем чуть позже.
                _queueTimer.Start()
                Return
            End If

            Dim batch = _queue.ToList()
            _queue.Clear()

            Try
                AutoSaveEngine.ProcessAddedItems(batch)
            Catch ex As Exception
                Trace.WriteLine("OutlookAutosave: processing added items failed: " & ex.Message)
            End Try
        End Sub

    End Module

    ''' <summary>Получатель событий одной папки.</summary>
    <ComVisible(True), ClassInterface(ClassInterfaceType.None)>
    Public NotInheritable Class ItemsSink
        Implements IItemsEvents

        Private ReadOnly _sourcePath As String
        Private ReadOnly _storeId As String

        Public Sub New(sourcePath As String, storeId As String)
            _sourcePath = sourcePath
            _storeId = storeId
        End Sub

        Public Sub ItemAdd(Item As Object) Implements IItemsEvents.ItemAdd
            FolderWatcher.OnItemAdded(Item, _sourcePath, _storeId)
        End Sub

        Public Sub ItemChange(Item As Object) Implements IItemsEvents.ItemChange
        End Sub

        Public Sub ItemRemove() Implements IItemsEvents.ItemRemove
        End Sub
    End Class

End Namespace
