Option Strict Off

Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms
Imports System.Runtime.InteropServices

Namespace Core

    ''' <summary>
    ''' Доступ к объектной модели Outlook через позднее связывание (как в VBA).
    ''' Так надстройка не зависит от конкретной версии Office PIA.
    ''' Весь код здесь должен вызываться только из главного потока Outlook.
    ''' </summary>
    Friend Module OutlookHost

        Friend Const OlMail As Integer = 43
        Friend Const OlMailItem As Integer = 0
        Friend Const OlFolderInbox As Integer = 6

        Private _application As Object

        Friend Property OutlookApp As Object
            Get
                Return _application
            End Get
            Set(value As Object)
                _application = value
            End Set
        End Property

        Friend Sub Release(ByRef comObject As Object)
            If comObject IsNot Nothing AndAlso Marshal.IsComObject(comObject) Then
                Try
                    Marshal.ReleaseComObject(comObject)
                Catch
                End Try
            End If

            comObject = Nothing
        End Sub

        Friend Function GetNamespace() As Object
            Return _application.GetNamespace("MAPI")
        End Function

        Friend Function GetActiveExplorer() As Object
            Try
                Return _application.ActiveExplorer()
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>Текущая открытая почтовая папка в активном окне Outlook.</summary>
        Friend Function GetCurrentFolder() As Object
            Dim explorer = GetActiveExplorer()

            If explorer Is Nothing Then
                Return Nothing
            End If

            Try
                Return explorer.CurrentFolder
            Catch
                Return Nothing
            End Try
        End Function

        Friend Function GetFolderPath(folder As Object) As String
            Try
                Return CStr(folder.FolderPath)
            Catch
                Return String.Empty
            End Try
        End Function

        Friend Function IsMailFolder(folder As Object) As Boolean
            Try
                Return CInt(folder.DefaultItemType) = OlMailItem
            Catch
                Return False
            End Try
        End Function

        Friend Function IsMailItem(item As Object) As Boolean
            Try
                Return CInt(item.Class) = OlMail
            Catch
                Return False
            End Try
        End Function

        ''' <summary>HWND активного окна Outlook для модальных диалогов.</summary>
        Friend Function GetOwnerWindow() As IWin32Window
            Try
                Dim activeWindow As Object = _application.ActiveWindow()
                Dim oleWindow = TryCast(activeWindow, Interop.IOleWindow)

                If oleWindow IsNot Nothing Then
                    Dim hwnd As IntPtr
                    oleWindow.GetWindow(hwnd)

                    If hwnd <> IntPtr.Zero Then
                        Return New WindowHandle(hwnd)
                    End If
                End If
            Catch ex As Exception
                Trace.WriteLine("OutlookAutosave: owner window not available: " & ex.Message)
            End Try

            Return Nothing
        End Function

        ''' <summary>Список почтовых папок для выбора исходной папки правила.</summary>
        Friend Function ListInboxFolders() As List(Of OutlookFolderInfo)
            Dim result As New List(Of OutlookFolderInfo)()
            Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            Dim ns = GetNamespace()
            Dim stores = ns.Stores

            For i As Integer = 1 To CInt(stores.Count)
                Dim store As Object = Nothing
                Dim inbox As Object = Nothing

                Try
                    store = stores.Item(i)
                    inbox = store.GetDefaultFolder(OlFolderInbox)
                Catch
                    inbox = Nothing
                End Try

                If inbox IsNot Nothing Then
                    Dim storeName As String = String.Empty

                    Try
                        storeName = CStr(store.DisplayName)
                    Catch
                    End Try

                    AddFolderRecursive(inbox, 0, storeName, result, seen)
                End If
            Next

            Return result
        End Function

        Private Sub AddFolderRecursive(folder As Object, level As Integer, storeName As String, result As List(Of OutlookFolderInfo), seen As HashSet(Of String))
            If Not ShouldShowMailFolder(folder) Then
                Return
            End If

            Dim folderPath = GetFolderPath(folder)

            If folderPath.Length > 0 AndAlso seen.Add(folderPath) Then
                Dim displayName = New String(" "c, level * 2) & CStr(folder.Name)

                If level = 0 AndAlso storeName.Length > 0 Then
                    displayName = storeName & " \ " & displayName
                End If

                result.Add(New OutlookFolderInfo(displayName, folderPath))
            End If

            Dim children As Object = Nothing

            Try
                children = folder.Folders
            Catch
                Return
            End Try

            For i As Integer = 1 To CInt(children.Count)
                Dim child As Object = Nothing

                Try
                    child = children.Item(i)
                Catch
                    Continue For
                End Try

                AddFolderRecursive(child, level + 1, storeName, result, seen)
            Next
        End Sub

        Friend Function ShouldShowMailFolder(folder As Object) As Boolean
            Try
                If Not IsMailFolder(folder) Then
                    Return False
                End If

                Dim folderName = CStr(folder.Name).ToLowerInvariant()

                If folderName.StartsWith("{", StringComparison.Ordinal) AndAlso folderName.EndsWith("}", StringComparison.Ordinal) Then
                    Return False
                End If

                Select Case folderName
                    Case "recipient cache", "peoplecentricconversation buddies", "organizational contacts", "gal contacts",
                         "contacts", "contactos", "calendar", "notes", "tasks", "journal",
                         "outbound", "inbound", "feeds"
                        Return False
                    Case Else
                        Return True
                End Select
            Catch
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Ищет почтовую папку по пути Outlook: сначала точное совпадение во всех ящиках, затем по пути
        ''' без имени ящика (если ящик переименован). Поиск без имени ящика срабатывает, только если такая
        ''' папка единственная: иначе можно взять одноименную папку другого ящика.
        ''' </summary>
        Friend Function FindMailFolderByPath(folderPath As String) As Object
            Dim normalized = TextUtil.NormalizeOutlookFolderPath(folderPath)

            If normalized.Length = 0 Then
                Return Nothing
            End If

            ' Быстрый путь: текущая открытая папка, только при точном совпадении пути.
            Dim current = GetCurrentFolder()

            If current IsNot Nothing AndAlso
               String.Equals(TextUtil.NormalizeOutlookFolderPath(GetFolderPath(current)), normalized, StringComparison.CurrentCultureIgnoreCase) Then
                Return current
            End If

            Dim roots = GetStoreRootFolders()

            For Each root In roots
                Dim found As New List(Of Object)()
                FindUnder(root, normalized, False, found)

                If found.Count > 0 Then
                    Return found(0)
                End If
            Next

            Dim suffix = TextUtil.GetOutlookFolderPathSuffix(normalized)

            If suffix.Length > 0 Then
                Dim matches As New List(Of Object)()

                For Each root In roots
                    FindUnder(root, suffix, True, matches)

                    If matches.Count > 1 Then
                        Exit For
                    End If
                Next

                If matches.Count = 1 Then
                    Trace.WriteLine("OutlookAutosave: source folder found without mailbox name: " & normalized & " -> " & GetFolderPath(matches(0)))
                    Return matches(0)
                End If

                If matches.Count > 1 Then
                    Trace.WriteLine("OutlookAutosave: source folder is ambiguous without mailbox name, skipped: " & normalized)
                    Return Nothing
                End If
            End If

            Trace.WriteLine("OutlookAutosave: source folder not found: " & normalized)
            Return Nothing
        End Function

        Private Function GetStoreRootFolders() As List(Of Object)
            Dim result As New List(Of Object)()
            Dim stores = GetNamespace().Stores

            For i As Integer = 1 To CInt(stores.Count)
                Try
                    Dim root As Object = stores.Item(i).GetRootFolder()

                    If root IsNot Nothing Then
                        result.Add(root)
                    End If
                Catch ex As Exception
                    Trace.WriteLine("OutlookAutosave: store skipped: " & ex.Message)
                End Try
            Next

            Return result
        End Function

        ''' <summary>
        ''' Добавляет в found почтовые папки, совпадающие с target по полному пути или (bySuffix) по пути без имени ящика.
        ''' Точный поиск останавливается на первой найденной папке, поиск без имени ящика — на второй
        ''' (этого достаточно, чтобы понять, что совпадение неоднозначно).
        ''' </summary>
        Private Sub FindUnder(folder As Object, target As String, bySuffix As Boolean, found As List(Of Object))
            Try
                If IsMailFolder(folder) Then
                    Dim path = TextUtil.NormalizeOutlookFolderPath(GetFolderPath(folder))
                    Dim compared = If(bySuffix, TextUtil.GetOutlookFolderPathSuffix(path), path)

                    If String.Equals(compared, target, StringComparison.CurrentCultureIgnoreCase) Then
                        found.Add(folder)
                    End If
                End If

                Dim children = folder.Folders

                For i As Integer = 1 To CInt(children.Count)
                    If IsSearchDone(found, bySuffix) Then
                        Return
                    End If

                    Dim child As Object = Nothing

                    Try
                        child = children.Item(i)
                    Catch ex As Exception
                        Trace.WriteLine("OutlookAutosave: folder skipped: " & ex.Message)
                        Continue For
                    End Try

                    FindUnder(child, target, bySuffix, found)
                Next
            Catch ex As Exception
                Trace.WriteLine("OutlookAutosave: folder search error: " & ex.Message)
            End Try
        End Sub

        Private Function IsSearchDone(found As List(Of Object), bySuffix As Boolean) As Boolean
            Return found.Count > If(bySuffix, 1, 0)
        End Function

    End Module

    Friend NotInheritable Class OutlookFolderInfo
        Public Sub New(displayName As String, folderPath As String)
            Me.DisplayName = displayName
            Me.FolderPath = folderPath
        End Sub

        Public ReadOnly Property DisplayName As String
        Public ReadOnly Property FolderPath As String

        Public Overrides Function ToString() As String
            Return DisplayName
        End Function
    End Class

    Friend NotInheritable Class WindowHandle
        Implements IWin32Window

        Private ReadOnly _handle As IntPtr

        Public Sub New(handle As IntPtr)
            _handle = handle
        End Sub

        Public ReadOnly Property Handle As IntPtr Implements IWin32Window.Handle
            Get
                Return _handle
            End Get
        End Property
    End Class

End Namespace
