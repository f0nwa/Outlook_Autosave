Option Strict On

Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms
Imports OutlookAutosave.Core

Namespace UI

    ''' <summary>Выбор одной исходной папки Outlook для правила (перенос FolderSelectionForm.PickSingleFolderPath).</summary>
    Friend NotInheritable Class OutlookFolderPickerForm
        Inherits Form

        Private ReadOnly _lstFolders As New ListBox()

        Public Property SelectedFolderPath As String = String.Empty

        Public Sub New()
            Text = "Выбор папки Outlook"
            FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
            StartPosition = FormStartPosition.CenterParent
            MaximizeBox = False
            MinimizeBox = False
            ShowInTaskbar = False
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            Font = SystemFonts.MessageBoxFont
            ClientSize = New Size(780, 560)
            Icon = UiHelper.AppIcon()

            Dim title As New Label With {.Text = "Выберите исходную папку Outlook", .Location = New Point(16, 12), .Size = New Size(560, 22)}
            title.Font = New Font(Font, FontStyle.Bold)
            Controls.Add(title)
            Controls.Add(New Label With {.Text = "Выберите одну папку и нажмите ""Выбрать"" или дважды щелкните по ней.", .Location = New Point(16, 38), .Size = New Size(748, 20)})

            With _lstFolders
                .Location = New Point(16, 64)
                .Size = New Size(748, 430)
                .SelectionMode = SelectionMode.One
                .IntegralHeight = False
                .HorizontalScrollbar = True
            End With

            AddHandler _lstFolders.DoubleClick, AddressOf OnPick
            Controls.Add(_lstFolders)

            Dim btnPick As New Button With {.Text = "Выбрать", .Location = New Point(16, 510), .Size = New Size(120, 30)}
            AddHandler btnPick.Click, AddressOf OnPick
            Controls.Add(btnPick)

            Dim btnCancel As New Button With {.Text = "Отмена", .Location = New Point(644, 510), .Size = New Size(120, 30), .DialogResult = System.Windows.Forms.DialogResult.Cancel}
            Controls.Add(btnCancel)

            AcceptButton = btnPick
            CancelButton = btnCancel

            For Each folder In OutlookHost.ListInboxFolders()
                _lstFolders.Items.Add(folder)
            Next
        End Sub

        Private Sub OnPick(sender As Object, e As EventArgs)
            Dim folder = TryCast(_lstFolders.SelectedItem, OutlookFolderInfo)

            If folder Is Nothing Then
                UiHelper.Info(Me, "Выберите папку.")
                Return
            End If

            SelectedFolderPath = folder.FolderPath
            Me.DialogResult = System.Windows.Forms.DialogResult.OK
            Close()
        End Sub

    End Class

End Namespace
