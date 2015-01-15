Imports System.Windows
Imports EnvDTE80
Imports Microsoft.VisualStudio.Shell.Interop
Imports System.Windows.Controls
Imports System.Globalization
Imports System.Threading.Tasks

Partial Public Class MyControl
    Inherits System.Windows.Controls.UserControl

    Private dte As DTE2

    'for type ahead search
    Private WithEvents myDispatcherTimer As System.Windows.Threading.DispatcherTimer
    Private tstart As Date

    Friend Event Hide()

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        dte = Utilities.GetDTE

        ' We use this timer to see when typing has "paused" so we can filter 
        myDispatcherTimer = New System.Windows.Threading.DispatcherTimer
        myDispatcherTimer.Interval = New TimeSpan(0, 0, 0, 0, 100)
        AddHandler myDispatcherTimer.Tick, AddressOf myDispatcherTimer_Tick

    End Sub

    Private Sub DataGrid1_PreviewKeyDown(sender As Object, e As System.Windows.Input.KeyEventArgs) Handles DataGrid1.PreviewKeyDown

        If e.Key = Input.Key.Enter And e.IsDown = True Then
            Call doCommand()
        ElseIf e.Key = Input.Key.Escape And e.IsDown = True Then
            RaiseEvent Hide()
        End If

    End Sub

    Private Sub DataGrid1_PreviewMouseDoubleClick(sender As Object, e As System.Windows.Input.MouseButtonEventArgs) Handles DataGrid1.PreviewMouseDoubleClick
        doCommand()
    End Sub

    Private Sub doCommand()
        Try

            Dim c As CommandTableEntry
            Dim wf As IVsWindowFrame

            c = TryCast(DataGrid1.SelectedItem, CommandTableEntry)
            If c Is Nothing Then
                Exit Sub
            End If

            wf = TryCast(Me.DataContext.CurrentDocWindowFrame, IVsWindowFrame)
            If wf IsNot Nothing Then
                wf.Show()
            End If

            dte.ExecuteCommand(c.CommandName)
            RaiseEvent Hide()

        Catch ex As Exception
            ' do nothing
        End Try

    End Sub

    Private Sub Grid_SizeChanged(sender As System.Object, e As System.Windows.SizeChangedEventArgs)
        DataGrid1.Height = e.NewSize.Height - txtSearch.ActualHeight
    End Sub

    Private Sub txtSearch_TextChanged(sender As Object, e As System.Windows.Controls.TextChangedEventArgs) Handles txtSearch.TextChanged
        tstart = Now
        myDispatcherTimer.IsEnabled = True
    End Sub

    Private Sub myDispatcherTimer_Tick(ByVal sender As Object, ByVal e As System.EventArgs)
        If Now.Ticks - tstart.Ticks > 2000000 Then
            myDispatcherTimer.IsEnabled = False
            Me.DataGrid1.ItemsSource = Me.DataContext.FilteredKeyBindings(Me.txtSearch.Text)
            Me.DataGrid1.SelectedIndex = 0
        End If
    End Sub

    Private Sub MyControl_PreviewKeyDown(sender As Object, e As System.Windows.Input.KeyEventArgs) Handles Me.PreviewKeyDown

        If e.KeyboardDevice.Modifiers = Input.ModifierKeys.Alt Then
            If e.SystemKey = Input.Key.K Then
                Debug.Print("alt+k pressed")
            End If
        End If

    End Sub


End Class

