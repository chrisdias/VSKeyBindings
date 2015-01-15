Imports System.Runtime.InteropServices
Imports EnvDTE
Imports EnvDTE80
Imports Microsoft.VisualStudio.Shell
Imports Microsoft.VisualStudio.Shell.Interop

<Guid("f7539fa7-dda1-4c70-a7dc-11f125a63092")>
Public Class VSKeyBindingsToolWindow
    Inherits ToolWindowPane

    Private WithEvents mc As MyControl
    Private WithEvents ViewModel As KeyBindingViewModel
    Private WithEvents Container As MyControl

    Friend Shadows ReadOnly Property Frame As IVsWindowFrame
        Get
            Return CType(MyBase.Frame, IVsWindowFrame)
        End Get
    End Property

    Public Sub New()
        MyBase.New(Nothing)

        Me.Caption = Resources.ToolWindowTitle ' Set the window title reading it from the resources.
        Me.BitmapResourceID = 301              ' The resource ID correspond to the one defined in the resx file
        Me.BitmapIndex = 2                     ' Index is the offset in the bitmap strip. Each image in the strip being 16x16

        Me.ViewModel = New KeyBindingViewModel
        Me.ViewModel.DTE = TryCast(VSKeyBindingsPackage.GetGlobalService(GetType(SDTE)), DTE2)

        Me.Container = New MyControl
        Me.Container.DataContext = Me.ViewModel
        Me.Content = Me.Container

    End Sub

    Public Sub InvokeSearchCommand(bForceLoadModel As Boolean)

        If bForceLoadModel = False Then
            If Me.ViewModel.IsDataLoaded = False Then
                Me.ViewModel.LoadModel()
            Else
                'get the current state and sort accordingly
                Me.ViewModel.LoadActiveKeyBindingScopes()
            End If
        Else
            Me.ViewModel.LoadModel()
        End If

    End Sub

    Private Sub Container_Hide() Handles Container.Hide

        If (Me.Frame IsNot Nothing) Then
            If IsDocked() = False Then
                Me.Frame.Hide()
            End If
        End If

    End Sub

    Public Function IsDocked() As Boolean
        Dim pos(1) As VSSETFRAMEPOS
        Dim x As Integer
        Dim y As Integer
        Dim w As Integer
        Dim h As Integer
        Dim unusedGuid As Guid

        pos(1) = New VSSETFRAMEPOS()
        Me.Frame.GetFramePos(pos, unusedGuid, x, y, w, h)
        Return (pos(0) And VSSETFRAMEPOS.SFP_fFloat) <> VSSETFRAMEPOS.SFP_fFloat

    End Function
End Class
