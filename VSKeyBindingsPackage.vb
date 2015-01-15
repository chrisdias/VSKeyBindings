Imports System.ComponentModel.Design
Imports System.Globalization
Imports System.Runtime.InteropServices
Imports Microsoft.VisualStudio
Imports Microsoft.VisualStudio.Shell
Imports Microsoft.VisualStudio.Shell.Interop

<PackageRegistration(UseManagedResourcesOnly:=True)>
<InstalledProductRegistration("#110", "#112", "1.0", IconResourceID:=400)>
<ProvideMenuResource("Menus.ctmenu", 1)>
<ProvideToolWindow(GetType(VSKeyBindingsToolWindow), Transient:=True)>
<ProvideAutoLoad(VSConstants.UICONTEXT.NoSolution_string)>
<ProvideAutoLoad(VSConstants.UICONTEXT.SolutionExists_string)>
<Guid(Globals.PackageGuidString)>
Public NotInheritable Class VSKeyBindingsPackage
    Inherits Package

    Public Shared Property Instance() As VSKeyBindingsPackage

    Private shellPropertyCookie As UInteger = 0
    Private idleCookie As UInteger = 0

    ' Default constructor of the package. Inside this method you can place any initialization code that does not require 
    ' any Visual Studio service because at this point the package object is created but not sited yet inside Visual Studio environment. 
    ' The place to do all the other initialization is the Initialize method.
    Public Sub New()
        'Trace.WriteLine(String.Format(CultureInfo.CurrentCulture, "Entering constructor for: {0}", Me.GetType().Name))
    End Sub

    ' Initialization of the package; this method is called right after the package is sited, so this is the place
    ' where you can put all the initilaization code that rely on services provided by VisualStudio.
    Protected Overrides Sub Initialize()
        Dim mcs As OleMenuCommandService

        VSKeyBindingsPackage.Instance = Me

        MyBase.Initialize()

        ' Add our command handlers for menu (commands must exist in the .vsct file)
        mcs = TryCast(GetService(GetType(IMenuCommandService)), OleMenuCommandService)
        If mcs IsNot Nothing Then
            mcs.AddCommand(New MenuCommand(New EventHandler(AddressOf OnKeyBindingsInvoked), Globals.KeyBindingsCommand))
            mcs.AddCommand(New MenuCommand(New EventHandler(AddressOf OnKeyBindingsBackInvoked), Globals.KeyBindingsBackCommand))
        End If

    End Sub

    Private Sub OnKeyBindingsInvoked(ByVal sender As Object, ByVal e As EventArgs)
        'creates tool window if it doesnt exist
        Dim window = TryCast(Me.FindToolWindow(GetType(VSKeyBindingsToolWindow), 0, True), VSKeyBindingsToolWindow)
        If window IsNot Nothing Then
            If window.Frame.IsVisible = 0 Then
                window.Frame.Hide()
            Else
                window.InvokeSearchCommand(False)
                window.Frame.Show()
            End If
        End If

    End Sub

    Private Sub OnKeyBindingsBackInvoked(ByVal sender As Object, ByVal e As EventArgs)
        'creates tool window if it doesnt exist
        Dim window = TryCast(Me.FindToolWindow(GetType(VSKeyBindingsToolWindow), 0, True), VSKeyBindingsToolWindow)
        window.InvokeSearchCommand(False)
        window.Frame.Show()
    End Sub

End Class
